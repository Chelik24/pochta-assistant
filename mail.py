"""
Локальный архив почты: докачка писем в базу и быстрые запросы к ней.

Зачем: раньше на каждый вопрос про почту скрипт заново выкачивал письма с
сервера и складывал в файл на сотню килобайт — это дорого по времени и по
токенам агента. Старые письма не меняются, поэтому их достаточно скачать
один раз в `Почта/архив.db`, а дальше отвечать из базы, не трогая сеть.

Только чтение, как и раньше: ящик открывается в режиме readonly, письма
забираются через BODY.PEEK и не помечаются прочитанными, ничего не удаляется,
не перемещается и не отправляется.

Запуск:
    python mail.py скачать              # докачать новое (первый раз — за месяц)
    python mail.py скачать --дней 90    # захватить период поглубже
    python mail.py сводка --дней 3      # одна строка на письмо, для разбора
    python mail.py поиск тендер         # поиск по всему архиву
    python mail.py письмо 142           # полный текст нужного письма
    python mail.py статистика           # что уже лежит в архиве

Логины и пароли берутся из .env (см. .env.example), сервер определяется по
адресу. Используется только стандартная библиотека Python 3.8+.
"""

import argparse
import email
import imaplib
import json
import re
import sqlite3
import sys
from datetime import datetime, timedelta
from email.policy import default
from email.utils import parseaddr, parsedate_to_datetime
from pathlib import Path

# Оттуда же берутся разбор письма, чистка имён и определение сервера —
# чтобы гарантия «только чтение» и обработка кодировок жили в одном месте.
# Импорт заодно переключает вывод на UTF-8.
import fetch_mail
from fetch_mail import (
    clean_header,
    get_text,
    imap_date,
    is_junk,
    load_accounts,
    load_env,
    load_ignore_rules,
    parse_date,
    safe_filename,
    save_attachments,
)

БАЗА = "Почта/архив.db"

# Сколько писем тянем одной командой FETCH. Больше — меньше обращений к
# серверу, но письмо с вложением на 20 МБ целиком лежит в памяти.
ПАРТИЯ = 25

# Длина текста письма в базе. Письма длиннее почти всегда рассылки с версткой.
МАКС_ТЕКСТ = 20000

# Сколько писем показываем без явной просьбы
ПРЕДЕЛ_СВОДКИ = 200
ПРЕДЕЛ_ПОИСКА = 30

# Папка отправленных называется у всех по-разному, а на mail.ru и Яндексе ещё
# и по-русски в кодировке IMAP. Сначала спрашиваем сервер (RFC 6154), и только
# если он молчит — сверяемся со списком известных имён.
ИМЕНА_ОТПРАВЛЕННЫХ = {
    "sent", "sent items", "sent mail", "sent messages",
    "inbox.sent", "[gmail]/sent mail",
}

СТРОКА_LIST = re.compile(rb'^\((?P<flags>[^)]*)\)\s+(?P<delim>"[^"]*"|NIL)\s+(?P<name>.*)$')

СХЕМА = """
CREATE TABLE IF NOT EXISTS mailboxes (
    id          INTEGER PRIMARY KEY,
    account     TEXT NOT NULL,
    folder      TEXT NOT NULL,
    label       TEXT NOT NULL,
    direction   TEXT NOT NULL,
    uidvalidity INTEGER,
    last_uid    INTEGER NOT NULL DEFAULT 0,
    covered_since TEXT,
    synced_at   TEXT,
    UNIQUE (account, folder)
);

CREATE TABLE IF NOT EXISTS messages (
    id          INTEGER PRIMARY KEY,
    mailbox_id  INTEGER NOT NULL REFERENCES mailboxes(id),
    uid         INTEGER NOT NULL,
    message_id  TEXT,
    ts          INTEGER NOT NULL,
    date_text   TEXT,
    sender      TEXT,
    sender_addr TEXT,
    recipient   TEXT,
    subject     TEXT,
    body        TEXT,
    attachments TEXT,
    unread      INTEGER NOT NULL DEFAULT 0,
    saved_at    TEXT,
    UNIQUE (mailbox_id, uid)
);

CREATE INDEX IF NOT EXISTS messages_ts ON messages(ts);
CREATE INDEX IF NOT EXISTS messages_sender ON messages(sender_addr);

CREATE VIRTUAL TABLE IF NOT EXISTS search USING fts5(
    subject, sender, body,
    content='messages', content_rowid='id', tokenize='{токенизатор}'
);

CREATE TRIGGER IF NOT EXISTS messages_ai AFTER INSERT ON messages BEGIN
    INSERT INTO search(rowid, subject, sender, body)
    VALUES (new.id, new.subject, new.sender, new.body);
END;

CREATE TRIGGER IF NOT EXISTS messages_ad AFTER DELETE ON messages BEGIN
    INSERT INTO search(search, rowid, subject, sender, body)
    VALUES ('delete', old.id, old.subject, old.sender, old.body);
END;

CREATE TRIGGER IF NOT EXISTS messages_au AFTER UPDATE ON messages BEGIN
    INSERT INTO search(search, rowid, subject, sender, body)
    VALUES ('delete', old.id, old.subject, old.sender, old.body);
    INSERT INTO search(rowid, subject, sender, body)
    VALUES (new.id, new.subject, new.sender, new.body);
END;
"""


# --- база -----------------------------------------------------------------


def открыть_базу(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(str(path))
    db.row_factory = sqlite3.Row
    # remove_diacritics 2 приравнивает «ё» к «е»: иначе «Королёв» не найдётся
    # по запросу «королев». Старые сборки SQLite такого варианта не знают.
    for токенизатор in ("unicode61 remove_diacritics 2", "unicode61"):
        try:
            db.executescript(СХЕМА.format(токенизатор=токенизатор))
            доработать_схему(db)
            return db
        except sqlite3.OperationalError as e:
            последняя = e
    sys.exit(
        "SQLite на этом компьютере собран без полнотекстового поиска (FTS5). "
        f"Нужен Python с python.org. Ошибка: {последняя}"
    )


def доработать_схему(db) -> None:
    """
    Дотягивает базу, собранную прежней версией, до текущей схемы.

    Архив живёт на машине клиента и пересобирать его при каждом обновлении
    нельзя: это часы скачивания заново.
    """
    столбцы = {строка[1] for строка in db.execute("PRAGMA table_info(mailboxes)")}
    if "covered_since" not in столбцы:
        db.execute("ALTER TABLE mailboxes ADD COLUMN covered_since TEXT")
        # За какой период собрана старая база, мы не записывали. Ближайшая
        # честная оценка — дата самого старого письма в папке.
        db.execute(
            """UPDATE mailboxes SET covered_since =
                 (SELECT date(MIN(m.ts), 'unixepoch') FROM messages m
                  WHERE m.mailbox_id = mailboxes.id)
               WHERE covered_since IS NULL"""
        )
        db.commit()


def найти_ящик(db, account: str, folder: str, label: str, direction: str) -> sqlite3.Row:
    db.execute(
        "INSERT OR IGNORE INTO mailboxes (account, folder, label, direction) VALUES (?, ?, ?, ?)",
        (account, folder, label, direction),
    )
    return db.execute(
        "SELECT * FROM mailboxes WHERE account = ? AND folder = ?", (account, folder)
    ).fetchone()


def собрано_с(ящик) -> datetime:
    """С какой даты архив по этой папке уже собран. Нет отметки — значит ни с какой."""
    значение = ящик["covered_since"] if "covered_since" in ящик.keys() else None
    if not значение:
        return None
    try:
        return datetime.strptime(значение, "%Y-%m-%d")
    except ValueError:
        return None


def отметить_ящик(db, mailbox_id: int, последний_uid: int, покрытие: datetime) -> None:
    db.execute(
        "UPDATE mailboxes SET last_uid = ?, covered_since = ?, synced_at = ? WHERE id = ?",
        (последний_uid, покрытие.strftime("%Y-%m-%d"),
         datetime.now().isoformat(timespec="seconds"), mailbox_id),
    )


def сбросить_ящик(db, mailbox_id: int) -> None:
    """UIDVALIDITY сменился — старые номера писем больше ничего не значат."""
    db.execute("DELETE FROM messages WHERE mailbox_id = ?", (mailbox_id,))
    db.execute("UPDATE mailboxes SET last_uid = 0 WHERE id = ?", (mailbox_id,))


def время_письма(msg, запасное: datetime) -> tuple:
    """(unix-время для сортировки, дата как в письме). Кривая дата — не повод терять письмо."""
    сырое = clean_header(msg.get("date"))
    try:
        дата = parsedate_to_datetime(msg.get("date"))
    except (TypeError, ValueError):
        дата = None
    if дата is None:
        return int(запасное.timestamp()), сырое
    return int(дата.timestamp()), сырое or дата.strftime("%d.%m.%Y %H:%M")


def сохранить_письмо(db, ящик, uid: int, msg, флаги: str, папка_вложений: Path,
                     база_проекта: Path, сохранять_вложения: bool,
                     перезаписать: bool = False) -> str:
    """Кладёт письмо в базу. Говорит, чем это было: «новых» или «обновлено»."""
    было = db.execute(
        "SELECT id, attachments FROM messages WHERE mailbox_id = ? AND uid = ?",
        (ящик["id"], uid),
    ).fetchone()
    if было and not перезаписать:
        return "пропущено"

    сейчас = datetime.now()
    ts, дата_текст = время_письма(msg, сейчас)
    текст = get_text(msg)
    if len(текст) > МАКС_ТЕКСТ:
        текст = текст[:МАКС_ТЕКСТ] + "\n…[текст обрезан]"

    if было:
        # Файлы с диска не перекачиваем: они уже там, а повторное сохранение
        # только наплодило бы копии вида «Договор_1.docx».
        вложения_json = было["attachments"]
    else:
        вложения = []
        if сохранять_вложения:
            for имя, путь in save_attachments(msg, папка_вложений):
                try:
                    показать = путь.relative_to(база_проекта).as_posix()
                except ValueError:  # папка вложений вне проекта
                    показать = путь.as_posix()
                вложения.append({"имя": clean_header(имя), "путь": показать})
        вложения_json = json.dumps(вложения, ensure_ascii=False) if вложения else None

    поля = {
        "message_id": clean_header(msg.get("message-id")),
        "ts": ts,
        "date_text": дата_текст,
        "sender": clean_header(msg.get("from")),
        "sender_addr": parseaddr(clean_header(msg.get("from")))[1].lower(),
        "recipient": clean_header(msg.get("to")),
        "subject": clean_header(msg.get("subject")) or "(без темы)",
        "body": текст,
        "attachments": вложения_json,
        "unread": 0 if "\\Seen" in флаги else 1,
        "saved_at": сейчас.isoformat(timespec="seconds"),
    }

    if было:
        # Именно UPDATE, а не «удалить и вставить»: номер письма остаётся
        # прежним. Агент показывает эти номера Денису, и после перечитывания
        # архива они не должны разъехаться.
        db.execute(
            "UPDATE messages SET " + ", ".join(f"{имя} = ?" for имя in поля) + " WHERE id = ?",
            [*поля.values(), было["id"]],
        )
        return "обновлено"

    столбцы = ["mailbox_id", "uid", *поля]
    db.execute(
        f"INSERT INTO messages ({', '.join(столбцы)}) "
        f"VALUES ({', '.join('?' * len(столбцы))})",
        [ящик["id"], uid, *поля.values()],
    )
    return "новых"


# --- разбор ответов IMAP --------------------------------------------------


def разобрать_папки(данные) -> list:
    """Ответ LIST → [(имя, ярлык, направление)] для Входящих и Отправленных."""
    папки = [("INBOX", "Входящие", "входящее")]
    отправленные = None
    запасной = None

    for строка in данные or []:
        if isinstance(строка, tuple):  # имя ушло отдельным блоком — редкость
            строка = строка[0]
        if not isinstance(строка, bytes):
            continue
        m = СТРОКА_LIST.match(строка.strip())
        if not m:
            continue
        флаги = m.group("flags").decode("ascii", "replace")
        имя = m.group("name").decode("ascii", "replace").strip().strip('"')
        if not имя or имя.upper() == "INBOX":
            continue
        if "\\Sent" in флаги:
            отправленные = имя
        elif имя.lower() in ИМЕНА_ОТПРАВЛЕННЫХ:
            запасной = запасной or имя

    имя = отправленные or запасной
    if имя:
        папки.append((имя, "Отправленные", "отправленное"))
    return папки


def разобрать_флаги(данные) -> dict:
    флаги = {}
    for часть in данные or []:
        строка = часть[0] if isinstance(часть, tuple) else часть
        if not isinstance(строка, bytes):
            continue
        uid = re.search(rb"UID\s+(\d+)", строка)
        найдены = re.search(rb"FLAGS\s+\(([^)]*)\)", строка)
        if uid:
            флаги[int(uid.group(1))] = найдены.group(1).decode("ascii", "replace") if найдены else ""
    return флаги


def разобрать_письма(данные) -> list:
    """Ответ FETCH → [(uid, сырые байты)]. UID сервер обязан вернуть сам."""
    письма = []
    for часть in данные or []:
        if not isinstance(часть, tuple) or часть[1] is None:
            continue
        заголовок = часть[0] if isinstance(часть[0], bytes) else b""
        uid = re.search(rb"UID\s+(\d+)", заголовок)
        if uid:
            письма.append((int(uid.group(1)), часть[1]))
    return письма


def в_кавычках(имя: str) -> str:
    """Имя папки для SELECT. Пробелы и скобки без кавычек сервер не поймёт."""
    return '"%s"' % имя.strip('"').replace("\\", "\\\\").replace('"', '\\"')


# --- докачка --------------------------------------------------------------


def скачать_папку(imap, db, account: str, папка: tuple, args, база_проекта: Path,
                  since: datetime) -> dict:
    имя, ярлык, направление = папка
    итог = {"ярлык": ярлык, "новых": 0, "обновлено": 0, "пропущено": 0, "ошибка": None}

    status, _ = imap.select(в_кавычках(имя), readonly=True)
    if status != "OK":
        итог["ошибка"] = f"не удалось открыть папку {ярлык}"
        return итог

    ящик = найти_ящик(db, account, имя, ярлык, направление)

    # UIDVALIDITY — метка нумерации писем в папке. Сменилась, значит прежние
    # UID указывают уже на другие письма, и архив по этой папке надо собрать заново.
    ответ = imap.response("UIDVALIDITY")[1]
    текущая = int(ответ[0]) if ответ and ответ[0] else None
    if текущая and ящик["uidvalidity"] and ящик["uidvalidity"] != текущая:
        print(f"  {ярлык}: сервер перенумеровал письма, качаю папку заново")
        сбросить_ящик(db, ящик["id"])
        ящик = найти_ящик(db, account, имя, ярлык, направление)
    if текущая:
        db.execute("UPDATE mailboxes SET uidvalidity = ? WHERE id = ?", (текущая, ящик["id"]))

    последний = 0 if args.заново else ящик["last_uid"]
    покрыто = собрано_с(ящик) if not args.заново else None

    # Обычно спрашиваем только то, что новее скачанного. Но если попросили
    # период глубже уже собранного — надо отдельно сходить и за старым:
    # у старых писем номера меньше, и в запрос «новее последнего» они не попадут.
    наборы = [("новые", ["UID", f"{последний + 1}:*"] if последний
                        else ["SINCE", imap_date(since)])]
    if покрыто and since.date() < покрыто.date():
        наборы.append(("старые", ["SINCE", imap_date(since), "BEFORE", imap_date(покрыто)]))
        print(f"  {ярлык}: архив собран с {покрыто:%d.%m.%Y}, добираю письма старее")

    uids = set()
    for вид, условия in наборы:
        status, данные = imap.uid("SEARCH", None, *условия)
        if status != "OK":
            итог["ошибка"] = f"сервер не принял поиск в папке {ярлык}"
            return итог
        найдены = [int(x) for x in (данные[0].split() if данные and данные[0] else [])]
        # Диапазон «N:*» сервер отдаёт даже когда новых писем нет — возвращает
        # последнее имеющееся. Поэтому всё, что не больше последнего, отбрасываем.
        if вид == "новые" and последний:
            найдены = [u for u in найдены if u > последний]
        uids.update(найдены)

    uids = sorted(uids)
    новое_покрытие = min(since, покрыто) if покрыто else since

    if not uids:
        отметить_ящик(db, ящик["id"], последний, новое_покрытие)
        return итог

    if len(uids) > 100:
        print(f"  {ярлык}: писем {len(uids)}, качаю…")

    папка_вложений = база_проекта / args.attach_dir / safe_filename(account)
    максимум = последний

    for начало in range(0, len(uids), ПАРТИЯ):
        партия = uids[начало:начало + ПАРТИЯ]
        набор = ",".join(str(u) for u in партия)

        status, данные = imap.uid("FETCH", набор, "(FLAGS)")
        флаги = разобрать_флаги(данные) if status == "OK" else {}

        # PEEK — письмо не станет прочитанным от того, что мы его забрали
        status, данные = imap.uid("FETCH", набор, "(BODY.PEEK[])")
        if status != "OK":
            итог["ошибка"] = f"сервер оборвал выдачу писем в папке {ярлык}"
            break

        for uid, сырое in разобрать_письма(данные):
            msg = email.message_from_bytes(сырое, policy=default)
            итог[сохранить_письмо(
                db, ящик, uid, msg, флаги.get(uid, ""),
                папка_вложений / f"{uid}", база_проекта, not args.без_вложений,
                перезаписать=args.заново,
            )] += 1
            максимум = max(максимум, uid)

        # Сохраняем после каждой партии: оборвалась связь — скачанное осталось,
        # следующий запуск продолжит с этого места, а не начнёт сначала.
        db.execute(
            "UPDATE mailboxes SET last_uid = ?, synced_at = ? WHERE id = ?",
            (максимум, datetime.now().isoformat(timespec="seconds"), ящик["id"]),
        )
        db.commit()
        if len(uids) > 100:
            print(f"  {ярлык}: {min(начало + ПАРТИЯ, len(uids))} из {len(uids)}")

    # Глубину отмечаем только когда всё скачалось: оборвались на середине —
    # значит период ещё не собран, и в следующий раз надо повторить.
    if not итог["ошибка"]:
        отметить_ящик(db, ящик["id"], максимум, новое_покрытие)
        db.commit()

    return итог


def скачать_ящик(account: dict, db, args, база_проекта: Path, since: datetime) -> dict:
    """Один ящик целиком. Упавший ящик не должен ронять остальные."""
    итог = {"user": account["user"], "новых": 0, "обновлено": 0, "ошибка": None, "папки": []}

    try:
        imap = imaplib.IMAP4_SSL(account["host"], fetch_mail.IMAP_PORT)
    except OSError as e:
        итог["ошибка"] = f"нет связи с сервером {account['host']} ({e})"
        return итог

    try:
        try:
            imap.login(account["user"], account["password"])
        except imaplib.IMAP4.error:
            итог["ошибка"] = (
                "сервер не принял логин или пароль. Нужен ПАРОЛЬ ДЛЯ ВНЕШНЕГО "
                "ПРИЛОЖЕНИЯ, а не обычный пароль от почты"
            )
            return итог

        папки = разобрать_папки(imap.list()[1])
        if args.только_входящие:
            папки = папки[:1]
        elif len(папки) == 1:
            print(f"  {account['user']}: папка отправленных не найдена, беру только входящие")

        for папка in папки:
            результат = скачать_папку(imap, db, account["user"], папка, args, база_проекта, since)
            итог["новых"] += результат["новых"]
            итог["обновлено"] += результат["обновлено"]
            итог["папки"].append(результат)
            if результат["ошибка"]:
                итог["ошибка"] = результат["ошибка"]
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    return итог


def команда_скачать(args, база_проекта: Path) -> int:
    accounts = load_accounts(load_env(база_проекта / ".env"))
    db = открыть_базу(база_проекта / args.база)
    since = parse_date(args.since) if args.since else datetime.now() - timedelta(days=args.days)

    пусто = not db.execute("SELECT 1 FROM messages LIMIT 1").fetchone()
    if пусто:
        print(f"Архив пуст, первая заливка с {since:%d.%m.%Y} — это может занять несколько минут.")

    результаты = [скачать_ящик(a, db, args, база_проекта, since) for a in accounts]
    db.commit()

    новых = sum(r["новых"] for r in результаты)
    обновлено = sum(r["обновлено"] for r in результаты)
    всего = db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    подробно = ", ".join(
        f"{p['ярлык'].lower()} {p['новых']}"
        for r in результаты for p in r["папки"] if p["новых"]
    )
    print(f"Докачано писем: {новых}" + (f" ({подробно})" if подробно else "")
          + (f", перечитано заново: {обновлено}" if обновлено else "")
          + f". Всего в архиве: {всего}.")

    ошибки = [r for r in результаты if r["ошибка"]]
    for r in ошибки:
        print(f"ВНИМАНИЕ: {r['user']} — {r['ошибка']}")
    db.close()
    return 1 if ошибки and len(ошибки) == len(результаты) else 0


# --- запросы к архиву -----------------------------------------------------


def условия_периода(args) -> tuple:
    """Кусок WHERE и параметры к нему по --дней / --с / --по."""
    куски, параметры = [], []
    if args.since:
        куски.append("m.ts >= ?")
        параметры.append(int(parse_date(args.since).timestamp()))
    elif args.days:
        куски.append("m.ts >= ?")
        параметры.append(int((datetime.now() - timedelta(days=args.days)).timestamp()))
    if args.before:
        куски.append("m.ts < ?")
        параметры.append(int((parse_date(args.before) + timedelta(days=1)).timestamp()))
    if args.account:
        куски.append("b.account LIKE ?")
        параметры.append(f"%{args.account}%")
    if args.sender:
        куски.append("(m.sender LIKE ? OR m.sender_addr LIKE ?)")
        параметры += [f"%{args.sender}%"] * 2
    if args.subject:
        куски.append("m.subject LIKE ?")
        параметры.append(f"%{args.subject}%")
    if getattr(args, "unread", False):
        куски.append("m.unread = 1")
    if getattr(args, "direction", None) == "входящие":
        куски.append("b.direction = 'входящее'")
    elif getattr(args, "direction", None) == "отправленные":
        куски.append("b.direction = 'отправленное'")
    return (" AND ".join(куски) or "1=1"), параметры


def кратко(текст: str, длина: int) -> str:
    текст = (текст or "").strip()
    return текст if len(текст) <= длина else текст[:длина - 1] + "…"


def когда(ts: int) -> str:
    дата = datetime.fromtimestamp(ts)
    if дата.year == datetime.now().year:
        return дата.strftime("%d.%m %H:%M")
    return дата.strftime("%d.%m.%Y")


def вложения_письма(строка) -> list:
    if not строка["attachments"]:
        return []
    try:
        return json.loads(строка["attachments"])
    except (ValueError, TypeError):
        return []


def строка_письма(r) -> str:
    метки = ""
    файлы = вложения_письма(r)
    if файлы:
        метки += f" 📎{len(файлы)}"
    if r["unread"]:
        метки += " ●"
    return (f"#{r['id']} {когда(r['ts'])} | {кратко(r['sender'], 45)} | "
            f"{кратко(r['subject'], 90)}{метки}")


def команда_сводка(args, база_проекта: Path) -> int:
    db = открыть_базу(база_проекта / args.база)
    где, параметры = условия_периода(args)
    строки = db.execute(
        f"""SELECT m.*, b.account, b.label, b.direction
            FROM messages m JOIN mailboxes b ON b.id = m.mailbox_id
            WHERE {где} ORDER BY m.ts DESC""",
        параметры,
    ).fetchall()

    правила = [] if args.всё else load_ignore_rules(база_проекта / args.ignore_file)
    мусорные = {r["id"] for r in строки if is_junk(r["sender"], r["subject"], правила)}
    мусор = len(мусорные)
    строки = [r for r in строки if r["id"] not in мусорные]

    период = f"за {args.days} дн." if args.days and not args.since else "за выбранный период"
    if not строки:
        print(f"Писем в архиве {период} нет."
              + (f" Отброшено по фильтру мусора: {мусор}." if мусор else "")
              + " Свежие письма — `mail.py скачать`.")
        db.close()
        return 0

    print(f"Писем {период}: {len(строки)}. Последняя докачка: {последняя_докачка(db)}")
    показано = 0
    for (аккаунт, ярлык), группа in группы(строки):
        видно = группа[:max(0, args.limit - показано)]
        if not видно:
            continue
        # В заголовке всегда полное число писем в папке, даже если ниже
        # показана только часть: иначе цифры в сводке перестают сходиться.
        print(f"\n## {аккаунт} — {ярлык} ({len(группа)})")
        for r in видно:
            print(строка_письма(r))
            показано += 1

    if показано < len(строки):
        print(f"\n…показано {показано} из {len(строки)}. Остальное — с `--сколько {len(строки)}`.")
    if мусор:
        print(f"\n⚪ Отброшено по фильтру мусора: {мусор} (показать — `--всё`).")
    print("\nТекст письма целиком: `mail.py письмо НОМЕР` (номера — это #142).")
    db.close()
    return 0


def группы(строки) -> list:
    """Письма по ящикам и папкам, порядок сохраняется."""
    порядок, собрано = [], {}
    for r in строки:
        ключ = (r["account"], r["label"])
        if ключ not in собрано:
            собрано[ключ] = []
            порядок.append(ключ)
        собрано[ключ].append(r)
    return [(ключ, собрано[ключ]) for ключ in порядок]


def последняя_докачка(db) -> str:
    значение = db.execute("SELECT MAX(synced_at) FROM mailboxes").fetchone()[0]
    if not значение:
        return "ещё не было"
    try:
        return datetime.fromisoformat(значение).strftime("%d.%m %H:%M")
    except ValueError:
        return значение


def запрос_fts(слова: list, точно: bool, любое: bool) -> str:
    """
    Слова человека → запрос FTS5.

    По умолчанию к каждому слову дописывается «*»: полнотекстовый поиск не знает
    русской морфологии, и без этого «тендер» не найдёт «тендеры» и «тендерная».
    Кавычки нужны, чтобы точка, собака или дефис не сошли за синтаксис запроса.
    """
    части = []
    for слово in слова:
        чистое = слово.strip().replace('"', '""')
        if not чистое:
            continue
        части.append(f'"{чистое}"' if точно else f'"{чистое}"*')
    return (" OR " if любое else " ").join(части)


def варианты_запроса(слова: list, точно: bool, любое: bool):
    """
    Запрос как есть, а если по нему пусто — он же по основам слов.

    «Тендер» приставкой «*» находит «тендеры», но «заявка» не найдёт «заявок»:
    расходятся не окончания, а буквы внутри. Человек с телефона пишет слово
    целиком, поэтому мы сами отрезаем по букве с конца и пробуем снова.
    """
    yield запрос_fts(слова, точно, любое), None
    if точно:
        return
    было = {tuple(слова)}
    for срез in (1, 2, 3):
        короче = [с[:-срез] if len(с) - срез >= 4 else с for с in слова]
        if tuple(короче) in было:
            continue
        было.add(tuple(короче))
        yield запрос_fts(короче, точно, любое), " ".join(короче)


def команда_поиск(args, база_проекта: Path) -> int:
    db = открыть_базу(база_проекта / args.база)
    где, параметры = условия_периода(args)

    # «Что приходило от Иванова» — вопрос без единого слова для поиска.
    # Тогда это просто список писем по фильтрам, без полнотекстового запроса.
    if not args.слова:
        строки = db.execute(
            f"""SELECT m.*, b.account, b.label FROM messages m
                JOIN mailboxes b ON b.id = m.mailbox_id
                WHERE {где} ORDER BY m.ts DESC LIMIT ?""",
            параметры + [args.limit],
        ).fetchall()
        if not строки:
            print(f"Таких писем в архиве нет (в нём {охват(db)}).")
        else:
            print(f"Найдено: {len(строки)}")
            for r in строки:
                print(строка_письма(r))
            print("\nТекст целиком: `mail.py письмо НОМЕР`.")
        db.close()
        return 0

    строки, по_основе = [], None
    for запрос, основа in варианты_запроса(args.слова, args.точно, args.любое):
        try:
            строки = db.execute(
                f"""SELECT m.*, b.account, b.label,
                           snippet(search, 2, '', '', '…', 14) AS кусок
                    FROM search
                    JOIN messages m ON m.id = search.rowid
                    JOIN mailboxes b ON b.id = m.mailbox_id
                    WHERE search MATCH ? AND {где}
                    ORDER BY m.ts DESC LIMIT ?""",
                [запрос] + параметры + [args.limit],
            ).fetchall()
        except sqlite3.OperationalError as e:
            db.close()
            sys.exit(f"Не понял запрос «{' '.join(args.слова)}»: {e}")
        if строки:
            по_основе = основа
            break

    if not строки:
        всего = db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
        print(f"По запросу «{' '.join(args.слова)}» в архиве ничего нет "
              f"(в нём {всего} писем, {охват(db)}).")
        print("Если письмо старее — сначала `mail.py скачать --дней 180`.")
        db.close()
        return 0

    print(f"Найдено: {len(строки)}"
          + (f" (показаны первые {args.limit})" if len(строки) == args.limit else "")
          + (f". Точного совпадения не было, искал по основе: «{по_основе}»" if по_основе else ""))
    for r in строки:
        print(строка_письма(r))
        кусок = " ".join((r["кусок"] or "").split())
        if кусок:
            print(f"    {кратко(кусок, 160)}")
    print("\nТекст целиком: `mail.py письмо НОМЕР`.")
    db.close()
    return 0


def команда_письмо(args, база_проекта: Path) -> int:
    db = открыть_базу(база_проекта / args.база)
    for номер in args.номера:
        r = db.execute(
            """SELECT m.*, b.account, b.label FROM messages m
               JOIN mailboxes b ON b.id = m.mailbox_id WHERE m.id = ?""",
            (номер,),
        ).fetchone()
        if not r:
            print(f"Письма #{номер} в архиве нет.")
            continue

        print(f"--- #{r['id']} — {r['label'].lower()}, ящик {r['account']}")
        print(f"Тема: {r['subject']}")
        print(f"От: {r['sender']}")
        if r["recipient"]:
            print(f"Кому: {кратко(r['recipient'], 200)}")
        print(f"Дата: {r['date_text']}")
        for файл in вложения_письма(r):
            print(f"Вложение: {файл['имя']} → {файл['путь']}")
        текст = r["body"] or "(пустое письмо)"
        if len(текст) > args.знаков:
            текст = текст[:args.знаков] + f"\n…[показано {args.знаков} знаков из {len(r['body'])}]"
        print()
        print(текст)
        print()
    db.close()
    return 0


def охват(db) -> str:
    границы = db.execute("SELECT MIN(ts), MAX(ts) FROM messages").fetchone()
    if not границы or not границы[0]:
        return "архив пуст"
    с, по = (datetime.fromtimestamp(x).strftime("%d.%m.%Y") for x in границы)
    return f"с {с} по {по}"


def команда_статистика(args, база_проекта: Path) -> int:
    путь = база_проекта / args.база
    db = открыть_базу(путь)
    всего = db.execute("SELECT COUNT(*) FROM messages").fetchone()[0]
    размер = путь.stat().st_size / 1024 / 1024 if путь.exists() else 0

    print(f"Архив: писем {всего}, {охват(db)}, {размер:.1f} МБ.")
    for r in db.execute(
        """SELECT b.account, b.label, COUNT(m.id) AS сколько, b.synced_at,
                  b.covered_since, SUM(m.unread) AS непрочитанных
           FROM mailboxes b LEFT JOIN messages m ON m.mailbox_id = b.id
           GROUP BY b.id ORDER BY b.account, b.label"""
    ):
        докачка = r["synced_at"][:16].replace("T", " ") if r["synced_at"] else "не было"
        непрочитано = f", непрочитанных {r['непрочитанных']}" if r["непрочитанных"] else ""
        собран = собрано_с(r)
        глубина = f", собран с {собран:%d.%m.%Y}" if собран else ""
        print(f"  {r['account']} — {r['label']}: {r['сколько']}{непрочитано}"
              f"{глубина}, докачка {докачка}")
    if not всего:
        print("Пока пусто. Заполнить: `mail.py скачать --дней 30`.")
    db.close()
    return 0


# --- точка входа ----------------------------------------------------------


def добавить_фильтры(p, дней_по_умолчанию=None) -> None:
    p.add_argument("--дней", "--days", dest="days", type=int, default=дней_по_умолчанию,
                   help="за сколько последних дней")
    p.add_argument("--с", "--since", dest="since", help="с какой даты: 05.09.2026")
    p.add_argument("--по", "--before", dest="before", help="по какую дату включительно")
    p.add_argument("--от", "--from", dest="sender", help="отправитель: имя, адрес или домен")
    p.add_argument("--тема", "--subject", dest="subject", help="слово в теме")
    p.add_argument("--ящик", "--mailbox", dest="account", help="если ящиков несколько")
    p.add_argument("--входящие", "--inbox", dest="direction",
                   action="store_const", const="входящие")
    p.add_argument("--отправленные", "--sent", dest="direction",
                   action="store_const", const="отправленные",
                   help="только то, что писал он сам")


def собрать_разбор() -> argparse.ArgumentParser:
    """
    Команды и ключи названы по-русски, но у каждого есть латинский двойник.
    Русское слово в .bat-файле не выживает: cmd.exe не читает UTF-8 в батниках.
    """
    parser = argparse.ArgumentParser(
        description="Локальный архив почты: докачка и запросы без обращения к серверу")
    parser.add_argument("--база", "--db", dest="база", default=БАЗА,
                        help=f"файл архива (по умолчанию {БАЗА})")
    команды = parser.add_subparsers(dest="команда", required=True)

    качать = команды.add_parser("скачать", aliases=["sync"],
                                help="докачать новые письма с сервера")
    качать.add_argument("--дней", "--days", dest="days", type=int, default=30,
                        help="глубина первой заливки (по умолчанию 30)")
    качать.add_argument("--с", "--since", dest="since", help="первая заливка с даты: 01.09.2026")
    качать.add_argument("--заново", "--refetch", dest="заново", action="store_true",
                        help="перечитать папки целиком, не доверяя отметке о последнем письме")
    качать.add_argument("--только-входящие", "--inbox-only", dest="только_входящие",
                        action="store_true", help="не трогать папку отправленных")
    качать.add_argument("--без-вложений", "--no-attachments", dest="без_вложений",
                        action="store_true", help="не сохранять файлы из писем на диск")
    качать.add_argument("--attach-dir", dest="attach_dir", default="Почта/вложения",
                        help="куда складывать вложения")
    качать.set_defaults(работа=команда_скачать)

    сводка = команды.add_parser("сводка", aliases=["summary"],
                                help="одна строка на письмо — для разбора почты")
    добавить_фильтры(сводка, дней_по_умолчанию=3)
    сводка.add_argument("--непрочитанные", "--unread", dest="unread", action="store_true")
    сводка.add_argument("--сколько", "--limit", dest="limit", type=int, default=ПРЕДЕЛ_СВОДКИ)
    сводка.add_argument("--всё", "--all", dest="всё", action="store_true",
                        help="не отбрасывать мусор по Шаблоны/игнор.txt")
    сводка.add_argument("--ignore-file", dest="ignore_file", default="Шаблоны/игнор.txt")
    сводка.set_defaults(работа=команда_сводка)

    поиск = команды.add_parser("поиск", aliases=["search"], help="поиск по всему архиву")
    поиск.add_argument("слова", nargs="*",
                       help="что искать; без слов — просто список по фильтрам")
    добавить_фильтры(поиск)
    поиск.add_argument("--точно", "--exact", dest="точно", action="store_true",
                       help="искать слово как есть, без учёта окончаний")
    поиск.add_argument("--любое", "--any", dest="любое", action="store_true",
                       help="достаточно любого из слов")
    поиск.add_argument("--сколько", "--limit", dest="limit", type=int, default=ПРЕДЕЛ_ПОИСКА)
    поиск.set_defaults(работа=команда_поиск)

    письмо = команды.add_parser("письмо", aliases=["show"], help="полный текст письма по номеру")
    письмо.add_argument("номера", nargs="+", type=int)
    письмо.add_argument("--знаков", "--chars", dest="знаков", type=int, default=4000,
                        help="сколько знаков текста показать")
    письмо.set_defaults(работа=команда_письмо)

    статистика = команды.add_parser("статистика", aliases=["stats"], help="что лежит в архиве")
    статистика.set_defaults(работа=команда_статистика)
    return parser


def main(argv=None) -> int:
    args = собрать_разбор().parse_args(argv)
    return args.работа(args, Path(__file__).resolve().parent)


if __name__ == "__main__":
    sys.exit(main())
