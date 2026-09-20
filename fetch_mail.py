"""
Выгрузка писем из нескольких почтовых ящиков по IMAP в markdown-файл для разбора ИИ.

Поддерживаются Mail.ru, Gmail, Яндекс, Outlook, Rambler — сервер определяется
по адресу автоматически.

Только чтение: ящик открывается в режиме readonly, письма НЕ помечаются
прочитанными, ничего не удаляется и не отправляется.

Запуск:
    python fetch_mail.py              # письма за последние сутки
    python fetch_mail.py --days 3     # за последние 3 дня
    python fetch_mail.py --unseen     # только непрочитанные (за указанный период)
    python fetch_mail.py --all        # не отбрасывать мусор из Шаблоны/игнор.txt

Логины и пароли берутся из файла .env рядом со скриптом (см. .env.example).
Используется только стандартная библиотека Python 3.8+.
"""

import argparse
import email
import html
import imaplib
import os
import re
import sys
from datetime import datetime, timedelta
from email.policy import default
from pathlib import Path

# Codex и планировщики читают вывод через пайп, где кодировка по умолчанию
# не UTF-8 — без этого все русские сообщения приедут абракадаброй.
for поток in (sys.stdout, sys.stderr):
    if hasattr(поток, "reconfigure"):
        поток.reconfigure(encoding="utf-8", errors="replace")

# Сервер IMAP по домену адреса — чтобы не заставлять заполнять его руками
IMAP_HOSTS = {
    "mail.ru": "imap.mail.ru",
    "inbox.ru": "imap.mail.ru",
    "bk.ru": "imap.mail.ru",
    "list.ru": "imap.mail.ru",
    "internet.ru": "imap.mail.ru",
    "gmail.com": "imap.gmail.com",
    "googlemail.com": "imap.gmail.com",
    "yandex.ru": "imap.yandex.ru",
    "yandex.com": "imap.yandex.ru",
    "ya.ru": "imap.yandex.ru",
    "outlook.com": "outlook.office365.com",
    "hotmail.com": "outlook.office365.com",
    "live.com": "outlook.office365.com",
    "rambler.ru": "imap.rambler.ru",
}
IMAP_PORT = 993

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Символы, запрещённые в именах файлов Windows, плюс управляющие
BAD_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
# Имена устройств Windows: такой файл не создать ни с каким расширением
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
MAX_NAME_LEN = 120

# Короче этого plain-часть считаем заглушкой вида «письмо содержит HTML»
PLAIN_STUB_LEN = 30

# Служебные «вложения», которые на самом деле вложениями не являются:
# криптоподписи писем. Корпоративная почта цепляет их почти к каждому письму.
SKIP_TYPES = {
    "application/pkcs7-signature",
    "application/x-pkcs7-signature",
    "application/pgp-signature",
}
SKIP_NAMES = {"smime.p7s", "signature.asc"}

# Значения из .env.example. Если их не заменили — это не ящик, а забытый
# кусок шаблона: пытаться войти с ним бессмысленно, а ошибка про пароль
# уводит в сторону.
ПРИМЕРЫ = {
    "name@mail.ru",
    "name@gmail.com",
    "name@yandex.ru",
    "пароль_для_внешнего_приложения",
    "шестнадцатьсимволов",
}


# --- настройки ------------------------------------------------------------


def load_env(path: Path) -> dict:
    if not path.exists():
        sys.exit(f"Нет файла {path}. Скопируйте .env.example в .env и заполните.")
    env = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def guess_host(user: str) -> str:
    domain = user.rpartition("@")[2].lower()
    host = IMAP_HOSTS.get(domain)
    if not host:
        sys.exit(
            f"Не знаю IMAP-сервер для адреса {user}. "
            f"Допишите в .env строку с адресом сервера, например MAIL_1_HOST=imap.example.ru"
        )
    return host


def из_примера(user: str, password: str) -> bool:
    return user.lower() in ПРИМЕРЫ or password in ПРИМЕРЫ


def load_accounts(env: dict) -> list:
    """Ящики из .env. Поддерживает и один ящик, и пронумерованные MAIL_1_, MAIL_2_…"""
    accounts = []
    пропущены = []

    # старый формат с одним ящиком
    if env.get("MAIL_USER"):
        if из_примера(env["MAIL_USER"], env.get("MAIL_APP_PASSWORD", "")):
            пропущены.append(env["MAIL_USER"])
        else:
            accounts.append({
                "user": env["MAIL_USER"],
                "password": env.get("MAIL_APP_PASSWORD", ""),
                "host": env.get("MAIL_HOST") or guess_host(env["MAIL_USER"]),
            })

    # пронумерованные ящики
    номера = sorted(
        {int(m.group(1)) for k in env if (m := re.fullmatch(r"MAIL_(\d+)_USER", k))}
    )
    for i in номера:
        user = env.get(f"MAIL_{i}_USER", "").strip()
        if not user:
            continue
        password = env.get(f"MAIL_{i}_PASSWORD") or env.get(f"MAIL_{i}_APP_PASSWORD") or ""
        if из_примера(user, password):
            пропущены.append(user)
            continue
        accounts.append({
            "user": user,
            "password": password,
            "host": env.get(f"MAIL_{i}_HOST") or guess_host(user),
        })

    for user in пропущены:
        print(f"Пропущен ящик из шаблона: {user} — замените или удалите эти строки в .env")

    if not accounts:
        if пропущены:
            sys.exit(
                "В .env остались только строки из примера. Впишите настоящий адрес "
                "и пароль для внешнего приложения вместо них."
            )
        sys.exit(
            "В .env не найдено ни одного ящика. Нужны строки вида "
            "MAIL_1_USER и MAIL_1_PASSWORD (см. .env.example)."
        )
    без_пароля = [a["user"] for a in accounts if not a["password"]]
    if без_пароля:
        sys.exit(f"Не заполнен пароль для ящика: {', '.join(без_пароля)}")
    return accounts


def load_ignore_rules(path: Path) -> list:
    """Строки-фильтры мусора. Пустой файл или его отсутствие — фильтра нет."""
    if not path.exists():
        return []
    rules = []
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            rules.append(line.lower())
    return rules


def is_junk(sender: str, subject: str, rules: list) -> bool:
    стог = f"{sender} {subject}".lower()
    return any(rule in стог for rule in rules)


# --- разбор письма --------------------------------------------------------


def html_to_text(text: str) -> str:
    text = re.sub(r"(?is)<(script|style).*?</\1>", "", text)
    text = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>|</li>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n\s*\n+", "\n\n", text)
    return text.strip()


def part_text(part) -> str:
    """Текст одной части письма с учётом её кодировки."""
    if part is None:
        return ""
    try:
        content = part.get_content()
    except Exception:
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset() or "utf-8"
        content = payload.decode(charset, errors="replace")
    if part.get_content_type() == "text/html":
        content = html_to_text(content)
    return content.strip()


def get_text(msg) -> str:
    body = msg.get_body(preferencelist=("plain", "html"))
    text = part_text(body)
    # Рассылки часто кладут в plain-часть заглушку, а всё содержимое — в HTML.
    if body is not None and body.get_content_type() == "text/plain" and len(text) < PLAIN_STUB_LEN:
        из_html = part_text(msg.get_body(preferencelist=("html",)))
        if len(из_html) > len(text):
            return из_html
    return text


def clean_header(value) -> str:
    """Заголовок в одну строку — чтобы не ломать разметку markdown."""
    return re.sub(r"\s+", " ", str(value or "")).strip()


# --- вложения -------------------------------------------------------------


def safe_filename(name: str) -> str:
    """Имя вложения из письма → имя, которое Windows примет без ошибки."""
    if not name:
        return "вложение"
    # Все разделители пути заменяются на "_", а не отрезаются: в именах от
    # заказчиков слеш встречается как часть названия («Форма КП / 2026.xlsx»),
    # и терять начало имени нельзя. Выйти из папки после замены невозможно.
    name = BAD_CHARS.sub("_", name)
    # Windows молча режет точки и пробелы в конце имени
    name = name.strip(" .")
    if not name:
        return "вложение"
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    if stem.upper() in RESERVED_NAMES:
        stem = f"_{stem}"
    if len(stem) > MAX_NAME_LEN:
        stem = stem[:MAX_NAME_LEN]
    return f"{stem}.{ext}" if ext else stem


def unique_path(directory: Path, name: str) -> Path:
    """Путь, который не затрёт уже лежащий файл: имя, имя_1, имя_2…"""
    candidate = directory / name
    if not candidate.exists():
        return candidate
    stem, dot, ext = name.rpartition(".")
    if not dot:
        stem, ext = name, ""
    n = 1
    while True:
        new_name = f"{stem}_{n}.{ext}" if ext else f"{stem}_{n}"
        candidate = directory / new_name
        if not candidate.exists():
            return candidate
        n += 1


def save_attachments(msg, directory: Path) -> list:
    """Сохраняет вложения письма на диск. Возвращает список (имя из письма, путь)."""
    saved = []
    for part in msg.iter_attachments():
        raw_name = part.get_filename()
        if not raw_name:
            continue
        if part.get_content_type().lower() in SKIP_TYPES or raw_name.lower() in SKIP_NAMES:
            continue  # подпись письма, а не вложение
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        directory.mkdir(parents=True, exist_ok=True)
        path = unique_path(directory, safe_filename(raw_name))
        path.write_bytes(payload)
        saved.append((raw_name, path))
    return saved


# --- выгрузка одного ящика ------------------------------------------------


def imap_date(dt: datetime) -> str:
    return f"{dt.day:02d}-{MONTHS[dt.month - 1]}-{dt.year}"


def parse_date(текст: str) -> datetime:
    """Дата из командной строки: 05.09.2026 или 2026-09-05."""
    for формат in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(текст, формат)
        except ValueError:
            continue
    sys.exit(f"Не понял дату «{текст}». Пишите так: 05.09.2026")


def build_criteria(args, since: datetime, before) -> tuple:
    """
    Условия поиска для IMAP.

    Возвращает (условия, литерал, доп_фильтры):
      условия     — список аргументов SEARCH;
      литерал     — русский текст, который уйдёт отдельным блоком (или None);
      доп_фильтры — что не влезло в запрос и фильтруется уже у нас.

    IMAP умеет принять только один текстовый блок с не-латиницей за запрос,
    и он обязан идти последним аргументом. Остальное досеиваем сами.
    """
    условия = []
    if args.unseen:
        условия.append("UNSEEN")
    условия += ["SINCE", imap_date(since)]
    if before:
        условия += ["BEFORE", imap_date(before)]

    литерал = None
    доп_фильтры = []
    for ключ, значение in (("FROM", args.from_addr),
                           ("SUBJECT", args.subject),
                           ("TEXT", args.search)):
        if not значение:
            continue
        if значение.isascii():
            условия += [ключ, f'"{значение}"']
        elif литерал is None:
            литерал = (ключ, значение)
        else:
            доп_фильтры.append((ключ, значение.lower()))

    if литерал:
        условия.append(литерал[0])  # литерал подставится следом
    return условия, литерал, доп_фильтры


def подходит_под_доп_фильтры(msg, доп_фильтры) -> bool:
    for ключ, значение in доп_фильтры:
        поле = "from" if ключ == "FROM" else "subject"
        if значение not in clean_header(msg.get(поле)).lower():
            return False
    return True


def fetch_account(account: dict, args, base: Path, attach_root: Path, rules: list,
                  since: datetime, before) -> dict:
    """Выгружает один ящик. Ошибка одного ящика не должна ронять остальные."""
    итог = {"user": account["user"], "lines": [], "писем": 0, "мусора": 0,
            "вложений": 0, "мусор_темы": [], "ошибка": None}

    criteria, литерал, доп_фильтры = build_criteria(args, since, before)

    try:
        imap = imaplib.IMAP4_SSL(account["host"], IMAP_PORT)
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

        status, _ = imap.select(args.folder, readonly=True)
        if status != "OK":
            итог["ошибка"] = f"не удалось открыть папку {args.folder}"
            return итог

        if литерал:
            # Русский текст уходит отдельным блоком: иначе imaplib не смог бы
            # его закодировать, а сервер — распознать.
            imap.literal = литерал[1].encode("utf-8")
            status, data = imap.search("UTF-8", *criteria)
        else:
            status, data = imap.search(None, *criteria)
        if status != "OK":
            итог["ошибка"] = "сервер не принял условия поиска"
            return итог
        ids = data[0].split() if data and data[0] else []

        # В кратком режиме тело не скачиваем — только заголовки
        часть = "(BODY.PEEK[HEADER])" if args.brief else "(BODY.PEEK[])"

        номер = 0
        for msg_id in reversed(ids):  # сначала новые
            status, msg_data = imap.fetch(msg_id, часть)  # PEEK — не помечать прочитанным
            if status != "OK":
                continue
            raw = next((part[1] for part in msg_data if isinstance(part, tuple)), None)
            if raw is None:
                continue
            msg = email.message_from_bytes(raw, policy=default)

            тема = clean_header(msg.get("subject")) or "(без темы)"
            отправитель = clean_header(msg.get("from"))

            if not подходит_под_доп_фильтры(msg, доп_фильтры):
                continue

            if not args.all and is_junk(отправитель, тема, rules):
                итог["мусора"] += 1
                итог["мусор_темы"].append(f"{отправитель} — {тема}")
                continue

            номер += 1
            if args.brief:
                дата = clean_header(msg.get("date"))
                итог["lines"].append(f"{номер}. {дата} — {отправитель} — {тема}")
                continue

            text = get_text(msg)
            if len(text) > args.max_chars:
                text = text[: args.max_chars] + "\n…[текст обрезан]"
            attachments = save_attachments(msg, attach_root / safe_filename(account["user"]) / f"{номер:02d}")
            итог["вложений"] += len(attachments)

            итог["lines"] += [
                "---",
                f"### {номер}. {тема}",
                f"- **От:** {отправитель}",
                f"- **Дата:** {clean_header(msg.get('date'))}",
            ]
            if attachments:
                итог["lines"].append("- **Вложения:**")
                for raw_name, path in attachments:
                    try:
                        shown = path.relative_to(base).as_posix()
                    except ValueError:  # папка вложений вне проекта
                        shown = path.as_posix()
                    итог["lines"].append(f"  - {clean_header(raw_name)} → `{shown}`")
            итог["lines"] += ["", text or "(пустое письмо)", ""]

        итог["писем"] = номер
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    return итог


# --- точка входа ----------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(description="Выгрузка писем для разбора ИИ")
    parser.add_argument("--days", type=int, default=1, help="за сколько дней (по умолчанию 1)")
    parser.add_argument("--since", help="с какой даты: 05.09.2026 (вместо --days)")
    parser.add_argument("--before", help="по какую дату включительно: 10.09.2026")
    parser.add_argument("--search", help="слово в тексте письма")
    parser.add_argument("--from-addr", dest="from_addr", help="адрес или домен отправителя")
    parser.add_argument("--subject", help="слово в теме")
    parser.add_argument("--brief", action="store_true",
                        help="только заголовки: от кого, когда, тема (для больших периодов)")
    parser.add_argument("--unseen", action="store_true", help="только непрочитанные")
    parser.add_argument("--folder", default="INBOX", help="папка ящика (по умолчанию INBOX)")
    parser.add_argument("--out", default="Почта/выгрузки", help="куда сохранять файл")
    parser.add_argument("--attach-dir", default="Почта/вложения", help="куда сохранять вложения")
    parser.add_argument("--ignore-file", default="Шаблоны/игнор.txt", help="список мусорных отправителей")
    parser.add_argument("--all", action="store_true", help="не отбрасывать мусор")
    parser.add_argument("--max-chars", type=int, default=3000, help="макс. длина текста письма")
    parser.add_argument("--open", action="store_true", help="открыть папку с выгрузкой после запуска")
    args = parser.parse_args()

    base = Path(__file__).resolve().parent
    # Рабочие папки не лежат в git — создаём их сами, чтобы после переноса
    # на другой компьютер структура была на месте.
    for папка in ("Почта/выгрузки", "Почта/вложения", "КП",
                  "Отчёты", "На проверку", "Шаблоны/примеры КП"):
        (base / папка).mkdir(parents=True, exist_ok=True)

    accounts = load_accounts(load_env(base / ".env"))
    rules = load_ignore_rules(base / args.ignore_file)

    now = datetime.now()
    since = parse_date(args.since) if args.since else now - timedelta(days=args.days)
    before = parse_date(args.before) + timedelta(days=1) if args.before else None
    attach_root = base / args.attach_dir / f"{now:%Y-%m-%d_%H-%M}"

    результаты = [fetch_account(a, args, base, attach_root, rules, since, before)
                  for a in accounts]

    всего = sum(r["писем"] for r in результаты)
    мусора = sum(r["мусора"] for r in результаты)
    вложений = sum(r["вложений"] for r in результаты)
    ошибки = [r for r in результаты if r["ошибка"]]

    период = f"с {since:%d.%m.%Y}" + (f" по {args.before}" if args.before else "")
    условия = [f"«{з}» в {п}" for п, з in
               (("отправителе", args.from_addr), ("теме", args.subject), ("тексте", args.search)) if з]

    lines = [
        f"# Почта: выгрузка {now:%d.%m.%Y %H:%M}",
        "",
        f"Период: {период}. Папка: {args.folder}. "
        f"{'Только непрочитанные. ' if args.unseen else ''}"
        + (f"Поиск: {', '.join(условия)}. " if условия else "")
        + (f"Краткий режим, без текстов. " if args.brief else "")
        + f"Ящиков: {len(accounts)}. Писем: {всего}."
        + (f" Отброшено как мусор: {мусора}." if мусора else ""),
        "",
    ]

    for r in ошибки:
        lines += [f"> ⚠️ **{r['user']}** — {r['ошибка']}", ""]

    for r in результаты:
        if r["ошибка"]:
            continue
        lines += [f"## Ящик: {r['user']} — писем: {r['писем']}", ""]
        lines += r["lines"] or ["(новых писем нет)", ""]

    мусор_всего = [t for r in результаты for t in r["мусор_темы"]]
    if мусор_всего:
        lines += [
            "---",
            "## Отброшено как мусор",
            "(по списку в `Шаблоны/игнор.txt`; чтобы увидеть целиком — запуск с `--all`)",
            "",
        ]
        lines += [f"- {t}" for t in мусор_всего]
        lines.append("")

    out_dir = base / args.out
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / f"{now:%Y-%m-%d_%H-%M}_почта.md"
    out_file.write_text("\n".join(lines), encoding="utf-8")

    хвост = f", мусора: {мусора}" if мусора else ""
    print(f"Готово: ящиков {len(accounts)}, писем {всего}{хвост}, вложений {вложений} → {out_file}")
    for r in ошибки:
        print(f"ВНИМАНИЕ: {r['user']} — {r['ошибка']}")

    if args.open:
        try:
            os.startfile(out_dir)  # только Windows, и только по явному флагу
        except Exception:
            pass

    if ошибки and len(ошибки) == len(accounts):
        sys.exit(1)  # ни один ящик не прочитался


if __name__ == "__main__":
    main()
