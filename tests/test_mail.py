"""
Офлайн-тесты локального архива почты. Реальный ящик не нужен: письма
собираются через email.message.EmailMessage, IMAP подменяется заглушкой.

Запуск:  python -m pytest tests -v
"""

import sqlite3
from email.message import EmailMessage

import pytest

import mail

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def письмо(тема, отправитель, текст, дата="Fri, 19 Sep 2026 10:00:00 +0300", кому=None):
    msg = EmailMessage()
    msg["Subject"] = тема
    msg["From"] = отправитель
    msg["Date"] = дата
    if кому:
        msg["To"] = кому
    msg.set_content(текст)
    return msg


ВХОДЯЩИЕ = {
    11: письмо(
        "Извещение о закупке: уборка офисов",
        "Отдел закупок <zakupki@zakazchik.ru>",
        "Приглашаем к участию. Срок подачи заявок — 25.09.2026 до 10:00.",
        дата="Wed, 17 Sep 2026 09:00:00 +0300",
    ),
    12: письмо(
        "Форма для заполнения",
        "Анна Кузнецова <anna@zakazchik.ru>",
        "Заполните форму и верните до конца недели.",
        дата="Thu, 18 Sep 2026 12:00:00 +0300",
    ),
    13: письмо(
        "Скидки сентября!",
        "Рассылка <news@spam-shop.ru>",
        "Только сегодня скидки на всё.",
        дата="Fri, 19 Sep 2026 18:30:00 +0300",
    ),
}
ВХОДЯЩИЕ[12].add_attachment(
    b"xlsx", maintype="application", subtype=XLSX.split("/")[1], filename="Форма КП.xlsx"
)

ОТПРАВЛЕННЫЕ = {
    5: письмо(
        "Re: Форма для заполнения",
        "Денис <denis@mail.ru>",
        "Добрый день! Форму заполнил, во вложении.",
        дата="Thu, 18 Sep 2026 15:00:00 +0300",
        кому="anna@zakazchik.ru",
    ),
}


class ЗаглушкаIMAP:
    """Вместо настоящего IMAP4_SSL. Помнит, как её вызывали."""

    последний = None
    папки = {"INBOX": ВХОДЯЩИЕ, "Sent": ОТПРАВЛЕННЫЕ}
    # Письма старее собранного периода: их отдают только на запрос с BEFORE
    старые = {}
    uidvalidity = {"INBOX": 111, "Sent": 222}

    def __init__(self, host, port):
        self.host = host
        self.readonly = None
        self.команды = []
        self.вышел = False
        self.папка = None
        ЗаглушкаIMAP.последний = self

    def login(self, user, password):
        return "OK", [b"Logged in"]

    def list(self, *a, **kw):
        return "OK", [
            br'(\HasNoChildren) "/" "INBOX"',
            br'(\HasNoChildren \Sent) "/" "Sent"',
            br'(\HasNoChildren \Trash) "/" "Trash"',
        ]

    def select(self, folder, readonly=False):
        self.папка = folder.strip('"')
        self.readonly = readonly
        return "OK", [str(len(self.папки[self.папка])).encode()]

    def response(self, ключ):
        if ключ == "UIDVALIDITY":
            return ключ, [str(self.uidvalidity[self.папка]).encode()]
        return ключ, [None]

    def uid(self, команда, *args):
        self.команды.append((команда.upper(), args))
        письма = self.папки[self.папка]

        if команда.upper() == "SEARCH":
            if "UID" in args:
                начало = int(args[args.index("UID") + 1].split(":")[0])
                найдены = [u for u in письма if u >= начало]
                # Так ведёт себя настоящий сервер: диапазон «N:*» всегда
                # отдаёт хотя бы последнее письмо, даже если новых нет.
                найдены = найдены or [max(письма)]
            elif "BEFORE" in args:
                найдены = sorted(self.старые.get(self.папка, {}))
            else:
                найдены = sorted(письма)
            return "OK", [b" ".join(str(u).encode() for u in найдены)]

        uids = [int(x) for x in args[0].split(",")]
        if "FLAGS" in args[1] and "BODY" not in args[1]:
            return "OK", [b"%d (UID %d FLAGS (\\Seen))" % (i, u) for i, u in enumerate(uids, 1)]

        все = {**self.старые.get(self.папка, {}), **письма}
        ответ = []
        for i, u in enumerate(uids, 1):
            сырое = все[u].as_bytes()
            ответ.append((b"%d (UID %d BODY[] {%d}" % (i, u, len(сырое)), сырое))
            ответ.append(b")")
        return "OK", ответ

    def logout(self):
        self.вышел = True
        return "BYE", [b"Logout"]

    # эти методы вызывать нельзя — почта только на чтение
    def store(self, *a, **kw):
        raise AssertionError("store() — попытка изменить флаги письма")

    def copy(self, *a, **kw):
        raise AssertionError("copy() — попытка переместить письмо")

    def expunge(self, *a, **kw):
        raise AssertionError("expunge() — попытка удалить письма")


@pytest.fixture
def архив(tmp_path, monkeypatch):
    """Готовая база с закачанными письмами. Возвращает (путь к базе, запускалку)."""
    monkeypatch.setattr(mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(
        mail,
        "load_env",
        lambda path: {"MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": "тестовый-пароль"},
    )
    база = tmp_path / "архив.db"

    def запустить(*аргументы):
        return mail.main(["--база", str(база), *аргументы])

    запустить("скачать", "--attach-dir", str(tmp_path / "вложения"))
    return база, запустить, tmp_path


def строки_базы(база, запрос, параметры=()):
    db = sqlite3.connect(str(база))
    db.row_factory = sqlite3.Row
    try:
        return db.execute(запрос, параметры).fetchall()
    finally:
        db.close()


# --- докачка --------------------------------------------------------------


def test_письма_попадают_в_архив(архив):
    база, _, _ = архив
    всего = строки_базы(база, "SELECT COUNT(*) AS n FROM messages")[0]["n"]
    assert всего == 4  # три входящих и одно отправленное


def test_входящие_и_отправленные_различаются(архив):
    база, _, _ = архив
    направления = {
        r["direction"]: r["n"]
        for r in строки_базы(
            база,
            """SELECT b.direction, COUNT(*) AS n FROM messages m
               JOIN mailboxes b ON b.id = m.mailbox_id GROUP BY b.direction""",
        )
    }
    assert направления == {"входящее": 3, "отправленное": 1}


def test_повторная_докачка_не_тянет_старое(архив):
    база, запустить, tmp = архив
    ЗаглушкаIMAP.последний.команды.clear()

    запустить("скачать", "--attach-dir", str(tmp / "вложения"))

    всего = строки_базы(база, "SELECT COUNT(*) AS n FROM messages")[0]["n"]
    assert всего == 4, "письма не должны задвоиться"

    поиски = [args for команда, args in ЗаглушкаIMAP.последний.команды if команда == "SEARCH"]
    assert all("UID" in args for args in поиски), "вторая докачка обязана идти от последнего UID"
    диапазоны = [args[args.index("UID") + 1] for args in поиски]
    assert диапазоны == ["14:*", "6:*"], "спрашиваем только то, что новее скачанного"


def test_докачка_подхватывает_новое(архив, monkeypatch):
    база, запустить, tmp = архив
    новое = dict(ВХОДЯЩИЕ)
    новое[14] = письмо(
        "Срочно: замечания по объекту",
        "Прораб <prorab@zakazchik.ru>",
        "На объекте нашли замечания, нужен ответ сегодня.",
        дата="Sat, 20 Sep 2026 08:00:00 +0300",
    )
    monkeypatch.setitem(ЗаглушкаIMAP.папки, "INBOX", новое)

    запустить("скачать", "--attach-dir", str(tmp / "вложения"))

    темы = [r["subject"] for r in строки_базы(база, "SELECT subject FROM messages")]
    assert "Срочно: замечания по объекту" in темы
    assert len(темы) == 5


def test_углубление_архива_добирает_старое(архив, monkeypatch, capsys):
    """«Скачать за полгода» на непустом архиве обязано сходить за старыми письмами."""
    база, запустить, tmp = архив
    monkeypatch.setitem(ЗаглушкаIMAP.старые, "INBOX", {
        7: письмо("Договор на весенний период", "Заказчик <old@zakazchik.ru>",
                  "Направляю договор.", дата="Mon, 16 Mar 2026 10:00:00 +0300"),
    })

    запустить("скачать", "--дней", "180", "--attach-dir", str(tmp / "вложения"))

    темы = [r["subject"] for r in строки_базы(база, "SELECT subject FROM messages")]
    assert "Договор на весенний период" in темы
    assert "добираю письма старее" in capsys.readouterr().out


def test_обычная_докачка_за_старым_не_лезет(архив, monkeypatch):
    """Ежедневная докачка не должна каждый раз перетряхивать весь ящик."""
    _, запустить, tmp = архив
    ЗаглушкаIMAP.последний.команды.clear()

    запустить("скачать", "--attach-dir", str(tmp / "вложения"))

    поиски = [args for команда, args in ЗаглушкаIMAP.последний.команды if команда == "SEARCH"]
    assert all("BEFORE" not in args for args in поиски)


def test_глубина_архива_запоминается(архив):
    база, _, _ = архив
    отметки = [r["covered_since"] for r in строки_базы(база, "SELECT covered_since FROM mailboxes")]
    assert all(о for о in отметки), "без отметки о глубине углубление не сработает"


def test_база_прежней_версии_дотягивается(архив):
    """Архив клиента не пересобирается при обновлении — колонка добавляется на месте."""
    база, запустить, _ = архив
    import sqlite3 as s
    d = s.connect(str(база))
    d.execute("CREATE TABLE старое AS SELECT * FROM mailboxes")
    d.execute("DROP TABLE mailboxes")
    d.execute("""CREATE TABLE mailboxes (id INTEGER PRIMARY KEY, account TEXT NOT NULL,
                 folder TEXT NOT NULL, label TEXT NOT NULL, direction TEXT NOT NULL,
                 uidvalidity INTEGER, last_uid INTEGER NOT NULL DEFAULT 0, synced_at TEXT,
                 UNIQUE (account, folder))""")
    d.execute("""INSERT INTO mailboxes (id, account, folder, label, direction, uidvalidity,
                 last_uid, synced_at) SELECT id, account, folder, label, direction,
                 uidvalidity, last_uid, synced_at FROM старое""")
    d.commit()
    d.close()

    запустить("статистика")

    отметки = [r["covered_since"] for r in строки_базы(база, "SELECT covered_since FROM mailboxes")]
    assert any(отметки), "глубина должна восстановиться по дате самого старого письма"


def test_заново_обновляет_текст_писем(архив, monkeypatch, capsys):
    """Правки в разборе письма должны доходить до уже скачанного."""
    база, запустить, tmp = архив
    исправленное = dict(ВХОДЯЩИЕ)
    исправленное[13] = письмо(
        "Скидки сентября!",
        "Рассылка <news@spam-shop.ru>",
        "Текст разобрали заново и почистили.",
        дата="Fri, 19 Sep 2026 18:30:00 +0300",
    )
    monkeypatch.setitem(ЗаглушкаIMAP.папки, "INBOX", исправленное)

    запустить("скачать", "--заново", "--attach-dir", str(tmp / "вложения"))

    строки = строки_базы(база, "SELECT body FROM messages WHERE uid = 13")
    assert len(строки) == 1, "перечитанное письмо не должно задвоиться"
    assert "почистили" in строки[0]["body"]
    assert "перечитано заново: 4" in capsys.readouterr().out


def test_заново_не_сдвигает_номера_писем(архив, monkeypatch):
    """Номера агент показывает Денису — после перечитывания они те же."""
    база, запустить, tmp = архив
    было = {r["uid"]: r["id"] for r in строки_базы(база, "SELECT uid, id FROM messages")}

    запустить("скачать", "--заново", "--attach-dir", str(tmp / "вложения"))

    стало = {r["uid"]: r["id"] for r in строки_базы(база, "SELECT uid, id FROM messages")}
    assert стало == было


def test_заново_не_плодит_копии_вложений(архив, monkeypatch):
    база, запустить, tmp = архив
    запустить("скачать", "--заново", "--attach-dir", str(tmp / "вложения"))

    файлы = [f.name for f in (tmp / "вложения").rglob("*.xlsx")]
    assert файлы == ["Форма КП.xlsx"], "файл на диске уже есть, второй раз не нужен"
    r = строки_базы(база, "SELECT attachments FROM messages WHERE uid = 12")[0]
    assert "Форма КП.xlsx" in r["attachments"], "ссылка на вложение не должна потеряться"


def test_смена_uidvalidity_пересобирает_папку(архив, monkeypatch):
    """Сервер перенумеровал письма — старые UID указывают уже не туда."""
    база, запустить, tmp = архив
    monkeypatch.setitem(ЗаглушкаIMAP.uidvalidity, "INBOX", 999)

    запустить("скачать", "--attach-dir", str(tmp / "вложения"))

    всего = строки_базы(база, "SELECT COUNT(*) AS n FROM messages")[0]["n"]
    assert всего == 4, "папка должна быть перечитана без задвоения"


def test_докачка_ничего_не_меняет_в_ящике(архив):
    """Главная проверка: только чтение и никаких пометок о прочтении."""
    imap = ЗаглушкаIMAP.последний

    assert imap.readonly is True, "ящик обязан открываться в readonly"
    assert imap.вышел, "сессия должна закрываться"
    выборки = [args for команда, args in imap.команды if команда == "FETCH"]
    тела = [args for args in выборки if "BODY" in args[1]]
    assert тела, "письма должны были скачаться"
    for args in тела:
        assert "BODY.PEEK[]" in args[1], "без PEEK письмо пометится прочитанным"


def test_вложение_сохраняется_на_диск(архив):
    база, _, tmp = архив
    файлы = list((tmp / "вложения").rglob("*.xlsx"))
    assert [f.name for f in файлы] == ["Форма КП.xlsx"]
    assert файлы[0].parent.name == "12", "вложения лежат в папке с UID письма"

    r = строки_базы(база, "SELECT attachments FROM messages WHERE uid = 12")[0]
    assert "Форма КП.xlsx" in r["attachments"]


def test_флаг_без_вложений(tmp_path, monkeypatch):
    monkeypatch.setattr(mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(
        mail, "load_env",
        lambda path: {"MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": "пароль"},
    )
    вложения = tmp_path / "вложения"
    mail.main(["--база", str(tmp_path / "а.db"), "скачать",
               "--без-вложений", "--attach-dir", str(вложения)])

    assert not вложения.exists()


def test_только_входящие(tmp_path, monkeypatch):
    monkeypatch.setattr(mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(
        mail, "load_env",
        lambda path: {"MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": "пароль"},
    )
    база = tmp_path / "а.db"
    mail.main(["--база", str(база), "скачать", "--только-входящие",
               "--attach-dir", str(tmp_path / "в")])

    ящики = строки_базы(база, "SELECT label FROM mailboxes")
    assert [r["label"] for r in ящики] == ["Входящие"]


# --- разбор ответов сервера ----------------------------------------------


def test_папка_отправленных_по_флагу_сервера():
    папки = mail.разобрать_папки([
        br'(\HasNoChildren) "/" "INBOX"',
        br'(\HasNoChildren \Sent) "/" "[Gmail]/Sent Mail"',
    ])
    assert папки == [("INBOX", "Входящие", "входящее"),
                     ("[Gmail]/Sent Mail", "Отправленные", "отправленное")]


def test_папка_отправленных_по_имени_если_флага_нет():
    папки = mail.разобрать_папки([
        br'(\HasNoChildren) "." "INBOX"',
        br'(\HasNoChildren) "." "INBOX.Sent"',
    ])
    assert папки[1][0] == "INBOX.Sent"


def test_без_папки_отправленных_остаются_входящие():
    assert mail.разобрать_папки([br'(\HasNoChildren) "/" "INBOX"']) == [
        ("INBOX", "Входящие", "входящее")
    ]


def test_имя_папки_берётся_в_кавычки():
    assert mail.в_кавычках("[Gmail]/Sent Mail") == '"[Gmail]/Sent Mail"'
    assert mail.в_кавычках('"INBOX"') == '"INBOX"'


def test_uid_выуживается_из_ответа():
    ответ = [(b"1 (UID 42 BODY[] {5}", b"hello"), b")"]
    assert mail.разобрать_письма(ответ) == [(42, b"hello")]


def test_непрочитанное_видно_по_флагам():
    assert mail.разобрать_флаги([b"1 (UID 7 FLAGS (\\Seen))"]) == {7: "\\Seen"}
    assert mail.разобрать_флаги([b"1 (UID 8 FLAGS ())"]) == {8: ""}


# --- поиск ----------------------------------------------------------------


def test_поиск_находит_другое_окончание(архив, capsys):
    """«Тендеры» по запросу «тендер» ловятся приставкой «*» к слову."""
    _, запустить, _ = архив
    запустить("поиск", "закупк")

    assert "Извещение о закупке" in capsys.readouterr().out


def test_поиск_сам_отрезает_окончание(архив, capsys):
    """Человек пишет «заявка», в письме «заявок» — падеж не должен мешать."""
    _, запустить, _ = архив
    запустить("поиск", "заявка")

    вывод = capsys.readouterr().out
    assert "Извещение о закупке" in вывод
    assert "искал по основе" in вывод, "подмену слова надо показать, а не делать молча"


def test_точный_поиск_не_подбирает_основу(архив, capsys):
    _, запустить, _ = архив
    запустить("поиск", "заявка", "--точно")

    assert "ничего нет" in capsys.readouterr().out


def test_поиск_по_тексту_письма(архив, capsys):
    _, запустить, _ = архив
    запустить("поиск", "уборк")

    вывод = capsys.readouterr().out
    assert "Извещение о закупке" in вывод


def test_поиск_показывает_кусок_текста(архив, capsys):
    _, запустить, _ = архив
    запустить("поиск", "подачи")

    вывод = capsys.readouterr().out
    assert "25.09.2026" in вывод, "рядом с находкой должен быть кусок текста"


def test_поиск_ничего_не_нашёл_говорит_прямо(архив, capsys):
    _, запустить, _ = архив
    запустить("поиск", "экскаватор")

    вывод = capsys.readouterr().out
    assert "ничего нет" in вывод
    assert "скачать" in вывод, "стоит подсказать, что архив можно углубить"


def test_поиск_не_ходит_в_сеть(архив, monkeypatch, capsys):
    """Запросы к архиву обязаны отвечать из базы, иначе экономии нет."""
    _, запустить, _ = архив

    def взорваться(*a, **kw):
        raise AssertionError("запрос к архиву полез в сеть")

    monkeypatch.setattr(mail.imaplib, "IMAP4_SSL", взорваться)
    запустить("поиск", "форма")
    запустить("сводка", "--дней", "30")
    запустить("статистика")

    assert "Форма для заполнения" in capsys.readouterr().out


def test_поиск_без_слов_это_список_по_фильтрам(архив, capsys):
    """«Что приходило от Анны» — вопрос без ключевого слова."""
    _, запустить, _ = архив
    запустить("поиск", "--от", "anna@zakazchik.ru")

    вывод = capsys.readouterr().out
    assert "Форма для заполнения" in вывод
    assert "Скидки сентября" not in вывод


def test_поиск_только_по_отправленным(архив, capsys):
    """Нужно для черновиков: посмотреть, как он сам обычно пишет."""
    _, запустить, _ = архив
    запустить("поиск", "--отправленные", "--дней", "30")

    вывод = capsys.readouterr().out
    assert "Re: Форма для заполнения" in вывод
    assert "Извещение о закупке" not in вывод


def test_фильтр_по_отправителю(архив, capsys):
    _, запустить, _ = архив
    запустить("сводка", "--дней", "30", "--от", "spam-shop.ru", "--всё")

    вывод = capsys.readouterr().out
    assert "Скидки сентября" in вывод
    assert "Извещение" not in вывод


@pytest.mark.parametrize(
    "слова, точно, любое, ожидание",
    [
        (["тендер"], False, False, '"тендер"*'),
        (["тендер"], True, False, '"тендер"'),
        (["счёт", "оплата"], False, False, '"счёт"* "оплата"*'),
        (["счёт", "оплата"], False, True, '"счёт"* OR "оплата"*'),
        (['кавычки"внутри'], True, False, '"кавычки""внутри"'),
    ],
)
def test_запрос_к_поиску_собирается(слова, точно, любое, ожидание):
    assert mail.запрос_fts(слова, точно, любое) == ожидание


def test_спецсимвол_в_запросе_не_ломает_поиск(архив, capsys):
    _, запустить, _ = архив
    запустить("поиск", "zakupki@zakazchik.ru")

    assert "Извещение о закупке" in capsys.readouterr().out


# --- сводка и чтение письма ----------------------------------------------


def test_сводка_одна_строка_на_письмо(архив, capsys):
    _, запустить, _ = архив
    запустить("сводка", "--дней", "30", "--всё")

    строки = [с for с in capsys.readouterr().out.splitlines() if с.startswith("#")
              and not с.startswith("##")]
    assert len(строки) == 4
    assert all(" | " in с for с in строки)


def test_сводка_помечает_вложения(архив, capsys):
    _, запустить, _ = архив
    запустить("сводка", "--дней", "30", "--всё")

    строка = next(с for с in capsys.readouterr().out.splitlines() if "Форма для заполнения" in с)
    assert "📎1" in строка


def test_сводка_отбрасывает_мусор(архив, capsys, tmp_path):
    _, запустить, _ = архив
    игнор = tmp_path / "игнор.txt"
    игнор.write_text("spam-shop.ru\n", encoding="utf-8")

    запустить("сводка", "--дней", "30", "--ignore-file", str(игнор))

    вывод = capsys.readouterr().out
    assert "Скидки сентября" not in вывод
    assert "Отброшено по фильтру мусора: 1" in вывод


def test_сводка_за_короткий_период_пуста(архив, capsys):
    """Письма в заглушке сентябрьские: за последние сутки их быть не должно."""
    _, запустить, _ = архив
    запустить("сводка", "--дней", "1")

    assert "нет" in capsys.readouterr().out


def test_письмо_целиком(архив, capsys):
    база, запустить, _ = архив
    номер = строки_базы(база, "SELECT id FROM messages WHERE uid = 11")[0]["id"]

    запустить("письмо", str(номер))

    вывод = capsys.readouterr().out
    assert "Срок подачи заявок — 25.09.2026 до 10:00." in вывод
    assert "zakupki@zakazchik.ru" in вывод


def test_письмо_показывает_путь_к_вложению(архив, capsys):
    база, запустить, _ = архив
    номер = строки_базы(база, "SELECT id FROM messages WHERE uid = 12")[0]["id"]

    запустить("письмо", str(номер))

    assert "Форма КП.xlsx" in capsys.readouterr().out


def test_нет_такого_письма(архив, capsys):
    _, запустить, _ = архив
    запустить("письмо", "99999")

    assert "в архиве нет" in capsys.readouterr().out


def test_статистика_показывает_охват(архив, capsys):
    _, запустить, _ = архив
    запустить("статистика")

    вывод = capsys.readouterr().out
    assert "писем 4" in вывод
    assert "17.09.2026" in вывод
    assert "Входящие" in вывод and "Отправленные" in вывод


# --- мелочи ---------------------------------------------------------------


def test_кривая_дата_не_теряет_письмо(tmp_path, monkeypatch):
    """Письмо с нечитаемой датой должно попасть в архив, а не пропасть."""
    monkeypatch.setattr(mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(
        mail, "load_env",
        lambda path: {"MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": "пароль"},
    )
    битое = письмо("Без даты", "kto@example.ru", "Текст", дата="позавчера")
    monkeypatch.setitem(ЗаглушкаIMAP.папки, "INBOX", {21: битое})

    база = tmp_path / "а.db"
    mail.main(["--база", str(база), "скачать", "--attach-dir", str(tmp_path / "в")])

    r = строки_базы(база, "SELECT subject, ts FROM messages")[0]
    assert r["subject"] == "Без даты"
    assert r["ts"] > 0


def test_латинские_псевдонимы_команд(архив, capsys):
    """Русское слово не выживает в .bat — оттуда команды зовутся по-латински."""
    _, запустить, _ = архив
    запустить("stats")
    запустить("summary", "--days", "30", "--all", "--limit", "2")

    вывод = capsys.readouterr().out
    assert "Архив: писем 4" in вывод
    assert "показано 2 из 4" in вывод
    assert "Входящие (3)" in вывод, "в заголовке папки — все письма, а не только показанные"


def test_обрезка_длинного_текста():
    assert mail.кратко("коротко", 20) == "коротко"
    assert mail.кратко("а" * 30, 10) == "а" * 9 + "…"
