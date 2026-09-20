"""
Офлайн-тесты выгрузки писем. Реальный ящик не нужен: письма собираются
через email.message.EmailMessage, IMAP подменяется заглушкой.

Запуск:  python -m pytest tests -v
"""

import email
from email.header import Header
from email.message import EmailMessage
from email.policy import default

import pytest

import fetch_mail

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def roundtrip(msg: EmailMessage):
    """Письмо → байты → разбор. Так же, как оно приедет из IMAP."""
    return email.message_from_bytes(msg.as_bytes(), policy=default)


def make_msg(subject, sender, body, date="Fri, 19 Sep 2026 10:00:00 +0300"):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["Date"] = date
    msg.set_content(body)
    return msg


# --- текст письма ---------------------------------------------------------


def test_кириллица_не_ломается():
    msg = make_msg(
        "Запрос КП на монтаж вентиляции",
        "Иван Петров <ivan@example.ru>",
        "Здравствуйте!\nПрошу подготовить КП на монтаж приточной установки.\nС уважением, Иван.",
    )
    parsed = roundtrip(msg)

    assert parsed["subject"] == "Запрос КП на монтаж вентиляции"
    assert "Иван Петров" in parsed["from"]
    text = fetch_mail.get_text(parsed)
    assert "приточной установки" in text
    assert "?" not in text  # признак сломанной кодировки


HTML_ПИСЬМО = (
    "<html><head><style>p{color:red}</style></head><body>"
    "<p>Добрый день!</p><p>Цена&nbsp;&mdash; 1500&nbsp;руб.</p>"
    "<script>alert(1)</script>"
    "<ul><li>Труба</li><li>Фитинг</li></ul>"
    "</body></html>"
)


def test_html_превращается_в_текст():
    msg = EmailMessage()
    msg["Subject"] = "Прайс на сентябрь"
    msg["From"] = "ООО Поставщик <sales@example.ru>"
    msg.set_content(HTML_ПИСЬМО, subtype="html")
    parsed = roundtrip(msg)
    text = fetch_mail.get_text(parsed)

    assert "Добрый день!" in text
    assert "1500" in text
    assert "Труба" in text and "Фитинг" in text
    assert "<p>" not in text and "<li>" not in text  # теги вычищены
    assert "alert" not in text  # скрипт вырезан
    assert "color:red" not in text  # стиль вырезан
    assert "&nbsp;" not in text and "&mdash;" not in text  # мнемоники раскрыты


def test_если_есть_и_plain_и_html_берётся_plain():
    """Обычное письмо: plain-часть осмысленная — её и читаем, HTML не трогаем."""
    msg = EmailMessage()
    msg["Subject"] = "Прайс на сентябрь"
    msg["From"] = "sales@example.ru"
    msg.set_content("Добрый день! Направляю прайс на сентябрь, цены действуют до 30.09.")
    msg.add_alternative(HTML_ПИСЬМО, subtype="html")

    text = fetch_mail.get_text(roundtrip(msg))

    assert "цены действуют до 30.09" in text
    assert "Фитинг" not in text


def test_заглушка_в_plain_подменяется_html():
    """Рассылки кладут в plain отписку, а содержимое — только в HTML."""
    msg = EmailMessage()
    msg["Subject"] = "Прайс на сентябрь"
    msg["From"] = "sales@example.ru"
    msg.set_content("Письмо в HTML")
    msg.add_alternative(HTML_ПИСЬМО, subtype="html")

    text = fetch_mail.get_text(roundtrip(msg))

    assert "Труба" in text and "1500" in text


def test_windows_1251_читается_правильно():
    """Так письма шлёт старый Outlook: тело в cp1251, тема в RFC2047."""
    subject = Header("Счёт на оплату", "windows-1251").encode()
    body = "Добрый день! Счёт во вложении. Сумма 45 200 рублей."
    raw = (
        f"Subject: {subject}\r\n"
        "From: buh@example.ru\r\n"
        "Date: Fri, 19 Sep 2026 10:00:00 +0300\r\n"
        "MIME-Version: 1.0\r\n"
        "Content-Type: text/plain; charset=windows-1251\r\n"
        "Content-Transfer-Encoding: 8bit\r\n"
        "\r\n"
    ).encode("ascii") + body.encode("windows-1251")

    parsed = email.message_from_bytes(raw, policy=default)

    assert parsed["subject"] == "Счёт на оплату"
    assert fetch_mail.get_text(parsed) == body


def test_пустое_письмо_без_текста_и_вложений(tmp_path):
    msg = EmailMessage()
    msg["Subject"] = "(без темы)"
    msg["From"] = "noreply@example.ru"
    parsed = roundtrip(msg)

    assert fetch_mail.get_text(parsed) == ""
    assert fetch_mail.save_attachments(parsed, tmp_path / "01") == []
    assert not (tmp_path / "01").exists()  # пустых папок не плодим


# --- имена файлов ---------------------------------------------------------


@pytest.mark.parametrize(
    "исходное, ожидаемое",
    [
        ("Форма КП / 2026.xlsx", "Форма КП _ 2026.xlsx"),
        ('смета<1>:"черновик".xlsx', "смета_1___черновик_.xlsx"),
        (r"C:\Users\Ivan\смета.xlsx", "C__Users_Ivan_смета.xlsx"),  # разделители глушатся
        ("../../важное.xlsx", "_.._важное.xlsx"),      # выход из папки невозможен
        ("документ.  ", "документ"),                   # Windows режет хвостовые точки
        ("CON.txt", "_CON.txt"),                       # зарезервированное имя
        ("", "вложение"),
    ],
)
def test_имя_файла_чистится(исходное, ожидаемое):
    assert fetch_mail.safe_filename(исходное) == ожидаемое


def test_длинное_имя_обрезается_с_сохранением_расширения():
    имя = "о" * 300 + ".xlsx"
    результат = fetch_mail.safe_filename(имя)
    assert результат.endswith(".xlsx")
    assert len(результат) <= fetch_mail.MAX_NAME_LEN + len(".xlsx")


# --- сохранение вложений --------------------------------------------------


def test_вложение_с_запрещёнными_символами_сохраняется(tmp_path):
    msg = make_msg("КП", "zakaz@example.ru", "Форма во вложении.")
    msg.add_attachment(
        b"\x50\x4b\x03\x04test",
        maintype="application",
        subtype=XLSX.split("/")[1],
        filename="Форма КП / 2026.xlsx",
    )
    parsed = roundtrip(msg)

    saved = fetch_mail.save_attachments(parsed, tmp_path / "01")

    assert len(saved) == 1
    имя_из_письма, путь = saved[0]
    assert имя_из_письма == "Форма КП / 2026.xlsx"  # в отчёт идёт исходное имя
    assert путь.name == "Форма КП _ 2026.xlsx"
    assert путь.exists()
    assert путь.read_bytes() == b"\x50\x4b\x03\x04test"


def test_два_вложения_с_одним_именем_не_затирают_друг_друга(tmp_path):
    msg = make_msg("Сметы", "zakaz@example.ru", "Две сметы.")
    for содержимое in ("первая смета".encode(), "вторая смета".encode()):
        msg.add_attachment(
            содержимое,
            maintype="application",
            subtype=XLSX.split("/")[1],
            filename="смета.xlsx",
        )
    parsed = roundtrip(msg)

    saved = fetch_mail.save_attachments(parsed, tmp_path / "01")

    assert len(saved) == 2
    пути = [путь for _, путь in saved]
    assert пути[0].name == "смета.xlsx"
    assert пути[1].name == "смета_1.xlsx"
    assert пути[0].read_bytes() == "первая смета".encode()
    assert пути[1].read_bytes() == "вторая смета".encode()
    assert len(list((tmp_path / "01").iterdir())) == 2


def test_криптоподпись_не_считается_вложением(tmp_path):
    """Корпоративная почта цепляет smime.p7s почти к каждому письму — это не вложение."""
    msg = make_msg("Счёт", "buh@example.ru", "Счёт во вложении.")
    msg.add_attachment(
        "данные счёта".encode(),
        maintype="application",
        subtype=XLSX.split("/")[1],
        filename="счёт 17.xlsx",
    )
    msg.add_attachment(
        "подпись".encode(),
        maintype="application",
        subtype="pkcs7-signature",
        filename="smime.p7s",
    )
    parsed = roundtrip(msg)

    saved = fetch_mail.save_attachments(parsed, tmp_path / "01")

    assert [путь.name for _, путь in saved] == ["счёт 17.xlsx"]


def test_вложение_с_именем_в_rfc2047(tmp_path):
    """Русское имя, закодированное =?utf-8?B?...?= — должно раскодироваться."""
    msg = make_msg("Договор", "zakaz@example.ru", "Договор во вложении.")
    msg.add_attachment(
        b"pdf-bytes",
        maintype="application",
        subtype="pdf",
        filename=("utf-8", "", "Договор №17 от 19.09.2026.pdf"),
    )
    parsed = roundtrip(msg)

    saved = fetch_mail.save_attachments(parsed, tmp_path / "01")

    assert len(saved) == 1
    assert saved[0][1].name == "Договор №17 от 19.09.2026.pdf"


# --- полный прогон main() с заглушкой IMAP --------------------------------


class ЗаглушкаIMAP:
    """Вместо настоящего IMAP4_SSL. Запоминает, как её вызывали."""

    последний = None
    экземпляры = []

    def __init__(self, host, port):
        self.host = host
        self.port = port
        self.залогинился = False
        self.readonly = None
        self.команды_fetch = []
        self.вышел = False
        ЗаглушкаIMAP.последний = self
        ЗаглушкаIMAP.экземпляры.append(self)

    def login(self, user, password):
        self.залогинился = True
        return "OK", [b"Logged in"]

    def select(self, folder, readonly=False):
        self.folder = folder
        self.readonly = readonly
        return "OK", [str(len(ПИСЬМА)).encode()]

    def search(self, charset, *criteria):
        self.criteria = criteria
        ids = b" ".join(str(i).encode() for i in range(1, len(ПИСЬМА) + 1))
        return "OK", [ids]

    def fetch(self, msg_id, parts):
        self.команды_fetch.append((msg_id, parts))
        сырое = ПИСЬМА[int(msg_id) - 1]
        return "OK", [(b"%s (BODY[] {%d}" % (msg_id, len(сырое)), сырое), b")"]

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


def _собрать_письма():
    первое = make_msg(
        "Старое письмо про сроки",
        "Пётр Сидоров <petr@example.ru>",
        "Когда будет готово КП?",
        date="Wed, 17 Sep 2026 09:00:00 +0300",
    )

    второе = make_msg(
        "Форма для заполнения",
        "Отдел закупок <zakaz@example.ru>",
        "Заполните форму и верните.",
        date="Thu, 18 Sep 2026 12:00:00 +0300",
    )
    второе.add_attachment(
        b"xlsx-1",
        maintype="application",
        subtype=XLSX.split("/")[1],
        filename="Форма КП / 2026.xlsx",
    )
    второе.add_attachment(
        b"xlsx-2",
        maintype="application",
        subtype=XLSX.split("/")[1],
        filename="Форма КП / 2026.xlsx",
    )

    третье = EmailMessage()
    третье["Subject"] = "Самое свежее письмо"
    третье["From"] = "Анна <anna@example.ru>"
    третье["Date"] = "Fri, 19 Sep 2026 18:30:00 +0300"
    третье.set_content("<p>Здравствуйте!</p><p>Цена 1500&nbsp;руб.</p>", subtype="html")

    return [m.as_bytes() for m in (первое, второе, третье)]


ПИСЬМА = _собрать_письма()


@pytest.fixture
def прогон(tmp_path, monkeypatch):
    """Запускает main() на заглушке. Реальный .env и реальный IMAP не трогаются."""
    monkeypatch.setattr(fetch_mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(
        fetch_mail,
        "load_env",
        lambda path: {"MAIL_1_USER": "test@mail.ru", "MAIL_1_PASSWORD": "тестовый-пароль"},
    )
    выгрузки = tmp_path / "выгрузки"
    вложения = tmp_path / "вложения"
    monkeypatch.setattr(
        "sys.argv",
        [
            "fetch_mail.py", "--days", "3",
            "--out", str(выгрузки),
            "--attach-dir", str(вложения),
            "--ignore-file", str(tmp_path / "нет-такого-файла.txt"),
        ],
    )

    fetch_mail.main()

    файлы = list(выгрузки.glob("*.md"))
    assert len(файлы) == 1, "должен появиться ровно один файл выгрузки"
    return файлы[0], файлы[0].read_text(encoding="utf-8"), вложения


def test_main_создаёт_файл_выгрузки(прогон):
    файл, текст, _ = прогон

    assert файл.name.endswith("_почта.md")
    assert "# Почта: выгрузка" in текст
    assert "Писем: 3." in текст


def test_main_письма_от_новых_к_старым(прогон):
    _, текст, _ = прогон

    assert текст.index("Самое свежее письмо") < текст.index("Форма для заполнения")
    assert текст.index("Форма для заполнения") < текст.index("Старое письмо про сроки")
    assert "### 1. Самое свежее письмо" in текст
    assert "### 3. Старое письмо про сроки" in текст


def test_main_пути_к_вложениям_в_отчёте(прогон):
    _, текст, вложения = прогон

    # письмо с вложениями идёт вторым сверху → папка 02
    сохранённые = sorted(p.name for p in вложения.rglob("*.xlsx"))
    assert сохранённые == ["Форма КП _ 2026.xlsx", "Форма КП _ 2026_1.xlsx"]

    assert "- **Вложения:**" in текст
    assert "Форма КП / 2026.xlsx" in текст          # имя как в письме
    assert "Форма КП _ 2026.xlsx`" in текст          # путь на диске
    assert "Форма КП _ 2026_1.xlsx`" in текст
    assert "/test@mail.ru/02/" in текст              # разложено по ящику и номеру письма


def test_main_html_письмо_попало_текстом(прогон):
    _, текст, _ = прогон

    assert "Здравствуйте!" in текст
    assert "1500" in текст
    assert "<p>" not in текст


def test_main_ничего_не_меняет_в_ящике(прогон):
    """Главная проверка: ящик открыт на чтение и письма не помечаются прочитанными."""
    imap = ЗаглушкаIMAP.последний

    assert imap.залогинился
    assert imap.readonly is True, "ящик обязан открываться в readonly"
    assert imap.вышел, "сессия должна закрываться"
    assert len(imap.команды_fetch) == 3
    for _, parts in imap.команды_fetch:
        assert "BODY.PEEK[]" in parts, "без PEEK письмо пометится прочитанным"
        assert "BODY[]" not in parts.replace("BODY.PEEK[]", "")
    assert imap.criteria[0] == "SINCE"


# --- несколько ящиков и настройки ----------------------------------------


@pytest.mark.parametrize(
    "адрес, сервер",
    [
        ("denis@mail.ru", "imap.mail.ru"),
        ("denis@bk.ru", "imap.mail.ru"),
        ("denis@gmail.com", "imap.gmail.com"),
        ("denis@yandex.ru", "imap.yandex.ru"),
        ("denis@outlook.com", "outlook.office365.com"),
    ],
)
def test_сервер_определяется_по_адресу(адрес, сервер):
    assert fetch_mail.guess_host(адрес) == сервер


def test_незнакомый_домен_требует_явного_сервера():
    with pytest.raises(SystemExit) as e:
        fetch_mail.guess_host("denis@своя-фирма.рф")
    assert "MAIL_1_HOST" in str(e.value)


def test_старый_формат_env_с_одним_ящиком():
    accounts = fetch_mail.load_accounts(
        {"MAIL_USER": "denis@mail.ru", "MAIL_APP_PASSWORD": "пароль"}
    )
    assert len(accounts) == 1
    assert accounts[0]["host"] == "imap.mail.ru"


def test_несколько_ящиков_из_env():
    accounts = fetch_mail.load_accounts({
        "MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": "п1",
        "MAIL_2_USER": "denis@gmail.com", "MAIL_2_PASSWORD": "п2",
        "MAIL_3_USER": "denis@своя.рф", "MAIL_3_PASSWORD": "п3",
        "MAIL_3_HOST": "imap.своя.рф",
    })
    assert [a["user"] for a in accounts] == [
        "denis@mail.ru", "denis@gmail.com", "denis@своя.рф",
    ]
    assert accounts[1]["host"] == "imap.gmail.com"
    assert accounts[2]["host"] == "imap.своя.рф"


def test_ящик_без_пароля_ловится_сразу():
    with pytest.raises(SystemExit) as e:
        fetch_mail.load_accounts({"MAIL_1_USER": "denis@mail.ru", "MAIL_1_PASSWORD": ""})
    assert "denis@mail.ru" in str(e.value)


def test_пустой_env_понятно_сообщает():
    with pytest.raises(SystemExit) as e:
        fetch_mail.load_accounts({})
    assert "MAIL_1_USER" in str(e.value)


def test_фильтр_мусора(tmp_path):
    файл = tmp_path / "игнор.txt"
    файл.write_text(
        "# комментарий\n@rassylka.example.ru\nскидка\n\npromo@\n",
        encoding="utf-8",
    )
    rules = fetch_mail.load_ignore_rules(файл)

    assert fetch_mail.is_junk("news@rassylka.example.ru", "Новости", rules)
    assert fetch_mail.is_junk("shop@example.ru", "СКИДКА 50%", rules)   # регистр не важен
    assert fetch_mail.is_junk("promo@example.ru", "Акция", rules)
    assert not fetch_mail.is_junk("ivan@zakazchik.ru", "Смета на объект", rules)


def test_нет_файла_игнора_значит_нет_фильтра(tmp_path):
    assert fetch_mail.load_ignore_rules(tmp_path / "нет.txt") == []


def _прогнать(tmp_path, monkeypatch, env, доп_аргументы=()):
    ЗаглушкаIMAP.экземпляры = []
    monkeypatch.setattr(fetch_mail.imaplib, "IMAP4_SSL", ЗаглушкаIMAP)
    monkeypatch.setattr(fetch_mail, "load_env", lambda path: env)
    выгрузки = tmp_path / "выгрузки"
    monkeypatch.setattr("sys.argv", [
        "fetch_mail.py", "--days", "3",
        "--out", str(выгрузки),
        "--attach-dir", str(tmp_path / "вложения"),
        "--ignore-file", str(tmp_path / "игнор.txt"),
        *доп_аргументы,
    ])
    fetch_mail.main()
    return next(выгрузки.glob("*.md")).read_text(encoding="utf-8")


def test_два_ящика_в_одной_выгрузке(tmp_path, monkeypatch):
    текст = _прогнать(tmp_path, monkeypatch, {
        "MAIL_1_USER": "work@mail.ru", "MAIL_1_PASSWORD": "п1",
        "MAIL_2_USER": "work@gmail.com", "MAIL_2_PASSWORD": "п2",
    })

    assert "Ящиков: 2. Писем: 6." in текст
    assert "## Ящик: work@mail.ru — писем: 3" in текст
    assert "## Ящик: work@gmail.com — писем: 3" in текст
    серверы = sorted(i.host for i in ЗаглушкаIMAP.экземпляры)
    assert серверы == ["imap.gmail.com", "imap.mail.ru"]


def test_мусор_отбрасывается_но_остаётся_виден(tmp_path, monkeypatch):
    (tmp_path / "игнор.txt").write_text("самое свежее\n", encoding="utf-8")

    текст = _прогнать(tmp_path, monkeypatch, {
        "MAIL_1_USER": "work@mail.ru", "MAIL_1_PASSWORD": "п1",
    })

    assert "Отброшено как мусор: 1." in текст
    assert "## Отброшено как мусор" in текст
    assert "Самое свежее письмо" in текст          # тема видна в списке отброшенных
    assert "### 1. Форма для заполнения" in текст  # нумерация не сбилась


def test_флаг_all_отключает_фильтр(tmp_path, monkeypatch):
    (tmp_path / "игнор.txt").write_text("самое свежее\n", encoding="utf-8")

    текст = _прогнать(tmp_path, monkeypatch, {
        "MAIL_1_USER": "work@mail.ru", "MAIL_1_PASSWORD": "п1",
    }, доп_аргументы=["--all"])

    assert "Писем: 3." in текст
    assert "Отброшено" not in текст


def test_упавший_ящик_не_ломает_остальные(tmp_path, monkeypatch):
    class ЗаглушкаСОшибкой(ЗаглушкаIMAP):
        def login(self, user, password):
            if user.endswith("@gmail.com"):
                raise __import__("imaplib").IMAP4.error("AUTHENTICATIONFAILED")
            return super().login(user, password)

    ЗаглушкаIMAP.экземпляры = []
    monkeypatch.setattr(fetch_mail.imaplib, "IMAP4_SSL", ЗаглушкаСОшибкой)
    monkeypatch.setattr(fetch_mail, "load_env", lambda path: {
        "MAIL_1_USER": "work@mail.ru", "MAIL_1_PASSWORD": "п1",
        "MAIL_2_USER": "work@gmail.com", "MAIL_2_PASSWORD": "неверный",
    })
    выгрузки = tmp_path / "выгрузки"
    monkeypatch.setattr("sys.argv", [
        "fetch_mail.py", "--out", str(выгрузки),
        "--attach-dir", str(tmp_path / "вложения"),
        "--ignore-file", str(tmp_path / "нет.txt"),
    ])

    fetch_mail.main()  # не должно падать

    текст = next(выгрузки.glob("*.md")).read_text(encoding="utf-8")
    assert "⚠️ **work@gmail.com**" in текст
    assert "ПАРОЛЬ ДЛЯ ВНЕШНЕГО ПРИЛОЖЕНИЯ" in текст
    assert "## Ящик: work@mail.ru — писем: 3" in текст  # рабочий ящик выгрузился
