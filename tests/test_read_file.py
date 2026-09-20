"""
Тесты чтения вложений. Файлы .docx и .xlsx собираются здесь же вручную —
это обычные zip-архивы с XML внутри, поэтому внешние пакеты не нужны.
"""

import zipfile

import pytest

import read_file

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def сделать_docx(path, body_xml):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(
            "word/document.xml",
            f'<?xml version="1.0"?><w:document xmlns:w="{W_NS}">'
            f"<w:body>{body_xml}</w:body></w:document>",
        )
    return path


def абзац(текст):
    return f"<w:p><w:r><w:t>{текст}</w:t></w:r></w:p>"


def сделать_xlsx(path, sheet_xml, shared=(), имя_листа="Лист1"):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(
            "xl/workbook.xml",
            f'<?xml version="1.0"?><workbook xmlns="{S_NS}" xmlns:r="{R_NS}">'
            f'<sheets><sheet name="{имя_листа}" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_NS}">'
            '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        if shared:
            элементы = "".join(f"<si><t>{s}</t></si>" for s in shared)
            z.writestr(
                "xl/sharedStrings.xml",
                f'<?xml version="1.0"?><sst xmlns="{S_NS}">{элементы}</sst>',
            )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            f'<?xml version="1.0"?><worksheet xmlns="{S_NS}">'
            f"<sheetData>{sheet_xml}</sheetData></worksheet>",
        )
    return path


# --- Word ------------------------------------------------------------------


def test_docx_читается(tmp_path):
    путь = сделать_docx(
        tmp_path / "Договор.docx",
        абзац("Договор подряда №17") + абзац("Срок выполнения — 30 дней."),
    )

    текст = read_file.read_any(путь)

    assert "Договор подряда №17" in текст
    assert "Срок выполнения — 30 дней." in текст


def test_docx_абзац_из_нескольких_кусков(tmp_path):
    """Word режет строку на куски при правке — склеиваться должно без пробелов."""
    путь = сделать_docx(
        tmp_path / "а.docx",
        "<w:p><w:r><w:t>Уведомление </w:t></w:r>"
        "<w:r><w:t>участнику </w:t></w:r>"
        "<w:r><w:t>тендера</w:t></w:r></w:p>",
    )

    assert read_file.read_any(путь) == "Уведомление участнику тендера"


def test_docx_таблица_превращается_в_строки(tmp_path):
    таблица = (
        "<w:tbl>"
        "<w:tr><w:tc>" + абзац("Наименование") + "</w:tc>"
        "<w:tc>" + абзац("Цена") + "</w:tc></w:tr>"
        "<w:tr><w:tc>" + абзац("Труба 50мм") + "</w:tc>"
        "<w:tc>" + абзац("1500") + "</w:tc></w:tr>"
        "</w:tbl>"
    )
    путь = сделать_docx(tmp_path / "смета.docx", таблица)

    текст = read_file.read_any(путь)

    assert "Наименование | Цена" in текст
    assert "Труба 50мм | 1500" in текст


def test_битый_docx_понятно_ругается(tmp_path):
    путь = tmp_path / "ломаный.docx"
    путь.write_bytes(b"not a zip at all")

    with pytest.raises(SystemExit) as e:
        read_file.read_any(путь)
    assert "повреждён" in str(e.value)


# --- Excel -----------------------------------------------------------------


def test_xlsx_читается(tmp_path):
    путь = сделать_xlsx(
        tmp_path / "прайс.xlsx",
        '<row r="1"><c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" t="inlineStr"><is><t>Цена</t></is></c></row>'
        '<row r="2"><c r="A2" t="s"><v>1</v></c>'
        '<c r="B2"><v>1500</v></c></row>',
        shared=("Наименование", "Труба 50мм"),
    )

    текст = read_file.read_any(путь)

    assert "Наименование | Цена" in текст   # общие и встроенные строки
    assert "Труба 50мм | 1500" in текст     # строка и число


def test_xlsx_пустые_строки_пропускаются(tmp_path):
    путь = сделать_xlsx(
        tmp_path / "дыры.xlsx",
        '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
        '<row r="2"><c r="A2"/></row>'
        '<row r="3"><c r="A3" t="s"><v>1</v></c></row>',
        shared=("Первая", "Третья"),
    )

    строки = read_file.read_any(путь).splitlines()

    assert строки == ["Первая", "Третья"]


# --- CSV и текст -----------------------------------------------------------


def test_csv_с_точкой_с_запятой(tmp_path):
    путь = tmp_path / "реестр.csv"
    путь.write_text("Дата;Сумма;Назначение\n19.09.2026;45200;Оплата счёта\n",
                    encoding="utf-8")

    текст = read_file.read_any(путь)

    assert "Дата | Сумма | Назначение" in текст
    assert "19.09.2026 | 45200 | Оплата счёта" in текст


def test_csv_в_windows_1251(tmp_path):
    """Выгрузки из 1С и банков часто приходят в cp1251."""
    путь = tmp_path / "банк.csv"
    путь.write_bytes("Контрагент;Сумма\nООО Ромашка;12000\n".encode("cp1251"))

    текст = read_file.read_any(путь)

    assert "ООО Ромашка | 12000" in текст


def test_txt_в_windows_1251(tmp_path):
    путь = tmp_path / "заметка.txt"
    путь.write_bytes("Перезвонить Иванову".encode("cp1251"))

    assert "Перезвонить Иванову" in read_file.read_any(путь)


# --- Чего не умеем ---------------------------------------------------------


@pytest.mark.parametrize("имя, подсказка", [
    ("файл.doc", ".docx"),
    ("файл.xls", ".xlsx"),
    ("файл.pdf", "вручную"),
])
def test_неподдерживаемые_форматы_объясняют_что_делать(tmp_path, имя, подсказка):
    путь = tmp_path / имя
    путь.write_bytes(b"whatever")

    with pytest.raises(SystemExit) as e:
        read_file.read_any(путь)
    assert подсказка in str(e.value)
