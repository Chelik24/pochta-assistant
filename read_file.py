"""
Достаёт текст из вложений, чтобы ассистент мог их прочитать.

Поддерживает .docx, .xlsx, .csv, .txt и другие текстовые файлы.
Работает на одной стандартной библиотеке — ставить ничего не нужно.

Запуск:
    run.cmd read_file.py "Почта/вложения/2026-09-20_20-53/01/Договор.docx"
    run.cmd read_file.py "Шаблоны/прайс.xlsx" --max-chars 5000
"""

import argparse
import csv
import io
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

for поток in (sys.stdout, sys.stderr):
    if hasattr(поток, "reconfigure"):
        поток.reconfigure(encoding="utf-8", errors="replace")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_R = "{http://schemas.openxmlformats.org/package/2006/relationships}"

# Кодировки по порядку: сначала самые вероятные для русских файлов
ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "cp866")


# --- Word ------------------------------------------------------------------


def _para_text(p) -> str:
    """Текст одного абзаца Word со всеми вложенными кусками."""
    parts = []
    for node in p.iter():
        if node.tag == f"{W}t":
            parts.append(node.text or "")
        elif node.tag == f"{W}tab":
            parts.append("\t")
        elif node.tag in (f"{W}br", f"{W}cr"):
            parts.append("\n")
    return "".join(parts).strip()


def docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(f"{W}body")
    if body is None:
        return ""

    lines = []
    for child in body:
        if child.tag == f"{W}p":
            lines.append(_para_text(child))
        elif child.tag == f"{W}tbl":
            # Таблицы в договорах и тендерных формах несут основную суть,
            # поэтому строки склеиваем через | — так видно колонки.
            for tr in child.iter(f"{W}tr"):
                cells = []
                for tc in tr.findall(f"{W}tc"):
                    текст = " ".join(_para_text(p) for p in tc.iter(f"{W}p"))
                    cells.append(" ".join(текст.split()))
                if any(cells):
                    lines.append(" | ".join(cells))
    return "\n".join(lines)


# --- Excel -----------------------------------------------------------------


def _shared_strings(z: zipfile.ZipFile) -> list:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(f"{S}t")) for si in root.iter(f"{S}si")]


def _sheet_files(z: zipfile.ZipFile) -> list:
    """Листы в том порядке, в каком они в книге, с их именами."""
    names = z.namelist()
    if "xl/workbook.xml" not in names:
        return []

    цели = {}
    if "xl/_rels/workbook.xml.rels" in names:
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        for rel in rels.iter(f"{PKG_R}Relationship"):
            цели[rel.get("Id")] = rel.get("Target", "").lstrip("/")

    листы = []
    book = ET.fromstring(z.read("xl/workbook.xml"))
    for sheet in book.iter(f"{S}sheet"):
        target = цели.get(sheet.get(f"{R}id"), "")
        if not target:
            continue
        path = target if target.startswith("xl/") else f"xl/{target}"
        if path in names:
            листы.append((sheet.get("name", "Лист"), path))
    return листы


def _cell_value(c, shared: list) -> str:
    тип = c.get("t")
    if тип == "inlineStr":
        return "".join(t.text or "" for t in c.iter(f"{S}t")).strip()
    v = c.find(f"{S}v")
    if v is None or v.text is None:
        # Файл, собранный программой, а не Excel: формула есть, а её значение
        # ещё не посчитано. Лучше показать саму формулу, чем пустую ячейку.
        f = c.find(f"{S}f")
        if f is not None and f.text:
            return f"={f.text}"
        return ""
    if тип == "s":
        try:
            return shared[int(v.text)].strip()
        except (ValueError, IndexError):
            return ""
    return v.text.strip()


def xlsx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        shared = _shared_strings(z)
        листы = _sheet_files(z)
        out = []
        for имя, файл in листы:
            root = ET.fromstring(z.read(файл))
            строки = []
            for row in root.iter(f"{S}row"):
                ячейки = [_cell_value(c, shared) for c in row.findall(f"{S}c")]
                while ячейки and not ячейки[-1]:
                    ячейки.pop()
                if any(ячейки):
                    строки.append(" | ".join(ячейки))
            if строки:
                if len(листы) > 1:
                    out.append(f"--- Лист: {имя} ---")
                out.extend(строки)
        return "\n".join(out)


# --- Текстовые файлы -------------------------------------------------------


def plain_text(path: Path) -> str:
    data = path.read_bytes()
    for enc in ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def csv_text(path: Path) -> str:
    сырое = plain_text(path)
    # Разделитель бывает и запятой, и точкой с запятой — выясняем по первой строке
    первая = сырое.split("\n", 1)[0]
    разделитель = ";" if первая.count(";") > первая.count(",") else ","
    строки = []
    for row in csv.reader(io.StringIO(сырое), delimiter=разделитель):
        if any(поле.strip() for поле in row):
            строки.append(" | ".join(поле.strip() for поле in row))
    return "\n".join(строки)


# --- Точка входа -----------------------------------------------------------

ЧИТАЛКИ = {
    ".docx": docx_text,
    ".xlsx": xlsx_text,
    ".xlsm": xlsx_text,
    ".csv": csv_text,
}
КАК_ТЕКСТ = {".txt", ".md", ".log", ".json", ".xml", ".htm", ".html", ".eml"}

НЕ_УМЕЕМ = {
    ".doc": "старый формат Word (.doc). Попросите прислать .docx или откройте в Word и пересохраните",
    ".xls": "старый формат Excel (.xls). Попросите прислать .xlsx",
    ".pdf": "PDF. Пока не поддерживается — откройте вручную",
    ".rtf": "RTF. Пока не поддерживается — откройте вручную",
}


def read_any(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in НЕ_УМЕЕМ:
        sys.exit(f"Не могу прочитать: {НЕ_УМЕЕМ[suffix]}.")
    читалка = ЧИТАЛКИ.get(suffix)
    if читалка:
        try:
            return читалка(path)
        except (zipfile.BadZipFile, ET.ParseError):
            sys.exit(
                f"Файл {path.name} повреждён или это не {suffix} на самом деле."
            )
    if suffix in КАК_ТЕКСТ or suffix == "":
        return plain_text(path)
    sys.exit(f"Не знаю, как читать файлы {suffix}.")


def main():
    parser = argparse.ArgumentParser(description="Текст из вложения для разбора ИИ")
    parser.add_argument("path", help="путь к файлу")
    parser.add_argument("--max-chars", type=int, default=20000,
                        help="макс. длина вывода (по умолчанию 20000)")
    args = parser.parse_args()

    path = Path(args.path)
    if not path.exists():
        sys.exit(f"Нет файла: {path}")

    текст = read_any(path).strip()
    if not текст:
        print(f"[{path.name}: файл пустой или в нём нет текста]")
        return

    if len(текст) > args.max_chars:
        текст = текст[: args.max_chars] + f"\n…[обрезано, всего {len(текст)} знаков]"
    print(текст)


if __name__ == "__main__":
    main()
