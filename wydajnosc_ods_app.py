# -*- coding: utf-8 -*-
"""Rejestr wydajności dopasowany do firmowego skoroszytu ODS.

Pliki wymagane w jednym folderze:
    wydajnosc_ods_app.py
    SZABLON_WYDAJNOSC.ods

Uruchomienie:
    python wydajnosc_ods_app.py

Program nie wymaga instalowania bibliotek zewnętrznych. Korzysta wyłącznie
z Pythona, tkintera oraz standardowego formatu ZIP/XML plików ODS.
Zawiera wybór zmiany I/II, osobne przyczyny przerw, ochronę formuł oraz
pamięć układu arkuszy przyspieszającą kolejne uruchomienia. Przy każdym
zapisie usuwa zapisane wyniki formuł, aby program arkuszowy przeliczył je
ponownie po otwarciu pliku.
"""

from __future__ import annotations

import copy
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Callable, Iterable


APP_DIR = Path(__file__).resolve().parent
TEMPLATE_FILENAME = "SZABLON_WYDAJNOSC.ods"
ASSIGNMENTS_FILENAME = "przydzialy_pracownikow.json"
LAYOUT_CACHE_FILENAME = "uklad_arkusza_cache.json"
LAYOUT_CACHE_VERSION = 3

POLSKIE_MIESIACE = {
    1: "STYCZEŃ",
    2: "LUTY",
    3: "MARZEC",
    4: "KWIECIEŃ",
    5: "MAJ",
    6: "CZERWIEC",
    7: "LIPIEC",
    8: "SIERPIEŃ",
    9: "WRZESIEŃ",
    10: "PAŹDZIERNIK",
    11: "LISTOPAD",
    12: "GRUDZIEŃ",
}

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
}

FORMULA_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:of:1.2"

OFFICE = f"{{{NS['office']}}}"
TABLE = f"{{{NS['table']}}}"
TEXT = f"{{{NS['text']}}}"
CALCEXT = f"{{{NS['calcext']}}}"

ROW_TAG = TABLE + "table-row"
CELL_TAG = TABLE + "table-cell"
COVERED_CELL_TAG = TABLE + "covered-table-cell"
ROW_REPEAT = TABLE + "number-rows-repeated"
COL_REPEAT = TABLE + "number-columns-repeated"

FORMULA_CACHE_ATTRIBUTES = (
    OFFICE + "value-type",
    OFFICE + "value",
    OFFICE + "date-value",
    OFFICE + "time-value",
    OFFICE + "boolean-value",
    OFFICE + "string-value",
    OFFICE + "currency",
    CALCEXT + "value-type",
)

for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)


class ExistingRecordError(RuntimeError):
    """Rekord pracownika z podaną datą już istnieje."""


@dataclass(frozen=True)
class CellView:
    text: str = ""
    value: str | None = None
    date_value: str | None = None
    formula: str | None = None


@dataclass(frozen=True)
class EmployeeBlock:
    start_row: int
    end_row: int
    detected_owner: str = ""


@dataclass(frozen=True)
class MachineLayout:
    sheet_name: str
    header_row: int
    date_col: int
    employee_col: int
    qty_col: int
    area_col: int
    setups_col: int
    work_hours_col: int
    break_cols: tuple[tuple[str, int], ...]
    shift_col: int | None
    reset_cols: tuple[int, ...]
    blocks: tuple[EmployeeBlock, ...]


@dataclass(frozen=True)
class SaveResult:
    path: Path
    sheet_name: str
    row: int
    quantity: int
    area_m2: float
    work_hours: float
    breaks_hours: float
    break_hours: tuple[tuple[str, float], ...]
    shift: str
    setups: int
    employee_block: int
    new_employee_assignment: bool


def qname(prefix: str, local: str) -> str:
    return f"{{{NS[prefix]}}}{local}"


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    value = re.sub(r"\s+", " ", value.strip())
    return value.casefold()


def normalize_header(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.upper().replace("²", "2").replace("³", "3")
    value = re.sub(r"\s+", " ", value.strip())
    return value


def display_number(value: float, decimals: int = 4) -> str:
    text = f"{value:.{decimals}f}".rstrip("0").rstrip(".")
    return text.replace(".", ",")


def parse_date(value: str) -> date:
    value = value.strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass
    raise ValueError("Data musi mieć format RRRR-MM-DD lub DD.MM.RRRR.")


def parse_clock(value: str) -> datetime:
    value = value.strip()
    for fmt in ("%H:%M", "%H.%M", "%H"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    raise ValueError(f"Nieprawidłowa godzina: {value}")


def split_interval(value: str) -> tuple[datetime, datetime]:
    parts = re.split(r"\s*[-–—]\s*", value.strip())
    if len(parts) != 2:
        raise ValueError("Przedział wpisz np. 8-16 albo 8:00-16:00.")
    start, end = map(parse_clock, parts)
    if end <= start:
        end += timedelta(days=1)
    return start, end


def interval_hours(value: str) -> float:
    start, end = split_interval(value)
    return (end - start).total_seconds() / 3600


def shift_from_interval(value: str) -> str:
    start, _ = split_interval(value)
    return "I" if (start.hour, start.minute) < (14, 0) else "II"


def break_item_hours(item: str) -> float:
    item = item.strip().lower().replace(" ", "")
    if not item:
        return 0.0

    if re.search(r"[-–—]", item):
        return interval_hours(item)

    match = re.fullmatch(r"(\d+(?:[.,]\d+)?)h", item)
    if match:
        return float(match.group(1).replace(",", "."))

    match = re.fullmatch(r"(\d+(?:[.,]\d+)?)(?:m|min)", item)
    if match:
        return float(match.group(1).replace(",", ".")) / 60

    match = re.fullmatch(r"(\d+):(\d{1,2})", item)
    if match:
        hours, minutes = map(int, match.groups())
        if minutes >= 60:
            raise ValueError(f"Nieprawidłowy czas przerwy: {item}")
        return hours + minutes / 60

    if re.fullmatch(r"\d+(?:[.,]\d+)?", item):
        return float(item.replace(",", ".")) / 60

    raise ValueError(
        f"Nieprawidłowa przerwa: {item}. Użyj np. 8:55-9:01, 15m, "
        "90min, 1,5h albo 1:30."
    )


def total_break_hours(text: str) -> float:
    # Zachowaj przecinek dziesiętny w zapisie typu 1,5h, ale pozwól
    # używać przecinków do oddzielania kolejnych przerw.
    normalized = re.sub(r"(?<=\d),(?=\d+\s*h\b)", ".", text, flags=re.IGNORECASE)
    items = [part for part in re.split(r"[,;\n]+", normalized) if part.strip()]
    return sum(break_item_hours(item) for item in items)


def monthly_path(work_day: date) -> Path:
    filename = f"WYDAJNOŚCI {POLSKIE_MIESIACE[work_day.month]} {work_day.year}.ods"
    return APP_DIR / filename


def is_valid_ods(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as archive:
            return (
                archive.read("mimetype").decode("ascii", errors="ignore").strip()
                == "application/vnd.oasis.opendocument.spreadsheet"
                and "content.xml" in archive.namelist()
            )
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def find_template() -> Path:
    preferred = APP_DIR / TEMPLATE_FILENAME
    if is_valid_ods(preferred):
        return preferred

    candidates = [
        path
        for path in APP_DIR.glob("*.ods")
        if not path.name.upper().startswith("WYDAJNOŚCI ") and is_valid_ods(path)
    ]
    if len(candidates) == 1:
        return candidates[0]
    if not candidates:
        raise FileNotFoundError(
            f"Brak pliku {TEMPLATE_FILENAME}. Umieść szablon ODS w tym samym folderze co program."
        )
    raise FileNotFoundError(
        f"Znaleziono kilka szablonów ODS. Zmień nazwę właściwego pliku na {TEMPLATE_FILENAME}."
    )


def register_document_namespaces(xml_bytes: bytes) -> None:
    try:
        for _event, item in ET.iterparse(io.BytesIO(xml_bytes), events=("start-ns",)):
            prefix, uri = item
            if prefix != "xml":
                ET.register_namespace(prefix or "", uri)
    except (ET.ParseError, ValueError):
        pass


def extract_root_namespace_declarations(xml_bytes: bytes) -> dict[str, str]:
    """Zachowuje deklaracje xmlns, także te używane tylko wewnątrz formuł.

    ElementTree usuwa przestrzenie nazw, które nie występują w nazwach tagów lub
    atrybutów. W ODS prefiks ``of`` występuje głównie w wartości atrybutu
    ``table:formula``. Bez ``xmlns:of`` Excel może uznać formuły za uszkodzone.
    """
    match = re.search(rb"<office:document-content\b[^>]*>", xml_bytes)
    if not match:
        return {}
    opening_tag = match.group(0)
    declarations: dict[str, str] = {}
    for item in re.finditer(
        rb"\s+xmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=([\"'])(.*?)\2",
        opening_tag,
    ):
        prefix = (item.group(1) or b"").decode("utf-8")
        declarations[prefix] = item.group(3).decode("utf-8")
    return declarations


def restore_root_namespace_declarations(
    xml_bytes: bytes, declarations: dict[str, str]
) -> bytes:
    """Dodaje do głównego tagu deklaracje xmlns utracone przy serializacji."""
    match = re.search(rb"<office:document-content\b[^>]*>", xml_bytes)
    if not match or not declarations:
        return xml_bytes

    opening_tag = match.group(0)
    additions: list[bytes] = []
    for prefix, uri in declarations.items():
        attribute = b"xmlns" + (b":" + prefix.encode("utf-8") if prefix else b"")
        if re.search(rb"\s+" + re.escape(attribute) + rb"\s*=", opening_tag):
            continue
        escaped_uri = (
            uri.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        additions.append(b" " + attribute + b'="' + escaped_uri.encode("utf-8") + b'"')

    if not additions:
        return xml_bytes
    repaired_tag = opening_tag[:-1] + b"".join(additions) + b">"
    return xml_bytes[: match.start()] + repaired_tag + xml_bytes[match.end() :]


def cell_text(cell: ET.Element) -> str:
    paragraphs = []
    for paragraph in cell.findall(".//" + TEXT + "p"):
        text = "".join(paragraph.itertext()).strip()
        if text:
            paragraphs.append(text)
    if paragraphs:
        return " | ".join(paragraphs)
    return "".join(cell.itertext()).strip()


def read_grid(table_element: ET.Element, max_rows: int = 180, max_cols: int = 40) -> dict[int, dict[int, CellView]]:
    grid: dict[int, dict[int, CellView]] = {}
    logical_row = 1

    for row_element in table_element.findall(ROW_TAG):
        row_repeat = int(row_element.get(ROW_REPEAT, "1"))
        if logical_row > max_rows:
            break

        row_values: dict[int, CellView] = {}
        logical_col = 1
        for cell in list(row_element):
            if cell.tag not in (CELL_TAG, COVERED_CELL_TAG):
                continue
            col_repeat = int(cell.get(COL_REPEAT, "1"))
            if logical_col <= max_cols:
                view = CellView(
                    text=cell_text(cell),
                    value=cell.get(OFFICE + "value") or cell.get(OFFICE + "string-value"),
                    date_value=cell.get(OFFICE + "date-value"),
                    formula=cell.get(TABLE + "formula"),
                )
                if view.text or view.value or view.date_value or view.formula:
                    for offset in range(min(col_repeat, max_cols - logical_col + 1)):
                        row_values[logical_col + offset] = view
            logical_col += col_repeat

        for offset in range(min(row_repeat, max_rows - logical_row + 1)):
            if row_values:
                grid[logical_row + offset] = dict(row_values)
        logical_row += row_repeat

    return grid


def is_employee_name(value: str) -> bool:
    normalized = normalize_header(value)
    if not normalized or normalized in {"0", "-", "XX", "XXX", "XXXX"}:
        return False
    if "RAZEM" in normalized or normalized == "SUMA":
        return False
    return any(ch.isalpha() for ch in value)


def detect_layouts(root: ET.Element) -> dict[str, MachineLayout]:
    spreadsheet = root.find("office:body/office:spreadsheet", NS)
    if spreadsheet is None:
        raise ValueError("Plik nie zawiera arkusza kalkulacyjnego.")

    layouts: dict[str, MachineLayout] = {}
    for table_element in spreadsheet.findall("table:table", NS):
        sheet_name = table_element.get(TABLE + "name", "")
        if "WZOR" in normalize_header(sheet_name):
            continue
        grid = read_grid(table_element)
        header_row = None
        headers: dict[int, str] = {}

        for row_number, row_cells in sorted(grid.items()):
            candidate = {col: normalize_header(view.text) for col, view in row_cells.items()}
            values = set(candidate.values())
            if {
                "DATA",
                "NAZWISKO",
                "GODZ. W PRACY",
                "INNE",
            }.issubset(values) and list(candidate.values()).count("SZT") >= 2:
                if "M2" in values:
                    header_row = row_number
                    headers = candidate
                    break

        if header_row is None:
            continue

        def first_col(label: str) -> int:
            return min(col for col, text in headers.items() if text == label)

        def last_col(label: str) -> int:
            return max(col for col, text in headers.items() if text == label)

        setups_col = next(
            (
                col
                for col, text in headers.items()
                if text.startswith("NASTAWY-") or text in {"NASTAWY", "NASTAWY-ILOSC"}
            ),
            None,
        )
        if setups_col is None:
            continue

        starts: list[int] = []
        for row_number in range(header_row + 1, 160):
            first = grid.get(row_number, {}).get(1)
            if first and first.text.strip() == "1" and not first.formula:
                starts.append(row_number)

        blocks: list[EmployeeBlock] = []
        employee_col = first_col("NAZWISKO")

        def is_next_row_formula(view: CellView, row_number: int) -> bool:
            """Rozpoznaje formułę numeracji nawet po usunięciu jej cache."""
            if not view.formula:
                return False
            compact = re.sub(r"\s+", "", view.formula).upper()
            match = re.fullmatch(
                r"(?:OF:)?=\[\.\$?A\$?(\d+)\]\+1(?:\.0+)?",
                compact,
            )
            return bool(match and int(match.group(1)) == row_number - 1)

        for start in starts:
            expected = 1
            end = start - 1
            for row_number in range(start, min(start + 36, 180)):
                first = grid.get(row_number, {}).get(1)
                if not first:
                    break
                try:
                    number = int(float(first.text.replace(",", ".")))
                except (ValueError, AttributeError):
                    if expected > 1 and is_next_row_formula(first, row_number):
                        number = expected
                    else:
                        break
                if number != expected:
                    break
                end = row_number
                expected += 1

            if end - start + 1 < 15:
                continue

            names = [
                grid.get(row_number, {}).get(employee_col, CellView()).text.strip()
                for row_number in range(start, end + 1)
            ]
            names = [name for name in names if is_employee_name(name)]
            owner = Counter(names).most_common(1)[0][0] if names else ""
            blocks.append(EmployeeBlock(start, end, owner))

        if not blocks:
            continue

        reset_headers = {
            "DATA",
            "SZT",
            "M2",
            "NASTAWY-ILOSC",
            "GODZ. W PRACY",
            "BRAK TOWARU",
            "AWARIA",
            "SERIA 0",
            "ZEBRANIE",
            "INNE",
            "KONSERW.",
            "KONSERW",
            "ZMIANA",
        }
        reset_cols = set()
        for col, header in headers.items():
            if header in reset_headers or "RAFAL" in header or "KOSTEK" in header:
                reset_cols.add(col)
        # Nie czyścimy pierwszych kolumn planu SZT/m2. Czyścimy wyłącznie kolumny wykonania.
        reset_cols.discard(first_col("SZT"))
        reset_cols.discard(first_col("M2"))
        reset_cols.add(first_col("DATA"))
        reset_cols.add(last_col("SZT"))
        reset_cols.add(last_col("M2"))
        reset_cols.add(setups_col)
        reset_cols.add(first_col("GODZ. W PRACY"))
        reset_cols.add(first_col("INNE"))
        reset_cols.add(employee_col)

        # W części arkuszy przed właściwą kolumną ZMIANA znajduje się dodatkowa
        # kolumna opisana nazwiskami kierowników (np. RAFAŁ / KOSTEK). Poprzednia
        # wersja omyłkowo wybierała tę pierwszą kolumnę. Najpierw szukamy więc
        # dokładnego nagłówka ZMIANA, a przy kilku kandydatach wybieramy kolumnę,
        # w której istniejące dane najczęściej mają wartości I lub II.
        exact_shift_cols = [col for col, text in headers.items() if text == "ZMIANA"]
        prefixed_shift_cols = [
            col for col, text in headers.items()
            if text.startswith("ZMIANA") and col not in exact_shift_cols
        ]
        shift_candidates = exact_shift_cols or prefixed_shift_cols

        def shift_column_score(col: int) -> tuple[int, int, int]:
            shift_values = 0
            nonempty_values = 0
            for block in blocks:
                for row_number in range(block.start_row, block.end_row + 1):
                    value = normalize_header(
                        grid.get(row_number, {}).get(col, CellView()).text
                    )
                    if value:
                        nonempty_values += 1
                    if value in {"I", "II"}:
                        shift_values += 1
            # Więcej poprawnych wartości I/II jest najważniejsze. Przy remisie
            # wybieramy kolumnę bardziej na prawo, bo w pliku właściwa ZMIANA
            # występuje za kolumną kierownika.
            return shift_values, -nonempty_values, col

        shift_col = max(shift_candidates, key=shift_column_score) if shift_candidates else None
        layouts[sheet_name] = MachineLayout(
            sheet_name=sheet_name,
            header_row=header_row,
            date_col=first_col("DATA"),
            employee_col=employee_col,
            qty_col=last_col("SZT"),
            area_col=last_col("M2"),
            setups_col=setups_col,
            work_hours_col=first_col("GODZ. W PRACY"),
            break_cols=tuple(
                (label, col)
                for label in ("BRAK TOWARU", "AWARIA", "SERIA 0", "ZEBRANIE", "INNE", "KONSERW.", "KONSERW")
                for col, header in headers.items()
                if header == label
            ),
            shift_col=shift_col,
            reset_cols=tuple(sorted(reset_cols)),
            blocks=tuple(blocks),
        )

    return layouts


def _template_fingerprint(template_path: Path) -> dict[str, int | str]:
    stat = template_path.stat()
    return {
        "name": template_path.name,
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def save_layout_cache(template_path: Path, layouts: dict[str, MachineLayout]) -> None:
    payload = {
        "version": LAYOUT_CACHE_VERSION,
        "template": _template_fingerprint(template_path),
        "layouts": {
            name: {
                "sheet_name": layout.sheet_name,
                "header_row": layout.header_row,
                "date_col": layout.date_col,
                "employee_col": layout.employee_col,
                "qty_col": layout.qty_col,
                "area_col": layout.area_col,
                "setups_col": layout.setups_col,
                "work_hours_col": layout.work_hours_col,
                "break_cols": [[label, col] for label, col in layout.break_cols],
                "shift_col": layout.shift_col,
                "reset_cols": list(layout.reset_cols),
                "blocks": [
                    {
                        "start_row": block.start_row,
                        "end_row": block.end_row,
                        "detected_owner": block.detected_owner,
                    }
                    for block in layout.blocks
                ],
            }
            for name, layout in layouts.items()
        },
    }
    path = APP_DIR / LAYOUT_CACHE_FILENAME
    temp = path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def load_layout_cache(template_path: Path) -> dict[str, MachineLayout] | None:
    path = APP_DIR / LAYOUT_CACHE_FILENAME
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        if payload.get("version") != LAYOUT_CACHE_VERSION:
            return None
        if payload.get("template") != _template_fingerprint(template_path):
            return None

        result: dict[str, MachineLayout] = {}
        for name, item in payload.get("layouts", {}).items():
            blocks = tuple(
                EmployeeBlock(
                    int(block["start_row"]),
                    int(block["end_row"]),
                    str(block.get("detected_owner", "")),
                )
                for block in item.get("blocks", [])
            )
            break_cols = tuple(
                (str(pair[0]), int(pair[1]))
                for pair in item.get("break_cols", [])
                if isinstance(pair, list) and len(pair) == 2
            )
            if not blocks or not break_cols:
                return None
            result[name] = MachineLayout(
                sheet_name=str(item["sheet_name"]),
                header_row=int(item["header_row"]),
                date_col=int(item["date_col"]),
                employee_col=int(item["employee_col"]),
                qty_col=int(item["qty_col"]),
                area_col=int(item["area_col"]),
                setups_col=int(item["setups_col"]),
                work_hours_col=int(item["work_hours_col"]),
                break_cols=break_cols,
                shift_col=(None if item.get("shift_col") is None else int(item["shift_col"])),
                reset_cols=tuple(int(value) for value in item.get("reset_cols", [])),
                blocks=blocks,
            )
        return result or None
    except (OSError, ValueError, TypeError, KeyError, json.JSONDecodeError):
        return None


def load_layouts_fast(template_path: Path) -> dict[str, MachineLayout]:
    cached = load_layout_cache(template_path)
    if cached:
        return cached
    document = OdsDocument(template_path)
    save_layout_cache(template_path, document.layouts)
    return document.layouts


class OdsDocument:
    def __init__(self, path: Path):
        self.path = Path(path)
        with zipfile.ZipFile(self.path, "r") as archive:
            xml_bytes = archive.read("content.xml")
        self.namespace_declarations = extract_root_namespace_declarations(xml_bytes)
        # Naprawia również pliki utworzone przez wcześniejszą, wadliwą wersję.
        self.namespace_declarations.setdefault("of", FORMULA_NAMESPACE)
        register_document_namespaces(xml_bytes)
        self.root = ET.fromstring(xml_bytes)
        self.spreadsheet = self.root.find("office:body/office:spreadsheet", NS)
        if self.spreadsheet is None:
            raise ValueError("Nieprawidłowy plik ODS: brak arkusza kalkulacyjnego.")
        self.layouts = detect_layouts(self.root)

    def table(self, sheet_name: str) -> ET.Element:
        for table_element in self.spreadsheet.findall("table:table", NS):
            if table_element.get(TABLE + "name") == sheet_name:
                return table_element
        raise KeyError(f"Nie znaleziono arkusza: {sheet_name}")

    @staticmethod
    def _split_repeated_element(
        parent: ET.Element,
        element: ET.Element,
        element_index: int,
        repeat_attr: str,
        offset: int,
    ) -> ET.Element:
        repeat = int(element.get(repeat_attr, "1"))
        if repeat == 1:
            return element

        before_count = offset
        after_count = repeat - offset - 1
        parts: list[ET.Element] = []

        if before_count:
            before = copy.deepcopy(element)
            before.set(repeat_attr, str(before_count))
            parts.append(before)

        target = copy.deepcopy(element)
        target.attrib.pop(repeat_attr, None)
        parts.append(target)

        if after_count:
            after = copy.deepcopy(element)
            after.set(repeat_attr, str(after_count))
            parts.append(after)

        parent.remove(element)
        for part in reversed(parts):
            parent.insert(element_index, part)
        return target

    def row_element(self, table_element: ET.Element, row_number: int) -> ET.Element:
        logical_row = 1
        for element_index, element in enumerate(list(table_element)):
            if element.tag != ROW_TAG:
                continue
            repeat = int(element.get(ROW_REPEAT, "1"))
            if logical_row <= row_number < logical_row + repeat:
                return self._split_repeated_element(
                    table_element,
                    element,
                    element_index,
                    ROW_REPEAT,
                    row_number - logical_row,
                )
            logical_row += repeat
        raise IndexError(f"Brak wiersza {row_number}.")

    def cell_element(self, sheet_name: str, row_number: int, col_number: int) -> ET.Element:
        table_element = self.table(sheet_name)
        row = self.row_element(table_element, row_number)
        logical_col = 1
        children = list(row)

        for element_index, cell in enumerate(children):
            if cell.tag not in (CELL_TAG, COVERED_CELL_TAG):
                continue
            repeat = int(cell.get(COL_REPEAT, "1"))
            if logical_col <= col_number < logical_col + repeat:
                target = self._split_repeated_element(
                    row,
                    cell,
                    element_index,
                    COL_REPEAT,
                    col_number - logical_col,
                )
                if target.tag == COVERED_CELL_TAG:
                    raise ValueError(
                        f"Komórka {sheet_name}!R{row_number}C{col_number} jest częścią scalonego zakresu."
                    )
                return target
            logical_col += repeat

        while logical_col <= col_number:
            cell = ET.Element(CELL_TAG)
            row.append(cell)
            if logical_col == col_number:
                return cell
            logical_col += 1
        raise IndexError("Nie można utworzyć komórki.")

    def cell_view(self, sheet_name: str, row_number: int, col_number: int) -> CellView:
        cell = self.cell_element(sheet_name, row_number, col_number)
        return CellView(
            text=cell_text(cell),
            value=cell.get(OFFICE + "value") or cell.get(OFFICE + "string-value"),
            date_value=cell.get(OFFICE + "date-value"),
            formula=cell.get(TABLE + "formula"),
        )

    @staticmethod
    def _clear_value(cell: ET.Element) -> None:
        for attr in (
            OFFICE + "value-type",
            OFFICE + "value",
            OFFICE + "date-value",
            OFFICE + "time-value",
            OFFICE + "string-value",
            CALCEXT + "value-type",
            TABLE + "formula",
        ):
            cell.attrib.pop(attr, None)
        for child in list(cell):
            if child.tag == TEXT + "p":
                cell.remove(child)

    def set_blank(self, sheet_name: str, row_number: int, col_number: int) -> None:
        self._clear_value(self.cell_element(sheet_name, row_number, col_number))

    def set_string(self, sheet_name: str, row_number: int, col_number: int, value: str) -> None:
        cell = self.cell_element(sheet_name, row_number, col_number)
        self._clear_value(cell)
        cell.set(OFFICE + "value-type", "string")
        cell.set(CALCEXT + "value-type", "string")
        paragraph = ET.SubElement(cell, TEXT + "p")
        paragraph.text = value

    def set_number(self, sheet_name: str, row_number: int, col_number: int, value: float | int) -> None:
        cell = self.cell_element(sheet_name, row_number, col_number)
        self._clear_value(cell)
        cell.set(OFFICE + "value-type", "float")
        cell.set(CALCEXT + "value-type", "float")
        cell.set(OFFICE + "value", str(value))
        paragraph = ET.SubElement(cell, TEXT + "p")
        if isinstance(value, int) or float(value).is_integer():
            paragraph.text = str(int(value))
        else:
            paragraph.text = display_number(float(value))

    def set_date(self, sheet_name: str, row_number: int, col_number: int, value: date) -> None:
        cell = self.cell_element(sheet_name, row_number, col_number)
        self._clear_value(cell)
        cell.set(OFFICE + "value-type", "date")
        cell.set(CALCEXT + "value-type", "date")
        cell.set(OFFICE + "date-value", value.isoformat())
        paragraph = ET.SubElement(cell, TEXT + "p")
        paragraph.text = value.strftime("%-d.%-m.%Y") if os.name != "nt" else f"{value.day}.{value.month}.{value.year}"

    def clear_formula_cache(self) -> tuple[int, int, int]:
        """Usuwa zapisane wyniki formuł, nie zmieniając samych formuł."""
        formula_cells = 0
        attributes_removed = 0
        paragraphs_removed = 0

        for cell in self.root.iter(CELL_TAG):
            if not cell.get(TABLE + "formula"):
                continue
            formula_cells += 1
            for attr in FORMULA_CACHE_ATTRIBUTES:
                if attr in cell.attrib:
                    attributes_removed += 1
                    cell.attrib.pop(attr)
            for child in list(cell):
                if child.tag == TEXT + "p":
                    paragraphs_removed += 1
                    cell.remove(child)

        return formula_cells, attributes_removed, paragraphs_removed

    def save(self, destination: Path | None = None, make_backup: bool = True) -> Path:
        destination = Path(destination or self.path)
        destination.parent.mkdir(parents=True, exist_ok=True)

        if make_backup and destination.exists() and destination.resolve() == self.path.resolve():
            backup = destination.with_suffix(destination.suffix + ".bak")
            shutil.copy2(destination, backup)

        self.clear_formula_cache()
        xml_bytes = ET.tostring(self.root, encoding="utf-8", xml_declaration=True)
        xml_bytes = restore_root_namespace_declarations(
            xml_bytes, self.namespace_declarations
        )
        temp_fd, temp_name = tempfile.mkstemp(prefix=destination.stem + "_", suffix=".ods", dir=destination.parent)
        os.close(temp_fd)
        temp_path = Path(temp_name)

        try:
            with zipfile.ZipFile(self.path, "r") as source, zipfile.ZipFile(temp_path, "w") as target:
                mimetype = source.read("mimetype")
                target.writestr("mimetype", mimetype, compress_type=zipfile.ZIP_STORED)
                for info in source.infolist():
                    if info.filename in {"mimetype", "content.xml"}:
                        continue
                    target.writestr(info, source.read(info.filename))
                target.writestr("content.xml", xml_bytes, compress_type=zipfile.ZIP_DEFLATED)
            os.replace(temp_path, destination)
        finally:
            temp_path.unlink(missing_ok=True)

        self.path = destination
        return destination


def parse_cell_date(view: CellView, default_year: int) -> date | None:
    if view.date_value:
        try:
            return date.fromisoformat(view.date_value[:10])
        except ValueError:
            pass

    text = view.text.strip()
    if not text:
        return None
    for fmt in ("%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    for fmt in ("%d-%m", "%d.%m"):
        try:
            partial = datetime.strptime(text, fmt)
            return date(default_year, partial.month, partial.day)
        except ValueError:
            pass
    return None


def load_assignments(template: Path, layouts: dict[str, MachineLayout]) -> dict[str, list[str]]:
    path = APP_DIR / ASSIGNMENTS_FILENAME
    result = {
        sheet_name: [block.detected_owner for block in layout.blocks]
        for sheet_name, layout in layouts.items()
    }

    if path.exists():
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            for sheet_name, owners in data.items():
                if sheet_name not in result or not isinstance(owners, list):
                    continue
                for index, owner in enumerate(owners[: len(result[sheet_name])]):
                    if isinstance(owner, str) and owner.strip():
                        result[sheet_name][index] = owner.strip()
        except (OSError, json.JSONDecodeError):
            pass

    save_assignments(result)
    return result


def save_assignments(assignments: dict[str, list[str]]) -> None:
    path = APP_DIR / ASSIGNMENTS_FILENAME
    temp = path.with_suffix(".json.tmp")
    temp.write_text(json.dumps(assignments, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def initialize_month_file(template_path: Path, destination: Path) -> None:
    shutil.copy2(template_path, destination)
    document = OdsDocument(destination)
    for layout in document.layouts.values():
        for block in layout.blocks:
            for row_number in range(block.start_row, block.end_row + 1):
                for col_number in layout.reset_cols:
                    document.set_blank(layout.sheet_name, row_number, col_number)
    document.save(destination, make_backup=False)


def calculate_dimensions(dimensions: Iterable[tuple[float, float, int]]) -> tuple[int, float, int]:
    items = list(dimensions)
    total_qty = sum(qty for _, _, qty in items)
    total_area = sum(a * b * qty / 1_000_000 for a, b, qty in items)
    return total_qty, total_area, len(items)


def find_employee_block(
    document: OdsDocument,
    layout: MachineLayout,
    employee: str,
    assignments: dict[str, list[str]],
) -> tuple[int, bool]:
    normalized = normalize_text(employee)
    owners = assignments.setdefault(layout.sheet_name, [""] * len(layout.blocks))
    if len(owners) < len(layout.blocks):
        owners.extend([""] * (len(layout.blocks) - len(owners)))

    for index, owner in enumerate(owners):
        if owner and normalize_text(owner) == normalized:
            return index, False

    # Dodatkowe zabezpieczenie: odczyt aktualnych nazw bezpośrednio z arkusza.
    for index, block in enumerate(layout.blocks):
        for row_number in range(block.start_row, block.end_row + 1):
            current = document.cell_view(layout.sheet_name, row_number, layout.employee_col).text
            if current and normalize_text(current) == normalized:
                owners[index] = employee
                save_assignments(assignments)
                return index, False

    for index, owner in enumerate(owners):
        if not owner.strip():
            owners[index] = employee
            save_assignments(assignments)
            return index, True

    raise ValueError(
        f"W arkuszu {layout.sheet_name} nie ma wolnego bloku dla nowego pracownika {employee}. "
        f"Edytuj plik {ASSIGNMENTS_FILENAME} albo zwolnij nieużywany blok."
    )


def find_target_row(
    document: OdsDocument,
    layout: MachineLayout,
    block_index: int,
    employee: str,
    work_day: date,
) -> tuple[int, bool]:
    block = layout.blocks[block_index]
    first_free = None

    for row_number in range(block.start_row, block.end_row + 1):
        view = document.cell_view(layout.sheet_name, row_number, layout.date_col)
        existing_date = parse_cell_date(view, work_day.year)
        if existing_date == work_day:
            return row_number, True
        if existing_date is None and first_free is None:
            first_free = row_number

    if first_free is None:
        raise ValueError(
            f"Blok pracownika {employee} w arkuszu {layout.sheet_name} jest pełny."
        )
    return first_free, False


def repair_misplaced_shifts(
    document: OdsDocument,
    layout: MachineLayout,
) -> int:
    """Przenosi I/II zapisane przez wadliwą wersję do właściwej kolumny ZMIANA.

    Naprawa jest zachowawcza: działa tylko wtedy, gdy właściwa komórka zmiany
    jest pusta, a sąsiednia kolumna o nagłówku ZMIANA lub RAFAŁ/KOSTEK zawiera
    dokładnie I albo II.
    """
    if layout.shift_col is None:
        return 0

    table_element = document.table(layout.sheet_name)
    max_row = max(block.end_row for block in layout.blocks)
    grid = read_grid(table_element, max_rows=max_row, max_cols=40)
    headers = {
        col: normalize_header(view.text)
        for col, view in grid.get(layout.header_row, {}).items()
    }
    source_cols = [
        col
        for col, header in headers.items()
        if col != layout.shift_col
        and (header == "ZMIANA" or "RAFAL" in header or "KOSTEK" in header)
    ]
    source_cols.sort(key=lambda col: (abs(layout.shift_col - col), -col))

    repaired = 0
    for block in layout.blocks:
        for row_number in range(block.start_row, block.end_row + 1):
            target_value = normalize_header(
                document.cell_view(
                    layout.sheet_name, row_number, layout.shift_col
                ).text
            )
            if target_value in {"I", "II"}:
                continue

            for source_col in source_cols:
                source_value = normalize_header(
                    document.cell_view(
                        layout.sheet_name, row_number, source_col
                    ).text
                )
                if source_value not in {"I", "II"}:
                    continue
                document.set_string(
                    layout.sheet_name, row_number, layout.shift_col, source_value
                )
                document.set_blank(layout.sheet_name, row_number, source_col)
                repaired += 1
                break

    return repaired


def save_record(
    *,
    template_path: Path,
    machine: str,
    employee: str,
    dimensions: Iterable[tuple[float, float, int]],
    break_entries: dict[str, str],
    work_interval: str,
    shift: str,
    work_day: date,
    assignments: dict[str, list[str]],
    overwrite: bool = False,
) -> SaveResult:
    quantity, area_m2, setups = calculate_dimensions(dimensions)
    work_hours = interval_hours(work_interval)
    shift = shift.strip().upper()
    if shift not in {"I", "II"}:
        raise ValueError("Wybierz zmianę I albo II.")
    parsed_breaks = {label: total_break_hours(text) for label, text in break_entries.items()}
    breaks_hours = sum(parsed_breaks.values())

    output = monthly_path(work_day)
    if not output.exists():
        initialize_month_file(template_path, output)

    lock_file = output.parent / f".~lock.{output.name}#"
    if lock_file.exists():
        raise PermissionError(
            f"Plik {output.name} jest otwarty. Zamknij go w Excelu lub LibreOffice i spróbuj ponownie."
        )

    document = OdsDocument(output)
    if machine not in document.layouts:
        raise ValueError(f"Arkusz {machine} nie ma rozpoznanego układu wydajności.")
    layout = document.layouts[machine]
    # Automatycznie poprawia wpisy I/II pozostawione w złej kolumnie przez
    # wcześniejszą wersję programu. Zmiany zostaną utrwalone przy bieżącym zapisie.
    repair_misplaced_shifts(document, layout)

    block_index, new_assignment = find_employee_block(
        document, layout, employee, assignments
    )
    row_number, existing = find_target_row(
        document, layout, block_index, employee, work_day
    )
    if existing and not overwrite:
        raise ExistingRecordError(
            f"W arkuszu {machine} istnieje już wpis pracownika {employee} z dnia "
            f"{work_day.strftime('%d.%m.%Y')} (wiersz {row_number})."
        )

    document.set_date(machine, row_number, layout.date_col, work_day)
    document.set_string(machine, row_number, layout.employee_col, employee.strip())
    document.set_number(machine, row_number, layout.qty_col, quantity)
    document.set_number(machine, row_number, layout.area_col, round(area_m2, 4))
    document.set_number(machine, row_number, layout.setups_col, setups)
    document.set_number(machine, row_number, layout.work_hours_col, round(work_hours, 4))
    for label, col in layout.break_cols:
        canonical = "KONSERWACJA" if label.startswith("KONSERW") else label
        document.set_number(machine, row_number, col, round(parsed_breaks.get(canonical, 0.0), 4))
    if layout.shift_col is None:
        raise ValueError(
            f"W arkuszu {machine} nie znaleziono kolumny ZMIANA. "
            "Usuń plik uklad_arkusza_cache.json i uruchom program ponownie."
        )
    document.set_string(machine, row_number, layout.shift_col, shift)

    # Kontrola przed zapisem: wybrana zmiana musi znajdować się dokładnie w
    # wykrytej kolumnie ZMIANA. Dzięki temu błąd układu nie przechodzi bez komunikatu.
    written_shift = normalize_header(
        document.cell_view(machine, row_number, layout.shift_col).text
    )
    if written_shift != shift:
        raise ValueError(
            f"Nie udało się zapisać zmiany {shift} w arkuszu {machine}."
        )

    document.save(output, make_backup=True)
    return SaveResult(
        path=output,
        sheet_name=machine,
        row=row_number,
        quantity=quantity,
        area_m2=area_m2,
        work_hours=work_hours,
        breaks_hours=breaks_hours,
        break_hours=tuple(parsed_breaks.items()),
        shift=shift,
        setups=setups,
        employee_block=block_index + 1,
        new_employee_assignment=new_assignment,
    )


def open_folder(path: Path) -> None:
    folder = str(path.parent)
    if sys.platform.startswith("win"):
        os.startfile(folder)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", folder])
    else:
        subprocess.Popen(["xdg-open", folder])


class ProductivityApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Rejestr wydajności — firmowy arkusz ODS")
        self.geometry("900x760")
        self.minsize(790, 660)

        try:
            self.template_path = find_template()
            self.layouts = load_layouts_fast(self.template_path)
            if not self.layouts:
                raise ValueError("W szablonie nie znaleziono arkuszy maszyn z kolumnami wydajności.")
            self.assignments = load_assignments(self.template_path, self.layouts)
        except (OSError, ValueError, zipfile.BadZipFile) as exc:
            messagebox.showerror("Błąd szablonu", str(exc))
            self.destroy()
            return

        self.machine_names = list(self.layouts)
        self.machine_var = tk.StringVar(value=self.machine_names[0])
        self.employee_var = tk.StringVar()
        self.work_interval_var = tk.StringVar(value="8-16")
        self.shift_var = tk.StringVar(value="I")
        self.date_var = tk.StringVar(value=date.today().isoformat())
        self.break_vars = {
            label: tk.StringVar()
            for label in ("BRAK TOWARU", "AWARIA", "SERIA 0", "ZEBRANIE", "INNE", "KONSERWACJA")
        }
        self.status_var = tk.StringVar(value=f"Szablon: {self.template_path.name}")
        self.dimension_rows: list[tuple[ttk.Entry, ttk.Entry, ttk.Entry, ttk.Frame]] = []
        self.last_saved_path: Path | None = None

        self._build_ui()
        self.add_dimension_row()
        self.update_employee_values()

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=16)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(9, weight=1)

        ttk.Label(main, text="Maszyna:").grid(row=0, column=0, sticky="w", pady=5)
        self.machine_box = ttk.Combobox(
            main,
            textvariable=self.machine_var,
            values=self.machine_names,
            state="readonly",
        )
        self.machine_box.grid(row=0, column=1, sticky="ew", pady=5)
        self.machine_box.bind("<<ComboboxSelected>>", lambda _event: self.update_employee_values())

        ttk.Label(main, text="Pracownik:").grid(row=1, column=0, sticky="w", pady=5)
        self.employee_box = ttk.Combobox(main, textvariable=self.employee_var, state="normal")
        self.employee_box.grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Label(
            main,
            text="Wybierz istniejącego pracownika albo wpisz nowego. Nowy pracownik otrzyma pierwszy wolny blok w arkuszu.",
            wraplength=680,
        ).grid(row=2, column=1, sticky="w")

        ttk.Label(main, text="Data:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.date_var).grid(row=3, column=1, sticky="ew", pady=5)

        ttk.Label(main, text="Zmiana:").grid(row=4, column=0, sticky="w", pady=5)
        shift_frame = ttk.Frame(main)
        shift_frame.grid(row=4, column=1, sticky="w", pady=5)
        ttk.Radiobutton(shift_frame, text="I zmiana", variable=self.shift_var, value="I").pack(side="left")
        ttk.Radiobutton(shift_frame, text="II zmiana", variable=self.shift_var, value="II").pack(side="left", padx=18)

        ttk.Label(main, text="Przedział pracy:").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Entry(main, textvariable=self.work_interval_var).grid(row=5, column=1, sticky="ew", pady=5)
        ttk.Label(main, text="np. 8-16, 6:00-14:00 lub 14-22").grid(row=6, column=1, sticky="w")

        ttk.Separator(main).grid(row=7, column=0, columnspan=2, sticky="ew", pady=12)
        ttk.Label(main, text="Wymiary", font=("Segoe UI", 11, "bold")).grid(
            row=8, column=0, columnspan=2, sticky="w"
        )

        dimensions_outer = ttk.Frame(main)
        dimensions_outer.grid(row=9, column=0, columnspan=2, sticky="nsew", pady=6)
        dimensions_outer.rowconfigure(0, weight=1)
        dimensions_outer.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(dimensions_outer, highlightthickness=0, height=220)
        scrollbar = ttk.Scrollbar(dimensions_outer, orient="vertical", command=self.canvas.yview)
        self.dimensions_frame = ttk.Frame(self.canvas)
        self.dimensions_frame.bind(
            "<Configure>",
            lambda _event: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas_window = self.canvas.create_window((0, 0), window=self.dimensions_frame, anchor="nw")
        self.canvas.bind(
            "<Configure>",
            lambda event: self.canvas.itemconfigure(self.canvas_window, width=event.width),
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        header = ttk.Frame(self.dimensions_frame)
        header.pack(fill="x")
        for text, width in (("a [mm]", 18), ("b [mm]", 18), ("ilość", 18), ("", 9)):
            ttk.Label(header, text=text, width=width).pack(side="left", padx=3)

        ttk.Button(main, text="Dodaj kolejne", command=self.add_dimension_row).grid(
            row=10, column=0, columnspan=2, sticky="w", pady=6
        )

        ttk.Label(main, text="Przerwy według przyczyny", font=("Segoe UI", 11, "bold")).grid(
            row=11, column=0, columnspan=2, sticky="w", pady=(10, 4)
        )
        breaks_frame = ttk.Frame(main)
        breaks_frame.grid(row=12, column=0, columnspan=2, sticky="ew")
        breaks_frame.columnconfigure(1, weight=1)
        for idx, label in enumerate(self.break_vars):
            ttk.Label(breaks_frame, text=label.title() + ":").grid(row=idx, column=0, sticky="w", pady=2)
            ttk.Entry(breaks_frame, textvariable=self.break_vars[label]).grid(row=idx, column=1, sticky="ew", pady=2)
        ttk.Label(
            main,
            text="W każdym polu można podać przedział lub czas, np. 8:55-9:01, 15m, 90min, 1,5h. Kilka przerw oddziel przecinkiem lub średnikiem.",
            wraplength=780,
        ).grid(row=13, column=0, columnspan=2, sticky="w", pady=(4, 0))

        buttons = ttk.Frame(main)
        buttons.grid(row=14, column=0, columnspan=2, sticky="ew", pady=(18, 0))
        ttk.Button(buttons, text="Zapisz do arkusza", command=self.submit).pack(side="left")
        ttk.Button(buttons, text="Wyczyść formularz", command=self.clear_form).pack(side="left", padx=8)
        ttk.Button(buttons, text="Otwórz folder wynikowy", command=self.open_output_folder).pack(side="left")

        ttk.Label(main, textvariable=self.status_var, wraplength=830).grid(
            row=15, column=0, columnspan=2, sticky="w", pady=12
        )

    def update_employee_values(self) -> None:
        machine = self.machine_var.get()
        employees = [name for name in self.assignments.get(machine, []) if name.strip()]
        self.employee_box.configure(values=employees)
        current = self.employee_var.get().strip()
        if current and all(normalize_text(current) != normalize_text(name) for name in employees):
            self.employee_var.set("")

    def add_dimension_row(self) -> None:
        row_frame = ttk.Frame(self.dimensions_frame)
        row_frame.pack(fill="x", pady=3)
        a_entry = ttk.Entry(row_frame, width=18)
        b_entry = ttk.Entry(row_frame, width=18)
        qty_entry = ttk.Entry(row_frame, width=18)
        a_entry.pack(side="left", padx=3)
        b_entry.pack(side="left", padx=3)
        qty_entry.pack(side="left", padx=3)
        remove = ttk.Button(row_frame, text="Usuń", width=9)
        remove.pack(side="left", padx=3)
        record = (a_entry, b_entry, qty_entry, row_frame)
        remove.configure(command=lambda: self.remove_dimension_row(record))
        self.dimension_rows.append(record)
        a_entry.focus_set()

    def remove_dimension_row(self, record) -> None:
        if len(self.dimension_rows) == 1:
            messagebox.showinfo("Informacja", "Musi pozostać co najmniej jeden zestaw wymiarów.")
            return
        self.dimension_rows.remove(record)
        record[3].destroy()

    def read_dimensions(self) -> list[tuple[float, float, int]]:
        result = []
        for index, (a_entry, b_entry, qty_entry, _) in enumerate(self.dimension_rows, start=1):
            a_text = a_entry.get().strip().replace(",", ".")
            b_text = b_entry.get().strip().replace(",", ".")
            qty_text = qty_entry.get().strip()
            if not any((a_text, b_text, qty_text)):
                continue
            if not all((a_text, b_text, qty_text)):
                raise ValueError(f"Uzupełnij wszystkie pola w zestawie wymiarów nr {index}.")
            try:
                a = float(a_text)
                b = float(b_text)
                qty = int(qty_text)
            except ValueError as exc:
                raise ValueError(f"Nieprawidłowa liczba w zestawie wymiarów nr {index}.") from exc
            if a <= 0 or b <= 0 or qty <= 0:
                raise ValueError(f"Wartości w zestawie nr {index} muszą być większe od zera.")
            result.append((a, b, qty))
        if not result:
            raise ValueError("Dodaj co najmniej jeden kompletny zestaw wymiarów.")
        return result

    def _save(self, overwrite: bool) -> SaveResult:
        machine = self.machine_var.get().strip()
        employee = self.employee_var.get().strip()
        if not machine:
            raise ValueError("Wybierz maszynę.")
        if not employee:
            raise ValueError("Wybierz lub wpisz pracownika.")

        return save_record(
            template_path=self.template_path,
            machine=machine,
            employee=employee,
            dimensions=self.read_dimensions(),
            break_entries={label: var.get().strip() for label, var in self.break_vars.items()},
            work_interval=self.work_interval_var.get().strip(),
            shift=self.shift_var.get(),
            work_day=parse_date(self.date_var.get()),
            assignments=self.assignments,
            overwrite=overwrite,
        )

    def submit(self) -> None:
        try:
            try:
                result = self._save(overwrite=False)
            except ExistingRecordError as exc:
                overwrite = messagebox.askyesno(
                    "Wpis już istnieje",
                    f"{exc}\n\nCzy zastąpić istniejące wartości?",
                )
                if not overwrite:
                    return
                result = self._save(overwrite=True)

            self.last_saved_path = result.path
            self.update_employee_values()
            assignment_note = (
                f" | nowy przydział: blok {result.employee_block}"
                if result.new_employee_assignment
                else f" | blok pracownika: {result.employee_block}"
            )
            self.status_var.set(
                f"Zapisano: {result.path.name} | arkusz: {result.sheet_name} | "
                f"wiersz: {result.row} | szt.: {result.quantity} | "
                f"m²: {result.area_m2:.4f} | zmiana: {result.shift} | praca: {result.work_hours:.2f} h | "
                f"przerwy: {result.breaks_hours:.2f} h | nastawy: {result.setups}"
                f"{assignment_note}"
            )
            messagebox.showinfo(
                "Zapisano",
                f"Dane zapisano do:\n{result.path}\n\n"
                f"Arkusz: {result.sheet_name}\nWiersz: {result.row}",
            )
        except PermissionError as exc:
            messagebox.showerror("Plik jest zajęty", str(exc))
        except (ValueError, OSError, zipfile.BadZipFile, ET.ParseError) as exc:
            messagebox.showerror("Błąd", str(exc))

    def clear_form(self) -> None:
        self.employee_var.set("")
        self.work_interval_var.set("8-16")
        self.shift_var.set("I")
        self.date_var.set(date.today().isoformat())
        for var in self.break_vars.values():
            var.set("")
        for record in self.dimension_rows[:]:
            record[3].destroy()
        self.dimension_rows.clear()
        self.add_dimension_row()
        self.status_var.set(f"Formularz wyczyszczony. Szablon: {self.template_path.name}")

    def open_output_folder(self) -> None:
        try:
            open_folder(self.last_saved_path or self.template_path)
        except OSError as exc:
            messagebox.showerror("Błąd", f"Nie można otworzyć folderu: {exc}")


if __name__ == "__main__":
    app = ProductivityApp()
    if app.winfo_exists():
        app.mainloop()
