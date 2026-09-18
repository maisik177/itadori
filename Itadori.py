# -*- coding: utf-8 -*-
"""Czyści zapisane wyniki (cache) formuł w arkuszach ODS.

Itadori nie zmienia treści formuł ani komórek wejściowych. Usuwa wyłącznie
wartości wynikowe zapisane przy komórkach zawierających ``table:formula``.
Oryginał pozostaje bez zmian; domyślnie powstaje kopia z dopiskiem
``_f1``.

Uruchomienie bez argumentów otwiera okno wyboru pliku. Można też użyć konsoli:

    python Itadori.py "WYDAJNOŚCI SIERPIEŃ 2026.ods"
    python Itadori.py wejscie.ods --output wyjscie.ods
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox


ODS_MIMETYPE = b"application/vnd.oasis.opendocument.spreadsheet"

# Komórki nie mogą zawierać innych komórek, dlatego można bezpiecznie wycinać
# je jako zamknięte fragmenty XML. Operujemy na bajtach, aby nie przepisywać
# całego content.xml i nie zmieniać formatowania ani deklaracji przestrzeni nazw.
FORMULA_CELL_PATTERN = re.compile(
    rb"<table:table-cell\b(?=[^>]*\btable:formula\s*=)"
    rb"(?:[^>]*?/>|[^>]*>.*?</table:table-cell\s*>)",
    re.DOTALL,
)
OPENING_TAG_PATTERN = re.compile(rb"\A<table:table-cell\b[^>]*>", re.DOTALL)
FORMULA_ATTRIBUTE_PATTERN = re.compile(
    rb"\btable:formula\s*=\s*([\"'])(.*?)\1", re.DOTALL
)

# Są to wyłącznie pola z zapamiętanym wynikiem formuły. Nie usuwamy
# table:formula, stylu, powtórzeń, ochrony komórki ani żadnych innych atrybutów.
CACHED_RESULT_ATTRIBUTES = (
    b"office:value-type",
    b"office:value",
    b"office:date-value",
    b"office:time-value",
    b"office:boolean-value",
    b"office:string-value",
    b"office:currency",
    b"calcext:value-type",
)
CACHED_ATTRIBUTE_PATTERN = re.compile(
    rb"\s+(?:"
    + b"|".join(re.escape(name) for name in CACHED_RESULT_ATTRIBUTES)
    + rb")\s*=\s*([\"']).*?\1",
    re.DOTALL,
)

# Tokenizer służy tylko do rozpoznania akapitów będących bezpośrednimi dziećmi
# komórki. Dzięki temu nie usuwamy np. treści komentarza office:annotation.
XML_TOKEN_PATTERN = re.compile(
    rb"<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|<[^>]+>", re.DOTALL
)
TAG_NAME_PATTERN = re.compile(rb"</?\s*([^\s/>]+)")


@dataclass(frozen=True)
class CleanResult:
    path: Path
    formula_cells: int
    attributes_removed: int
    paragraphs_removed: int


def _is_closing_tag(token: bytes) -> bool:
    return token.startswith(b"</")


def _is_self_closing_tag(token: bytes) -> bool:
    return token.rstrip().endswith(b"/>")


def _tag_name(token: bytes) -> bytes:
    match = TAG_NAME_PATTERN.match(token)
    return match.group(1) if match else b""


def _remove_direct_result_paragraphs(body: bytes) -> tuple[bytes, int]:
    """Usuwa bezpośrednie ``text:p`` z wynikiem, zachowując zagnieżdżone opisy."""
    removals: list[tuple[int, int]] = []
    depth = 0
    paragraph_start: int | None = None
    paragraph_depth = 0

    for match in XML_TOKEN_PATTERN.finditer(body):
        token = match.group(0)
        if token.startswith((b"<!--", b"<![CDATA[", b"<?", b"<!")):
            continue

        name = _tag_name(token)
        closing = _is_closing_tag(token)
        self_closing = _is_self_closing_tag(token)

        if closing:
            depth = max(0, depth - 1)
            if paragraph_start is not None and name == b"text:p" and depth == paragraph_depth:
                removals.append((paragraph_start, match.end()))
                paragraph_start = None
            continue

        if name == b"text:p" and depth == 0:
            if self_closing:
                removals.append((match.start(), match.end()))
            else:
                paragraph_start = match.start()
                paragraph_depth = depth

        if not self_closing:
            depth += 1

    if not removals:
        return body, 0

    pieces: list[bytes] = []
    cursor = 0
    for start, end in removals:
        pieces.append(body[cursor:start])
        cursor = end
    pieces.append(body[cursor:])
    return b"".join(pieces), len(removals)


def _clean_formula_cell(cell: bytes) -> tuple[bytes, int, int]:
    opening_match = OPENING_TAG_PATTERN.match(cell)
    if not opening_match:
        raise ValueError("Nie udało się odczytać komórki z formułą w content.xml.")

    opening = opening_match.group(0)
    cleaned_opening, attributes_removed = CACHED_ATTRIBUTE_PATTERN.subn(b"", opening)

    if opening.rstrip().endswith(b"/>"):
        return cleaned_opening, attributes_removed, 0

    closing_start = cell.rfind(b"</table:table-cell")
    if closing_start < opening_match.end():
        raise ValueError("Nieprawidłowo zamknięta komórka z formułą w content.xml.")

    body = cell[opening_match.end() : closing_start]
    cleaned_body, paragraphs_removed = _remove_direct_result_paragraphs(body)
    cleaned = cleaned_opening + cleaned_body + cell[closing_start:]
    return cleaned, attributes_removed, paragraphs_removed


def _formula_values(xml_bytes: bytes) -> list[bytes]:
    values: list[bytes] = []
    for cell_match in FORMULA_CELL_PATTERN.finditer(xml_bytes):
        formula_match = FORMULA_ATTRIBUTE_PATTERN.search(cell_match.group(0))
        if not formula_match:
            raise ValueError("Wykryto komórkę formuły bez czytelnego atrybutu table:formula.")
        values.append(formula_match.group(2))
    return values


def clean_content_xml(xml_bytes: bytes) -> tuple[bytes, int, int, int]:
    """Zwraca XML bez cache oraz liczniki wykonanych, kontrolowanych zmian."""
    formulas_before = _formula_values(xml_bytes)
    attributes_removed = 0
    paragraphs_removed = 0

    def replace(match: re.Match[bytes]) -> bytes:
        nonlocal attributes_removed, paragraphs_removed
        cleaned, attr_count, paragraph_count = _clean_formula_cell(match.group(0))
        attributes_removed += attr_count
        paragraphs_removed += paragraph_count
        return cleaned

    cleaned_xml, formula_cells = FORMULA_CELL_PATTERN.subn(replace, xml_bytes)
    formulas_after = _formula_values(cleaned_xml)

    if formulas_after != formulas_before:
        raise RuntimeError("Kontrola bezpieczeństwa wykryła zmianę formuły. Przerwano zapis.")
    if formula_cells != len(formulas_before):
        raise RuntimeError("Kontrola bezpieczeństwa wykryła niepełną obróbkę formuł.")

    return cleaned_xml, formula_cells, attributes_removed, paragraphs_removed


def _default_destination(source: Path) -> Path:
    candidate = source.with_name(source.stem + "_f1.ods")
    index = 2
    while candidate.exists():
        candidate = source.with_name(f"{source.stem}_f1_{index}.ods")
        index += 1
    return candidate


def clean_file(source: Path, destination: Path | None = None) -> CleanResult:
    source = Path(source).resolve()
    if source.suffix.lower() != ".ods":
        raise ValueError("Wybierz plik z rozszerzeniem .ods.")
    if not source.is_file():
        raise FileNotFoundError(source)

    destination = Path(destination).resolve() if destination else _default_destination(source)
    if destination.suffix.lower() != ".ods":
        raise ValueError("Plik wynikowy musi mieć rozszerzenie .ods.")
    if destination == source:
        raise ValueError("Itadori nie nadpisuje oryginału. Wybierz inną nazwę pliku wynikowego.")
    if destination.exists():
        raise FileExistsError(f"Plik wynikowy już istnieje: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source, "r") as archive:
        if "mimetype" not in archive.namelist() or "content.xml" not in archive.namelist():
            raise ValueError("Plik nie jest prawidłowym arkuszem ODS.")
        if archive.read("mimetype").strip() != ODS_MIMETYPE:
            raise ValueError("Plik nie jest arkuszem OpenDocument Spreadsheet.")
        content = archive.read("content.xml")

    cleaned, formula_cells, attributes_removed, paragraphs_removed = clean_content_xml(content)
    if formula_cells == 0:
        raise ValueError("W pliku nie znaleziono żadnych komórek z formułami.")

    fd, temp_name = tempfile.mkstemp(
        prefix=destination.stem + "_", suffix=".ods", dir=destination.parent
    )
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        with zipfile.ZipFile(source, "r") as input_zip, zipfile.ZipFile(temp_path, "w") as output_zip:
            # W ODS plik mimetype powinien być pierwszy i nieskompresowany.
            mimetype_info = input_zip.getinfo("mimetype")
            output_zip.writestr(mimetype_info, input_zip.read("mimetype"), compress_type=zipfile.ZIP_STORED)

            for info in input_zip.infolist():
                if info.filename in {"mimetype", "content.xml"}:
                    continue
                output_zip.writestr(info, input_zip.read(info.filename))

            content_info = input_zip.getinfo("content.xml")
            output_zip.writestr(content_info, cleaned, compress_type=zipfile.ZIP_DEFLATED)

        # Weryfikacja gotowego archiwum przed udostępnieniem go użytkownikowi.
        with zipfile.ZipFile(temp_path, "r") as verified_zip:
            verified_content = verified_zip.read("content.xml")
            if verified_zip.read("mimetype").strip() != ODS_MIMETYPE:
                raise RuntimeError("Kontrola pliku wynikowego wykryła błędny mimetype.")
        if _formula_values(verified_content) != _formula_values(content):
            raise RuntimeError("Kontrola pliku wynikowego wykryła zmianę formuł.")

        os.replace(temp_path, destination)
    finally:
        temp_path.unlink(missing_ok=True)

    return CleanResult(
        path=destination,
        formula_cells=formula_cells,
        attributes_removed=attributes_removed,
        paragraphs_removed=paragraphs_removed,
    )


def run_gui() -> None:
    root = tk.Tk()
    root.withdraw()
    source_name = filedialog.askopenfilename(
        title="Itadori — wybierz plik ODS do wyczyszczenia cache",
        filetypes=[("Arkusze OpenDocument", "*.ods"), ("Wszystkie pliki", "*.*")],
    )
    if not source_name:
        root.destroy()
        return

    try:
        result = clean_file(Path(source_name))
        messagebox.showinfo(
            "Itadori — cache wyczyszczony",
            "Wyczyszczono zapisane wyniki formuł. Same formuły i dane wejściowe "
            "nie zostały zmienione.\n\n"
            f"Komórki z formułami: {result.formula_cells}\n"
            f"Usunięte pola cache: {result.attributes_removed}\n"
            f"Usunięte wyświetlane wyniki: {result.paragraphs_removed}\n\n"
            f"Nowy plik:\n{result.path}",
        )
    except Exception as exc:
        messagebox.showerror("Itadori — błąd", str(exc))
    finally:
        root.destroy()


def parse_args(arguments: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Czyści cache wyników formuł ODS bez zmieniania formuł i danych wejściowych."
    )
    parser.add_argument("source", nargs="?", type=Path, help="źródłowy plik .ods")
    parser.add_argument("--output", "-o", type=Path, help="nazwa nowego pliku .ods")
    return parser.parse_args(arguments)


def main(arguments: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if arguments is None else arguments)
    if args.source is None:
        if args.output is not None:
            print("Błąd: --output wymaga podania pliku źródłowego.", file=sys.stderr)
            return 2
        run_gui()
        return 0

    try:
        result = clean_file(args.source, args.output)
    except Exception as exc:
        print(f"Błąd: {exc}", file=sys.stderr)
        return 1

    print(
        f"Gotowe: {result.path} | formuły: {result.formula_cells} | "
        f"usunięte pola cache: {result.attributes_removed} | "
        f"usunięte wyniki tekstowe: {result.paragraphs_removed}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
