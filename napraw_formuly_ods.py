# -*- coding: utf-8 -*-
"""Naprawa formuł w plikach ODS zapisanych przez wadliwą wersję programu.

Program nie zmienia danych użytkownika. Przywraca deklaracje przestrzeni nazw
XML wymagane przez formuły ODF (głównie xmlns:of), tworząc nowy plik
z dopiskiem _NAPRAWIONY. Oryginał pozostaje bez zmian.
"""

from __future__ import annotations

import re
import sys
import tempfile
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

FORMULA_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:of:1.2"
ROOT_PATTERN = re.compile(rb"<office:document-content\b[^>]*>")
XMLNS_PATTERN = re.compile(
    rb"\s+xmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=([\"'])(.*?)\2"
)


def extract_declarations(xml_bytes: bytes) -> dict[str, str]:
    match = ROOT_PATTERN.search(xml_bytes)
    if not match:
        return {}
    result: dict[str, str] = {}
    for item in XMLNS_PATTERN.finditer(match.group(0)):
        prefix = (item.group(1) or b"").decode("utf-8")
        result[prefix] = item.group(3).decode("utf-8")
    return result


def restore_declarations(xml_bytes: bytes, declarations: dict[str, str]) -> bytes:
    match = ROOT_PATTERN.search(xml_bytes)
    if not match:
        raise ValueError("Nie znaleziono głównego elementu office:document-content.")

    opening_tag = match.group(0)
    additions: list[bytes] = []
    for prefix, uri in declarations.items():
        attribute = b"xmlns" + (b":" + prefix.encode("utf-8") if prefix else b"")
        if re.search(rb"\s+" + re.escape(attribute) + rb"\s*=", opening_tag):
            continue
        escaped = (
            uri.replace("&", "&amp;")
            .replace('"', "&quot;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        additions.append(b" " + attribute + b'=\"' + escaped.encode("utf-8") + b'\"')

    if not additions:
        return xml_bytes
    repaired_tag = opening_tag[:-1] + b"".join(additions) + b">"
    return xml_bytes[: match.start()] + repaired_tag + xml_bytes[match.end() :]


def formula_count(xml_bytes: bytes) -> int:
    return len(re.findall(rb"\btable:formula\s*=", xml_bytes))


def repair_file(source: Path) -> tuple[Path, int, bool]:
    source = source.resolve()
    if source.suffix.lower() != ".ods":
        raise ValueError("Wybierz plik z rozszerzeniem .ods.")
    if not source.exists():
        raise FileNotFoundError(source)

    with zipfile.ZipFile(source, "r") as archive:
        if "content.xml" not in archive.namelist():
            raise ValueError("Plik nie jest prawidłowym arkuszem ODS.")
        content = archive.read("content.xml")

    declarations = extract_declarations(content)

    # Pobieramy pełny zestaw deklaracji z firmowego szablonu, jeśli jest obok.
    template = source.parent / "SZABLON_WYDAJNOSC.ods"
    if template.exists() and template.resolve() != source:
        try:
            with zipfile.ZipFile(template, "r") as archive:
                declarations.update(extract_declarations(archive.read("content.xml")))
        except (OSError, zipfile.BadZipFile, KeyError):
            pass

    declarations.setdefault("of", FORMULA_NAMESPACE)
    before = formula_count(content)
    repaired_content = restore_declarations(content, declarations)
    changed = repaired_content != content

    destination = source.with_name(source.stem + "_NAPRAWIONY.ods")
    index = 2
    while destination.exists():
        destination = source.with_name(f"{source.stem}_NAPRAWIONY_{index}.ods")
        index += 1

    fd, temp_name = tempfile.mkstemp(
        prefix=destination.stem + "_", suffix=".ods", dir=destination.parent
    )
    Path(temp_name).unlink(missing_ok=True)

    try:
        with zipfile.ZipFile(source, "r") as input_zip, zipfile.ZipFile(temp_name, "w") as output_zip:
            mimetype = input_zip.read("mimetype")
            output_zip.writestr("mimetype", mimetype, compress_type=zipfile.ZIP_STORED)
            for info in input_zip.infolist():
                if info.filename in {"mimetype", "content.xml"}:
                    continue
                output_zip.writestr(info, input_zip.read(info.filename))
            output_zip.writestr(
                "content.xml", repaired_content, compress_type=zipfile.ZIP_DEFLATED
            )
        Path(temp_name).replace(destination)
    finally:
        Path(temp_name).unlink(missing_ok=True)

    with zipfile.ZipFile(destination, "r") as archive:
        verified = archive.read("content.xml")
    if formula_count(verified) != before:
        destination.unlink(missing_ok=True)
        raise RuntimeError("Liczba formuł zmieniła się podczas naprawy. Plik nie został zapisany.")
    if before and not re.search(
        rb"xmlns:of=[\"']urn:oasis:names:tc:opendocument:xmlns:of:1\.2[\"']",
        verified,
    ):
        destination.unlink(missing_ok=True)
        raise RuntimeError("Nie udało się przywrócić przestrzeni nazw formuł.")

    return destination, before, changed


def run_gui() -> None:
    root = tk.Tk()
    root.withdraw()
    source_name = filedialog.askopenfilename(
        title="Wybierz uszkodzony plik wydajności ODS",
        filetypes=[("Arkusze OpenDocument", "*.ods"), ("Wszystkie pliki", "*.*")],
    )
    if not source_name:
        return
    try:
        destination, count, changed = repair_file(Path(source_name))
        note = (
            "Przywrócono brakujące deklaracje formuł."
            if changed
            else "Deklaracje formuł były już obecne; utworzono sprawdzoną kopię."
        )
        messagebox.showinfo(
            "Naprawa zakończona",
            f"{note}\n\nLiczba zachowanych formuł: {count}\n\nPlik wynikowy:\n{destination}",
        )
    except Exception as exc:
        messagebox.showerror("Błąd naprawy", str(exc))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        try:
            for argument in sys.argv[1:]:
                output, count, changed = repair_file(Path(argument))
                print(f"Naprawiono: {output} | formuły: {count} | zmiana XML: {changed}")
        except Exception as exc:
            print(f"Błąd: {exc}", file=sys.stderr)
            raise SystemExit(1)
    else:
        run_gui()
