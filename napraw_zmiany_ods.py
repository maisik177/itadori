# -*- coding: utf-8 -*-
"""Naprawia wpisy I/II umieszczone przez starszą wersję w złej kolumnie ODS."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import zipfile
import xml.etree.ElementTree as ET

from wydajnosc_ods_app import OdsDocument, repair_misplaced_shifts


def repair_file(path: Path) -> int:
    document = OdsDocument(path)
    repaired = 0
    for layout in document.layouts.values():
        repaired += repair_misplaced_shifts(document, layout)
    if repaired:
        document.save(path, make_backup=True)
    return repaired


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    selected = filedialog.askopenfilename(
        title="Wybierz plik WYDAJNOŚCI do naprawy zmian",
        filetypes=[("Arkusz ODS", "*.ods"), ("Wszystkie pliki", "*.*")],
    )
    if not selected:
        return

    path = Path(selected)
    try:
        repaired = repair_file(path)
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        messagebox.showerror("Błąd naprawy", str(exc))
        return

    if repaired:
        messagebox.showinfo(
            "Naprawiono",
            f"Przeniesiono {repaired} wpisów I/II do właściwej kolumny ZMIANA.\n\n"
            f"Plik: {path}\n"
            f"Kopia bezpieczeństwa: {path.name}.bak",
        )
    else:
        messagebox.showinfo(
            "Brak zmian",
            "Nie znaleziono wpisów I/II umieszczonych w niewłaściwej kolumnie.",
        )


if __name__ == "__main__":
    main()
