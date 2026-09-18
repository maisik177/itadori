# Itadori — rejestr wydajności ODS

Itadori to lokalny zestaw narzędzi do rejestrowania wydajności pracowników w firmowym skoroszycie ODS oraz bezpiecznego czyszczenia zapisanych wyników formuł. Główna aplikacja zapisuje dane do miesięcznych arkuszy, zachowując układ szablonu, formuły i podział na maszyny oraz pracowników.

## Główna aplikacja

`wydajnosc_ods_app.py` pozwala:

- wybrać maszynę i pracownika;
- przypisać nowego pracownika do pierwszego wolnego bloku;
- podać datę, zmianę i przedział pracy;
- wprowadzić wiele wymiarów `a × b` i ilości;
- rozdzielić przerwy na: brak towaru, awarię, serię 0, zebranie, inne i konserwację;
- obliczyć sztuki, m², czas pracy, przerwy i nastawy;
- wykryć istniejący wpis i zapytać przed zastąpieniem;
- utworzyć brakujący plik miesięczny z szablonu.

## Wymagania

- Python 3.10+ z `tkinter`;
- Windows, macOS lub Linux;
- LibreOffice Calc lub inny program obsługujący ODS.

Aplikacja używa wyłącznie standardowej biblioteki Pythona — dodatkowe pakiety nie są wymagane.

## Uruchomienie rejestru

Obok programu muszą znajdować się `wydajnosc_ods_app.py` i `SZABLON_WYDAJNOSC.ods`.

```powershell
py wydajnosc_ods_app.py
```

### Praca z formularzem

1. Wybierz maszynę.
2. Wybierz pracownika albo wpisz nowego.
3. Podaj datę jako `RRRR-MM-DD` lub `DD.MM.RRRR`.
4. Wybierz zmianę I albo II.
5. Podaj czas pracy, np. `8-16`, `6:00-14:00` lub `14-22`.
6. Dodaj przynajmniej jeden zestaw wymiarów i ilość.
7. Opcjonalnie wpisz przerwy.
8. Kliknij **Zapisz do arkusza**.

Przerwy można zapisywać jako przedział (`8:55-9:01`), minuty (`15m`, `90min`), godziny (`1,5h`) lub czas (`1:30`). Kilka wartości oddziel przecinkiem, średnikiem albo nowym wierszem.

## Pliki miesięczne

Dane trafiają obok programu do:

```text
WYDAJNOŚCI <MIESIĄC> <ROK>.ods
```

Jeżeli pliku nie ma, program kopiuje `SZABLON_WYDAJNOSC.ods`. Szablon musi zawierać rozpoznawalne arkusze maszyn, nagłówki i bloki pracowników. Skoroszyt należy zamknąć w LibreOffice przed zapisem.

## Konfiguracja

- `przydzialy_pracownikow.json` — przypisania pracowników do bloków maszyn;
- `uklad_arkusza_cache.json` — cache wykrytego układu, przyspieszający start;
- `SZABLON_WYDAJNOSC.ods` — źródłowy układ nowych miesięcy.

Po zmianie struktury szablonu układ powinien zostać ponownie rozpoznany.

## Czyszczenie cache formuł

`Itadori.py` usuwa z **kopii** ODS zapisane wartości wynikowe formuł, zachowując formuły i dane wejściowe. Dzięki temu LibreOffice przelicza skoroszyt po otwarciu.

Tryb graficzny:

```powershell
py Itadori.py
```

Tryb wiersza poleceń:

```powershell
py Itadori.py "WYDAJNOŚCI SIERPIEŃ 2026.ods"
py Itadori.py "plik.ods" --output "plik_po_czyszczeniu.ods"
```

Bez `--output` powstaje plik z końcówką `_f1.ods`. Oryginał nigdy nie jest nadpisywany. Przed publikacją wyniku narzędzie sprawdza strukturę ODS i niezmienność formuł.

## Narzędzia serwisowe

- `napraw_formuly_ods.py` — naprawa/odświeżenie formuł w kopii ODS;
- `napraw_zmiany_ods.py` — naprawa danych zmian w plikach ODS.

Przed użyciem narzędzi serwisowych zrób kopię i zamknij arkusz. Do samego czyszczenia cache preferowany jest `Itadori.py`.

## Bezpieczeństwo danych

- zapis używa pliku tymczasowego;
- formuły przelicza program arkuszowy, nie Python;
- cache wyników jest usuwany, aby nie pokazywać starych wartości;
- duplikat daty wymaga potwierdzenia zastąpienia;
- aplikacja działa lokalnie i nie wysyła danych do Internetu.

## Struktura repozytorium

- `wydajnosc_ods_app.py` — główny rejestr;
- `Itadori.py` — czyszczenie cache formuł;
- `napraw_formuly_ods.py`, `napraw_zmiany_ods.py` — narzędzia serwisowe;
- `SZABLON_WYDAJNOSC.ods` — szablon;
- `WYDAJNOŚCI ... .ods` — skoroszyty robocze/przykładowe;
- pliki JSON — przypisania i cache układu.

## Rozwiązywanie problemów

- **Brak szablonu:** umieść `SZABLON_WYDAJNOSC.ods` obok programu.
- **Kilka szablonów:** właściwy nazwij dokładnie `SZABLON_WYDAJNOSC.ods`.
- **Plik jest zajęty:** zamknij go w LibreOffice.
- **Nie znaleziono maszyn:** sprawdź nagłówki i bloki w szablonie.
- **Stare wyniki formuł:** otwórz plik w LibreOffice i przelicz albo użyj `Itadori.py` na kopii.
