# Projekt Kiciorowy - System Raportowania Refundacji

Automatyzacja generowania i przetwarzania raportów refundacji przy użyciu Excel (VBA), Python oraz skryptów pomocniczych VBS.

## Funkcjonalności

*   **Czyszczenie danych**: Automatyczne poprawianie numerów faktur, czyszczenie znaków diakrytycznych w nazwach, formatowanie dat i numerów kont.
*   **Agregacja**: Sumowanie kwot dla tych samych beneficjentów i kont, łączenie opisów faktur.
*   **Wyróżnianie**: Automatyczne podświetlanie faktur z lat 2025 i 2026 oraz kwot powyżej 500 PLN.
*   **Dzielenie raportu**: Automatyczny podział dużych raportów na części co 2000 wierszy z nazewnictwem opartym na cyfrach rzymskich (cz. I, cz. II itd.).
*   **Obsługa załączników**: Dynamiczne wykrywanie i zachowywanie danych z kolumn załączników.

## Struktura projektu

### Główne skrypty
*   `GenerateReportModule.bas`: Główny moduł VBA do Excela integrujący wszystkie funkcjonalności (czyszczenie, agregacja, podział).
*   `generate_report.py`: Odpowiednik systemu w Pythonie korzystający z biblioteki `openpyxl`.

### Skrypty pomocnicze
*   `fill_template.py`: Generator danych testowych do szablonu.
*   `format_template.py`: Skrypt do masowej korekty formatowania kolumn.
*   `update_test_chars.py`: Wstawianie do raportu nazwisk z rzadkimi znakami obcymi (test dekompozycji).
*   `get_headers_v3.vbs` / `find_zalacznik.vbs`: Narzędzia do inspekcji struktury plików Excel bez ich otwierania w GUI.

### Inne
*   `skrypt makro.txt`: Oryginalna wersja makra agregującego.

## Instalacja i użycie

### Excel (VBA)
1. Otwórz Excela i przejdź do edytora VBA (`Alt + F11`).
2. Zaimportuj plik `GenerateReportModule.bas` (lub skopiuj jego zawartość do nowego modułu).
3. Uruchom makro `GenerateReport`.

### Python
1. Zainstaluj bibliotekę `openpyxl`: `pip install openpyxl`.
2. Uruchom skrypt: `python generate_report.py [Data]`.
3. Jeśli nie podasz daty jako argumentu, skrypt zapyta o nią w konsoli.
