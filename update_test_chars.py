import openpyxl
import os

def update_test_chars(file_path='wzór.xlsx'):
    if not os.path.exists(file_path):
        print(f"Błąd: Plik {file_path} nie istnieje.")
        return

    print(f"Wczytywanie {file_path}...")
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Kolumna T (index 20)
    col_t_idx = 20
    
    # Przykładowe nazwiska z obcymi znakami
    test_names = [
        "José Ångström",
        "Müller Gröz",
        "René François",
        "Søren Hjorth",
        "Björn Svensson"
    ]

    print("Wstawianie danych testowych do Kolumny T...")
    # Wstawiamy w losowe miejsca lub po prostu w pierwsze wolne po wzorcach
    for i, name in enumerate(test_names):
        row_idx = 4 + i # Zaczynamy od 4. wiersza
        ws.cell(row=row_idx, column=col_t_idx).value = name
        # Nie drukujemy 'name', żeby uniknąć błędów kodowania w konsoli Windows
        print(f"  - Zaktualizowano wiersz {row_idx}")

    wb.save(file_path)
    print("Zapisano dane testowe.")

if __name__ == "__main__":
    update_test_chars()
