import openpyxl
import json

excel_file = "law_books.xlsx"
json_file = "books.json"

workbook = openpyxl.load_workbook(excel_file, data_only=True)
sheet = workbook["OPAC"]

books = []

for row in sheet.iter_rows(min_row=2, values_only=True):

    if not any(row):
        continue

    book = {
        "accNo": str(row[0] or ""),
        "author": str(row[1] or ""),
        "title": str(row[2] or ""),
        "volume": str(row[3] or ""),
        "edition": str(row[4] or ""),
        "publisher": str(row[5] or ""),
        "year": str(row[6] or ""),
        "pages": str(row[7] or ""),
        "binding": str(row[8] or ""),
        "class": str(row[9] or ""),
        "rate": str(row[10] or ""),
        "isbn": str(row[11] or "")
    }

    books.append(book)

with open(json_file, "w", encoding="utf-8") as file:
    json.dump(books, file, ensure_ascii=False, indent=4)

print(f"Successfully converted {len(books)} books.")
print("Created: books.json")