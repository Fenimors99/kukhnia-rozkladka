import pdfplumber

PDF = "/Users/vladandrieiev/Desktop/Кухня/2_5364157966494768420.pdf"

with pdfplumber.open(PDF) as pdf:
    page = pdf.pages[0]
    tables = page.extract_tables()
    print(f"Знайдено таблиць: {len(tables)}")
    if tables:
        t = tables[0]
        print(f"Рядків: {len(t)}, Колонок: {len(t[0]) if t else 0}")
        print("\n--- Перші 5 рядків ---")
        for row in t[:5]:
            print(row)
        print("\n--- Рядок 1 (заголовки колонок, перші 10) ---")
        if len(t) > 1:
            print(t[1][:10])
