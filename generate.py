import pdfplumber
import os

PDF     = "/Users/vladandrieiev/Desktop/Кухня/2_5364157966494768420.pdf"
OUT_DIR = "/Users/vladandrieiev/Desktop/Кухня/output"
os.makedirs(OUT_DIR, exist_ok=True)

def rev(text):
    if not text:
        return ''
    parts = text.split('\n')
    parts.reverse()
    return ' '.join(p[::-1] for p in parts).strip()

# ── Читаємо таблицю ──────────────────────────────────────────────────────────
with pdfplumber.open(PDF) as pdf:
    raw = pdf.pages[0].extract_tables()[0]

HEADER_ROW = raw[1]
ING_NAMES = [rev(HEADER_ROW[i]) if (i < len(HEADER_ROW) and HEADER_ROW[i]) else ''
             for i in range(4, 91)]  # 87 інгредієнтів

# ── Парсимо рядки даних ──────────────────────────────────────────────────────
DAYS_ORDER = ['Понеділок','Вівторок','Середа','Четвер',"П'ятниця",'Субота','Неділя']
DAYS_DATES = {
    'Понеділок': '06.04.2026', 'Вівторок': '07.04.2026',
    'Середа':    '08.04.2026', 'Четвер':   '09.04.2026',
    "П'ятниця":  '10.04.2026', 'Субота':   '11.04.2026',
    'Неділя':    '12.04.2026',
}

days = {d: [] for d in DAYS_ORDER}
cur_day  = None
cur_meal = None

for row in raw[2:]:
    col1 = (row[1] or '').strip()

    if col1 == 'Усього за день':
        ing_vals = [(row[i] or '').strip() if i < len(row) else '' for i in range(4, 91)]
        if cur_day:
            days[cur_day].append({
                'meal': '__total__', 'dish': 'Усього за день',
                'pct': '', 'total': (row[91] or '').strip() if len(row) > 91 else '',
                'meat': (row[92] or '').strip() if len(row) > 92 else '',
                'ings': ing_vals,
            })
        continue

    if row[0]:
        day_text = rev(row[0])
        for d in DAYS_ORDER:
            if d in day_text:
                cur_day = d
                break

    if col1:
        meal_text = rev(col1)
        if meal_text:
            cur_meal = meal_text

    if not cur_day:
        continue

    dish = (row[2] or '').replace('\n', ' ').strip()
    if not dish:
        continue

    ing_vals = [(row[i] or '').strip() if i < len(row) else '' for i in range(4, 91)]
    days[cur_day].append({
        'meal':  cur_meal or '',
        'dish':  dish,
        'pct':   (row[3] or '').strip(),
        'total': (row[91] or '').strip() if len(row) > 91 else '',
        'meat':  (row[92] or '').strip() if len(row) > 92 else '',
        'ings':  ing_vals,
    })

# ── CSS ──────────────────────────────────────────────────────────────────────
CSS = """
@page { size: A3 landscape; margin: 4mm; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 6pt; background: white; }
h2 { text-align: center; font-size: 8pt; font-weight: bold; margin-bottom: 2mm; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th, td { border: 0.4pt solid #000; padding: 1px 1px; text-align: center;
         vertical-align: middle; line-height: 1.2; font-size: 6pt; }

/* Назва страви */
td.c-name { text-align: left; font-size: 8pt; padding-left: 3px; }

/* Вертикальні заголовки (інгредієнти + фіксовані) */
th.v-hdr { height: 42mm; padding: 0; overflow: visible; }
th.v-hdr span {
    display: inline-block;
    writing-mode: vertical-rl;
    transform: rotate(180deg);
    font-size: 5pt;
    line-height: 1;
    height: 42mm;
    word-break: normal;
    overflow-wrap: break-word;
}

/* Spanning заголовок "Найменування продуктів..." */
th.ing-main-hdr { font-size: 6pt; font-weight: bold; padding: 2px; }

/* Колонка "Прийняття їжі" в tbody — вертикальний текст */
td.meal-cell {
    writing-mode: vertical-rl;
    transform: rotate(180deg);
    text-align: center;
    font-weight: bold;
    font-size: 7pt;
    background: #e0e0e0;
    padding: 2px 1px;
    white-space: nowrap;
}

/* Рядок "Усього за день" */
tr.total td { background: #e8e8e8; font-weight: bold; }

@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
"""

JS_AUTOFONT = """<script>
document.querySelectorAll('td:not(.c-name):not(.meal-cell)').forEach(function(td) {
    var size = 6;
    while (td.scrollWidth > td.offsetWidth + 1 && size > 3.5) {
        size -= 0.5;
        td.style.fontSize = size + 'pt';
    }
});
</script>"""

MEAL_ORDER = ['Сніданок', 'Обід', 'Вечеря']

def generate_day(day_name, rows, inner_only=False):
    date = DAYS_DATES[day_name]

    total_row = next((r for r in rows if r['meal'] == '__total__'), None)
    data_rows = [r for r in rows if r['meal'] != '__total__']

    # Які інгредієнти використовуються цього дня
    used_idx = [i for i in range(87)
                if any(r['ings'][i] for r in data_rows)]

    # Ширина колонок: фіксовані + інгредієнти
    # A3 landscape ~408mm usable (трохи менше для запасу)
    # Фіксовані: meal 8mm + назва 34mm + % 6mm + маса 8mm + м'ясо 8mm = 64mm
    name_w = 34
    fix_w  = 8 + name_w + 6 + 11 + 11   # meal+name+%+total+meat = 70mm
    avail  = 408 - fix_w
    ing_w  = round(avail / len(used_idx), 1) if used_idx else 7

    if inner_only:
        lines = [f'<h2>РОЗКЛАДКА ПРОДУКТІВ — {day_name}, {date} — Підрозділ Т0920</h2>', '<table>']
    else:
        lines = [
            '<!DOCTYPE html><html lang="uk"><head>',
            '<meta charset="UTF-8">',
            f'<title>Розкладка — {day_name} {date}</title>',
            f'<style>{CSS}</style>',
            '</head><body>',
            f'<h2>РОЗКЛАДКА ПРОДУКТІВ — {day_name}, {date} — Підрозділ Т0920</h2>',
            '<table>',
        ]

    # Групуємо страви по прийомах їжі
    from itertools import groupby
    meal_groups = []
    for meal, group in groupby(data_rows, key=lambda r: r['meal']):
        meal_groups.append((meal, list(group)))

    total_cols = 5 + len(used_idx)  # meal + name + % + total + meat + ings

    # colgroup: meal(8mm) + name + %(6mm) + ings + total(11mm) + meat(11mm)
    lines.append('<colgroup>')
    lines.append('<col style="width:8mm">')   # Прийняття їжі
    lines.append(f'<col style="width:{name_w}mm">')
    lines.append('<col style="width:6mm">')   # %
    for _ in used_idx:
        lines.append(f'<col style="width:{ing_w}mm">')
    lines.append('<col style="width:11mm">')  # маса готової
    lines.append('<col style="width:11mm">')  # м'ясо
    lines.append('</colgroup>')

    # thead — два рядки
    lines.append('<thead>')

    # Рядок 1: фіксовані (rowspan=2) + spanning "Найменування продуктів..." + total/meat (rowspan=2)
    lines.append('<tr>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Прийняття їжі</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Найменування страв</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>% страви за типом</span></th>')
    lines.append(f'<th class="ing-main-hdr" colspan="{len(used_idx)}">Найменування продуктів та маса їх в грамах на одну особу</th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Загальна маса готової страви, г</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Маса м\'ясних та рибних порцій, г</span></th>')
    lines.append('</tr>')

    # Рядок 2: вертикальні заголовки інгредієнтів
    lines.append('<tr>')
    for i in used_idx:
        lines.append(f'<th class="v-hdr"><span>{ING_NAMES[i]}</span></th>')
    lines.append('</tr>')
    lines.append('</thead>')

    # tbody
    lines.append('<tbody>')
    for meal, dishes in meal_groups:
        for idx, r in enumerate(dishes):
            lines.append('<tr>')
            # Клітинка прийому їжі — лише для першого рядку групи (rowspan)
            if idx == 0:
                lines.append(f'<td class="meal-cell" rowspan="{len(dishes)}">{meal}</td>')
            lines.append(f'<td class="c-name">{r["dish"]}</td>')
            lines.append(f'<td>{r["pct"]}</td>')
            for i in used_idx:
                lines.append(f'<td>{r["ings"][i]}</td>')
            lines.append(f'<td>{r["total"]}</td>')
            lines.append(f'<td>{r["meat"]}</td>')
            lines.append('</tr>')

    # Рядок "Усього за день"
    if total_row:
        lines.append('<tr class="total">')
        lines.append(f'<td colspan="2" class="c-name">Усього за день</td>')
        lines.append(f'<td></td>')
        for i in used_idx:
            lines.append(f'<td>{total_row["ings"][i]}</td>')
        lines.append(f'<td>{total_row["total"]}</td>')
        lines.append(f'<td>{total_row["meat"]}</td>')
        lines.append('</tr>')

    lines.append('</tbody></table>')
    if not inner_only:
        lines.append(JS_AUTOFONT)
        lines.append('</body></html>')
    return '\n'.join(lines)

# ── Генеруємо файли ──────────────────────────────────────────────────────────
for day in DAYS_ORDER:
    rows = days[day]
    if not rows:
        print(f"⚠️  {day}: немає рядків")
        continue
    html = generate_day(day, rows)
    safe = day.lower().replace("'", '')
    fname = os.path.join(OUT_DIR, f"rozkladka_{safe}.html")
    with open(fname, 'w', encoding='utf-8') as f:
        f.write(html)
    used = len([i for i in range(87) if any(r['ings'][i] for r in rows if r['meal'] != '__total__')])
    print(f"✅  {day} ({DAYS_DATES[day]}): {len(rows)-1} страв, {used} інгредієнтів → {fname}")

print("\nГотово! Файли у папці:", OUT_DIR)

# ── Генеруємо об'єднаний файл (всі 7 днів, по одному A3 на сторінку) ──────────
CSS_COMBINED = CSS + """
.day-wrap { break-after: page; page-break-after: always; }
.day-wrap:last-child { break-after: avoid; page-break-after: avoid; }
"""

combined_lines = [
    '<!DOCTYPE html><html lang="uk"><head>',
    '<meta charset="UTF-8">',
    '<title>Розкладка продуктів — 06.04–12.04.2026 — Т0920</title>',
    f'<style>{CSS_COMBINED}</style>',
    '</head><body>',
]
for day in DAYS_ORDER:
    rows = days[day]
    if not rows:
        continue
    combined_lines.append('<div class="day-wrap">')
    combined_lines.append(generate_day(day, rows, inner_only=True))
    combined_lines.append('</div>')
combined_lines.append(JS_AUTOFONT)
combined_lines.append('</body></html>')

combined_path = os.path.join(OUT_DIR, 'rozkladka_all.html')
with open(combined_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(combined_lines))
print(f"✅  Об'єднаний файл → {combined_path}")
