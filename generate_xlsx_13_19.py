import openpyxl
import os
from itertools import groupby

XLSX    = "/Users/vladandrieiev/Downloads/Telegram Desktop/Продрозкладка_РМТЗ_13.04-19.04.2026-1.xlsx"
OUT_DIR = "/Users/vladandrieiev/Desktop/Кухня/output"
os.makedirs(OUT_DIR, exist_ok=True)

def fmt(v):
    if v is None or v == '':
        return ''
    if isinstance(v, float):
        if v == 0.0:
            return '0'
        s = f'{v:.3f}'.rstrip('0').rstrip('.')
        return s.replace('.', ',')
    if isinstance(v, int):
        return str(v)
    return str(v).strip()

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb['Аркуш1']
raw = list(ws.iter_rows(values_only=True))

ING_NAMES = [str(raw[11][i]).strip() if raw[11][i] else '' for i in range(4, 89)]
N_ING = len(ING_NAMES)

DAYS_ORDER = ['Понеділок', 'Вівторок', 'Середа', 'Четвер', "П'ятниця", 'Субота', 'Неділя']
DAYS_DATES = {
    'Понеділок': '13.04.2026', 'Вівторок': '14.04.2026',
    'Середа':    '15.04.2026', 'Четвер':   '16.04.2026',
    "П'ятниця":  '17.04.2026', 'Субота':   '18.04.2026',
    'Неділя':    '19.04.2026',
}

days = {d: [] for d in DAYS_ORDER}
cur_day  = None
cur_meal = None

for row in raw[12:]:
    col0 = str(row[0]).strip() if row[0] else ''
    col1 = str(row[1]).strip() if row[1] else ''
    col2 = str(row[2]).strip() if row[2] else ''

    if col0:
        for d in DAYS_ORDER:
            if col0.startswith(d):
                cur_day = d
                break

    if col1 and col1 != 'Усього за день':
        cur_meal = col1

    if not cur_day:
        continue

    if col1 == 'Усього за день':
        ing_vals = [fmt(row[i]) if i < len(row) else '' for i in range(4, 89)]
        days[cur_day].append({
            'meal': '__total__', 'dish': 'Усього за день',
            'pct': '', 'total': fmt(row[89]), 'meat': fmt(row[90]),
            'ings': ing_vals,
        })
        continue

    if not col2:
        continue

    ing_vals = [fmt(row[i]) if i < len(row) else '' for i in range(4, 89)]
    days[cur_day].append({
        'meal':  cur_meal or '',
        'dish':  col2,
        'pct':   str(row[3]).strip() if row[3] else '',
        'total': fmt(row[89]),
        'meat':  fmt(row[90]),
        'ings':  ing_vals,
    })

CSS = """
@page { size: A3 landscape; margin: 4mm; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 6pt; background: white; }
h2 { text-align: center; font-size: 8pt; font-weight: bold; margin-bottom: 2mm; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th, td { border: 0.4pt solid #000; padding: 1px 1px; text-align: center;
         vertical-align: middle; line-height: 1.2; font-size: 6pt; }

td.c-name { text-align: left; font-size: 8pt; padding-left: 3px; }

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

th.ing-main-hdr { font-size: 6pt; font-weight: bold; padding: 2px; }

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

    used_idx = [i for i in range(N_ING) if any(r['ings'][i] for r in data_rows)]

    name_w = 32
    fix_w  = 8 + name_w + 6 + 11 + 11
    avail  = 398 - fix_w
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

    meal_groups = []
    for meal, group in groupby(data_rows, key=lambda r: r['meal']):
        meal_groups.append((meal, list(group)))

    lines.append('<colgroup>')
    lines.append('<col style="width:8mm">')
    lines.append(f'<col style="width:{name_w}mm">')
    lines.append('<col style="width:6mm">')
    for _ in used_idx:
        lines.append(f'<col style="width:{ing_w}mm">')
    lines.append('<col style="width:11mm">')
    lines.append('<col style="width:11mm">')
    lines.append('</colgroup>')

    lines.append('<thead>')
    lines.append('<tr>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Прийняття їжі</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Найменування страв</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>% страви за типом</span></th>')
    lines.append(f'<th class="ing-main-hdr" colspan="{len(used_idx)}">Найменування продуктів та маса їх в грамах на одну особу</th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Загальна маса готової страви, г</span></th>')
    lines.append('<th class="v-hdr" rowspan="2"><span>Маса м\'ясних та рибних порцій, г</span></th>')
    lines.append('</tr>')
    lines.append('<tr>')
    for i in used_idx:
        lines.append(f'<th class="v-hdr"><span>{ING_NAMES[i]}</span></th>')
    lines.append('</tr>')
    lines.append('</thead>')

    lines.append('<tbody>')
    for meal, dishes in meal_groups:
        for idx, r in enumerate(dishes):
            lines.append('<tr>')
            if idx == 0:
                lines.append(f'<td class="meal-cell" rowspan="{len(dishes)}">{meal}</td>')
            lines.append(f'<td class="c-name">{r["dish"]}</td>')
            lines.append(f'<td>{r["pct"]}</td>')
            for i in used_idx:
                lines.append(f'<td>{r["ings"][i]}</td>')
            lines.append(f'<td>{r["total"]}</td>')
            lines.append(f'<td>{r["meat"]}</td>')
            lines.append('</tr>')

    if total_row:
        lines.append('<tr class="total">')
        lines.append('<td colspan="2" class="c-name">Усього за день</td>')
        lines.append('<td></td>')
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

# Окремі файли по днях
for day in DAYS_ORDER:
    rows = days[day]
    if not rows:
        print(f"  {day}: немає рядків")
        continue
    html = generate_day(day, rows)
    safe = day.lower().replace("'", '')
    fname = os.path.join(OUT_DIR, f"rozkladka_{safe}_13_19.html")
    with open(fname, 'w', encoding='utf-8') as f:
        f.write(html)
    used = len([i for i in range(N_ING) if any(r['ings'][i] for r in rows if r['meal'] != '__total__')])
    print(f"  {day} ({DAYS_DATES[day]}): {len(rows)-1} страв, {used} інгредієнтів → {fname}")

print("\nГотово! Файли у папці:", OUT_DIR)

# Об'єднаний файл (всі 7 днів, по одному A3 на сторінку)
CSS_COMBINED = CSS + """
.day-wrap { break-after: page; page-break-after: always; }
.day-wrap:last-child { break-after: avoid; page-break-after: avoid; }
"""

combined_lines = [
    '<!DOCTYPE html><html lang="uk"><head>',
    '<meta charset="UTF-8">',
    '<title>Розкладка продуктів — 13.04–19.04.2026 — Т0920</title>',
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

combined_path = os.path.join(OUT_DIR, 'rozkladka_all_13_19.html')
with open(combined_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(combined_lines))
print(f"  Об'єднаний файл → {combined_path}")
