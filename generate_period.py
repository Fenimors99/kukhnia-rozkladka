import pdfplumber
import os

PDF    = "/Users/vladandrieiev/Desktop/Кухня/2_5364157966494768420.pdf"
OUT    = "/Users/vladandrieiev/Desktop/Кухня/output/rozkladka_period.html"

def rev(text):
    if not text:
        return ''
    parts = text.split('\n')
    parts.reverse()
    return ' '.join(p[::-1] for p in parts).strip()

# ── Читаємо таблицю ──────────────────────────────────────────────────────────
with pdfplumber.open(PDF) as pdf:
    raw = pdf.pages[0].extract_tables()[0]

# Назви інгредієнтів (cols 4–90)
ING_NAMES = [rev(raw[1][i]) if (i < len(raw[1]) and raw[1][i]) else ''
             for i in range(4, 91)]

# ── Рядки "Усього за день" (col1='Усього за день') ───────────────────────────
DAYS = [
    ('Понеділок', '06.04'),
    ('Вівторок',  '07.04'),
    ('Середа',    '08.04'),
    ('Четвер',    '09.04'),
    ("П'ятниця",  '10.04'),
    ('Субота',    '11.04'),
    ('Неділя',    '12.04'),
]

day_totals = []
for row in raw:
    if (row[1] or '').strip() == 'Усього за день':
        ing_vals = [(row[i] or '').strip() if i < len(row) else '' for i in range(4, 91)]
        total_w  = (row[91] or '').strip() if len(row) > 91 else ''
        meat_w   = (row[92] or '').strip() if len(row) > 92 else ''
        day_totals.append({'ings': ing_vals, 'total': total_w, 'meat': meat_w})

# ── Рядок "Усього за період" (col0='Усього за період') ───────────────────────
period_total = None
for row in raw:
    if (row[0] or '').strip().startswith('Усього за період'):
        period_total = {
            'ings':  [(row[i] or '').strip() if i < len(row) else '' for i in range(4, 91)],
            'total': (row[91] or '').strip() if len(row) > 91 else '',
            'meat':  (row[92] or '').strip() if len(row) > 92 else '',
        }
        break

print(f"Знайдено 'Усього за день': {len(day_totals)} (очікується 7)")
print(f"'Усього за період': {'знайдено' if period_total else 'НЕ ЗНАЙДЕНО'}")

# Перевірка
if len(day_totals) != 7:
    print("⚠️  Кількість днів не відповідає 7, перевір PDF")

# ── Визначаємо які інгредієнти мають хоч якісь дані ─────────────────────────
all_cols = [day_totals[d]['ings'] for d in range(len(day_totals))]
if period_total:
    all_cols.append(period_total['ings'])

used_idx = [i for i in range(87) if any(col[i] for col in all_cols)]
print(f"Інгредієнтів з даними: {len(used_idx)}")

# ── CSS ──────────────────────────────────────────────────────────────────────
CSS = """
@page { size: A3 landscape; margin: 4mm; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; font-size: 6pt; background: white; }
h2 { text-align: center; font-size: 8pt; font-weight: bold; margin-bottom: 2mm; }

.doc-header {
    display: flex;
    justify-content: space-between;
    margin-bottom: 2mm;
    font-size: 6.5pt;
    line-height: 1.5;
}
.doc-header-left { flex: 1; }
.doc-header-right { text-align: right; min-width: 75mm; }
.doc-header-left .title { font-size: 9pt; font-weight: bold; }
.doc-header-right .zatv { font-size: 8pt; font-weight: bold; }
.doc-header .hint { font-size: 5.5pt; color: #555; }

.doc-footer {
    margin-top: 3mm;
    font-size: 6.5pt;
    line-height: 1.6;
}
.doc-footer-row {
    display: flex;
    justify-content: space-between;
    margin-top: 1mm;
}
.doc-footer-col { flex: 1; }
.doc-footer .hint { font-size: 5.5pt; color: #555; }
table { border-collapse: collapse; width: 100%; table-layout: fixed; }
th, td {
    border: 0.4pt solid #000;
    padding: 0px 2px;
    text-align: center;
    vertical-align: middle;
    line-height: 1.1;
    font-size: 5.5pt;
}
td.name, th.name {
    text-align: left;
    font-size: 6pt;
    padding-left: 3px;
}
th.day-hdr {
    background: #d0d0d0;
    font-weight: bold;
    font-size: 7pt;
}
th.period-hdr {
    background: #a0a0a0;
    font-weight: bold;
    font-size: 7pt;
}
tr.param td { background: #f0f0f0; font-style: italic; }
tr.param td.name { font-style: italic; }
td.period-val { background: #e0e0e0; font-weight: bold; }
tr:nth-child(even) td:not(.period-val) { background: #fafafa; }
tr:nth-child(even) td.name { background: #fafafa; }
tr.param:nth-child(even) td { background: #f0f0f0; }
@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
"""

# ── Будуємо HTML ─────────────────────────────────────────────────────────────
lines = [
    '<!DOCTYPE html><html lang="uk"><head>',
    '<meta charset="UTF-8">',
    '<title>Розкладка — Усього за період 06.04–12.04.2026</title>',
    f'<style>{CSS}</style>',
    '</head><body>',
    '<div class="doc-header">',
    '  <div class="doc-header-left">',
    '    <div class="title">РОЗКЛАДКА ПРОДУКТІВ</div>',
    '    <div>за нормами пайка</div>',
    '    <div>(на одну особу на добу)</div>',
    '    <div><b>Т0920, Т0920</b></div>',
    '    <div class="hint">(військова частина, підрозділ)</div>',
    '    <div>на час з 06.04 по 12.04.2026</div>',
    '  </div>',
    '  <div class="doc-header-right">',
    '    <div class="zatv">ЗАТВЕРДЖУЮ</div>',
    '    <div>_______________________________</div>',
    '    <div class="hint">(посада)</div>',
    '    <div>_______________________________</div>',
    '    <div class="hint">(військове звання, підпис, прізвище)</div>',
    '    <div>"___" _______________ 20__ р.</div>',
    '  </div>',
    '</div>',
    '<h2>Усього за період — 06.04–12.04.2026 — Підрозділ Т0920</h2>',
    '<table>',
]

# colgroup: 1 назва (70mm) + 7 днів (30mm) + 1 період (35mm)
lines.append('<colgroup>')
lines.append('<col style="width:70mm">')
for _ in DAYS:
    lines.append('<col style="width:30mm">')
lines.append('<col style="width:35mm">')
lines.append('</colgroup>')

# thead
lines.append('<thead><tr>')
lines.append('<th class="name">Найменування продуктів та маса їх в грамах на одну особу</th>')
for name, date in DAYS:
    lines.append(f'<th class="day-hdr">{name}<br>{date}.2026</th>')
lines.append('<th class="period-hdr">Усього за період</th>')
lines.append('</tr></thead>')

# tbody
lines.append('<tbody>')

PARAMS = [
    ('Загальна маса готової страви, г', 'total'),
    ("Маса м'ясних та рибних порцій, г", 'meat'),
]

# Параметри (маса страви, маса м'яса)
for label, key in PARAMS:
    lines.append(f'<tr class="param"><td class="name">{label}</td>')
    for d in range(len(day_totals)):
        lines.append(f'<td>{day_totals[d][key]}</td>')
    pv = period_total[key] if period_total else ''
    lines.append(f'<td class="period-val">{pv}</td>')
    lines.append('</tr>')

# Рядки інгредієнтів
for i in used_idx:
    lines.append(f'<tr><td class="name">{ING_NAMES[i]}</td>')
    for d in range(len(day_totals)):
        lines.append(f'<td>{day_totals[d]["ings"][i]}</td>')
    pv = period_total['ings'][i] if period_total else ''
    lines.append(f'<td class="period-val">{pv}</td>')
    lines.append('</tr>')

lines.append('</tbody></table>')
lines += [
    '<div class="doc-footer">',
    '  <div>Заступник командира військової частини з тилу (матеріально-технічного забезпечення)</div>',
    '  <div class="hint">(військове звання, підпис, прізвище)</div>',
    '  <div class="doc-footer-row" style="margin-top:2mm">',
    '    <div class="doc-footer-col">',
    '      <div>Начальник продовольчої служби</div>',
    '      <div class="hint">(військове звання, підпис, прізвище)</div>',
    '    </div>',
    '    <div class="doc-footer-col" style="text-align:center">',
    '      <div>Начальник медичної служби</div>',
    '      <div class="hint">(старший лікар)</div>',
    '    </div>',
    '    <div class="doc-footer-col" style="text-align:right">',
    '      <div>&nbsp;</div>',
    '      <div class="hint">(військове звання, підпис, прізвище)</div>',
    '    </div>',
    '  </div>',
    '</div>',
    '</body></html>',
]

# ── Записуємо файл ───────────────────────────────────────────────────────────
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

print(f"✅  Збережено: {OUT}")
