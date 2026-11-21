# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pylib'))
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
SRC, XLSX, MD = sys.argv[1:4]
ws = load_workbook(SRC, data_only=False)['구매 목록']
rows = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not isinstance(r[0], int):
        continue
    no, shop, name, opt, q, unit, p, _, ship, tax, _, url, chk, bid = r[:14]
    total = q * p + (ship or 0) + (tax or 0)
    nm = name + (f' ({opt})' if opt and len(opt) < 40 else '')
    memo = []
    if ship: memo.append(f'배송비 {ship:,}원 포함')
    if tax: memo.append(f'관세 {tax:,}원 포함')
    if chk: memo.append(chk.split('. ')[0][:70])
    rows.append((nm, url, f'{q}{"" if unit == "개" else " × " + unit}', total, ' / '.join(memo), shop))
wb = Workbook(); s = wb.active; s.title = '부품 리스트'
for i, (h, w) in enumerate(zip(['No', '부품명', '링크', '수량', '가격 (원)', '비고', '판매처'], [5, 46, 50, 10, 12, 60, 16]), 1):
    c = s.cell(row=1, column=i, value=h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
    s.column_dimensions[c.column_letter].width = w
for j, (nm, url, q, t, m, shop) in enumerate(rows, 2):
    for i, v in enumerate([j - 1, nm, url, q, t, m, shop], 1):
        c = s.cell(row=j, column=i, value=v); c.alignment = Alignment(wrap_text=True, vertical='top')
        if i == 3 and v: c.hyperlink = v; c.font = Font(color='0563C1', underline='single')
        if i == 5: c.number_format = '#,##0'
n = len(rows) + 2
s.cell(row=n, column=2, value='합계').font = Font(bold=True)
s.cell(row=n, column=5, value=f'=SUM(E2:E{n - 1})').number_format = '#,##0'
s.freeze_panes = 'B2'
wb.save(XLSX)
tot = sum(r[3] for r in rows)
with open(MD, 'w', encoding='utf-8') as f:
    f.write(f'# 구매 부품 리스트 (2026-11-21)\n\n총 {len(rows)}개 품목, 합계 {tot:,}원 (VAT·배송비 포함). 가격 = 수량 × 단가 + 배송비.\n\n')
    f.write('| No | 부품명 | 링크 | 수량 | 가격 | 비고 |\n|---|---|---|---|---|---|\n')
    for j, (nm, url, q, t, m, shop) in enumerate(rows, 1):
        f.write(f'| {j} | {nm} | [{shop}]({url}) | {q} | {t:,} | {m} |\n')
print(len(rows), f'{tot:,}')
