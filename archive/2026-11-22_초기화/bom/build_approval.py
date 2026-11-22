# -*- coding: utf-8 -*-
"""BOM v0.3 → 구매 품의서 (판매처별 묶음, 결재란)"""
import sys, os, math
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pylib'))
from collections import OrderedDict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, OUT = sys.argv[1], sys.argv[2]
src = load_workbook(SRC)
items, hold = [], []
for sh in ('BOM 전체', '1단계 시험'):
    ws = src[sh]
    hdr = [c.value for c in ws[3]]
    ix = {h: i for i, h in enumerate(hdr)}
    for r in ws.iter_rows(min_row=4, values_only=True):
        if not r[1]:
            continue
        need, spare, moq, price, ship = r[9] or 0, r[10] or 0, r[11] or 1, r[13], r[15] or 0
        if need + spare == 0:
            continue
        q = math.ceil((need + spare) / moq)
        st = r[23]
        opt = r[ix['선택 옵션']] or ''
        eta = r[ix['도착 예정']] or ''
        note = (r[27] or '').replace('[결정 11-21] ', '').split(' | 이전:')[0]
        if price is None or st in ('견적 필요', '가격 미조사'):
            hold.append((r[1], r[4], st, note[:150]))
            continue
        if price == 0:
            continue  # 다른 행(키트)에 포함
        tax = 0
        if r[25] == '해외':
            base = price if r[26] == 'Y' else q * price
            if base / 1336.3 > 150:
                tax = round((q * price + ship) * (1.08 * 1.1 - 1))
        check = ''
        if st in ('확인필요', '도착 위험', '결정 필요'):
            check = ('도착 위험(' + eta + '). ' if st == '도착 위험' else '') + note[:120]
        shop = (r[19] or '').split(' (')[0].split(' +')[0]
        items.append(dict(id=r[1], shop=shop, name=r[4], opt=opt, q=q, unit=(f'{moq}개입' if moq > 1 else '개'),
                          p=price, ship=ship, tax=tax, url=r[20] or '', check=check, split=(r[26] == 'Y')))

shops = OrderedDict()
for it in sorted(items, key=lambda x: (-sum(1 for y in items if y['shop'] == x['shop']), x['shop'], x['id'])):
    shops.setdefault(it['shop'], []).append(it)

wb = Workbook()
B = Side(style='thin', color='BFBFBF'); BOX = Border(left=B, right=B, top=B, bottom=B)
HF = PatternFill('solid', fgColor='1F3864'); SF = PatternFill('solid', fgColor='D9E1F2'); WF = PatternFill('solid', fgColor='FCE4D6')
AL = Alignment(wrap_text=True, vertical='center')
MONEY = '#,##0'

def cell(ws, r, c, v, bold=False, fill=None, fmt=None, color=None):
    x = ws.cell(row=r, column=c, value=v)
    x.border = BOX; x.alignment = AL
    if bold or color:
        x.font = Font(bold=bold, color=color)
    if fill:
        x.fill = fill
    if fmt:
        x.number_format = fmt
    return x

# ---- 구매 목록
L = wb.active; L.title = '구매 목록'
cols = ['No', '판매처', '품목', '옵션·형번', '수량', '단위', '단가', '금액', '배송비', '관세·VAT', '합계', '링크', '주문 전 확인', 'BOM ID']
widths = [5, 16, 32, 30, 6, 7, 10, 11, 9, 10, 12, 38, 52, 8]
for i, (c, w) in enumerate(zip(cols, widths), 1):
    cell(L, 1, i, c, True, HF, color='FFFFFF')
    L.column_dimensions[L.cell(row=1, column=i).column_letter].width = w
L.freeze_panes = 'C2'
r, n, shop_rows = 2, 0, []
for shop, its in shops.items():
    head = r
    for c in range(1, 15):
        cell(L, r, c, '', fill=SF)
    cell(L, r, 2, shop, True, SF); cell(L, r, 3, f'{len(its)}개 품목', fill=SF)
    r += 1; first = r
    for it in its:
        n += 1
        fill = WF if it['check'] else None
        chk = it['check'] + (' / 하루 1개씩 분할 결제' if it['split'] else '')
        vals = [n, shop, it['name'], it['opt'], it['q'], it['unit'], it['p'], f'=E{r}*G{r}', it['ship'], it['tax'], f'=H{r}+I{r}+J{r}', it['url'], chk, it['id']]
        for c, v in enumerate(vals, 1):
            x = cell(L, r, c, v, fill=fill, fmt=MONEY if c in (7, 8, 9, 10, 11) else None)
            if c == 12 and v:
                x.hyperlink = v; x.font = Font(color='0563C1', underline='single')
        r += 1
    cell(L, head, 11, f'=SUM(K{first}:K{r - 1})', True, SF, MONEY)
    shop_rows.append((shop, head, len(its)))
last = r - 1
for c in range(1, 15):
    cell(L, r, c, '', fill=SF)
cell(L, r, 2, '총계', True, SF)
cell(L, r, 11, f'=SUMPRODUCT((A2:A{last}<>"")*K2:K{last})', True, SF, MONEY)
total_row = r

# ---- 요약 (맨 앞)
S = wb.create_sheet('요약', 0)
for col, w in zip('ABCD', [34, 16, 10, 56]):
    S.column_dimensions[col].width = w
S['A1'] = 'Azunyang 로봇개 부품 구매 품의'; S['A1'].font = Font(bold=True, size=16)
S['A2'] = '작성일 2026-11-21 · 근거 BOM v0.3 (가격 조사 2026-11-21) · 금액 단위 원, VAT 포함'
cell(S, 4, 1, '결재', True, SF)
for i, t in enumerate(['담당', '검토', '승인'], 2):
    cell(S, 4, i, t, True, SF)
for i in range(1, 5):
    cell(S, 5, i, '')
S.row_dimensions[5].height = 42
row = 7
cell(S, row, 1, '구매 총액 (VAT·배송비·예상 관세 포함)', True, SF); cell(S, row, 2, f"='구매 목록'!K{total_row}", True, SF, MONEY)
cell(S, row, 3, '', fill=SF); cell(S, row, 4, '보류 품목·판금 외주는 빠짐 → 실제는 더 늘어남', fill=SF); row += 1
cell(S, row, 1, '품목 / 주문처'); cell(S, row, 2, f'{len(items)}개 / {len(shops)}곳'); cell(S, row, 3, ''); cell(S, row, 4, f'보류 {len(hold)}개는 "보류" 시트'); row += 2
for i, t in enumerate(['판매처', '금액', '품목 수', '비고'], 1):
    cell(S, row, i, t, True, HF, color='FFFFFF')
row += 1
for shop, head, k in shop_rows:
    memo = ''
    if 'steadywin' in shop:
        memo = '모터 14개, 하루 1개씩 분할 결제 (면세 한도 US$150)'
    elif 'AliExpress' in shop:
        memo = '해외 배송, 12/1 넘을 수 있음'
    cell(S, row, 1, shop); cell(S, row, 2, f"='구매 목록'!K{head}", fmt=MONEY); cell(S, row, 3, k); cell(S, row, 4, memo)
    row += 1
row += 1
for t in ['주황 행 = 주문 전에 확인할 것이 있는 품목 ("구매 목록" 시트 "주문 전 확인" 열).',
          '도착 위험 품목(모터, 벤치 전원 등)은 먼저 주문.',
          '근거 자료: bom/Azunyang_BOM_v0.3.xlsx (사양·수량 산출·배선표·체결 상세)']:
    cell(S, row, 1, t); S.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    S.row_dimensions[row].height = 22; row += 1

# ---- 보류
H = wb.create_sheet('보류')
for i, (c, w) in enumerate(zip(['BOM ID', '품목', '상태', '이유·다음 할 일'], [8, 32, 12, 90]), 1):
    cell(H, 1, i, c, True, HF, color='FFFFFF')
    H.column_dimensions[H.cell(row=1, column=i).column_letter].width = w
for j, h in enumerate(hold, 2):
    for c, v in enumerate(h, 1):
        cell(H, j, c, v)

wb.save(OUT)
tot = sum(it['q'] * it['p'] + it['ship'] + it['tax'] for it in items)
print('items', len(items), 'shops', len(shops), 'hold', len(hold), 'total', f'{tot:,}')
for shop, its in shops.items():
    print(f"  {shop}: {len(its)} {sum(i['q'] * i['p'] + i['ship'] + i['tax'] for i in its):,}")
