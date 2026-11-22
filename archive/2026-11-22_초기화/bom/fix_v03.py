# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pylib'))
from openpyxl import load_workbook
P = sys.argv[1]
wb = load_workbook(P)
for sh in ('BOM 전체', '1단계 시험'):
    ws = wb[sh]
    hdr = [c.value for c in ws[3]]
    oc = hdr.index('선택 옵션') + 1; nc = hdr.index('페이지 상품명 (11-21 조사)') + 1
    for r in range(4, ws.max_row + 1):
        bid = ws.cell(row=r, column=2).value
        if not bid: continue
        shop = ws.cell(row=r, column=20).value or ''
        if ws.cell(row=r, column=14).value is not None and shop and shop != '없음':
            ws.cell(row=r, column=26).value = '해외' if 'AliExpress' in shop else '국내'
        fix = {
            'F-02': dict(spare=28), 'F-08': dict(spare=72), 'F-12': dict(spare=48),
            'F-13': dict(opt='AMTEC 황동 인서트 M3×5 (상세 페이지에서 선택)', name='AMTEC 황동 열융착 인서트 M3'),
            'P-01': dict(opt='DOGCOM 6S 5500mAh 150C XT90S', name='DOGCOM 6셀 22.2V 5500mAh 150C 리포배터리 (XT90S)'),
            'D-04': dict(opt='AFDF□-45 (15mm 벨트용 폭으로 선택)'),
        }.get(bid)
        if fix:
            if 'spare' in fix: ws.cell(row=r, column=11).value = fix['spare']
            if 'opt' in fix: ws.cell(row=r, column=oc).value = fix['opt']
            if 'name' in fix: ws.cell(row=r, column=nc).value = fix['name']
wb.save(P)
print('fixed')
