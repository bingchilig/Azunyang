# -*- coding: utf-8 -*-
import sys, math; sys.path.insert(0, 'pylib')
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
P = 'C:/Users/mso/Azunyang/bom/Azunyang_BOM_v0.3.xlsx'
wb = load_workbook(P); ws = wb['BOM 전체']
G, Y = PatternFill('solid', fgColor='E2EFDA'), PatternFill('solid', fgColor='FFF2CC')
row = {ws.cell(row=r, column=2).value: r for r in range(4, ws.max_row + 1)}
def upd(i, fill, **k):
    r = row[i]; cols = dict(spec=6, need=10, spare=11, moq=12, p=14, orig=15, s=16, shop=20, url=21, alt=22, st=24, note=28)
    for key, val in k.items(): ws.cell(row=r, column=cols[key]).value = val
    for c in range(1, ws.max_column + 1): ws.cell(row=r, column=c).fill = fill
# 1. 배터리: DOGCOM 5500 채택
upd('P-01', Y, st='확인필요', note='[결정 11-21] DOGCOM 6S 5500mAh 150C XT90S 채택: 12/1 전에 받을 수 있는 유일한 후보. 926g → 무게 예산 750g보다 +176~212g (전체 약 9.4kg, 토크 계산 9kg 기준 대비 +4%). 주문 전 치수가 트레이 180×80에 들어가는지 확인. XT90S가 달려 있음 → P-07은 로봇 쪽 1개만 필요. CNHL 5000(714g)은 이후 교체 후보')
upd('P-07', G, need=1, spare=0, note='[결정 11-21] 배터리에 XT90S가 달려 있어 로봇 쪽(반대 성별) 1개만. 성별은 배터리 받고 확인')
# 2. 백아이들러: 미스미 Ø45 + 핀 Ø10
upd('D-04', Y, st='확인필요', spec='미스미 아이들러 센터 베어링 타입 외경 Ø45, 양쪽 플랜지, 보어 Ø10 (6200ZZ 내장), 벨트폭 15mm 이상 대응 폭 선택',
    note='[결정 11-21] 미스미 Ø45 채택 (GT3 뒷면 최소 Ø45 충족, 출하 11/24). 조사된 AFDF19-45는 벨트폭 12.7 대응 → 15mm 벨트용 폭 형번을 같은 페이지에서 골라 주문. 핀은 Ø10으로 변경 (F-06), 625ZZ(D-05) 불필요. CAD: 아이들러 Ø45 유지, 핀 구멍 Ø5→Ø10, 지지판 장공 폭 Ø10 기준으로 수정')
upd('D-05', G, need=0, spare=0, st='조사완료', note='[결정 11-21] 아이들러에 베어링 내장 → 불필요')
upd('F-06', Y, spec='숄더볼트 Ø10 (M8 나사), 숄더 길이 = 지지 간격 25.5 + 판 두께 (CAD), 10.9 이상', p=None, orig='', st='확인필요', url='https://kr.misumi-ec.com/vona2/detail/110100143940/',
    note='[결정 11-21] 아이들러 보어 Ø10에 맞춰 Ø5→Ø10으로 변경. 같은 미스미 스트리퍼 볼트(MSB10-길이)에서 형번 선택. Ø10이라 굽힘 강도는 Ø5보다 8배(지름³), SF 1.887 → 크게 증가. 너트는 M8 나일론 너트 4개 추가 필요')
# 3. 인서트: 국내 황동 AMTEC
upd('F-13', Y, moq=1, p=550, orig='500원/개 (v0.1 목록가, VAT 별도 가정)', s=0, shop='디바이스마트 (AMTEC 황동 인서트 M3×5)', url='https://www.devicemart.co.kr/goods/catalog?code=00150007', st='확인필요',
    alt='프리쉽 구매대행 22,400원/50개 (재질 불명, 7~21일)',
    note='[결정 11-21] 국내 황동 인서트로 변경: 재질 확실, 1~3일 배송. 60개(50+10). 상품 상세 링크·외경(Ø4.6~5.0)·열융착용인지 확인 후 주문. CAD 구멍 지름을 이 제품 권장값에 맞춤')
wb.save(P)
# 합계
tot = 0
for r in ws.iter_rows(min_row=4, values_only=True):
    if not r[1]: continue
    need, spare, moq, price, ship = r[9] or 0, r[10] or 0, r[11] or 1, r[13], r[15] or 0
    q = 0 if need + spare == 0 else math.ceil((need + spare) / moq)
    sub = q * price if price else 0; tax = 0
    if r[25] == '해외' and price:
        base = price if r[26] == 'Y' else sub
        if base / 1336.3 > 150: tax = round((sub + ship) * (1.08 * 1.1 - 1))
    tot += sub + ship + tax if price is not None else 0
print('본체 합계', f'{tot:,}')
