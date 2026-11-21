# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pylib'))
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
OUT = sys.argv[1]
# (ID, 부품명, 링크, 수량, 단가, 배송비 합, 도착, 통과, 비고)
R = [
 ('A', 'GIM6010-8 + GDS68 + DE 모터 (steadywin 공식 스토어, motor-GDS68-DE)', 'https://ko.aliexpress.com/item/1005006339337286.html', 14, 147600, 7714 * 14, '11/26 - 12/02 (스탠다드)', 'X',
  '통과 후보 없음. 공식 스토어로 선택 (하루 1개씩 결제). 익스프레스 비공식 스토어(₩158,600 + ₩27,022/건)도 끝 날짜 12/01이라 이득 작음'),
 ('D-01', '타이밍 풀리 GT 5GT 20T 맞춤 (Drive Experts)', 'https://ko.aliexpress.com/item/1005013295770403.html', 5, 13950, 27022, '11/27 - 12/01 (익스프레스)', 'X',
  '주문 메모로 20T·보어12·15mm·양 플랜지 지정. 국내 미스미 정품(출하 12/1)과 비교해 결정'),
 ('D-02', '타이밍 풀리 GT 5GT 30T 맞춤 (Drive Experts)', 'https://ko.aliexpress.com/item/1005013295770403.html', 5, 13950, 0, '11/27 - 12/01 (익스프레스)', 'X',
  'D-01과 한 주문으로 묶으면 익스프레스 배송비 1회 (추정, 주문 시 확인). 30T 메모 지정'),
 ('S-04', 'Bambu Lab TPU 95A HF 1kg 블랙(51100) [한국 발송]', 'https://ko.aliexpress.com/item/1005010452623427.html', 1, 66000, 0, '11/23 - 11/26', 'O',
  '통과. 공식 브랜드, 1kg (기존 ERYONE 500g ₩14,500 대신)'),
 ('E-03', 'GPIO 2×20 스태킹 헤더 키트 (Geekbord 공식)', 'https://ko.aliexpress.com/item/1005001467856634.html', 2, 10900, 0, '11/26 - 12/02', 'X',
  '핀 길이 표기 없음. 국내 디바이스마트(해외 2주)보다는 빠름'),
 ('P-02', '충전기 ToolkitRC M7AC (youliang Store, official set)', 'https://ko.aliexpress.com/item/1005004677574079.html', 1, 83600, 15772, '11/27 - 12/01 (익스프레스)', 'X',
  '비공식. 배터리 충전은 3단계 조립 때 필요 → 12/1 도착도 무방하면 스탠다드 YunFei ₩81,400 + ₩7,428'),
 ('P-08', 'DC 접촉기 CZW 100A 24V 1NO (YUEREL Power)', 'https://ko.aliexpress.com/item/1005005457160871.html', 1, 17700, 10792, '11/26 - 12/02', 'X',
  '비공식, 국내 대안 없음. 18V 유지 여부 미확인'),
 ('P-16', 'DC-DC 8-36V → 5V 5A (RCNUN 공식, 12V-24V to 5V 5A)', 'https://ko.aliexpress.com/item/32326484181.html', 2, 3280, 28679, '11/27 - 12/01 (익스프레스)', 'X',
  '입력 8-36V로 사양 충족, 역극성 보호 미표기. 국내 Pololu(디바이스마트 77,000원×2, 해외 1주)와 비교: 품질은 Pololu'),
 ('T-02', 'XT30(2+2) 수-암 연장 케이블 30cm (EasyWired)', 'https://ko.aliexpress.com/item/1005012775760038.html', 2, 5170, 0, '11/26 - 12/02', 'X', '한국 발송 후보 없음'),
 ('T-03a', 'HX711 24bit AD 컨버터 모듈 [한국 발송] (NEATO)', 'https://ko.aliexpress.com/item/1005006954280796.html', 1, 1000, 3000, '11/23 - 11/25', 'O', '통과'),
 ('T-03b', 'S형 로드셀 0-20kg (XNQ Electric)', 'https://ko.aliexpress.com/item/32769359959.html', 1, 28450, 6185, '11/26 - 12/02', 'X',
  '20kg 옵션 (50kg와 같은 가격, 분해능 유리). 통과 후보 없음'),
 ('T-04', '벤치 전원 HANMATEK HM-310T 30V 10A (공식)', 'https://ko.aliexpress.com/item/1005004833435425.html', 1, 128000, 39575, '11/27 - 12/01 (익스프레스)', 'X',
  '1단계 시험 첫날 필요 → 익스프레스 권장. 스탠다드 무료는 12/02'),
]
wb = Workbook(); s = wb.active; s.title = 'AliExpress 도착 위험 품목'
H = ['No', 'BOM ID', '부품명', '링크', '수량', '가격 (원)', '도착 예정', '11/30 통과', '비고']
for i, (h, w) in enumerate(zip(H, [5, 8, 48, 46, 6, 12, 24, 9, 70]), 1):
    c = s.cell(row=1, column=i, value=h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
    s.column_dimensions[c.column_letter].width = w
OK = PatternFill('solid', fgColor='E2EFDA'); NG = PatternFill('solid', fgColor='FCE4D6')
for j, (bid, nm, url, q, p, ship, eta, ok, memo) in enumerate(R, 2):
    vals = [j - 1, bid, nm, url, q, q * p + ship, eta, ok, (f'단가 {p:,}원 × {q}' + (f' + 배송비 {ship:,}원' if ship else '') + ' / ' + memo)]
    for i, v in enumerate(vals, 1):
        c = s.cell(row=j, column=i, value=v); c.alignment = Alignment(wrap_text=True, vertical='top'); c.fill = OK if ok == 'O' else NG
        if i == 4: c.hyperlink = v; c.font = Font(color='0563C1', underline='single')
        if i == 6: c.number_format = '#,##0'
n = len(R) + 2
s.cell(row=n, column=3, value='합계').font = Font(bold=True)
s.cell(row=n, column=6, value=f'=SUM(F2:F{n - 1})').number_format = '#,##0'
s.cell(row=n + 2, column=3, value='조사 2026-11-21 · 가격은 옵션 선택 후 표시가 (쿠폰 제외) · 통과 = 11/30까지 도착 확정 (한국 발송 또는 끝 날짜 ≤ 11/30)')
s.freeze_panes = 'C2'
wb.save(OUT)
print(sum(q * p + ship for _, _, _, q, p, ship, *_ in R))
