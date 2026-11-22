# -*- coding: utf-8 -*-
"""Azunyang BOM v0.2 generator. Run: python -X utf8 build_bom_v02.py <out.xlsx>"""
import sys, math
from collections import OrderedDict
sys.path.insert(0, __file__.rsplit('\\', 1)[0] + '\\pylib')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]
wb = Workbook()
HDR = PatternFill('solid', fgColor='1F3864'); HFONT = Font(bold=True, color='FFFFFF')
SEC = PatternFill('solid', fgColor='D9E1F2'); KEEP = PatternFill('solid', fgColor='E2EFDA')
NEW = PatternFill('solid', fgColor='FFF2CC'); WARN = PatternFill('solid', fgColor='FCE4D6')
BLUE = Font(color='0000FF'); BOLD = Font(bold=True)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')

def header(ws, row, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.fill = HDR; cell.font = HFONT; cell.alignment = WRAP; cell.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=3)

def put(ws, row, vals, fill=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.alignment = WRAP; c.border = BOX
        if fill: c.fill = fill

def section(ws, row, text, ncol):
    put(ws, row, [text] + [''] * (ncol - 1), SEC)
    ws.cell(row=row, column=1).font = BOLD

# ---------------------------------------------------------------- 표지
cv = wb.active; cv.title = '표지'
cv.column_dimensions['A'].width = 46; cv.column_dimensions['B'].width = 22; cv.column_dimensions['C'].width = 90
cover = [
    ('Azunyang 로봇개 부품 목록 (BOM) v0.2', '', ''),
    ('v0.1(2026-11-20)을 상세화. 웹 가격 조사는 v0.1 값만 사용, 새 품목 가격은 비워 둠 (상태 "가격 미조사").', '', ''),
    ('가정 (파란 글씨 = 입력값)', '', ''),
    ('조사 날짜 (v0.1 가격)', '2026-11-20', 'v0.1 BOM의 AliExpress KR·국내몰 확인일'),
    ('환율 (원/US$)', 1336.3, 'v0.1과 동일'),
    ('면세 한도 (US$)', 150, '목록통관·소액면세. 같은 판매자·같은 날은 합산'),
    ('부가세율', 0.1, '수입 VAT 10%'),
    ('기본 관세율', 0.08, 'HS 8501 8%, 나머지 8% 일괄 가정'),
    ('', '', ''),
    ('총액 (원, 수식 — Excel에서 열면 계산됨)', '', ''),
    ('로봇 본체 (BOM 전체 시트)', "='BOM 전체'!S2", '가격 미조사 품목은 0으로 계산됨 → 실제보다 작음'),
    ('1단계 시험 장비', "='1단계 시험'!S2", ''),
    ('공구', "='공구'!M2", '사용자 보유 공구는 수량 0으로 바꾸면 빠짐'),
    ('가격 미조사·견적 필요 품목 수 (본체)', "=COUNTIF('BOM 전체'!X:X,\"가격 미조사\")+COUNTIF('BOM 전체'!X:X,\"견적 필요\")", ''),
    ('', '', ''),
    ('v0.1 → v0.2 바뀐 점', '', ''),
    ('1. 유지 (v0.1 그대로)', '초록 행', 'GIM6010-8 액추에이터 3행, USB-CAN(CANable Pro), 시험용 XT30(2+2) 수-암 케이블, S형 로드셀+HX711'),
    ('2. 배선', '배선표 시트 신설', '전선 한 가닥마다 시작·끝·신호·전류·AWG·색·재단 길이·양끝 단자·보호재를 지정. 커넥터·단자 수량은 배선표에서 집계'),
    ('3. 전원 보호 회로', '품목 추가', '코일 퓨즈·플라이백 다이오드·프리차지 버튼·TVS·다리별 퓨즈·E-stop 커버 분리 커넥터·LiPo 경보기·배터리 스트랩'),
    ('4. 체결', '체결 상세 시트 신설', '"M3 혼합 250개" 대신 위치별 규격·길이 산출식·강도 등급·체결 토크. 모터 체결은 12.9 강재(예압 계산)'),
    ('5. 축·아이들러', '사양 변경 제안', '축은 절단·홈 가공 주문(미스미 등, CNC 불필요). 아이들러 핀은 Ø5 숄더볼트(ISO 7379)로 바꿔 절단 작업 제거'),
    ('6. 전장', '품목 추가', 'GPIO 스태킹 헤더·M2.5 스페이서·IMU 케이블·USB-C 전원 피그테일·게임패드. DC-DC는 입력 범위가 명시된 제품으로 교체 제안'),
    ('7. 2단계 다리 지그', '품목 추가', '2020 프로파일 프레임 + MGN12 수직 레일 (지그 CAD 없음 → 추정)'),
    ('8. 출력물·공구', '시트 신설', '출력 부품 목록과 필라멘트 산출, 필요한 공구 목록'),
    ('', '', ''),
    ('색 범례', '', ''),
    ('초록', '유지(v0.1)', '사용자 요청으로 v0.1 선정 그대로'),
    ('노랑', '신규·변경', 'v0.2에서 새로 넣거나 사양을 바꾼 품목'),
    ('주황', '결정·확인 필요', '실물·CAD·데이터시트 확인 전에는 사면 안 되는 품목'),
    ('', '', ''),
    ('계산 규칙', '', ''),
    ('구매 수량', '', '=ROUNDUP((필요+예비)÷MOQ). MOQ = 구매 단위당 개수'),
    ('관세·VAT', '', '해외 행에서 물품가(분할결제 Y면 단가) ÷ 환율 > US$150 이면 (소계+배송비)×((1+관세율)(1+VAT)−1)'),
    ('길이 단위', '', '전선 재단 길이는 mm. 구매는 m 단위 (배선표 합계 → BOM 전선 행)'),
    ('근거 표기', '', 'CAD = Fusion 로그 실측, 보고서 = 설계보고서 v0.1, 계산 = 이 파일 안의 식, 추정 = 근거 약함 (실물·CAD 확인 필요)'),
]
for r, (a, b, c) in enumerate(cover, 1):
    cv.cell(row=r, column=1, value=a); cv.cell(row=r, column=2, value=b); cv.cell(row=r, column=3, value=c)
    for col in (1, 2, 3): cv.cell(row=r, column=col).alignment = WRAP
for r in (1, 3, 10, 16, 26, 31): cv.cell(row=r, column=1).font = BOLD
cv['A1'].font = Font(bold=True, size=14)
for r in range(4, 9): cv.cell(row=r, column=2).font = BLUE
RATE, VAT, DUTY = "'표지'!$B$5", "'표지'!$B$7", "'표지'!$B$8"

# ---------------------------------------------------------------- 체결 상세 (먼저 계산해서 BOM 수량에 씀)
# (위치, 체결 대상, 규격, 길이 산출, 등급, 개수/단위, 단위 수, 토크, 근거/비고)
FAST = [
 ('외전 모터 뒷면', '모터 8-M3(깊이5) ↔ 몸통 외전 마운트 3mm', 'M3×8 DIN912', '판 3 + 와셔 0.5 + 결합 4.5 = 8', '12.9 흑색', 8, 4, '1.2 N·m + Loctite 243', 'D 로그: PCD74 8×Ø3.4, 마운트 3mm'),
 ('고관절 모터 뒷면', '모터 8-M3 ↔ 고관절 장착판', 'M3×8 DIN912', '판 3(추정) + 0.5 + 4.5 = 8', '12.9 흑색', 8, 4, '1.2 N·m + 243', 'C2 hip_mount_plate, 두께 CAD 확인'),
 ('무릎 모터 뒷면', '모터 8-M3 ↔ 무릎 장착판', 'M3×8 DIN912', '판 3(추정) + 0.5 + 4.5 = 8', '12.9 흑색', 8, 4, '1.2 N·m + 243', 'C5-221951: 무릎 하우징 판 8구멍'),
 ('외전 모터 출력', '출력 6-M3(PCD25) ↔ 외전 스페이서 + U프레임 2mm', 'M3×10 DIN912', '스페이서 t + 2 + 0.5 + 4.5 (t=3 가정 → 10)', '12.9 흑색', 6, 4, '1.2 N·m + 243', 'AZN_C5_abad_spacer 두께 CAD 확인. 출력 3-Ø4 핀은 판에 Ø4.1 구멍'),
 ('고관절 모터 출력', '출력 6-M3 ↔ 고관절 출력 스페이서 + 상퇴 안판 1.5mm', 'M3×10 DIN912', '스페이서 t + 1.5 + 0.5 + 4.5 (t=3 가정 → 9.5, 와셔 1장 추가로 10)', '12.9 흑색', 6, 4, '1.2 N·m + 243', '토크 전달: 마찰 필요 예압 765 N/볼트 (C_plan 계산) ≤ 1.2 N·m 시 약 2,000 N (계산 F=T/(0.2d))'),
 ('무릎 모터 출력', '출력 6-M3 ↔ 구동축 허브(20T 쪽)', 'M3×10 DIN912', '허브 플랜지 두께 확인 후 결정 (추정)', '12.9 흑색', 6, 4, '1.2 N·m + 243', '구동축-모터 결합 방식이 로그에 없음 → 확인 필요 6번'),
 ('30T 풀리 ↔ 하퇴', '풀리 측면 Ø3.4 드릴 4개 ↔ 하퇴 PETG M3 인서트', 'M3×12 DIN912', '풀리 허브 폭 + 인서트 4 (추정)', 'A2-70 SUS', 4, 4, '0.6 N·m (PETG 인서트)', 'C_plan: 30T는 하퇴에 볼트 고정. 탁상 드릴 가공'),
 ('상퇴 횡연결 상·하', '횡연결판 ↔ 안판·바깥판 1.5mm', 'M3×8 DIN912', '1.5 + 판 3 + 와셔 2장 1 + 너트 2.4 = 7.9', 'A2-70 SUS', 8, 4, '0.8 N·m', 'C3 cross_upper/lower, 개수 추정'),
 ('6806 리테이너 내·외', '리테이너 3mm ↔ 바깥판 1.5mm', 'M3×10 DIN912', '3 + 1.5 + 와셔 1 + 너트 2.4 = 7.9 → 10 (여유)', 'A2-70 SUS', 6, 4, '0.8 N·m', 'C_plan 75행, 개수 추정'),
 ('스핀들 플랜지', '스핀들 플랜지 3mm ↔ 외전 U프레임 2mm', 'M3×8 DIN912', '3 + 2 + 1 + 2.4 = 8.4 → 10 권장', 'A2-70 SUS', 4, 4, '0.8 N·m', 'C2 spindle_flange, 개수 추정'),
 ('라이저·탑 브리지', '라이저/브리지 3mm ↔ U프레임·장착판', 'M3×10 DIN912', '3 + 3 + 1 + 2.4 = 9.4 → 10', 'A2-70 SUS', 8, 4, '0.8 N·m', 'C2 riser/top_bridge, 개수 추정'),
 ('아이들러 지지판 장공', '내·외 지지판 장공 클램프', 'M3×12 DIN912', '판 2 + 상퇴 1.5 + 와셔 1 + 너트 2.4 + 여유 = 12', 'A2-70 SUS', 4, 4, '0.8 N·m', 'C5-224605: 지지판 각 2개 M3'),
 ('아이들러 축', 'Ø5 숄더볼트 (지지 간격 25.5)', 'ISO 7379 Ø5×30 (M4 나사)', '숄더 = 지지 간격 25.5 + 판 두께 합 (CAD 확인)', '12.9', 1, 4, '2.0 N·m + 243', 'Ø5 리니어 샤프트 절단 대신. 12.9 항복 1,080 MPa > 설계 207 MPa'),
 ('6001 무릎 베어링 리테이너', '리테이너 ↔ 하퇴 PETG 인서트', 'M3×10 DIN912', '리테이너 3 + 인서트 결합 5 + 와셔', 'A2-70 SUS', 4, 4, '0.6 N·m', '개수 추정'),
 ('하퇴 ↔ 발', 'TPU 발 ↔ 하퇴 끝 인서트', 'M3×16 DIN912', '발 관통 + 인서트 5 (추정)', 'A2-70 SUS', 1, 4, '0.4 N·m', 'C5 FL_foot Ø30×16'),
 ('외전 마운트 ↔ 섀시', '마운트 3mm ↔ 섀시 3mm', 'M4×12 DIN912', '3 + 3 + 와셔 2장 1.6 + 너트 3.2 = 10.8 → 12', 'A2-70 SUS', 4, 4, '1.8 N·m', 'D 로그 mount_bolt_x/y=130/30 → 2×2 패턴 추정'),
 ('커버 ↔ 섀시', '커버 3mm ↔ 섀시 3mm', 'M4×12 DIN912', '위와 같음', 'A2-70 SUS', 8, 1, '1.8 N·m', '개수 추정'),
 ('전장 선반 ↔ 섀시', '선반 3mm ↔ 섀시', 'M3×10 DIN912', '3 + 3 + 1 + 2.4 = 9.4 → 10', 'A2-70 SUS', 4, 2, '0.8 N·m', 'D 로그 shelf ×2'),
 ('배터리 레지·서포트', 'PETG ledge ×4·support ×4 ↔ 섀시', 'M3×10 DIN912', 'PETG 4 + 3 + 와셔 + 너트', 'A2-70 SUS', 1, 8, '0.5 N·m', 'D 로그 ledge×4, support×4'),
 ('배터리 래치', 'PETG 래치 ↔ 트레이 인서트', 'M3×10 DIN912', '래치 1.2 + 인서트', 'A2-70 SUS', 2, 2, '0.4 N·m', 'D 로그 latch×2'),
 ('접근 덮개 손나사', '덮개 ↔ 섀시 M4 리벳너트', 'M4 널링 손나사 L8', 'D 로그 knob Ø14×4, 축 Ø4×7.2', 'SUS', 1, 4, '손 체결', 'D 로그 hand fastener ×4'),
 ('Pi 5 + CAN HAT', 'Pi 5 ↔ 선반, HAT 적층', 'M2.5×6 + M2.5 스페이서', 'HAT 적층 높이 = 스태킹 헤더 길이에 맞춤', '나일론/황동', 8, 1, '손 체결', '확인 필요 7번 (쿨러·HAT 간섭)'),
 ('BNO085', '보드 4구멍 Ø2.5 ↔ 선반', 'M2.5×6 + 나일론 스페이서 5mm', '', '나일론', 4, 1, '손 체결', '진동 절연은 스펀지 테이프 추가 검토'),
 ('DC-DC', '컨버터 ↔ 선반', 'M2×6 또는 M3×8 (제품 구멍에 맞춤)', '', 'SUS', 2, 1, '0.3 N·m', '제품 확정 후'),
 ('접촉기', 'CZW100 바닥 장착 구멍 ↔ 선반', 'M4×12 DIN912', '실물 장착 구멍 확인, 나일론 너트', 'A2-70 SUS', 2, 1, '1.5 N·m', '확인 필요 9번'),
 ('MIDI 퓨즈 홀더', '홀더 ↔ 선반', 'M4×12 DIN912', '나일론 너트', 'A2-70 SUS', 2, 1, '1.5 N·m', ''),
 ('분배 버스바', '버스바 ↔ 선반', 'M3×10 DIN912', '나일론 너트', 'A2-70 SUS', 4, 1, '0.8 N·m', ''),
 ('프리차지 저항', '알루미늄 하우징 저항 ↔ 섀시 (방열)', 'M3×10 DIN912', '나일론 너트', 'A2-70 SUS', 2, 1, '0.8 N·m', ''),
]
# 집계: (규격, 등급) → 수량
agg = OrderedDict()
for f in FAST:
    key = (f[2], f[4]); agg[key] = agg.get(key, 0) + f[5] * f[6]
# 너트/와셔/인서트 집계
NYLOC_M3 = sum(f[5]*f[6] for f in FAST if f[2].startswith('M3') and ('너트' in f[3] or '너트' in f[2]) )
NYLOC_M3 = 8*4 + 6*4 + 4*4 + 8*4 + 4*4 + 4*2 + 8 + 4 + 2  # 횡연결,리테이너,스핀들,라이저,아이들러,선반,레지,버스바,저항
NYLOC_M4 = 4*4 + 8 + 2 + 2 + 4  # 마운트, 커버, 접촉기, 퓨즈홀더, 숄더볼트
INSERT_M3 = 4*4 + 4*4 + 1*4 + 2*2 + 10  # 30T, 6001 리테이너, 발, 래치, 기타 PETG(전장 캐리어 등)
WASH_M3 = sum(f[5]*f[6] for f in FAST if f[2].startswith('M3')) * 1
WASH_M4 = 4*4*2 + 8*2 + 2 + 2

fs = wb.create_sheet('체결 상세')
cols = ['번호', '위치', '체결 대상', '나사 규격', '길이 산출 (mm)', '강도 등급', '개수/단위', '단위 수', '로봇 1대 수량', '체결 토크 (추정)', '근거·비고']
header(fs, 1, cols, [5, 20, 34, 22, 36, 11, 8, 8, 10, 16, 48])
r = 2
for i, f in enumerate(FAST, 1):
    put(fs, r, [i, f[0], f[1], f[2], f[3], f[4], f[5], f[6], f'=G{r}*H{r}', f[7], f[8]]); r += 1
r += 1
section(fs, r, '규격별 집계 (예비 20%는 BOM 전체 시트에서 더함)', 11); r += 1
put(fs, r, ['', '규격', '등급', '', '', '', '', '', '수량', '', '']); r += 1
for (spec, grade), q in agg.items():
    put(fs, r, ['', spec, grade, '', '', '', '', '', q, '', '']); r += 1
for name, q, note in [('M3 나일론 너트', NYLOC_M3, '관통 체결 행 합'), ('M4 나일론 너트', NYLOC_M4, ''), ('M3 평와셔', WASH_M3, 'M3 볼트 1개당 1장'), ('M4 평와셔', WASH_M4, '관통 체결 양쪽'), ('M3 황동 열융착 인서트', INSERT_M3, 'PETG 체결부')]:
    put(fs, r, ['', name, '', '', note, '', '', '', q, '', '']); r += 1
r += 1
for t in ['규칙 1. 모터 나사구멍은 깊이 5mm (Standard 도면). 볼트 끝이 바닥에 닿지 않게 나사 결합 길이를 4~4.8mm로 맞춤 → 길이 = 판 두께 합 + 와셔 + 4.5.',
          '규칙 2. 모터 체결(뒷면·출력)은 12.9 강재 + Loctite 243. 출력 6-M3만으로 11.47 N·m을 마찰로 넘기려면 볼트당 예압 765 N 필요 (C_plan, μ=0.2). M3 1.2 N·m ≈ 2,000 N (F = T ÷ 0.2d, 추정). A2-70 SUS는 체결 토크를 올리기 어려워 모터 체결에는 쓰지 않음.',
          '규칙 3. 알루미늄 모터 하우징 나사산 보호: 1.2 N·m 넘기지 말 것 (추정). 토크 드라이버 사용 (공구 시트).',
          '규칙 4. 개수에 "추정"이 붙은 행은 Fusion에서 구멍 수를 세어 수정. 판 두께는 판금 외주 시트와 같아야 함.']:
    put(fs, r, [t] + [''] * 10); fs.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11); fs.row_dimensions[r].height = 32; r += 1

# ---------------------------------------------------------------- 배선표
# (ID, 회로, 시작, 끝, 신호, 연속/피크 A, AWG, 색, 경로 추정 mm, 재단 mm, 수량, 시작 단자, 끝 단자, 보호재, 근거·비고)
W = []
def w(*a): W.append(a)
# 주 전원 (몸통 내부, 1대)
w('P01', '주 전원', 'XT90-S (로봇 쪽)', '접촉기 단자 A', '+BAT', '≈15 / 50', 12, '적', 160, 200, 1, 'XT90-S 납땜 컵', '링 단자 4mm² (접촉기 스터드 크기, M6 추정)', '열수축 Ø6 3:1 접착', '보고서: 메인 12AWG, 피크 약 50A')
w('P02', '주 전원', '접촉기 단자 B', 'MIDI 퓨즈 홀더 IN', '+BAT (차단 후)', '≈15 / 50', 12, '적', 80, 100, 1, '링 4mm² (M6 추정)', '링 4mm² (M5 추정)', '열수축 Ø6', '')
w('P03', '주 전원', 'MIDI 퓨즈 홀더 OUT', '분배 버스바 +', '+24V 버스', '≈15 / 50', 12, '적', 100, 120, 1, '링 4mm² (M5 추정)', '링 4mm² (버스바 스터드)', '열수축 Ø6', '')
w('P04', '주 전원', 'XT90-S (로봇 쪽) −', '분배 버스바 −', 'GND', '≈15 / 50', 12, '흑', 260, 300, 1, 'XT90-S 납땜 컵', '링 4mm² (버스바 스터드)', '열수축 Ø6', '')
w('P05', '프리차지', '접촉기 단자 A', '프리차지 저항 → 순간 버튼 → 접촉기 단자 B', '+BAT 프리차지', '<0.5 / 0.5', 20, '적', 3 * 100, 3 * 120, 1, '링 0.5~1.5mm² (M6, 접촉기 스터드 공유)', '저항 탭·버튼 단자 납땜', '열수축 Ø3', '3토막 (A→R, R→버튼, 버튼→B). 47Ω일 때 피크 0.54A (25.2÷47, 계산)')
w('P06', '과전압', '분배 버스바 +', '분배 버스바 −', 'TVS 다이오드', '-', '리드', '-', 30, 30, 1, 'TVS 리드 직결', 'TVS 리드 직결', '열수축 Ø3', '리드 길이 최소 (≤30mm)')
# 코일 회로
w('K01', 'E-stop 코일', '접촉기 단자 A (+BAT 상시)', '인라인 퓨즈 2A → 커버 분리 커넥터', '코일 +', '≈0.3 / 1', 20, '적', 220, 260, 1, '링 0.5~1.5mm² (M6)', 'JST-SM 2P 수 (몸통 쪽)', '열수축 Ø3', '코일 전류 미확인 (확인 필요 9번)')
w('K02', 'E-stop 코일', '커버 분리 커넥터 (커버 쪽)', 'E-stop NC 접점 왕복', '코일 + (루프)', '≈0.3 / 1', 20, '백', 180, 220, 2, 'JST-SM 2P 암 (커버 쪽)', 'E-stop 나사 단자 (페룰 0.5mm²)', '편조 슬리브 Ø6', '커버 열 때 커넥터만 분리')
w('K03', 'E-stop 코일', '커버 분리 커넥터 (몸통 쪽, 복귀)', '접촉기 코일 +', '코일 +', '≈0.3 / 1', 20, '백', 150, 180, 1, 'JST-SM 2P 수 (같은 하우징 2번 핀)', '코일 단자 (링 M4 또는 스페이드, 실물 확인)', '열수축 Ø3', '')
w('K04', 'E-stop 코일', '접촉기 코일 −', '분배 버스바 −', 'GND', '≈0.3 / 1', 20, '흑', 200, 240, 1, '코일 단자', '링 0.5~1.5mm² (버스바)', '열수축 Ø3', '')
w('K05', 'E-stop 코일', '접촉기 코일 + 단자', '접촉기 코일 − 단자', '플라이백 다이오드', '-', '리드', '-', 20, 20, 1, '다이오드 캐소드 → 코일 +', '애노드 → 코일 −', '열수축 Ø3', '코일 끊을 때 역기전력 흡수 (E-stop 접점 보호)')
# 다리 분기 (×4)
for leg, L in [('FL', 220), ('FR', 220), ('RL', 220), ('RR', 220)]:
    w(f'B-{leg}', '다리 분기', '분배 버스바 +/−', f'{leg} XT60 암 (몸통 쪽)', '+24V / GND', '≈4 / 15', 14, '적+흑', L, int(L * 1.2), 1, '링 2.5mm² (버스바) / + 선에 ATO 인라인 퓨즈 25A', 'XT60 암 납땜', '편조 슬리브 Ø6', '보고서: 분기 14AWG. 다리당 피크는 미확인 (GDS68 최대 선전류 30A/개)')
# 다리 하네스 (같은 구성 ×4)
w('LH-01', '다리 하네스 ×4', 'XT60 수 (다리 쪽)', '스플라이스 (외전 모터 옆)', '+24V / GND', '≈4 / 15', 14, '적+흑', 60, 80, 4, 'XT60 수 납땜', '3분기 납땜 스플라이스', '접착 열수축 Ø10 3:1', '스플라이스 = 14AWG 1가닥 + 16AWG 3가닥')
w('LH-02', '다리 하네스 ×4', '스플라이스', '외전 모터 XT30(2+2) 전원 핀', '+24V / GND', '≈1.5 / 10', 16, '적+흑', 80, 100, 4, '스플라이스', 'XT30(2+2) (모터와 반대 성별) 전원 컵', '열수축 Ø4', '외전 관절을 지나지 않음')
w('LH-03', '다리 하네스 ×4', '스플라이스', '고관절 모터 XT30(2+2) 전원 핀', '+24V / GND', '≈1.5 / 10', 16, '적+흑', 170, 250, 4, '스플라이스', 'XT30(2+2) 전원 컵', '편조 슬리브 Ø6 (전원+CAN 함께)', '외전 관절 통과, 서비스 루프 80mm 포함 (외전 ±20° 시험 자세)')
w('LH-04', '다리 하네스 ×4', '스플라이스', '무릎 모터 XT30(2+2) 전원 핀', '+24V / GND', '≈1.5 / 10', 16, '적+흑', 250, 330, 4, '스플라이스', 'XT30(2+2) 전원 컵', '편조 슬리브 Ø6', '고관절 모터(y 65~105) 바깥 무릎 모터(y 139~179)까지 (보고서 배치)')
# CAN (버스 2개, 같은 구성)
for bus, a, b in [('can0', 'FL', 'FR'), ('can1', 'RL', 'RR')]:
    seg = [(f'{bus}-01', 'HAT ' + bus + ' 단자', f'{a} 외전 T', 260, 300),
           (f'{bus}-02', f'{a} 외전 T', f'{a} 고관절 T', 170, 220),
           (f'{bus}-03', f'{a} 고관절 T', f'{a} 무릎 T', 80, 100),
           (f'{bus}-04', f'{a} 무릎 T', f'{b} 외전 T', 320, 380),
           (f'{bus}-05', f'{b} 외전 T', f'{b} 고관절 T', 170, 220),
           (f'{bus}-06', f'{b} 고관절 T', f'{b} 무릎 T (버스 끝)', 80, 100)]
    for sid, s, e, est, cut in seg:
        w(sid, 'CAN ' + bus + ' 간선', s, e, 'CANH / CANL', '-', 24, '황(H)+녹(L) 꼬임', est, cut, 1,
          '페룰 0.25mm² (HAT)' if sid.endswith('01') else 'T 납땜 (간선 2 + 스텁 1)', 'T 납땜', '열수축 Ø2/Ø3, 하네스는 전원선과 같은 슬리브 안에서 반대쪽', '꼬임 30회/m 이상. 외전 관절 통과 구간은 서비스 루프')
    w(f'{bus}-S', 'CAN ' + bus + ' 스텁', '각 T', '모터 XT30(2+2) 신호 핀 (6개)', 'CANH / CANL', '-', 24, '황+녹', 40, 50, 6, 'T 납땜', 'XT30(2+2) 신호 컵', '열수축 Ø2', '스텁 ≤ 50mm (1Mbps ISO 11898 권장 0.3m 이하)')
    w(f'{bus}-G', 'CAN ' + bus + ' 기준', f'HAT {bus} GND 단자', '분배 버스바 −', 'CAN GND', '-', 24, '흑', 260, 300, 1, '페룰 0.25mm²', '링 0.5~1.5mm²', '열수축 Ø2', 'HAT CAN 쪽은 절연형 → 모터 GND와 기준을 맞춤. XT30(2+2)에는 GND 신호핀이 없음')
# 5V / 제어
w('V01', '5V 전원', '분배 버스바 +', 'DC-DC IN+ (인라인 퓨즈 3A 경유)', '+24V', '≈1.3 / 1.5', 18, '적', 180, 220, 1, '링 0.5~1.5mm²', 'DC-DC 단자 (납땜 또는 페룰 0.75mm²)', '열수축 Ø3', '보고서: 24V측 약 1.2A (Pi+HAT 25W 가정)')
w('V02', '5V 전원', '분배 버스바 −', 'DC-DC IN−', 'GND', '≈1.3 / 1.5', 18, '흑', 180, 220, 1, '링 0.5~1.5mm²', 'DC-DC 단자', '열수축 Ø3', '')
w('V03', '5V 전원', 'DC-DC OUT', 'Pi 5 USB-C', '+5V / GND', '≈4 / 5', '20 (제품)', '피그테일 제품', 120, 150, 1, 'DC-DC 출력 단자 납땜', 'USB-C 수 (5A 피그테일)', '열수축 Ø3', 'PD 협상 없음 → config.txt usb_max_current_enable=1 필요')
w('V04', 'IMU', 'BNO085 핀 헤더', 'Pi 5 GPIO (UART 또는 I2C)', '3V3·GND·TX·RX(·P0)', '-', '점퍼', '-', 120, 150, 1, '2.54 암 점퍼', '2.54 암 점퍼', '-', 'Pi에서 BNO08x I2C 클럭 스트레칭 문제 → UART-RVC 권장 (추정, 확인 필요 11번)')

wsw = wb.create_sheet('배선표')
cols = ['ID', '회로', '시작', '끝', '신호', '전류 연속/피크 (A, 추정)', 'AWG', '색', '경로 추정 (mm)', '재단 길이 (mm/가닥)', '수량 (세트)', '시작 단자', '끝 단자', '보호재', '근거·비고', '총 전선 길이 (m, 가닥 수 반영)']
header(wsw, 1, cols, [10, 14, 26, 30, 16, 14, 7, 12, 10, 10, 8, 30, 30, 24, 46, 12])
def colours_of(c):
    if c == '적+흑': return ['적', '흑']
    if '꼬임' in c or ('황' in c and '녹' in c): return ['황', '녹']
    return [c]
r = 2
wire_total = {}  # (awg, colour) -> m
for row in W:
    cs_ = colours_of(row[7])
    put(wsw, r, list(row) + [f'=J{r}*K{r}*{len(cs_)}/1000'])
    if isinstance(row[6], int):
        for c in cs_:
            wire_total[(row[6], c)] = wire_total.get((row[6], c), 0) + row[9] * row[10] / 1000
    r += 1
r += 1
section(wsw, r, '전선 집계 (재단 길이 합. 구매량 = 합 × 1.3 여유, 0.5m 단위 올림)', 16); r += 1
WIRE_BUY = {}
for (awg, col), m in sorted(wire_total.items()):
    buy = math.ceil(m * 1.3 * 2) / 2
    WIRE_BUY[(awg, col)] = buy
    put(wsw, r, ['', f'{awg}AWG', col, f'재단 합 {m:.2f} m', '', '', awg, col, '', '', '', '', '', '', f'구매 {buy} m', round(m, 3)]); r += 1
r += 1
for t in ['경로 추정 = 보고서·CAD 좌표(외전 모터 x=±125, 고관절 x=±190, 외전 축 y=±45, 고관절 모터 y=65~105, 무릎 모터 y=139~179, 섀시 길이 284)로 잡은 꺾인 경로. 재단 길이 = 추정 × 약 1.2 + 서비스 루프. Fusion E 단계 배선 경로에서 3D 길이를 재서 고칠 것.',
          'CAN 배선은 선형 데이지 체인: HAT → 앞(뒤) 왼쪽 다리 3개 → 오른쪽 다리 3개 → 끝. 종단 120Ω은 HAT 점퍼 + 버스 끝 모터(ID 6, ID 12)만 enable_r120=True. 나머지 10개는 False로 설정 후 전원 OFF 상태에서 CANH–CANL ≈ 60Ω 확인.',
          'GDS68에 CAN 입출력 포트가 2개(통과형)이면 T 납땜 대신 모터→모터로 잇고 스텁을 없앤다 (확인 필요 2번).',
          '전원선과 CAN 꼬임선은 같은 편조 슬리브에 넣되 서로 꼬지 않는다. 외전 관절 통과부는 슬리브를 PETG 클립으로 고정하고 루프를 남긴다.',
          '색 규칙: 적 = +24V, 흑 = GND, 황 = CANH, 녹 = CANL, 백 = 코일 복귀.']:
    put(wsw, r, [t] + [''] * 15); wsw.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16); wsw.row_dimensions[r].height = 30; r += 1

# 커넥터·단자 집계 시트
CONN = [
 ('XT90-S 안티스파크 (수+암 1조)', 1, 1, 'P01·P04', '배터리 쪽 플러그 성별 확인 후 반대 성별을 로봇에. 안티스파크 저항은 XT90-S 수(핀) 쪽'),
 ('XT60H 수+암 1조', 4, 1, 'B-xx ↔ LH-01', '다리 분리용. 14AWG 납땜, 60A'),
 ('XT30(2+2) 커넥터 (모터와 반대 성별)', 12, 2, 'LH-02~04 + CAN 스텁', 'GDS68 쪽 커넥터 실물 확인 후 성별 결정 (확인 필요 1번)'),
 ('JST-SM 2P 수+암 1조', 1, 1, 'K01~K03', '커버 E-stop 분리용 (20AWG, 3A)'),
 ('링 단자 비절연 4mm² (12AWG) M6', 2, 2, 'P01, P02', '접촉기 스터드 실물 확인'),
 ('링 단자 비절연 4mm² (12AWG) M5', 4, 2, 'P02, P03, P04', 'MIDI 홀더·버스바'),
 ('링 단자 비절연 2.5mm² (14AWG) M5', 8, 2, 'B-xx ×4 (+/−)', '버스바'),
 ('링 단자 비절연 1.5mm² (18~20AWG) M6', 2, 2, 'P05, K01', '접촉기 스터드 공유'),
 ('링 단자 비절연 1.5mm² (18~24AWG) M5', 5, 2, 'K04, V01, V02, can0-G, can1-G', '버스바'),
 ('링/스페이드 단자 (코일 단자용)', 2, 2, 'K03, K04', '코일 단자 형태 실물 확인'),
 ('페룰 0.25mm² (24AWG)', 6, 6, 'HAT 단자: CANH·CANL·GND ×2 버스', ''),
 ('페룰 0.5mm² (20AWG)', 2, 4, 'E-stop 나사 단자', ''),
 ('페룰 0.75mm² (18AWG)', 2, 4, 'DC-DC 단자 (나사 단자형일 때)', ''),
 ('ATO 인라인 퓨즈 홀더 14AWG + 25A 퓨즈', 4, 4, 'B-xx + 선', '퓨즈 예비 4. 시험에서 끊기면 30A로 (권장 항목)'),
 ('미니 블레이드 인라인 홀더 20AWG + 2A 퓨즈', 1, 2, 'K01', '코일 회로 보호'),
 ('미니 블레이드 인라인 홀더 18AWG + 3A 퓨즈', 1, 2, 'V01', 'DC-DC 입력 보호'),
 ('USB-C 수 전원 피그테일 (5A, 20AWG)', 1, 1, 'V03', 'E-marker 불필요, 2선(VBUS/GND) 제품'),
 ('2.54 암-암 점퍼 150mm (10가닥 묶음)', 1, 0, 'V04', ''),
]
cs = wb.create_sheet('커넥터·단자')
header(cs, 1, ['번호', '커넥터·단자', '필요', '예비', '합계', '사용 배선 ID', '비고'], [5, 40, 7, 7, 7, 30, 60])
for i, c in enumerate(CONN, 2):
    put(cs, i, [i - 1, c[0], c[1], c[2], f'=C{i}+D{i}', c[3], c[4]])

# ---------------------------------------------------------------- BOM 전체
BCOLS = ['번호', 'ID', '분류', '하위 시스템', '부품명', '상세 규격 (정격·치수)', '선정 이유·조건', '사용 위치', '수량 산출 근거', '필요', '예비', 'MOQ', '구매 수량', '단가 (원)', '원래 통화 가격', '배송비 (원)', '관세·VAT (원)', '소계 (원)', '총액 (원)', '구매처', '링크', '대체품·대체 구매처', '구매 단계', '상태', '우선순위', '국내/해외', '분할결제', '근거·비고']
# col letters: J 필요 K 예비 L MOQ M 구매 N 단가 P 배송 Q 세금 R 소계 S 총액 X 상태 Z 국내/해외 AA 분할
def bom_sheet(title, items, total_col='S'):
    ws = wb.create_sheet(title)
    widths = [5, 8, 9, 14, 26, 44, 40, 12, 34, 6, 6, 6, 7, 10, 18, 9, 10, 11, 11, 22, 30, 32, 7, 11, 8, 8, 7, 44]
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=13)
    ws.cell(row=2, column=18, value='합계').font = BOLD
    header(ws, 3, BCOLS, widths)
    r = 4; first = r
    for it in items:
        if isinstance(it, tuple):
            section(ws, r, it[1], len(BCOLS)); r += 1; continue
        d = it
        fill = KEEP if d['st'] == '유지(v0.1)' else (WARN if d['st'] in ('결정 필요', '확인필요') else NEW)
        vals = ['', d['id'], d['cat'], d['sub'], d['name'], d['spec'], d['why'], d['where'], d['calc'], d['need'], d.get('spare', 0), d.get('moq', 1),
                f'=IF(J{r}+K{r}=0,0,ROUNDUP((J{r}+K{r})/L{r},0))', d.get('price'), d.get('orig', ''), d.get('ship', 0),
                f'=IF(Z{r}="해외",IF(IF(AA{r}="Y",N{r},R{r})/{RATE}>150,ROUND((R{r}+P{r})*((1+{DUTY})*(1+{VAT})-1),0),0),0)',
                f'=IF(N{r}="",0,M{r}*N{r})', f'=IF(N{r}="",0,R{r}+P{r}+Q{r})',
                d.get('shop', ''), d.get('url', ''), d.get('alt', ''), d['stage'], d['st'], d.get('pri', '필수'), d.get('dom', '국내'), d.get('split', 'N'), d.get('note', '')]
        put(ws, r, vals, fill)
        ws.cell(row=r, column=1, value=f'=ROW()-{first - 1}-COUNTBLANK($B${first}:B{r})')
        r += 1
    ws.cell(row=2, column=19, value=f'=SUM(S{first}:S{r - 1})').font = BOLD
    ws.auto_filter.ref = f'A3:{get_column_letter(len(BCOLS))}{r - 1}'
    return ws

def I(**k): return k

ACT_COMMON = dict(cat='액추에이터', sub='관절 구동', name='GIM6010-8 + GDS68 + DE', spec="SteadyWin GIM6010-8, 8:1, GDS68(SDC104) 드라이버, DE(출력축 2차 엔코더) / AliExpress 옵션 'motor-GDS68-DE'",
                  why='HANDOFF 결정표 (일체형, DE로 호밍 불필요)', where='다리', price=147600, orig='₩147,600 (AliExpress KR 표시가)', shop='AliExpress steadywin Store (공식)',
                  url='https://ko.aliexpress.com/item/1005006339337286.html', alt='ZennixTek US$103.00 (https://www.zennixtek.com/products/steadywin-gim6010-8-planetary-reducer-motor)',
                  st='유지(v0.1)', dom='해외', split='Y', note='하루 1개씩 다른 날 결제 → 개당 약 US$110 < $150 면세 계획 (세관 판단, 확인 불가). 배송비 ₩7,714/개')
B = []
B.append(('SEC', '액추에이터 (v0.1 유지)'))
B.append(I(id='A-01', calc='1단계 시험 2개', need=2, ship=7714 * 2, stage=1, **ACT_COMMON))
B.append(I(id='A-02', calc='2단계 다리 지그 3축 = 1단계 2 + 1', need=1, ship=7714, stage=2, **ACT_COMMON))
B.append(I(id='A-03', calc='12축 − 3 = 9, 예비 2', need=9, spare=2, ship=7714 * 11, stage=3, **ACT_COMMON))

B.append(('SEC', '무릎 벨트 구동 (다리 4개 기준)'))
B.append(I(id='D-01', cat='구동', sub='무릎 벨트', name='무릎 구동 풀리 20T', spec='GT3 5M(5MGT) 치형 20T, 피치원 Ø31.83, 벨트폭 15mm용(풀리 폭 ≥16), 보어 Ø12 H7, 양쪽 플랜지, 무두볼트 M4 2개(90°) 또는 클램프 허브',
           why='C5·HANDOFF 2026-11-18 확정 (GT3 15mm 20T/30T). "5GT" 표기 HTD 치형품은 제외', where='무릎 모터 구동축', calc='다리당 1 × 4', need=4, spare=1, price=13950, orig='₩13,950 (맞춤 옵션 기본가)',
           shop='AliExpress 맞춤 동기 풀리 (GT 5GT 옵션)', url='https://ko.aliexpress.com/item/1005013295770403.html', alt='미스미·성일기공 GT 5M 견적 / HTD5M 20T ₩5,901 (치형 다름)', stage=2, st='확인필요', dom='해외',
           note='주문 시 판매자에게 지정: 치형 GT3-5M(5MGT), 20T, 보어 12 H7, 폭 15mm용, 양쪽 플랜지, 무두볼트 위치. 축은 D컷 주문(D-06)과 맞춤'))
B.append(I(id='D-02', cat='구동', sub='무릎 벨트', name='무릎 피동 풀리 30T', spec='GT3 5M 30T, 피치원 Ø47.75, 폭 15mm용, 보어 Ø12, 양쪽 플랜지. 측면에 Ø3.4 구멍 4개 추가 가공(하퇴 체결)',
           why='1.5:1 감속 (20T→30T)', where='무릎 축, 하퇴에 볼트 고정', calc='다리당 1 × 4', need=4, spare=1, price=13950, orig='₩13,950 (맞춤 옵션 기본가)',
           shop='AliExpress 맞춤 동기 풀리 (GT 5GT 옵션)', url='https://ko.aliexpress.com/item/1005013295770403.html', alt='HTD5M 30T ₩9,297 (치형 다름)', stage=2, st='확인필요', dom='해외',
           note='하퇴 볼트 구멍 PCD는 CAD에서 확인 후 탁상 드릴 가공 (공구 시트). 판매자에게 가공 요청 가능하면 요청'))
B.append(I(id='D-03', cat='구동', sub='무릎 벨트', name='무릎 타이밍 벨트', spec='GT3 5M 530mm (106T), 폭 15mm, 고무+유리섬유 심선',
           why='두 풀리 경로 525.32mm보다 긴 최근접 표준 → 백아이들러로 흡수 (C2_resume)', where='무릎 (축간 200mm)', calc='다리당 1 × 4, 예비 2', need=4, spare=2, price=12667, orig='₩12,667 (15mm/530-5GT)', ship=1570,
           shop='AliExpress G5M 고무 타이밍 벨트', url='https://ko.aliexpress.com/item/1005012677058152.html', alt='POWGE 530-5GT 15mm ₩37,976', stage=2, st='조사완료', dom='해외',
           note='Gates 정품 아님. 피크 장력차 480 N 실물 시험 HOLD'))
B.append(I(id='D-04', cat='구동', sub='무릎 벨트', name='백아이들러', spec='평(무치) 아이들러, 외경 Ø45 (GT3 뒷면 최소), 폭 ≥16mm, 플랜지 양쪽, 베어링 내장 내경 Ø5',
           why='Gates: GT3 뒷면 아이들러 최소 Ø45. 기성품은 Ø35까지만 확인', where='상퇴 장공', calc='다리당 1 × 4', need=4, spare=1, price=None, orig='Ø35 기성품 ₩4,640 (기준 미달)',
           shop='결정 필요', url='https://ko.aliexpress.com/item/1005013205909644.html', alt='① Ø45 알루미늄 원판 외주 + 625ZZ 2개 ② 무치 아이들러 대신 GT3 30T 풀리를 치면 쪽 아이들러로 재배치', stage=2, st='결정 필요', dom='해외',
           note='Ø45 기성품이 없으면 ①(판금/선반 외주) 또는 ② 설계 변경 중 결정. 결정 전 구매 보류'))
B.append(I(id='D-05', cat='구동', sub='무릎 벨트', name='아이들러 베어링 625ZZ', spec='625ZZ 5×16×5, 강판 실드', why='D-04 ①안(원판 자작)일 때만 필요', where='백아이들러 내부, 2개/아이들러', calc='2 × 4', need=8, spare=2, price=1980, orig='1,800원 (VAT 별도)', ship=2700,
           shop='디바이스마트', url='https://www.devicemart.co.kr/goods/view?no=1241343', alt='원일베어링 NSK 625ZZ 1,890원', stage=2, st='결정 필요', note='D-04 결정 후'))
B.append(I(id='D-06', cat='구동', sub='축', name='무릎 구동축 Ø12 (절단·D컷 주문)', spec='Ø12 g6 강 축 (SUJ2 고주파 경화 또는 S45C), 길이 CAD 실측(약 45mm 추정), 20T 무두볼트 자리 D컷, 축 양끝 C형 스냅링 홈(DIN 471-12) 위치 지정',
           why='CNC 없음 → 미스미 등 "절단·홈·D컷 지정 주문"으로 가공을 맡김. 6801ZZ 2개 + 20T 풀리 관통', where='외전 링크 6806 스핀들 내부', calc='다리당 1 × 4', need=4, spare=1, price=None,
           shop='미스미코리아 (리니어 샤프트 홈·D컷 지정형)', url='https://kr.misumi-ec.com/', alt='AliExpress Ø12×100 크롬봉 ₩10,014 → 직접 절단·D컷 (그라인더)', stage=2, st='가격 미조사',
           note='설계는 중공 OD12/ID9.5 (C5-223343). 중공+홈 동시 주문이 안 되면 솔리드: 축당 약 +25g (7.85 g/cm³ × π/4 × 0.95² cm² × 4.5 cm, 계산)'))
B.append(I(id='D-07', cat='구동', sub='축', name='무릎 피벗 축 Ø12 (절단·홈 주문)', spec='Ø12 g6, 길이 CAD 실측(약 50mm 추정), 6001ZZ 2개 사이, 30T 풀리 D컷, 양끝 스냅링 홈',
           why='D-06과 같음', where='무릎 관절', calc='다리당 1 × 4', need=4, spare=1, price=None, shop='미스미코리아', url='https://kr.misumi-ec.com/', alt='D-06 대체와 같음', stage=2, st='가격 미조사', note=''))
B.append(I(id='D-08', cat='구동', sub='축', name='고정 중공 스핀들 Ø30/Ø21', spec='OD Ø30 h6 (6806ZZ 내륜 끼움), ID Ø21 이상 (구동축 6801 하우징 관통), 길이 20~30mm (CAD), 한쪽 플랜지 고정',
           why='6806 내륜 지지 + 구동축 관통 (C5)', where='외전 링크 (고관절 바깥)', calc='다리당 1 × 4', need=4, spare=1, price=None, shop='견적 필요: 미스미 정밀 가공품 / 국내 선반 가공소', url='https://kr.misumi-ec.com/vona2/detail/110300230060',
           alt='미스미 PISK30 ID20 파이프 + 리머 가공 (공차 미보장)', stage=2, st='견적 필요', note='CNC 없음 제약과 충돌 → 외주 가공품으로 견적. 도면(외경 공차 h6, 내경, 길이, 플랜지 볼트 구멍) 필요'))
B.append(I(id='D-09', cat='구동', sub='축 고정', name='축용 C형 스냅링 Ø12', spec='DIN 471-12, 스프링강, 두께 1.0mm', why='D-06·D-07 축방향 고정 (홈 주문 전제)', where='구동축·피벗 축 양끝',
           calc='축 2개 × 양끝 2 × 4다리', need=16, spare=4, price=None, orig='711원~ (VAT 별도, 목록)', shop='미스미코리아', url='https://kr.misumi-ec.com/vona2/mech_screw/M1809000000', alt='Ø12 클램프 칼라 (홈 없을 때, 폭 약 9mm → 스택 공간 확인)', stage=2, st='가격 미조사', note=''))
B.append(I(id='D-10', cat='구동', sub='축 고정', name='축용 C형 스냅링 Ø30', spec='DIN 471-30', why='6806 내륜 축방향 고정. AZN_C3 리테이너 판으로 고정되면 불필요', where='스핀들', calc='다리당 1 × 4', need=4, spare=2, price=None,
           shop='미스미코리아', url='https://kr.misumi-ec.com/vona2/mech_screw/M1809000000', alt='', stage=2, st='확인필요', note='CAD에서 리테이너 판이 내륜까지 잡는지 확인 후 결정'))
B.append(I(id='D-11', cat='구동', sub='예비안', name='[예비안] HTD5M 30T / 45T 풀리 + 590mm 벨트', spec='HTD5M 30T·45T 폭 15 보어 12, 벨트 590mm(118T)', why='HANDOFF 2026-11-18 대형 HTD 예비안', where='-', calc='채택 시에만 4세트', need=0, price=None, orig='30T ₩9,297 / 벨트 7,990원(VAT 별도) / 45T 미확인',
           shop='AliExpress Drive Experts / 나비엠알오', url='https://ko.aliexpress.com/item/1005012292562817.html', alt='', stage=3, st='조사완료', pri='선택', note='현재 수량 0 → 합계 제외'))

B.append(('SEC', '베어링'))
B.append(I(id='BR-01', cat='베어링', sub='고관절', name='고관절 지지 베어링 6806ZZ', spec='6806ZZ 30×42×7, NSK, Cr 5,200N / C0r 3,650N', why='고관절 바깥 지지 (외륜 = 상퇴 바깥판, 내륜 = 스핀들 D-08)', where='상퇴 바깥판', calc='다리당 1 × 4', need=4, spare=2, price=6160, orig='5,950 + ZZ 210원', ship=3500,
           shop='원일베어링 (NSK)', url='https://bearingshop.co.kr/product/일본산-nsk-볼-베어링-6806-6806zz-6806dd-6806vv/34320', alt='디바이스마트 7,600원(VAT 별도)', stage=2, st='조사완료', note='C_plan 9절'))
B.append(I(id='BR-02', cat='베어링', sub='무릎', name='무릎 피벗 베어링 6001ZZ', spec='6001ZZ 12×28×8, KBC, C0r 2,370N(NSK 기준)', why='무릎 피벗 양쪽 지지 (C5)', where='하퇴·상퇴 끝', calc='다리당 2 × 4', need=8, spare=2, price=1450, ship=3000,
           shop='베어링스토어 (KBC)', url='https://bearingstore.co.kr/product/국산-kbc-볼베어링-6001zz/12', alt='디바이스마트 1,870원(VAT 별도)', stage=2, st='조사완료', note='6801ZZ 경량화 후보는 CAD 미적용'))
B.append(I(id='BR-03', cat='베어링', sub='구동축', name='구동축 지지 베어링 6801ZZ', spec='6801ZZ 12×21×5, NSK Cr 2,110N / C0r 1,040N', why='구동축 Ø12 지지 (C5)', where='스핀들 내부', calc='다리당 2 × 4', need=8, spare=2, moq=10, price=6469, orig='₩6,469 / 10개',
           shop='AliExpress 6801ZZ 10개입', url='https://ko.aliexpress.com/item/32819549839.html', alt='원일베어링 NSK 3,010원/개 (품질 우선이면 권장)', stage=2, st='조사완료', dom='해외', note='중국산 브랜드 불명'))

B.append(('SEC', '구조 (판금·출력 재료)'))
B.append(I(id='S-01', cat='구조', sub='판금', name='알루미늄 판금 다리 1세트 (FL)', spec="Al 5052-H32, 1.5/2/3mm 레이저+절곡, 다리 1개분 ('판금 외주' 시트 다리 행 ÷ 4)", why='2단계 다리 지그로 공차·절곡 R·K(0.44 추정) 먼저 확인', where='다리', calc='1세트', need=1, price=None,
           shop='MakeIT', url='https://makeit.ai.kr/services/sheet-metal', alt='JLCCNC / CutON', stage=2, st='견적 필요', note='DXF 미출력 (D 로그). 판금 외주 시트 참고'))
B.append(I(id='S-02', cat='구조', sub='판금', name='알루미늄 판금 다리 3세트 + 몸통', spec='Al 5052-H32 몸통 3mm (섀시·외전 마운트 4·커버·전장 선반 2) + 다리 3세트', why='', where='공통', calc='1식', need=1, price=None,
           shop='MakeIT', url='https://makeit.ai.kr/services/sheet-metal', alt='CutON', stage=3, st='견적 필요', note=''))
B.append(I(id='S-03', cat='구조', sub='출력 재료', name='PETG 필라멘트', spec='PETG 1.75mm 1kg, 노즐 240~250°C. 색 1종(검정 권장, UV·자국 덜 보임)', why="출력물 시트 산출 약 1.3kg (실패·시험 지그 30% 포함)", where='하퇴·트레이·브래킷·지그', calc="출력물 시트 합계 ÷ 1kg 올림", need=2, price=15000, ship=3000,
           shop='Bambu Lab KR 공식', url='https://kr.store.bambulab.com/kr/products/petg-basic', alt='PLABS eSUN PETG 18,700원', stage=1, st='조사완료', note='₩15,000은 리필(스풀 없음) 옵션 추정'))
B.append(I(id='S-04', cat='구조', sub='출력 재료', name='TPU 95A 필라멘트', spec='TPU 95A 1.75mm 500g', why='발 패드 4개 약 55g (CAD 13.8g/개)', where='발', calc='1롤', need=1, price=14500, shop='AliExpress ERYONE', url='https://ko.aliexpress.com/item/1005012007471380.html',
           alt='Bambu TPU 95A HF ₩66,000', stage=2, st='조사완료', dom='해외'))

def fast_qty(spec, grade): return agg.get((spec, grade), 0)
B.append(('SEC', "체결 부품 ('체결 상세' 시트에서 집계, 예비 20%)"))
fid = 0
for (spec, grade), q in agg.items():
    if spec.startswith('M2') or spec.startswith('M4 널링'): continue
    fid += 1
    steel = grade.startswith('12.9')
    price = 170 if (grade.startswith('A2') and spec.startswith('M3')) else None
    B.append(I(id=f'F-{fid:02d}', cat='체결', sub='볼트', name=f'{spec} ({grade})', spec=f'{spec}, {grade}' + (', 흑색 산화' if steel else ''),
               why='모터 체결 예압 확보 (체결 상세 규칙 2)' if steel else '일반 구조 체결', where='체결 상세 시트 참고', calc=f'체결 상세 집계 {q}개 + 20%',
               need=q, spare=math.ceil(q * 0.2), price=price, orig='170원/개 (v0.1 M3 SUS 낱개)' if price else '', shop='굿나잇몰 (SUS 낱개)' if price else ('미스미코리아 / 볼트몰' if steel else '굿나잇몰'),
               url='https://goodnightmall.com/product/스텐-볼트너트-세트/301' if price else '', alt='', stage=2 if q else 3, st='조사완료' if price else '가격 미조사'))
for nm, spec, q, why, price in [('M3 나일론 너트 (SUS)', 'DIN 985 M3', NYLOC_M3, '관통 체결', 70), ('M4 나일론 너트 (SUS)', 'DIN 985 M4', NYLOC_M4, '관통 체결', 70),
                                ('M3 평와셔', 'DIN 125 M3 (SUS) / 모터 체결부는 경화 와셔', WASH_M3, '판 눌림 방지', 10), ('M4 평와셔', 'DIN 125 M4', WASH_M4, '', 10)]:
    fid += 1
    B.append(I(id=f'F-{fid:02d}', cat='체결', sub='너트·와셔', name=nm, spec=spec, why=why, where='체결 상세', calc=f'체결 상세 집계 {q} + 20%', need=q, spare=math.ceil(q * 0.2), price=price,
               orig=f'{price}원/개 (v0.1 목록가)', shop='11번가 / 디바이스마트', url='https://www.11st.co.kr/products/4236196158', stage=2, st='조사완료'))
B.append(I(id='F-13', cat='체결', sub='인서트', name='M3 황동 열융착 인서트', spec='M3, 외경 약 Ø4.6~5.0, 길이 5~5.7mm (CAD 구멍 Ø4.0~4.2에 맞춤), 널링형', why='PETG 체결부 (30T 풀리, 6001 리테이너, 발, 래치, 전장 캐리어)', where='PETG 출력물',
           calc=f'체결 상세 집계 {INSERT_M3} + 20%', need=INSERT_M3, spare=math.ceil(INSERT_M3 * 0.2), moq=50, price=22400, orig='22,400원 / 50세트', ship=3000, shop='프리쉽 (구매대행)', url='https://freeship.co.kr/Goods/mini/25690133',
           alt='디바이스마트 AMTEC M3×5 500원/개', stage=1, st='조사완료', note='인서트 외경에 맞춘 CAD 구멍 지름 확인'))
B.append(I(id='F-14', cat='체결', sub='인서트', name='M4 리벳너트 (알루미늄)', spec='M4 블라인드 리벳너트, 판 두께 0.5~3mm용, 평머리', why='섀시 3mm에 접근 덮개 손나사 받침 (뒤에서 너트를 잡을 수 없음)', where='섀시', calc='손나사 4', need=4, spare=6, moq=10, price=None,
           shop='국내 공구몰', url='', alt='M4 클린칭 너트 (프레스 필요)', stage=3, st='가격 미조사', note='리벳너트 공구 필요 (공구 시트)'))
B.append(I(id='F-15', cat='체결', sub='볼트', name='M4 널링 손나사', spec='M4 × L8, 머리 Ø14×4 (D 로그 envelope)', why='배터리 접근판 공구 없이 탈착', where='커버', calc='D 로그 ×4', need=4, spare=2, price=None, shop='확인 불가 (v0.1 미발견)', url='', alt='PETG 노브 출력 + M4×10 볼트 압입', stage=3, st='가격 미조사'))
B.append(I(id='F-16', cat='체결', sub='전장 장착', name='M2.5 스페이서·볼트 키트', spec='M2.5 황동 스페이서 (6/10/11/12/16/20mm) + M2.5×6 볼트·너트 키트', why='Pi 5 + CAN HAT 적층, BNO085 장착 (체결 상세)', where='전장 선반', calc='Pi 4 + HAT 4 + IMU 4 = 12개 + 키트', need=1, price=None,
           shop='디바이스마트 / 엘레파츠', url='', alt='', stage=2, st='가격 미조사'))
B.append(I(id='F-17', cat='체결', sub='소모품', name='나사 고정제 Loctite 243', spec='중강도 10ml', why='모터 체결·진동부 (체결 상세)', where='공통', calc='1', need=1, price=10120, ship=3000, shop='툴앤샵', url='https://www.toolnshop.kr/viewgs?goods=4968&gs=108432', alt='다우몰 7,900원', stage=1, st='조사완료'))

B.append(('SEC', '전장 (제어)'))
B.append(I(id='E-01', cat='전장', sub='제어', name='Raspberry Pi 5 8GB', spec='BCM2712, 8GB, 5V/5A USB-C', why='ROS2 + CHAMP 상위 제어 (HANDOFF 3.3)', where='전장 선반 위층', calc='1', need=1, price=177518, orig='161,380원 (VAT 별도)', ship=2700,
           shop='디바이스마트 (칩원스탑)', url='https://www.devicemart.co.kr/goods/view?no=15392299', alt='Adafruit #5813 US$200', stage=2, st='조사완료', note='평균 준비 2주 → 일찍 주문'))
B.append(I(id='E-02', cat='전장', sub='제어', name='Pi 5 액티브 쿨러', spec='Raspberry Pi 공식 Active Cooler (SC1148)', why='Pi 5 방열 (몸통 밀폐)', where='Pi 5', calc='1', need=1, price=7000, shop='엘레파츠', url='', alt='Adafruit #5815 US$13.50', stage=2, st='확인필요', note='HAT과 높이 간섭 → E-03 스태킹 헤더 (확인 필요 7번)'))
B.append(I(id='E-03', cat='전장', sub='제어', name='GPIO 2×20 스태킹 헤더 (긴 핀)', spec='2×20 암 헤더, 핀 길이 약 10~13mm 연장형', why='액티브 쿨러 위로 CAN HAT 높이 확보', where='Pi ↔ HAT', calc='1 + 예비 1', need=1, spare=1, price=None, shop='디바이스마트 / 엘레파츠', url='', alt='Pi 5 전용 쿨러+HAT 호환 스페이서 세트', stage=2, st='가격 미조사', note='필요 높이는 쿨러 높이 실측 후'))
B.append(I(id='E-04', cat='전장', sub='제어', name='microSD 64GB', spec='Samsung EVO Plus 64GB U3 V30 A2 (국내 정품)', why='Pi 5 OS', where='Pi 5', calc='1', need=1, price=28115, ship=2603, shop='AliExpress (v0.1)', url='https://ko.aliexpress.com/item/1005012417449427.html', alt='국내 정품 판매처 권장 (위조품 위험)', stage=2, st='확인필요', dom='해외'))
B.append(I(id='E-05', cat='전장', sub='CAN', name='2-CH CAN FD HAT', spec='Waveshare 2-CH CAN FD HAT, MCP2518FD, 절연형, 120Ω 점퍼, 나사 단자', why='CAN 버스 2개 (can0 앞 / can1 뒤)', where='Pi 5 위', calc='1', need=1, price=81481, shop='AliExpress Waveshare 리셀러', url='https://ko.aliexpress.com/item/1005005315735215.html',
           alt='Waveshare 공식 US$46.99', stage=2, st='조사완료', dom='해외', note='외부 전원 입력 쓰지 말고 Pi 5V로 구동. 절연 쪽 GND는 배선표 canX-G로 모터 GND와 연결'))
B.append(I(id='E-06', cat='전장', sub='센서', name='IMU BNO085', spec='Adafruit 4754, 9-DOF, I2C/UART-RVC, 3.3V', why='몸통 자세', where='몸통 중심 근처 선반', calc='1', need=1, price=33675, shop='AliExpress (Adafruit 리셀러)', url='https://ko.aliexpress.com/item/1005010533705883.html', alt='Adafruit 직구 US$29.50', stage=3, st='조사완료', dom='해외',
           note='Pi에서는 UART-RVC 권장 (배선표 V04)'))
B.append(I(id='E-07', cat='전장', sub='조작', name='블루투스 게임패드', spec='Linux 호환 (PS4/PS5 듀얼센스 또는 8BitDo SN30 Pro), 블루투스', why='텔레옵 (보고서 그림 19, teleop_twist)', where='-', calc='1', need=1, price=None, shop='국내 쇼핑몰', url='', alt='', stage=3, st='가격 미조사', pri='권장', note='보유 패드 있으면 0'))

B.append(('SEC', '전원'))
B.append(I(id='P-01', cat='전원', sub='배터리', name='6S LiPo 5000mAh', spec='6S1P 22.2V (18~25.2V), 5000mAh, 70C, XT90 플러그, 약 714g, 149×51×49mm', why='보고서 전원안. 트레이 180×80 안에 들어감 (D 로그)', where='배터리 트레이', calc='1 (예비 배터리는 운용 후 결정)', need=1, price=None, orig='CNHL 공식 US$58.99 / AliExpress ₩234,154',
           shop='CNHL 공식몰', url='https://chinahobbyline.com/products/cnhl-gplus-series-5000mah-22-2v-6s-70c-lipo-battery-with-xt90-plug', alt='국내 RC점 (알씨뱅크 등) / Gens ace 6S 5000 60C', stage=2, st='확인필요', dom='해외',
           note='항공 배송 제한·배송비 미확인. 배터리 플러그 성별 확인 (커넥터 시트 XT90-S)'))
B.append(I(id='P-02', cat='전원', sub='배터리', name='밸런스 충전기', spec='AC 100W / DC 300W, 2~6S, XT60 출력', why='LiPo 충전 (AC만 쓰면 6S 약 4A)', where='-', calc='1', need=1, price=150741, ship=7428, shop='AliExpress ToolkitRC M7AC', url='https://ko.aliexpress.com/item/1005009325477156.html',
           alt='ISDT 608AC ₩85,125 + ₩7,139', stage=2, st='조사완료', dom='해외', note='충전기 XT60 ↔ 배터리 XT90 변환 케이블 필요 (P-04)'))
B.append(I(id='P-03', cat='전원', sub='배터리', name='LiPo 안전 가방', spec='215×155×115mm 이상 (6S 5000 수납)', why='충전·보관', where='-', calc='1', need=1, price=9215, shop='AliExpress', url='https://ko.aliexpress.com/item/1005012002096700.html', stage=2, st='조사완료', dom='해외'))
B.append(I(id='P-04', cat='전원', sub='배터리', name='충전 변환 케이블 XT60 → XT90', spec='XT60 수 ↔ XT90 암(배터리 반대 성별), 12AWG 10~15cm', why='충전기 출력 XT60, 배터리 XT90', where='-', calc='1', need=1, price=None, shop='수성RC / 국내 RC점', url='', alt='XT90 커넥터 + 12AWG로 자작', stage=2, st='가격 미조사'))
B.append(I(id='P-05', cat='전원', sub='배터리', name='LiPo 전압 경보기 (셀 체커)', spec='1~8S, 셀별 전압 표시 + 부저, 설정 3.5V/셀 이상', why='RC LiPo에는 BMS 없음. 소프트웨어 저전압 정지(약 21V)와 별도의 하드웨어 경보', where='배터리 밸런스 커넥터 (JST-XH 7P)', calc='1 + 예비 1', need=1, spare=1, price=None, shop='국내 RC점', url='', alt='', stage=2, st='가격 미조사', pri='권장'))
B.append(I(id='P-06', cat='전원', sub='배터리', name='배터리 고정 스트랩', spec='고무 미끄럼 방지 벨크로 20×300mm', why='트레이 안에서 배터리 흔들림 방지 (점프·낙하)', where='배터리 트레이', calc='2', need=2, spare=1, price=None, shop='국내 RC점', url='', stage=2, st='가격 미조사'))
B.append(I(id='P-07', cat='전원', sub='보호 회로', name='안티스파크 커넥터 XT90-S', spec='Amass XT90-S 수 + XT90 암 1조, 연속 90A', why='배터리 연결 시 돌입 스파크 방지', where='배터리 ↔ 로봇', calc='커넥터 시트', need=1, spare=1, price=3500, orig='2,500 + 1,000원', ship=3500, shop='수성RC', url='https://susungrc.com/product/detail.html?cate_no=86&display_group=1&product_no=1452', alt='AliExpress Amass 1쌍 ₩11,926', stage=3, st='조사완료'))
B.append(I(id='P-08', cat='전원', sub='보호 회로', name='DC 접촉기', spec='100A, 코일 24V (6S 18~25.2V에서 흡착·유지 확인), 1NO, 스터드 단자', why='E-stop이 코일을 끊어 주 전원 차단', where='전장 선반', calc='1', need=1, price=21071, ship=10792, shop='AliExpress CZW', url='https://ko.aliexpress.com/item/1005005457160871.html',
           alt='코일 9~36V 광범위형 접촉기 (예: Gigavac GX 계열 이코노마이저형, 가격 미조사)', stage=3, st='확인필요', dom='해외', note='CZW 24V 코일이 18V(방전 말기)에서 붙어 있는지 데이터시트 확인. 실패하면 광범위 코일형으로 교체'))
B.append(I(id='P-09', cat='전원', sub='보호 회로', name='플라이백 다이오드', spec='1N5408 (3A, 1000V) 또는 1N4007, 코일 단자 직결', why='코일 차단 시 역기전력으로 E-stop 접점 아크 방지', where='접촉기 코일 (K05)', calc='1 + 예비', need=1, spare=4, moq=5, price=None, shop='디바이스마트', url='', stage=3, st='가격 미조사', note='접촉기에 다이오드 내장이면 불필요'))
B.append(I(id='P-10', cat='전원', sub='보호 회로', name='비상 정지 스위치', spec='22mm 버섯형 래칭, 1NC 이상, 접점 정격 ≥ 코일 전류 × 2 (DC 30V)', why='접촉기 코일 차단', where='커버 윗면 (22mm 구멍)', calc='1', need=1, price=8010, ship=1950, shop='AliExpress 금속 E-stop', url='https://ko.aliexpress.com/item/1005009505038209.html', stage=3, st='확인필요', dom='해외', note='DC 접점 정격 표기 확인'))
B.append(I(id='P-11', cat='전원', sub='보호 회로', name='주 퓨즈 MIDI 60A + 홀더', spec='MIDI(AMI) 60A, DC 32V 이상 정격, 볼트형 홀더 (M5 스터드 추정)', why='주 전원 단락 보호. 60A < 접촉기 100A < XT90 90A', where='접촉기 뒤 (P02~P03)', calc='홀더 1 + 퓨즈 1 + 예비 퓨즈 1', need=1, spare=1, price=5560, shop='AliExpress ANS-H', url='https://ko.aliexpress.com/item/1005009732249061.html', stage=3, st='확인필요', dom='해외', note='정격 전압 표기·세트 구성 확인'))
B.append(I(id='P-12', cat='전원', sub='보호 회로', name='프리차지 저항', spec='알루미늄 하우징 권선 저항 47Ω 25W', why='E-stop 해제 전 드라이버 커패시터를 미리 충전해 접촉기 접점 용착 방지. 피크 0.54A, 순간 전력 13.5W (25.2²÷47, 계산)', where='섀시 (방열)', calc='1', need=1, price=None, shop='디바이스마트', url='', alt='mjbots power_dist (프리차지 내장, 약 US$149)', stage=3, st='가격 미조사', pri='권장', note='드라이버 총 커패시턴스 미확인 → 충전 시간 실측 (2~3초 누름)'))
B.append(I(id='P-13', cat='전원', sub='보호 회로', name='프리차지 순간 버튼', spec='16mm 또는 22mm 모멘터리 푸시버튼, 1NO, 접점 ≥ 1A DC 30V', why='P-12 직렬, 누르는 동안만 프리차지', where='커버 윗면 E-stop 옆', calc='1', need=1, price=None, shop='디바이스마트', url='', stage=3, st='가격 미조사', pri='권장'))
B.append(I(id='P-14', cat='전원', sub='보호 회로', name='TVS 다이오드', spec='5KP28A (단방향, 스탠드오프 28V > 만충 25.2V, 클램프 약 45V < GDS68 최대 56V) 또는 동급', why='접촉기 개방·회생 시 버스 과전압 흡수', where='분배 버스바 (P06)', calc='1 + 예비', need=1, spare=1, price=None, shop='디바이스마트 / 엘레파츠', url='', stage=3, st='가격 미조사', pri='권장', note='값은 계산 기반 추정'))
B.append(I(id='P-15', cat='전원', sub='분배', name='분배 버스바 (+/− 2개)', spec='4~6웨이 구리 버스바, M5 스터드 + M3/M4 단자, 연속 ≥ 60A, 개당 30g 이하', why='퓨즈 뒤 다리 4개 + DC-DC + 코일 − + CAN GND 분배', where='전장 선반', calc='+ 1, − 1', need=2, price=None,
           shop='국내 공구·전기몰', url='', alt='드론용 PDB (6S·정격 확인) / mjbots power_dist', stage=3, st='가격 미조사', note='단자 수: + 6 (다리 4, DC-DC, 여유 1), − 9 (주, 다리 4, DC-DC, 코일, CAN GND 2)'))
B.append(I(id='P-16', cat='전원', sub='5V', name='DC-DC 강압 24V→5V', spec='입력 6~50V(명시), 출력 5.1V 5.5A 이상, 역극성·과열 보호 (예: Pololu D36V50F5)', why='Pi 5 5V/5A. v0.1 RCNUN은 최대 입력 미표기 → 만충 25.2V 보장 제품으로 교체 제안', where='전장 선반', calc='1 + 예비 1', need=1, spare=1, price=None,
           shop='Pololu 리셀러 (디바이스마트 등)', url='', alt='v0.1 RCNUN 12/24V→5V 5A ₩6,975 (입력 범위 확인 시)', stage=2, st='가격 미조사', note='2단계 벤치는 27W 어댑터(1단계 시험 시트)로 대체 가능'))
B.append(I(id='P-17', cat='전원', sub='배선 보호', name='인라인 퓨즈 홀더·퓨즈 세트', spec='ATO 14AWG 홀더 4 + 25A 퓨즈 8 / 미니 블레이드 20AWG 홀더 1 + 2A 퓨즈 2 / 18AWG 홀더 1 + 3A 퓨즈 2', why='다리 분기·코일·DC-DC 각각 보호 (커넥터 시트)', where='배선표 B-xx, K01, V01', calc='커넥터 시트', need=1, price=None, shop='국내 자동차 전장몰', url='', stage=3, st='가격 미조사', pri='권장'))

B.append(('SEC', "배선 (전선 길이·커넥터는 '배선표'·'커넥터·단자' 시트에서 산출)"))
WSPEC = {12: '12AWG 실리콘선, 680가닥 이상 극세사, 200°C', 14: '14AWG 실리콘선, 400가닥', 16: '16AWG 실리콘선, 252가닥', 18: '18AWG 실리콘선', 20: '20AWG 실리콘선', 24: '24AWG 실리콘선 (외경 ≤1.4mm)'}
PAIR_PRICE = {12: (6000, '6,000원 (흑+적 1m 1조, v0.1 목록가)', 0), 14: (5600, '2,800원/1m × 2색 (v0.1 목록가)', 0), 18: (4510, '4,510원 (흑+적 1m, v0.1)', 4400)}
wid = 0
for awg in sorted({a for a, _ in WIRE_BUY}):
    cols_ = {c: WIRE_BUY[(a, c)] for a, c in WIRE_BUY if a == awg}
    rb = {c: cols_.pop(c) for c in ('적', '흑') if c in cols_}
    if rb:
        wid += 1; need = math.ceil(max(rb.values())); pp = PAIR_PRICE.get(awg)
        B.append(I(id=f'W-{wid:02d}', cat='배선', sub='전선', name=(f'실리콘선 {awg}AWG 적+흑 (1m 1조)' if len(rb) == 2 else f'실리콘선 {awg}AWG ' + list(rb)[0]), spec=WSPEC[awg], why='관절 통과부 굴곡 수명 (실리콘·극세사)', where='배선표',
                   calc='배선표 집계 ' + ', '.join(f'{c} {v} m' for c, v in rb.items()) + ' → 조 단위 올림', need=need, price=pp[0] if pp else None, orig=pp[1] if pp else '', ship=pp[2] if pp else 0,
                   shop='엑스캅터 / 팀에스아이(BUYRC) / 아트콥터', url='https://xcopter.com/category/실리콘-전선/6240', alt='AliExpress 2핀 실리콘 케이블', stage=3, st='조사완료' if pp else '가격 미조사'))
    for c, v in cols_.items():
        wid += 1
        B.append(I(id=f'W-{wid:02d}', cat='배선', sub='전선', name=f'실리콘선 {awg}AWG {c}', spec=WSPEC[awg] + (' — CANH/CANL 두 색을 드릴로 꼬아 사용' if c in ('황', '녹') else ''), why='CAN 꼬임 페어' if c in ('황', '녹') else '색 구분',
                   where='배선표', calc=f'배선표 집계 {c} → {v} m', need=math.ceil(v), price=None, shop='엑스캅터 / 아트콥터', url='', alt='CAN 전용 120Ω 꼬임 케이블 (굴곡성 확인)' if c in ('황', '녹') else '', stage=3, st='가격 미조사'))
B.append(I(id='W-10', cat='배선', sub='커넥터', name='XT60H 커넥터 수+암', spec='Amass XT60H, 60A, 12~14AWG 납땜', why='다리 분리 (커넥터 시트)', where='B-xx ↔ LH-01', calc='4 + 예비 1', need=4, spare=1, price=None, shop='디바이스마트 / 수성RC', url='', stage=3, st='가격 미조사'))
B.append(I(id='W-11', cat='배선', sub='커넥터', name='XT30(2+2) 커넥터 (모터 반대 성별)', spec='Amass XT30(2+2), 전원 2핀 16AWG + 신호 2핀 22~26AWG 납땜', why='모터 12개 전원+CAN', where='LH-02~04', calc='12 + 예비 2', need=12, spare=2, price=None,
           shop='디바이스마트', url='', alt='v0.1 완제품 연장 케이블 ₩4,905 (길이 고정이라 하네스에는 부적합)', stage=3, st='확인필요', note='1단계 모터 실물로 성별·핀 배치 확인 후 구매'))
B.append(I(id='W-12', cat='배선', sub='커넥터', name='JST-SM 2P 수+암', spec='JST-SM 2핀, 20~22AWG, 3A', why='커버 E-stop 분리 (K01~K03)', where='커버 ↔ 몸통', calc='1 + 예비 1', need=1, spare=1, price=None, shop='디바이스마트', url='', stage=3, st='가격 미조사'))
B.append(I(id='W-13', cat='배선', sub='단자', name='비절연 링 단자 세트', spec='동관 링 단자: 4mm² M5·M6, 2.5mm² M5, 1.5mm² M5·M6 (커넥터 시트 수량)', why='접촉기·퓨즈 홀더·버스바', where='P01~P05, B-xx, K, V', calc='커넥터 시트 합계', need=1, price=None, shop='국내 전기몰', url='', stage=3, st='가격 미조사', note='링 단자 압착기 필요 (공구 시트)'))
B.append(I(id='W-14', cat='배선', sub='단자', name='페룰 세트', spec='절연 페룰 0.25/0.5/0.75mm² 각 20개 이상', why='HAT·E-stop·DC-DC 나사 단자', where='커넥터 시트', calc='1세트', need=1, price=None, shop='국내 전기몰', url='', stage=3, st='가격 미조사'))
B.append(I(id='W-15', cat='배선', sub='단자', name='USB-C 전원 피그테일', spec='USB-C 수, 2선 (VBUS/GND) 20AWG, 5A 정격, 15cm', why='DC-DC → Pi 5 (V03)', where='전장', calc='1 + 예비 1', need=1, spare=1, price=None, shop='AliExpress / 국내', url='', stage=2, st='가격 미조사'))
B.append(I(id='W-16', cat='배선', sub='단자', name='점퍼 와이어 2.54 암-암', spec='15cm, 10가닥 묶음', why='BNO085 ↔ Pi GPIO (V04)', where='전장', calc='1묶음', need=1, price=None, shop='디바이스마트', url='', stage=3, st='가격 미조사'))
B.append(I(id='W-17', cat='배선', sub='보호재', name='열수축 튜브 세트 (일반)', spec='2:1 폴리올레핀 1~10mm 280개', why='납땜부 절연', where='공통', calc='1', need=1, price=8000, shop='디바이스마트', url='https://www.devicemart.co.kr/goods/view?no=1320937', alt='쿠팡 164개 3,800원', stage=1, st='조사완료'))
B.append(I(id='W-18', cat='배선', sub='보호재', name='접착식 열수축 3:1', spec='Ø6 1m + Ø10 1m + Ø12 1m (접착 내장)', why='12AWG 링 단자·다리 스플라이스 방수·인장 보강', where='P01~P04, LH-01', calc='각 1m', need=3, price=None, shop='디바이스마트 / 국내 전기몰', url='', stage=3, st='가격 미조사'))
B.append(I(id='W-19', cat='배선', sub='보호재', name='PET 편조 슬리브', spec='Ø6 확장형 3m + Ø10 1m', why='다리 하네스(전원 2 + CAN 2) 외전 관절 통과부 마모 방지', where='LH-03·04, B-xx', calc='다리당 0.6m × 4 = 2.4 → 3m', need=2, price=None, moq=1, shop='국내 전기몰', url='', stage=3, st='가격 미조사', note='수량 2 = Ø6 3m 1, Ø10 1m 1'))
B.append(I(id='W-20', cat='배선', sub='보호재', name='케이블 타이 + 접착 마운트', spec='나일론 케이블 타이 2.5×100mm 100개 + 접착 타이 마운트 20×20 20개', why='하네스 고정', where='공통', calc='1세트', need=1, price=None, shop='쿠팡 / 디바이스마트', url='', stage=3, st='가격 미조사'))

B.append(('SEC', '2단계 다리 지그 (CAD 없음 → 구성은 추정)'))
B.append(I(id='J-01', cat='지그', sub='프레임', name='알루미늄 프로파일 2020', spec='2020 V슬롯/T슬롯, 500mm × 4 (기둥·바닥) + 300mm × 4 (가로)', why='다리 1개 스쿼트·낙하 시험용 수직 프레임 (보고서 개발 순서 ②)', where='지그', calc='8본', need=8, price=None, shop='국내 프로파일몰 (절단 주문)', url='', stage=2, st='가격 미조사', note='높이 500 = 서 있음 280 + 낙하 여유 200 (추정)'))
B.append(I(id='J-02', cat='지그', sub='프레임', name='2020 코너 브래킷 + T너트 + 볼트', spec='코너 브래킷 12, M5 T너트 40, M5×10 볼트 40', why='J-01 조립', where='지그', calc='1세트', need=1, price=None, shop='국내 프로파일몰', url='', stage=2, st='가격 미조사'))
B.append(I(id='J-03', cat='지그', sub='수직 슬라이드', name='MGN12H 리니어 가이드 400mm', spec='MGN12 레일 400mm + MGN12H 캐리지 1', why='고관절 마운트를 수직으로만 움직이게 (낙하 시험)', where='지그 기둥', calc='1', need=1, price=None, shop='AliExpress / 국내', url='', stage=2, st='가격 미조사'))
B.append(I(id='J-04', cat='지그', sub='수직 슬라이드', name='슬라이드 캐리지판', spec='Al 5052 3mm (판금 외주에 추가) 또는 PETG 출력, 외전 마운트 PCD74 구멍', why='외전 모터(또는 외전 고정 블록)를 캐리지에 체결', where='지그', calc='1', need=1, price=None, shop='판금 외주 (S-01과 함께)', url='', stage=2, st='견적 필요'))
B.append(I(id='J-05', cat='지그', sub='안전', name='낙하 스토퍼·완충 패드', spec='EVA 폼 20mm 300×300 + 프로파일 스토퍼 블록 2', why='낙하 시험 시 하한 제한', where='지그 바닥', calc='1식', need=1, price=None, shop='국내', url='', stage=2, st='가격 미조사'))

bom_sheet('BOM 전체', B)

# ---------------------------------------------------------------- 1단계 시험
T = []
T.append(('SEC', 'v0.1 유지 (사용자 요청: USB-CAN·XT30(2+2) 수-암 케이블·무게 센서)'))
T.append(I(id='T-01', cat='시험', sub='통신', name='USB-CAN (절연)', spec='CANable Pro v2.0 (CAN FD, 절연)', why='PC ↔ 모터 CAN 시험. 24V 버스라 절연형', where='벤치', calc='1', need=1, price=None, ship=3500, shop='vctec(가치창조기술)',
           url='https://vctec.co.kr/product/usb-can-통신-canable-pro-v20-can-fd-절연-canable-pro-v20-usb-to-can-adapte/22335', alt='AliExpress Makerbase CANable 2.0 ₩34,026', stage=1, st='유지(v0.1)', note='v0.1에서 단가 칸 비어 있음 → 구매 전 가격 확인'))
T.append(I(id='T-02', cat='시험', sub='배선', name='시험용 XT30(2+2) 케이블', spec='XT30(2+2) 수-암 연장', why='1단계 모터 2개 연결', where='벤치', calc='2', need=2, price=4905, shop='AliExpress', url='https://ko.aliexpress.com/item/1005012775760038.html', stage=1, st='유지(v0.1)', dom='해외', note='GDS68 커넥터 형식 확인 후 구매'))
T.append(I(id='T-03', cat='시험', sub='토크', name='S형 로드셀 + HX711', spec='S형 로드셀 0~50kg + HX711 모듈', why='200mm 암 토크 측정 (11 N·m ÷ 0.2 m = 55 N)', where='토크 지그', calc='1', need=1, price=46300, shop='AliExpress S형 로드셀', url='https://ko.aliexpress.com/item/1005008003972994.html', stage=1, st='유지(v0.1)', dom='해외', note='20kg 옵션이 분해능에 유리'))
T.append(('SEC', '전원'))
T.append(I(id='T-04', cat='시험', sub='전원', name='벤치 전원 공급기', spec='CV/CC, 0~30V, ≥10A, 출력 ON/OFF 버튼, OVP. 24.0V 설정', why='전류 제한(CC)으로 첫 전원 투입 보호. 스톨 24A(상전류) 시험 시 버스 전류 ≈ 1.5×24²×0.22Ω÷24V ≈ 7.9A (보고서 상저항 0.22Ω 가정, 계산) → 10A로 가능',
           where='벤치', calc='1', need=1, price=278261, shop='AliExpress Hanmatek HM310T', url='https://ko.aliexpress.com/item/1005004833435425.html', alt='국내 30V 10A 리니어/SMPS 벤치 전원 (가격 미조사)', stage=1, st='확인필요', dom='해외', note='11/12 ₩127,800 → 11/20 ₩278,261 급등. 국내 대체품 우선 조사'))
T.append(I(id='T-05', cat='시험', sub='전원', name='회생 대책: 브레이크 저항', spec='알루미늄 하우징 권선 저항 2Ω 50W (ODrive 계열 표준값, 추정) + 연결선 18AWG', why='벤치 전원은 회생 전류를 못 받음 → GDS68 브레이크 저항 단자 또는 dc_max_negative_current 설정 (매뉴얼 p.38~39)', where='GDS68 브레이크 단자', calc='모터 1개씩 시험 → 1 + 예비 1', need=1, spare=1, price=None, shop='디바이스마트', url='', stage=1, st='확인필요', note='매뉴얼 권장 저항값 확인 후 구매'))
T.append(I(id='T-06', cat='시험', sub='전원', name='벤치 전원 → XT30(2+2) 피그테일', spec='XT30(2+2) 1개(모터 반대 성별) + 16AWG 적/흑 각 50cm + 4mm 바나나 플러그 적/흑 + CAN 신호선 24AWG 꼬임 50cm (끝 페룰)', why='T-02 연장 케이블과 벤치 전원·CANable을 잇는 브레이크아웃', where='벤치', calc='모터 2개 동시 → 2', need=2, price=None, shop='디바이스마트 (부품) / 자작', url='', stage=1, st='가격 미조사'))
T.append(I(id='T-07', cat='시험', sub='전원', name='벤치 인라인 퓨즈', spec='ATO 인라인 홀더 16AWG + 15A 퓨즈 3', why='배선 단락 보호 (벤치 전원 CC와 이중)', where='T-06 + 선', calc='2', need=2, price=None, shop='국내 자동차 전장몰', url='', stage=1, st='가격 미조사'))
T.append(('SEC', '토크·백래시·열 측정'))
T.append(I(id='T-08', cat='시험', sub='토크', name='토크 암 (알루미늄)', spec='Al 6061/5052 평철 3×20×250mm, 출력 PCD25 6×Ø3.4 + 3×Ø4.1 + 끝 Ø4 고리 구멍 (중심에서 200.0mm)', why='PETG보다 휨·크리프 적어 레버 길이 정확. 판금 외주 S-01에 1장 추가 가능', where='모터 출력', calc='1', need=1, price=None, shop='판금 외주 (S-01과 함께) 또는 평철 + 드릴', url='', stage=1, st='견적 필요', note='M3×8 12.9 6개는 F 행에서'))
T.append(I(id='T-09', cat='시험', sub='토크', name='HX711 판독용 마이컴', spec='Arduino Nano 호환(USB-C) 또는 ESP32 DevKit + USB 케이블 + 점퍼 10가닥', why='HX711은 MCU가 있어야 PC로 값 전송 (v0.1 누락)', where='벤치', calc='1', need=1, price=None, shop='디바이스마트', url='', stage=1, st='가격 미조사'))
T.append(I(id='T-10', cat='시험', sub='토크', name='로드셀 교정 분동', spec='2kg (또는 1kg×2), 등급 M1', why='로드셀 교정 (55 N ≈ 5.6kg 범위 내 2점 교정)', where='벤치', calc='1', need=1, price=None, shop='국내 계측몰', url='', alt='무게 잰 물병 (정확도 낮음)', stage=1, st='가격 미조사'))
T.append(I(id='T-11', cat='시험', sub='토크', name='로드셀 연결 하드웨어', spec='로드셀 나사(M6 또는 M8, 실물 확인) 아이볼트 2 + 와이어 로프 Ø1.5 1m + 와이어 클립 4', why='토크 암 끝 ↔ 로드셀 ↔ 고정점', where='토크 지그', calc='1세트', need=1, price=None, shop='국내 철물', url='', stage=1, st='가격 미조사'))
T.append(I(id='T-12', cat='시험', sub='고정', name='모터 고정 지그', spec='Al 3mm 판 (뒷면 PCD74 8×Ø3.4, 중앙 Ø62 개방) + 2020 프로파일 또는 PETG 블록, C클램프 100mm × 2', why='스톨 11 N·m 반력 고정', where='작업대', calc='1', need=1, price=None, shop='판금 외주 + 철물점', url='', stage=1, st='가격 미조사'))
T.append(I(id='T-13', cat='시험', sub='백래시', name='다이얼 게이지 + 마그네틱 베이스', spec='0.01mm 눈금, 측정 범위 10mm', why='백래시 15 arcmin → 반경 50mm에서 0.218mm (50×tan 0.25°, 계산) → 0.01mm 분해능 필요', where='토크 암 위', calc='1', need=1, price=None, shop='국내 공구몰', url='', stage=1, st='가격 미조사'))
T.append(I(id='T-14', cat='시험', sub='열', name='온도 측정 멀티미터 + K형 열전대', spec='DC 10A, 저항 200Ω 레인지, 온도(K형) 측정 기능 + 비드형 K형 열전대 1~2개 + 캡톤 테이프', why='5 N·m 30분 열 시험을 연속 기록. CAN 종단 60Ω 확인도 겸함 (v0.1 DT-830B·IR 온도계 대체)', where='모터 하우징', calc='1', need=1, price=None, shop='국내 계측몰', url='', alt='v0.1: DT-830B + 비접촉 IR 온도계 (하우징 반사율 오차)', stage=1, st='가격 미조사'))
T.append(I(id='T-15', cat='시험', sub='치수', name='디지털 캘리퍼스 (금속)', spec='150mm, 0.01mm, 스테인리스', why='플랜지·공차·판금 실측 (저가 플라스틱형은 정밀 부족)', where='-', calc='1', need=1, price=None, ship=2700, shop='디바이스마트 (블루텍 금속형)', url='https://m.devicemart.co.kr/goods/view/12333553', alt='Mitutoyo 500-181-30 135,630원', stage=1, st='가격 미조사'))
T.append(('SEC', '2단계 벤치 (Pi 5)'))
T.append(I(id='T-16', cat='시험', sub='전원', name='Pi 5 27W 전원 어댑터', spec='공식 27W USB-C (5.1V/5A PD)', why='벤치에서 Pi 5 구동 (배터리·DC-DC 없이)', where='-', calc='1', need=1, price=None, shop='디바이스마트', url='https://www.devicemart.co.kr/goods/view?no=15502402', alt='아이씨뱅큐 22,000원(VAT 별도)', stage=2, st='가격 미조사'))
bom_sheet('1단계 시험', T)

# ---------------------------------------------------------------- 공구
tl = wb.create_sheet('공구')
tl.cell(row=1, column=1, value='공구 (보유한 것은 수량 0으로)').font = Font(bold=True, size=13)
tl.cell(row=2, column=12, value='합계').font = BOLD
TCOLS = ['번호', '공구', '사양', '필요한 작업', '관련 행', '수량', '단가 (원)', '소계', '구매처', '상태', '우선순위', '', '']
header(tl, 3, TCOLS[:11], [5, 26, 44, 40, 18, 6, 10, 10, 18, 11, 8])
TOOLS = [
 ('납땜 인두 (온도 조절, 80W 이상)', '팁 교체형, 450°C, 큰 끌팁 포함', 'XT90·XT60·12AWG 납땜 (작은 인두는 열 부족)', 'P01, P04, W-10'),
 ('열융착 인서트 팁', 'M3 (필요시 M4), 인두 모델에 맞는 규격', 'PETG 인서트 압입', 'F-13'),
 ('와이어 스트리퍼 (자동형)', '10~24AWG', '전선 피복 제거', '배선표 전체'),
 ('링 단자 압착기 (래칫)', '비절연 동관 1.5~6mm² (육각 압착)', '12/14AWG 링 단자', 'W-13'),
 ('페룰 압착기', '0.25~6mm² 사각 압착', '나사 단자 페룰', 'W-14'),
 ('열풍기', '접착 열수축 수축용 (300°C 이상)', '열수축·접착 열수축', 'W-17, W-18'),
 ('육각 렌치 세트 (볼엔드)', '1.5/2/2.5/3/4mm', 'M2.5~M5 렌치볼트', '체결 상세'),
 ('토크 드라이버', '0.4~2.0 N·m, 육각 2.5mm 비트', '모터 M3 1.2 N·m (알루미늄 나사산 보호)', '체결 상세 규칙 3'),
 ('전동 드릴 + HSS 비트', 'Ø3.4, Ø4.1, Ø4.5 + 센터 펀치 + 바이스', '30T 풀리 측면 구멍, 토크 암 가공, CAN 선 꼬기', 'D-02, T-08'),
 ('리벳너트 공구', 'M4 대응 핸드 리베터', 'M4 리벳너트 (커버 손나사)', 'F-14'),
 ('축용 스냅링 플라이어', 'Ø10~25 축용', 'Ø12 스냅링 (Ø30은 큰 플라이어)', 'D-09, D-10'),
 ('줄·디버링 툴', '평줄 + 디버러', '판금 버 제거, 장공 다듬기', '판금'),
 ('보안경 + 내열 매트', '', '납땜·드릴·LiPo 작업', '공통'),
]
for i, t in enumerate(TOOLS, 4):
    put(tl, i, [i - 3, t[0], t[1], t[2], t[3], 1, None, f'=IF(G{i}="",0,F{i}*G{i})', '', '가격 미조사', '필수'])
tl.cell(row=2, column=13, value=f'=SUM(H4:H{3 + len(TOOLS)})').font = BOLD

# ---------------------------------------------------------------- 출력물
pr = wb.create_sheet('출력물')
header(pr, 1, ['번호', '부품', '재료', '수량', 'CAD 질량 (g/개)', '필라멘트 (g, ×1.15)', '출력 방향·설정', '근거'], [5, 34, 8, 6, 12, 14, 50, 40])
PRINTS = [
 ('하퇴 FL_knee_link', 'PETG', 4, 103.151, '3mm 중공벽, 1.2mm 표면층 + 코어 30~40% 인필, 길이 방향 XY', 'C5_20261118-130325'),
 ('발 패드 FL_foot', 'TPU', 4, 13.798, 'Ø30×16, 인필 30~50%', 'C5'),
 ('배터리 트레이', 'PETG', 1, 68.640, '바닥 XY', 'D 로그'),
 ('PETG 브래킷 (래치 2·레지 4·서포트 4)', 'PETG', 1, 40.100, '래치는 길이 방향 XY (굽힘 방향)', 'D 로그 (합계 질량)'),
 ('접근 덮개 access lid', 'PETG', 1, 50.089, '1.2mm 판, XY', 'D 로그'),
 ('손나사 노브 (F-15 대체 시)', 'PETG', 4, 1.48, '', 'D 로그 5.92g/4'),
 ('전장 캐리어·E-stop 리브 등 (E 단계)', 'PETG', 1, 100, '', '추정 (E 단계 질량 로그 없음)'),
 ('케이블 클립 (외전 관절 통과부)', 'PETG', 8, 2, '', '추정'),
 ('시험 지그 (모터 고정 블록 등)', 'PETG', 1, 150, '', '추정'),
 ('다리 지그 부품 (J-04 PETG 선택 시 포함)', 'PETG', 1, 150, '', '추정'),
]
for i, p in enumerate(PRINTS, 2):
    put(pr, i, [i - 1, p[0], p[1], p[2], p[3], f'=ROUND(D{i}*E{i}*1.15,0)', p[4], p[5]])
n = len(PRINTS) + 2
put(pr, n, ['', 'PETG 합계 (실패 여유 ×1.3 → BOM S-03)', 'PETG', '', '', f'=ROUND(SUMIF(C2:C{n-1},"PETG",F2:F{n-1})*1.3,0)', '', ''])
put(pr, n + 1, ['', 'TPU 합계 (×1.3 → BOM S-04)', 'TPU', '', '', f'=ROUND(SUMIF(C2:C{n-1},"TPU",F2:F{n-1})*1.3,0)', '', ''])

# ---------------------------------------------------------------- 판금 외주 (v0.1 + 두께·절곡 열)
sm = wb.create_sheet('판금 외주')
header(sm, 1, ['부품명 (CAD 이름)', '재료·두께', '수량 (1대)', '2단계 (다리 1세트)', '절곡', '필요 파일', '구분', '근거·비고'], [40, 26, 10, 12, 10, 18, 8, 60])
SM = [
 ('섀시 AZN_D_chassis', '5052-H32 3mm', 1, 0, '있음', 'DXF + STEP', '몸통', 'D 로그 646.96g → 경량화 898.5g(몸통 합). 측벽 254×75 포켓'),
 ('외전 마운트 AZN_D_*_abad_mount', '5052-H32 3mm', 4, 1, '있음', 'DXF + STEP', '몸통', 'PCD74 8×Ø3.4, 위상 22.38°. 2단계는 지그 장착용 1장'),
 ('커버 AZN_D_cover', '5052-H32 3mm', 1, 0, '있음', 'DXF + STEP', '몸통', '개구 190×150, E-stop Ø22 + 프리차지 버튼 구멍 추가 필요'),
 ('전장 선반 AZN_D_electronics', '5052-H32 3mm', 2, 0, '있음', 'DXF + STEP', '몸통', ''),
 ('외전 U형 프레임 AZN_C6_U_frame', '5052-H32 2mm', 4, 1, '있음', 'DXF + STEP', '다리', '판금 전개 HOLD'),
 ('상퇴 안판 AZN_C3_inner_plate', '5052-H32 1.5mm', 4, 1, '없음', 'DXF', '다리', '출력 3-Ø4 핀 구멍 Ø4.1 확인'),
 ('상퇴 바깥판 AZN_C3_outer_plate', '5052-H32 1.5mm', 4, 1, '없음', 'DXF', '다리', '6806 시트'),
 ('상퇴 횡연결 상·하', '5052-H32 1.5~3mm (CAD 확인)', 8, 2, '있음', 'DXF', '다리', ''),
 ('6806 리테이너 내·외', '5052-H32 3mm (추정)', 8, 2, '없음', 'DXF', '다리', ''),
 ('고관절·무릎 모터 장착판', '5052-H32 3mm (추정)', 8, 2, '없음', 'DXF', '다리', '볼트 길이 M3×8 전제 (체결 상세)'),
 ('라이저·탑 브리지', '5052-H32 3mm (추정)', 12, 3, '있음', 'DXF', '다리', 'F 단계 extension 간섭으로 경로 수정 중 (DEF 로그)'),
 ('스핀들 플랜지', '5052-H32 3mm (추정)', 4, 1, '없음', 'DXF', '다리', ''),
 ('아이들러 지지판 내·외', '5052-H32 2~3mm (추정), 장공', 8, 2, '없음', 'DXF', '다리', ''),
 ('출력·외전 스페이서', '5052-H32 (두께 확인)', 8, 2, '없음', 'DXF 또는 출력', '다리', '두께가 볼트 길이를 정함 (체결 상세 4·5행)'),
 ('[시험] 토크 암', '5052-H32 3mm', 0, 1, '없음', 'DXF', '시험', '1단계 시험 T-08'),
 ('[시험] 모터 고정판', '5052-H32 3mm', 0, 1, '없음', 'DXF', '시험', 'T-12'),
 ('[지그] 슬라이드 캐리지판', '5052-H32 3mm', 0, 1, '없음', 'DXF', '지그', 'J-04'),
]
for i, s in enumerate(SM, 2):
    put(sm, i, list(s))
n = len(SM) + 2
put(sm, n, ['합계', '', f'=SUM(C2:C{n-1})', f'=SUM(D2:D{n-1})', '', '', '', ''])
put(sm, n + 2, ['업체 후보: MakeIT(국내, 1개부터, STEP/DWG), CutON(DXF 즉시 견적), 레이저1번가, JLCCNC(해외, 5052 0.5~6mm). 공시 단가 없음 → 견적. 2단계 세트를 먼저 주문해 절곡 R·K값(0.44 추정)·공차 확인 후 3단계 주문.', '', '', '', '', '', '', ''])
sm.merge_cells(start_row=n + 2, start_column=1, end_row=n + 2, end_column=8); sm.row_dimensions[n + 2].height = 45

# ---------------------------------------------------------------- 확인 필요
q = wb.create_sheet('확인 필요')
header(q, 1, ['번호', '항목', '내용', '막는 품목', '확인 방법', '시점'], [5, 26, 70, 22, 40, 12])
QS = [
 ('GDS68 커넥터 형식·성별', 'XT30(2+2)로 가정. 판매 페이지에 언급 없음', 'W-11, T-02, T-06', '1단계 모터 실물', '1단계 수령'),
 ('GDS68 CAN 포트 수', '통과형 2포트면 T 납땜 대신 모터→모터 데이지 체인', '배선표 CAN', '실물 + 매뉴얼', '1단계 수령'),
 ('enable_r120 기본값', '10개는 OFF, ID6·ID12만 ON. 출고 기본값 미확인', '배선표 CAN', 'odrivetool 조회', '1단계'),
 ('브레이크 저항 권장값', '매뉴얼 p.38~39 확인 (2Ω 50W 추정)', 'T-05', '매뉴얼', '1단계 전'),
 ('GT3 풀리 치형', '"5GT" 표기 풀리 대부분 HTD5M 치형', 'D-01, D-02', '판매자 견적·사진 (치형 단면)', '2단계 전'),
 ('구동축 ↔ 무릎 모터 결합 방식', 'Fusion 로그에 허브/플랜지 상세 없음 → 볼트 길이·축 길이 미정', 'D-06, 체결 상세 6행', 'Fusion에서 단면 확인', '2단계 전'),
 ('Pi 5 쿨러 + HAT 높이', '액티브 쿨러 위로 HAT이 들어가는지, 스태킹 헤더 길이', 'E-02, E-03', '부품 치수', '2단계'),
 ('백아이들러 Ø45', '기성품 없음 → 원판 외주 또는 설계 변경', 'D-04, D-05', '사용자 결정', '2단계 전'),
 ('접촉기 코일·스터드', 'CZW 24V 코일의 흡착/유지 전압, 코일 전류, 스터드 크기(M5/M6/M8)', 'P-08, W-13', '데이터시트', '3단계 전'),
 ('배터리 플러그 성별', 'XT90-S 어느 쪽을 로봇에 다는지', 'P-07, P-04', '배터리 실물', '2단계'),
 ('BNO085 연결 방식', 'Pi에서 I2C 클럭 스트레칭 문제 보고 → UART-RVC 권장 (추정)', 'V04, E-06', 'Adafruit 문서', '3단계'),
 ('축 중공 + 홈 동시 주문', '미스미에서 OD12/ID9.5 중공 + 스냅링 홈 + D컷 가능한지', 'D-06, D-07', '미스미 견적', '2단계 전'),
 ('스페이서·판 두께', '볼트 길이 계산 입력값 (체결 상세 4~6행)', '체결 상세', 'Fusion 실측', '2단계 전'),
 ('하네스 3D 길이', '배선표 길이는 꺾인 직선 추정', '배선표, W 행', 'Fusion E 단계 경로 측정', '3단계 전'),
 ('다리 분기 퓨즈 25A', '점프 순간 다리 피크 미확인 → 오동작 가능', 'P-17', '2단계 지그 전류 측정', '2단계'),
 ('프리차지 필요성', '드라이버 커패시턴스 미확인. XT90-S가 플러그 때만 보호', 'P-12, P-13', '접촉기 투입 돌입 전류 측정', '3단계'),
 ('DC-DC 제품', 'Pololu D36V50F5 등 입력 범위 명시품 국내 가격', 'P-16', '가격 조사', '2단계'),
 ('v0.1 남은 항목', '분할결제 면세(세관 125), 품목별 관세율, Pi 5 납기 2주, 가격 급등(충전기·HAT·벤치 전원)', 'A, P-02, E-05, T-04', 'v0.1 확인 필요 시트', '구매 전'),
]
for i, s in enumerate(QS, 2):
    put(q, i, [i - 1] + list(s))

order = ['표지', 'BOM 전체', '1단계 시험', '배선표', '커넥터·단자', '체결 상세', '출력물', '판금 외주', '공구', '확인 필요']
wb._sheets = [wb[n] for n in order]
wb.save(OUT)
print('saved', OUT, 'BOM items', sum(isinstance(b, dict) for b in B), 'test', sum(isinstance(t, dict) for t in T), 'wires', len(W))
print('fastener agg', dict(agg)); print('nyloc', NYLOC_M3, NYLOC_M4, 'wash', WASH_M3, WASH_M4, 'insert', INSERT_M3)
print('wire buy', WIRE_BUY)
