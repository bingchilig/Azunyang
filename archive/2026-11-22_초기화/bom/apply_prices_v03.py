# -*- coding: utf-8 -*-
"""BOM v0.2 + 가격 조사(2026-11-21) → BOM v0.3"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pylib'))
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC, OUT = sys.argv[1], sys.argv[2]
V = 1.1  # VAT 별도 → 포함
def v(x): return round(x * V)

# ID: dict(p=단가(VAT 포함, 구매 단위당), s=배송비, shop, name=페이지 상품명, opt=옵션, url, eta=도착, st=상태, note, need/spare/moq 덮어쓰기, alt=다른 후보)
C = {
 'A-01': dict(p=147600, s=7714 * 2, shop='AliExpress steadywin Store', name='스테디윈 GIM6010-8 5N.M 기어 모터 24V 8:1 감속기', opt='motor-GDS68-DE', url='https://ko.aliexpress.com/item/1005006339337286.html', eta='11/25~12/01', st='도착 위험', note='12/1 도착 표시 → 11/30 초과 위험. 1개씩 분할 결제', alt='ZennixTek US$103 (도착일 불명, 비공식)'),
 'A-02': dict(p=147600, s=7714, shop='AliExpress steadywin Store', name='스테디윈 GIM6010-8 5N.M 기어 모터 24V 8:1 감속기', opt='motor-GDS68-DE', url='https://ko.aliexpress.com/item/1005006339337286.html', eta='11/25~12/01', st='도착 위험', note='', alt=''),
 'A-03': dict(p=147600, s=7714 * 11, shop='AliExpress steadywin Store', name='스테디윈 GIM6010-8 5N.M 기어 모터 24V 8:1 감속기', opt='motor-GDS68-DE', url='https://ko.aliexpress.com/item/1005006339337286.html', eta='11/25~12/01', st='도착 위험', note='', alt=''),
 'D-01': dict(p=v(34298), s=0, shop='미스미코리아', name='하이토크 타이밍풀리 5GT형', opt='GPA20GT5150-A-P12', url='https://kr.misumi-ec.com/vona2/detail/110302193650/?HissuCode=GPA20GT5150-A-P12', eta='참고 출하 12/01', st='도착 위험', note='GT 치형 정품계. 15mm용·보어12·M4 무두볼트 2. 출하 12/1이라 11/30 초과 위험', alt='AliExpress 맞춤 ₩13,950 (치형 지정 필요, 비공식)'),
 'D-02': dict(p=v(45617), s=0, shop='미스미코리아', name='하이토크 타이밍풀리 5GT형', opt='GPA30GT5150-A-P12', url='https://kr.misumi-ec.com/vona2/detail/110302193650/?HissuCode=GPA30GT5150-A-P12', eta='참고 출하 12/01', st='도착 위험', note='Ø3.4 구멍 4개는 직접 가공', alt='AliExpress 맞춤 ₩13,950'),
 'D-03': dict(p=v(14179), s=0, shop='미스미코리아', name='타이밍 벨트 EV5GT 타입', opt='GBN530EV5GT-150', url='https://kr.misumi-ec.com/vona2/detail/110302652240/?HissuCode=GBN530EV5GT-150', eta='참고 출하 11/26', st='조사완료', note='530mm 106T 15mm 유리섬유. 5~12개 단가. Gates 정품은 못 찾음', alt='AliExpress Ruiyong ₩11,400 + ₩1,570'),
 'D-04': dict(p=v(8813), s=0, shop='미스미코리아', name='아이들러 센터 베어링 타입', opt='AFDF19-45', url='https://kr.misumi-ec.com/vona2/detail/110302651810/?HissuCode=AFDF19-45', eta='참고 출하 11/24', st='결정 필요', note='OD45·플랜지 OK, 그러나 보어 Ø10(6200ZZ)·벨트폭 12.7 → Ø5 핀 설계와 다름. 채택하면 아이들러 핀을 Ø10으로 변경', alt='AliExpress TPBSN45-20 ₩15,950/4개 (내경 미확인)'),
 'D-05': dict(p=1890, s=0, shop='원일베어링', name='625ZZ 철커버 일본 NSK 소형 볼 베어링 5x16x5', opt='단일', url='https://bearingshop.co.kr/product/625zz-철커버-일본-nsk-소형-볼-베어링-5x16x5/26987/', eta='도착일 불명', st='결정 필요', note='D-04가 Ø10 보어품이면 불필요', alt='NSK 국산 950원'),
 'D-06': dict(p=v(5610), s=0, shop='미스미코리아', name='회전축 고정링홈', opt='PSFRR12-45-B5-S5 (+D컷 FC 추가가공)', url='https://kr.misumi-ec.com/vona2/detail/110300087010/?HissuCode=PSFRR12-45-B5-S5', eta='참고 출하 11/25', st='조사완료', note='S45C 상당 g6, 양끝 스냅링 홈. D컷 FC 추가 시 단가·출하 재확인. 솔리드(중공 아님)', alt='MakeIT 견적 (6영업일)'),
 'D-07': dict(p=v(5610), s=0, shop='미스미코리아', name='회전축 고정링홈', opt='PSFRR12-50-B5-S5 (+D컷 FC)', url='https://kr.misumi-ec.com/vona2/detail/110300087010/?HissuCode=PSFRR12-50-B5-S5', eta='참고 출하 11/25', st='조사완료', note='D-06과 같음', alt='미스미 meviy'),
 'D-08': dict(p=None, s=0, shop='미스미 meviy (3D 업로드 자동 견적)', name='', opt='', url='https://meviy.misumi-ec.com/ko-kr/', eta='최단 1일째 출하 표시', st='견적 필요', note='STEP 올려 견적. 대안 MakeIT CNC (6영업일)', alt='https://makeit.ai.kr/services/cnc'),
 'D-09': dict(p=v(775), s=0, shop='미스미코리아', name='C형상 스냅링(축용)', opt='STWN12', url='https://kr.misumi-ec.com/vona2/detail/110300258330/?HissuCode=STWN12', eta='참고 출하 11/24', st='조사완료', note='JIS B 2804 (DIN 471과 홈 치수 확인). D-06 홈과 같은 미스미 규격이라 맞음', alt=''),
 'D-10': dict(p=v(2174), s=0, shop='미스미코리아', name='C형상 스냅링(축용)', opt='STWN30', url='https://kr.misumi-ec.com/vona2/detail/110300258330/?HissuCode=STWN30', eta='참고 출하 11/24', st='확인필요', note='리테이너 판으로 대체되면 불필요', alt=''),
 'BR-01': dict(p=6160, s=0, shop='원일베어링', name='일본산 NSK 볼 베어링 6806 6806ZZ 6806DD 6806VV', opt='철커버형(ZZ)', url='https://bearingshop.co.kr/product/일본산-nsk-볼-베어링-6806-6806zz-6806dd-6806vv/34320/', eta='도착일 불명', st='조사완료', note='원일베어링 합계 10만 원 이상 → 무료배송', alt=''),
 'BR-02': dict(p=4050, s=0, shop='원일베어링', name='일제 NSK 볼 베어링 6001 6001ZZ 6001DD', opt='6001ZZ(철커버형)', url='https://bearingshop.co.kr/product/일제-nsk-볼-베어링-6001-6001zz-6001dd/12145/', eta='도착일 불명', st='조사완료', note='', alt='KOYO 3,910원'),
 'BR-03': dict(p=4130, s=3500, shop='원일베어링', name='일본산 NSK 볼 베어링 6801 6801ZZ 6801DD 6801VV', opt='철커버형(ZZ)', url='https://bearingshop.co.kr/product/일본산-nsk-볼-베어링-6801-6801zz-6801dd-6801vv/34315/', eta='도착일 불명', st='조사완료', note='낱개 구매로 변경 (10개). 배송비는 무료 기준 확인 전 1회 반영', moq=1, alt=''),
 'S-03': dict(p=15000, s=0, shop='Bambu Lab KR 공식', name='PETG Basic', opt='블랙 (30105) / 리필 / 1kg', url='https://kr.store.bambulab.com/kr/products/petg-basic', eta='당일 발송, 1~3영업일', st='확인필요', note='블랙 품절 가능 → 다른 색 허용 여부', alt=''),
 'S-04': dict(p=14500, s=0, shop='AliExpress ERYONE Store', name='ERYONE TPU 95A 3D 프린터 필라멘트 1.75mm, 유연한 95A 소프트 고탄성 500g 스풀', opt='black', url='https://ko.aliexpress.com/item/1005012007471380.html', eta='11/26~12/01', st='도착 위험', note='', alt=''),
 'F-01': dict(p=v(50), s=0, shop='미스미코리아', name='육각 홀 붙이 볼트【금형용 표준부품】', opt='CB3-8', url='https://kr.misumi-ec.com/vona2/detail/110100141750/?HissuCode=CB3-8', eta='참고 출하 11/24', st='확인필요', note='100개 이상 50원(별도). 강도 12.9·흑색 표기 페이지에서 확인', alt=''),
 'F-02': dict(p=v(50), s=0, shop='미스미코리아', name='육각 홀 붙이 볼트【금형용 표준부품】', opt='CB3-10', url='https://kr.misumi-ec.com/vona2/detail/110100141750/?HissuCode=CB3-10', eta='참고 출하 11/24', st='확인필요', note='100개 단가가 87개보다 쌈 → 100개', spare=13, alt=''),
 'F-03': dict(p=v(116), s=0, shop='미스미코리아', name='육각홀붙이볼트-스테인리스-', opt='SCB3-12', url='https://kr.misumi-ec.com/vona2/detail/110300239250/?HissuCode=SCB3-12', eta='참고 출하 11/24', st='조사완료', note='A2-70. 100개 단가가 더 쌈 → 100개', spare=68, alt='굿나잇몰 40원 (A2-70 표기 없음)'),
 'F-04': dict(p=v(95), s=0, shop='미스미코리아', name='육각홀붙이볼트-스테인리스-', opt='SCB3-8', url='https://kr.misumi-ec.com/vona2/detail/110300239250/?HissuCode=SCB3-8', eta='참고 출하 11/24', st='조사완료', note='100개', spare=52, alt='굿나잇몰 30원'),
 'F-05': dict(p=v(116), s=0, shop='미스미코리아', name='육각홀붙이볼트-스테인리스-', opt='SCB3-10', url='https://kr.misumi-ec.com/vona2/detail/110300239250/?HissuCode=SCB3-10', eta='참고 출하 11/24', st='조사완료', note='', alt='굿나잇몰 40원'),
 'F-06': dict(p=v(2003), s=0, shop='미스미코리아', name='바깥나사식 스트리퍼 볼트', opt='MSB5-30', url='https://kr.misumi-ec.com/vona2/detail/110100143940/?HissuCode=MSB5-30', eta='참고 출하 11/24', st='확인필요', note='10.9·무처리 (요구 12.9 흑색). 10.9 항복 약 900 MPa라 SF 계산(207 MPa 기준)에는 충분', alt=''),
 'F-07': dict(p=v(378), s=0, shop='미스미코리아', name='육각홀붙이볼트-스테인리스-', opt='SCB3-16', url='https://kr.misumi-ec.com/vona2/detail/110300239250/?HissuCode=SCB3-16', eta='참고 출하 11/24', st='조사완료', note='', alt='굿나잇몰 90원'),
 'F-08': dict(p=v(147), s=0, shop='미스미코리아', name='육각홀붙이볼트-스테인리스-', opt='SCB4-12', url='https://kr.misumi-ec.com/vona2/detail/110300239250/?HissuCode=SCB4-12', eta='참고 출하 11/24', st='조사완료', note='100개', spare=66, alt='굿나잇몰 80원'),
 'F-09': dict(p=70, s=3500, shop='굿나잇몰', name='스텐 나일론너트', opt='스텐 나일론너트 / M3', url='https://goodnightmall.com/product/%EC%8A%A4%ED%85%90-%EB%82%98%EC%9D%BC%EB%A1%A0-%EB%84%88%ED%8A%B8/116/', eta='15시 전 당일 발송', st='조사완료', note='미스미 SUNCO는 1,180원×171 ≈ 22만 원이라 제외 (품질 차이 작음)', alt='미스미 HNTL1N-SUS-M3 1,180원(별도)'),
 'F-10': dict(p=70, s=0, shop='굿나잇몰', name='스텐 나일론너트', opt='스텐 나일론너트 / M4', url='https://goodnightmall.com/product/%EC%8A%A4%ED%85%90-%EB%82%98%EC%9D%BC%EB%A1%A0-%EB%84%88%ED%8A%B8/116/', eta='당일 발송', st='조사완료', note='굿나잇몰 묶음', alt='미스미 1,260원(별도)'),
 'F-11': dict(p=v(42), s=0, shop='미스미코리아', name='평와셔', opt='SPWF3', url='https://kr.misumi-ec.com/vona2/detail/110300252910/?HissuCode=SPWF3', eta='참고 출하 11/24', st='조사완료', note='SUS304 3.2×7×0.5', alt=''),
 'F-12': dict(p=v(50), s=0, shop='미스미코리아', name='평와셔', opt='SPWF4', url='https://kr.misumi-ec.com/vona2/detail/110300252910/?HissuCode=SPWF4', eta='참고 출하 11/24', st='조사완료', note='100개', spare=37, alt=''),
 'F-13': dict(p=22400, s=3000, shop='프리쉽 (구매대행)', name='상품명 미확인 (YUEYUN QING 인서트)', opt='(173)M3X5X5 (50Set)', url='https://freeship.co.kr/Goods/mini/25690133', eta='7~21일', st='확인필요', note='정품 브랜드 못 찾음. 재질 "탄소 강철" 표기 → 황동 아닐 수 있음. 11/30 초과 위험 → 국내 대체 재조사', alt=''),
 'F-14': dict(p=v(322), s=0, shop='미스미코리아 (POP)', name='리벳너트 - 스탠다드 너트 평두 ＡＰＨ', opt='APH-425', url='https://kr.misumi-ec.com/vona2/detail/221000765935/', eta='5일째 출하', st='확인필요', note='판두께 1.5~2.5 → 섀시 3mm에 안 맞음. 3mm용 형번(그립 2.5~4) 다시 고를 것', moq=1, spare=6, alt=''),
 'F-15': dict(p=v(1057), s=0, shop='미스미코리아 (SUNCO)', name='널링볼트 - 일자 붙이 널링 비스', opt='CSMKN-SUS-M4-8', url='https://kr.misumi-ec.com/vona2/detail/221000545975/?HissuCode=CSMKN-SUS-M4-8', eta='참고 출하 11/24', st='조사완료', note='머리 φ10×5 (CAD는 Ø14×4) → CAD 머리 자리 수정', alt=''),
 'F-16': dict(p=v(14000), s=0, shop='디바이스마트', name='M2.5 황동 서포트 키트 (15종) [NT-KIT-M012]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=1360152', eta='약 3일', st='조사완료', note='디바이스마트 합계 66,000원 이상 → 무료배송', alt=''),
 'F-17': dict(p=10120, s=3000, shop='툴앤샵', name='록타이트 나사고정제 243 (1331537)', opt='243 (1331537) / 10㎖', url='https://www.toolnshop.kr/viewgs?goods=4968&gs=108432', eta='11/23 발송', st='조사완료', note='정품', alt=''),
 'E-01': dict(p=v(293000), s=0, shop='아이씨뱅큐', name='라즈베리파이 5 , Raspberry pi 5 Model 8GB', opt='단품', url='https://www.icbanq.com/P015956550', eta='1~2일', st='조사완료', note='디바이스마트와 같은 가격, 배송비 무료. 기존 디바이스마트 링크는 판매 중지', alt='디바이스마트 가이드북 포함 293,000원(별도)'),
 'E-02': dict(p=v(7000), s=0, shop='디바이스마트', name='[정품] 라즈베리파이5 액티브 쿨러 (Raspberry Pi Active Cooler)', opt='단품(SC1148)', url='https://www.devicemart.co.kr/goods/view?no=15276241', eta='1~2일', st='조사완료', note='', alt=''),
 'E-03': dict(p=v(5400), s=0, shop='디바이스마트', name='GPIO Stacking Header for Pi A+/B+/Pi 2/Pi 3 - Extra-long 2x20 Pins [ada-2223]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=14601264', eta='[해외] 2주일', st='도착 위험', note='핀 길이 미확인. 2주 → 11/30 초과', alt=''),
 'E-04': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='확인필요', note='EVO Plus 64GB는 U1/V10/A1 → 사양을 "Samsung PRO Plus 또는 EVO Plus 128GB(U3/A2)"로 바꿔 재조사', alt=''),
 'E-05': dict(p=v(91000), s=3000, shop='엘레파츠', name='[Waveshare] 2채널 CAN 통신 모듈 2-CH CAN FD HAT', opt='단품', url='https://www.eleparts.co.kr/goods/view?no=10417875', eta='24시간 내 발송', st='조사완료', note='국내 정품 유통', alt='AliExpress ₩66,000 + ₩4,463 (12/2 도착, 비공식)'),
 'E-06': dict(p=53800, s=3500, shop='로봇사이언스몰', name='[로봇사이언스몰][Adafruit][에이다프루트] Adafruit 9-DOF Orientation IMU Fusion Breakout - BNO085 (BNO080) - STEMMA QT / Qwiic ID:4754', opt='단품', url='https://robotscience.kr/goods/view?no=15723', eta='국내 재고 8, 발송일 미표시', st='조사완료', note='정품', alt='엘레파츠 44,455원(별도), 5.5일'),
 'E-07': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='보유 패드 있으면 수량 0', alt=''),
 'P-01': dict(p=249000, s=3500, shop='수성RC', name='DOGCOM 6셀 22.2V 5500mAh 150C 리포배터리 (XT90S)', opt='단품', url='https://susungrc.com/product/detail.html?product_no=2672', eta='재고 있음', st='결정 필요', note='5500mAh·926g (예산 750g보다 +210g, 전체 질량 영향). 트레이 180×80 치수 확인. CNHL 5000(714g)은 배송 15~25일', alt='CNHL 공식 US$55.99 (11/30 초과)'),
 'P-02': dict(p=81400, s=7428, shop='AliExpress YunFei FPV Store', name='ToolkitRC M7AC 2-6S 모델 항공기 리튬 배터리 충전기 AC100w/DC300W 듀얼 입력 XT30/XT60 듀얼 출력 신호 테스트', opt='M7AC EU', url='https://ko.aliexpress.com/item/1005009325477156.html', eta='11/26~12/02', st='도착 위험', note='공식 판매처 못 찾음. 비공식. 12/2 → 11/30 초과', alt=''),
 'P-03': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='v0.1 AliExpress ₩9,215 있었음', alt='https://ko.aliexpress.com/item/1005012002096700.html'),
 'P-04': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='XT90 + 12AWG로 자작 가능 (W-01 여유분)', alt=''),
 'P-05': dict(p=2900, s=3500, shop='수성RC', name='리포알람 (셀체커, 저전압알람)', opt='하우징 미포함', url='https://susungrc.com/product/detail.html?product_no=93', eta='발송일 미표시', st='확인필요', note='1~8S 지원 여부 확인', alt='CellMeter8 16,000원'),
 'P-06': dict(p=13000, s=0, shop='수성RC', name='VEGA 케블라 배터리 스트랩 (30cm, 3pcs)', opt='단품', url='https://susungrc.com/product/detail.html?product_no=1365', eta='발송일 미표시', st='조사완료', note='3개 1팩', moq=3, spare=0, alt=''),
 'P-07': dict(p=3500, s=0, shop='수성RC', name='Amass XT90S 안티스파크 커넥터 (하우징포함, 단품, 1pcs)', opt='XT90S 암놈 + XT90S 숫놈', url='https://susungrc.com/product/detail.html?product_no=1452', eta='발송일 미표시', st='조사완료', note='수성RC 묶음', alt=''),
 'P-08': dict(p=17700, s=10792, shop='AliExpress YUEREL Power 스토어', name='CZW50A 100A 200A 400A 전원 릴레이 1NO 1NC DC 접촉기 …', opt='CZW 100A / 24V / 1아니요', url='https://ko.aliexpress.com/item/1005005457160871.html', eta='11/26~12/02', st='도착 위험', note='정품 브랜드 국내 재고 없음. 18V 유지 여부 미확인', alt=''),
 'P-09': dict(p=v(60) * 5, s=0, shop='디바이스마트', name='1N5408', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=26', eta='1~2일', st='조사완료', note='5개 단가', alt=''),
 'P-10': dict(p=v(13800), s=0, shop='디바이스마트', name='비상정지누름버튼스위치-KH-3065EB-11R', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=12303398', eta='4~5일', st='확인필요', note='Autonics 정품(30,500원)은 품절·3주. 건흥 제품: DC 접점 정격·래칭 확인', alt='Autonics SF2ER-E2RAB 30,500원(별도), 품절'),
 'P-11': dict(p=16950 + 3230 * 5, s=0, shop='휴즈천국', name='498 Fuse Holder (0498900.TXN) + 498 Series MIDI 60A (Littelfuse)', opt='홀더 1 + 60A 퓨즈 5 (최소 주문)', url='https://www.paparts.co.kr/shop/item.php?it_id=498FuseHolder', eta='발송일 미표시', st='조사완료', note='퓨즈 링크 https://www.paparts.co.kr/shop/item.php?it_id=498 . 배송비 주문 시 결정', need=1, spare=0, alt=''),
 'P-12': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='', alt=''),
 'P-13': dict(p=v(14500), s=0, shop='디바이스마트', name='Metal Pushbutton with Wires - Momentary (16mm, Blue) [COM-11967]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=15650575', eta='[해외] 1주일', st='확인필요', note='1NO·접점 정격 확인', alt=''),
 'P-14': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='', alt=''),
 'P-15': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='', alt=''),
 'P-16': dict(p=v(70000), s=0, shop='디바이스마트', name='5V, 5.5A Step-Down Voltage Regulator D36V50F5 #4091', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=12536510', eta='[해외] 1주일', st='도착 위험', note='Pololu 정품, 입력 5.5~50V. 역극성 보호 표기 없음', alt=''),
 'P-17': dict(p=v(1700) * 4, s=0, shop='디바이스마트', name='자동차 퓨즈홀더 287 AUTO / ATO 전용 방수 (소) [SZH-FU005]', opt='단품 ×4', url='https://www.devicemart.co.kr/goods/view?no=12170196', eta='1주일', st='확인필요', note='ATO 홀더 4개분만. 25A/2A/3A 퓨즈와 미니 블레이드 홀더 2개는 못 찾음', alt=''),
 'W-01': dict(p=6000, s=3000, shop='엑스캅터', name='12AWG 실리콘 전선 (블랙, 레드 1m / Silicon Wire)', opt='단품', url='https://www.xcopter.com/product/detail.html?product_no=2310', eta='발송일 미표시', st='조사완료', note='AMASS 200°C. 엑스캅터 묶음 배송 1회', alt=''),
 'W-02': dict(p=4400, s=0, shop='엑스캅터', name='14AWG 실리콘 전선 (블랙, 레드 1m / Silicon Wire)', opt='단품', url='https://www.xcopter.com/product/detail.html?product_no=2311', eta='', st='조사완료', note='', alt=''),
 'W-03': dict(p=3000, s=0, shop='엑스캅터', name='16AWG 실리콘 전선 (블랙, 레드 1m / Silicon Wire)', opt='단품', url='https://www.xcopter.com/product/detail.html?product_no=2312', eta='', st='조사완료', note='', alt=''),
 'W-04': dict(p=2000, s=0, shop='엑스캅터', name='18AWG 실리콘 전선 (블랙, 레드 1m / Silicon Wire)', opt='단품', url='https://www.xcopter.com/product/detail.html?product_no=2313', eta='', st='조사완료', note='', alt=''),
 'W-05': dict(p=1500, s=0, shop='엑스캅터', name='20AWG 실리콘 전선 (블랙, 레드 1m / Silicon Wire)', opt='단품', url='https://www.xcopter.com/product/detail.html?product_no=2308', eta='', st='조사완료', note='', alt=''),
 'W-06': dict(p=v(34500), s=0, shop='디바이스마트', name='6색 실리콘 와이어 키트 (20AWG, CBAZY)', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=15080687', eta='[해외] 1~2일', st='조사완료', note='흰색 단품 없음 → 6색 각 6m 키트 (흰색 1m만 필요해 남음)', alt=''),
 'W-07': dict(p=v(24000), s=0, shop='디바이스마트', name='6색 실리콘 와이어 키트 (24AWG, CBAZY, 36M)', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=12375538', eta='1~2일', st='조사완료', note='키트 1개로 W-07·08·09 (흑·녹·황 각 6m)', need=1, spare=0, alt=''),
 'W-08': dict(p=0, s=0, shop='디바이스마트 (W-07 키트)', name='W-07 키트에 포함', opt='', url='https://www.devicemart.co.kr/goods/view?no=12375538', eta='', st='조사완료', note='W-07 키트에 포함', alt=''),
 'W-09': dict(p=0, s=0, shop='디바이스마트 (W-07 키트)', name='W-07 키트에 포함', opt='', url='https://www.devicemart.co.kr/goods/view?no=12375538', eta='', st='조사완료', note='W-07 키트에 포함', alt=''),
 'W-10': dict(p=1200, s=0, shop='엑스캅터', name='아마스 Amass XT60H-F 드론 커넥터 암 1개 + XT60H-M 수 1개 (하우징 포함)', opt='암 1 + 수 1', url='https://www.xcopter.com/product/detail.html?product_no=27715', eta='', st='조사완료', note='수 링크 https://www.xcopter.com/product/detail.html?product_no=29799', alt=''),
 'W-11': dict(p=v(700), s=0, shop='디바이스마트', name='XT30(2+2)-F.G.B', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=15856243', eta='1~2일', st='확인필요', note='Amass 정품, 케이블 쪽 암. 모터 커넥터 성별 확인 후 구매', alt='AliExpress 5개 ₩5,040 (비공식)'),
 'W-12': dict(p=v(500), s=0, shop='디바이스마트', name='[NW3 (중국)] 하네스 커넥터 SM 2핀-2핀 케이블', opt='SM-F/M-20cm', url='https://www.devicemart.co.kr/goods/view?no=1342790', eta='2~3일', st='조사완료', note='JST 정품 아님, AWG22', alt=''),
 'W-13': dict(p=v(260) * 4, s=0, shop='디바이스마트', name='[이에스단자] 동관단자 1홀 4SQ (홀 M5)', opt='단품 ×4', url='https://www.devicemart.co.kr/goods/view?no=12380190', eta='4~5일', st='확인필요', note='4mm² M5만 찾음. M6·2.5mm²·1.5mm²는 못 찾음', alt=''),
 'W-14': dict(p=v(9000), s=0, shop='디바이스마트', name='유럽식 페룰 단자 1200PCS 혼합 세트 [케이스 포함 / 관형 예비 절연 터미널 / E0508~E10-12]PV-2', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=16055797', eta='1주일', st='확인필요', note='0.25mm² 없음 (HAT CAN 선용)', alt=''),
 'W-15': dict(p=v(100), s=0, shop='디바이스마트', name='USB 3.0 C타입 전원용 2PIN 케이블(25CM) 수 (REXBOT)', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=16097207', eta='1~2일', st='확인필요', note='AWG·정격 전류 미표기 → 5A 보장 안 됨. Pi 5 전원은 GPIO 5V 핀 직결 대안', alt=''),
 'W-16': dict(p=v(2600), s=0, shop='디바이스마트', name='아두이노 점퍼케이블 3P F/F 15cm 옐로우 (10개) [CAB380107]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=1386058', eta='[해외] 1주일', st='확인필요', note='3P 하우징 (1P 낱선 필요하면 다시)', alt=''),
 'W-17': dict(p=v(8000), s=0, shop='디바이스마트', name='NT- DIY 열수축튜브 280세트', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=1320937', eta='3~4일', st='조사완료', note='', alt=''),
 'W-18': dict(p=v(1800 + 2500 + 2900), s=0, shop='디바이스마트', name='접착식 열수축튜브 6파이 + 9파이 + 12파이 /검정(1M)', opt='각 1M', url='https://www.devicemart.co.kr/goods/view?no=30646', eta='3~4일', st='조사완료', note='9파이 https://www.devicemart.co.kr/goods/view?no=30647 / 12파이 no=30648. Ø10 대신 9파이. 3:1 표기 없음', need=1, spare=0, alt=''),
 'W-19': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='롤 단위만 있음', alt=''),
 'W-20': dict(p=v(400 + 1500), s=0, shop='디바이스마트', name='티테크놀로지 케이블타이 100개 [T-C2510B] + 케이블메이트 접착 마운트 20x20 50개', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=12511035', eta='3~4일', st='조사완료', note='마운트 https://www.devicemart.co.kr/goods/view?no=10925025', alt=''),
 'J-01': dict(p=round((4 * (80 + 5720) + 4 * (80 + 3400)) / 8), s=6000, shop='효성 (hyosung.godomall.com)', name='2020 알루미늄프로파일 블랙 절단 가공 제작 축양장', opt='401~500mm +5,720원 ×4 / 201~300mm +3,400원 ×4', url='http://hyosung.godomall.com/goods/goods_view.php?goodsNo=1000000196', eta='절단품 익일 발송', st='조사완료', note='단가는 8본 평균. 슬롯 형상 확인 (T너트 맞춤)', alt='시화프로파일 2,940원/500mm'),
 'J-02': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='J-01 판매처에서 같이 찾기', alt=''),
 'J-03': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='HIWIN 국내 판매 없음', alt=''),
 'J-05': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='가격 미조사', note='', alt=''),
 # 1단계 시험
 'T-01': dict(p=v(45000), s=0, shop='디바이스마트', name='CANable V2.0 Pro USB to CAN 모듈 (C타입)', opt='단품 (Makerbase)', url='https://www.devicemart.co.kr/goods/view?no=15775527', eta='[해외] 1주일', st='유지(v0.1)', note='선정 유지, 가격만 채움. vctec 43,500원(별도)은 재고 문의', alt='AliExpress ₩25,250 (비공식)'),
 'T-02': dict(p=5170, s=0, shop='AliExpress EasyWired Store', name='XT30(2+2) 수-암 연장 케이블, AMS 플러그, XT30(2+2) - F 커넥터 …', opt='Male to Female / 30CM', url='https://ko.aliexpress.com/item/1005012775760038.html', eta='11/26~12/02', st='유지(v0.1)', note='선정 유지(같은 링크), 가격 갱신. 12/2 도착 표시', alt='AliExpress Amass 표기 ₩7,570 + ₩4,201'),
 'T-03': dict(p=36700 + 1606, s=0, shop='AliExpress RAYTEI Load Cell Store + Realpoy', name='합금강 S형 로드셀 … 50kg + HX711 Mini 모듈', opt='50KG (20KG 같은 가격) + Mini', url='https://ko.aliexpress.com/item/1005005915852645.html', eta='11/26~12/02', st='유지(v0.1)', note='선정 유지(S형 + HX711). HX711 https://ko.aliexpress.com/item/1005006574602903.html . 20kg 옵션이 분해능에 유리', alt='XNQ ₩28,450 + ₩6,185'),
 'T-04': dict(p=128000, s=0, shop='AliExpress HANMATEK Official Store', name='HANMATEK DC 전원 공급기 30V 10A 다기능 디스플레이', opt='HM-310T 30V 10A', url='https://ko.aliexpress.com/item/1005004833435425.html', eta='11/26~12/02', st='도착 위험', note='브랜드 공식 스토어. ₩278,261 → ₩128,000으로 내림', alt=''),
 'T-05': dict(p=None, s=0, shop='없음', name='', opt='', url='', eta='', st='확인필요', note='', alt=''),
 'T-06': dict(p=v(900), s=0, shop='디바이스마트', name='티테크놀로지 바나나 플러그, 싱글 바나나잭, T-BPS41 [블랙/레드] [1set-2개]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=15155291', eta='3~4일', st='조사완료', note='바나나 플러그만. XT30·전선은 W-03·W-08·W-11 여유분', alt=''),
 'T-07': dict(p=v(1700), s=0, shop='디바이스마트', name='자동차 퓨즈홀더 287 AUTO / ATO 전용 방수 (소) [SZH-FU005]', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=12170196', eta='1주일', st='확인필요', note='15A 퓨즈 못 찾음', alt=''),
 'T-09': dict(p=v(5900 + 850), s=0, shop='디바이스마트', name='아두이노 나노 호환보드 V3.0 CH340 C타입 [THC-NAO001] + 소켓 점퍼 40P M/F 20cm', opt='납땜 버전', url='https://www.devicemart.co.kr/goods/view?no=15315999', eta='2~3일', st='조사완료', note='점퍼 https://www.devicemart.co.kr/goods/view?no=1321195', alt=''),
 'T-10': dict(p=8800 + 62700, s=3300, shop='랩앤샵', name='M1급 표준분동 낱개 1g~20kg', opt='용량 2kg', url='https://www.labnshop.co.kr/goods/goods_view.php?goodsNo=1000000808', eta='발송일 미표시', st='조사완료', note='OIML M1', alt='파인드툴 한성계기 65,000원 (SOLD OUT 문구)'),
 'T-11': dict(p=700 * 2 + 3300, s=0, shop='굿나잇몰', name='스텐 아이볼트 ×2 + 스텐 와이어 10M (1.6mm)', opt='M6 또는 M8 / 1.6mm', url='https://goodnightmall.com/product/%EC%8A%A4%ED%85%90-%EC%95%84%EC%9D%B4%EB%B3%BC%ED%8A%B8/184', eta='14시 전 당일 발송', st='조사완료', note='와이어 https://goodnightmall.com/product/%EC%8A%A4%ED%85%90-%EC%99%80%EC%9D%B4%EC%96%B4-10m/63/ . 와이어 클립 못 찾음', alt=''),
 'T-12': dict(p=None, s=0, shop='77g 철물점&목공소', name='C 클램프 4인치 (100mm)', opt='D5B5135', url='https://www.77g.com/products/H347?cd=D5136&co=D5B5135', eta='', st='가격 미조사', note='가격 확인 못 함', alt=''),
 'T-13': dict(p=67230 + v(94000), s=3000, shop='공구명가 + 미스미코리아', name='미쓰도요 다이얼 인디케이터 2046A + 마그네틱 스탠드 7010S-10', opt='2046A / 7010S-10', url='https://yugatool.co.kr/product/%EB%AF%B8%EC%93%B0%EB%8F%84%EC%9A%94-%EB%8B%A4%EC%9D%B4%EC%96%BC-%EC%9D%B8%EB%94%94%EC%BC%80%EC%9D%B4%ED%84%B0/987', eta='발송일 미표시', st='조사완료', note='Mitutoyo 정품. 베이스 https://kr.misumi-ec.com/vona2/detail/223304174613?HissuCode=7010S-10', alt=''),
 'T-14': dict(p=v(85000), s=0, shop='디바이스마트', name='휴대용 디지털 멀티미터 UT61D+', opt='단품', url='https://www.devicemart.co.kr/goods/view?no=1075512', eta='1~2일', st='확인필요', note='온도(K형) 측정 기능 표기 없음 → 온도 기능 모델 재확인. 열전대·캡톤 별도', alt=''),
 'T-15': dict(p=141500, s=3500, shop='캘리퍼스샵', name='Mitutoyo 디지털버니어캘리퍼스 150mm, 표준형, 전문가용, 방진 방수형', opt='표준형 150mm / 500-181-30 (정품)', url='https://www.calipers.shop/goods/goods_view.php?goodsNo=1000000202', eta='오늘 발송', st='조사완료', note='같은 정품 중 저렴한 곳', alt='대광계측기 165,000원'),
 'T-16': dict(p=v(16400), s=0, shop='디바이스마트', name='[공식 정품] 라즈베리파이 27W USB-C 아답터 흰색', opt='흰색', url='https://www.devicemart.co.kr/goods/view?no=15502402', eta='1~2일', st='조사완료', note='', alt=''),
}
# 판금·견적: S-01/S-02/J-04/T-08 → 링크만
QUOTE = {
 'S-01': ('MakeIT', 'https://makeit.ai.kr/services/sheet-metal', 'CutON DXF 즉시 견적 https://cuton.co.kr/'),
 'S-02': ('MakeIT', 'https://makeit.ai.kr/services/sheet-metal', 'JLCCNC https://jlccnc.com/kr/sheet-metal-quote'),
 'J-04': ('MakeIT (판금 주문에 추가)', 'https://makeit.ai.kr/services/sheet-metal', 'CutON'),
 'T-08': ('한국미스미 온라인 도면가공센터', 'https://fa-mdrs.misumi.co.kr/', 'MakeIT'),
}

def setv(ws, r, c, val):
    ws.cell(row=r, column=c).value = val

wb = load_workbook(SRC)
FILL = {'조사완료': 'E2EFDA', '도착 위험': 'FCE4D6', '확인필요': 'FFF2CC', '결정 필요': 'FCE4D6', '가격 미조사': 'FFFFFF', '견적 필요': 'FFF2CC', '유지(v0.1)': 'E2EFDA'}
hits = set()
for sh in ('BOM 전체', '1단계 시험'):
    ws = wb[sh]
    # 새 열: 페이지 상품명 / 옵션 / 도착 예정
    base = ws.max_column
    for i, h in enumerate(['페이지 상품명 (11-21 조사)', '선택 옵션', '도착 예정'], 1):
        c = ws.cell(row=3, column=base + i, value=h); c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1F3864')
        ws.column_dimensions[c.column_letter].width = [40, 26, 16][i - 1]
    for r in range(4, ws.max_row + 1):
        id_ = ws.cell(row=r, column=2).value
        if id_ in C:
            d = C[id_]; hits.add(id_)
            setv(ws, r, 14, d['p'])
            setv(ws, r, 15, '11-21 조사 (VAT 포함 환산)' if d['p'] else '')
            setv(ws, r, 16, d['s'])
            setv(ws, r, 20, d['shop'])
            setv(ws, r, 21, d['url'])
            setv(ws, r, 22, d['alt'])
            setv(ws, r, 24, d['st'])
            for k, col in (('need', 10), ('spare', 11), ('moq', 12)):
                if k in d: setv(ws, r, col, d[k])
            old = ws.cell(row=r, column=28).value or ''
            setv(ws, r, 28, (d['note'] + (' | 이전: ' + old if old else '')).strip(' |'))
            setv(ws, r, base + 1, d['name']); setv(ws, r, base + 2, d['opt']); setv(ws, r, base + 3, d['eta'])
            for col in range(1, base + 4):
                ws.cell(row=r, column=col).fill = PatternFill('solid', fgColor=FILL.get(d['st'], 'FFFFFF'))
                ws.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical='top')
        elif id_ in QUOTE:
            q = QUOTE[id_]; hits.add(id_)
            setv(ws, r, 20, q[0]); setv(ws, r, 21, q[1]); setv(ws, r, 22, q[2])
    ws.title = ws.title
missing = [k for k in list(C) + list(QUOTE) if k not in hits]
print('not found in sheets:', missing)

cv = wb['표지']
cv['A1'] = 'Azunyang 로봇개 부품 목록 (BOM) v0.3'
cv['A2'] = 'v0.2 + 2026-11-21 가격·링크 조사 반영. 단가는 VAT 포함으로 환산(별도 표기 × 1.1). 배송비는 판매처별 1회만 넣음 (디바이스마트·원일베어링은 무료 기준 넘음).'
cv['A4'] = '조사 날짜'; cv['B4'] = '2026-11-21'
row = cv.max_row + 2
for t in ['v0.2 → v0.3 (2026-11-21)',
          '선택 기준: 품질 우선(추천 후보). 같은 상품이면 더 싸거나 배송 빠른 곳. 추천이 너무 비싸면 최저가 (예: 나일론 너트 미스미 22만 원 → 굿나잇몰).',
          '색: 초록 = 조사완료, 주황 = 도착 위험(11/30 초과 가능) 또는 결정 필요, 노랑 = 확인 필요, 흰색 = 가격 미조사',
          '새 열 3개: 페이지 상품명 / 선택 옵션 / 도착 예정 → 링크를 열어 이 셋이 맞는지 확인',
          '미스미 수량 할인: 100개 단가가 더 싼 볼트·와셔는 구매 수량을 100으로 올림 (F-02, F-03, F-04, F-08, F-12)']:
    cv.cell(row=row, column=1, value=t).alignment = Alignment(wrap_text=True); row += 1
cv.cell(row=row - 5, column=1).font = Font(bold=True)
wb.save(OUT)
print('saved', OUT)
