import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]
FONT = '맑은 고딕'

# ---------- 27년 SW 예산 (정부·국회 심의자료) ----------
B27 = [
    ('모바일 서비스 구축', 60, '기존 서비스 기반 모바일 환경 확장, 하이브리드 UI/UX 고도화, 질의응답 및 조회 기능 모바일 최적화'),
    ('RAG 데이터 구축 및 고도화', 180, '내부 문서 수집·전처리 체계 고도화, 메타데이터 및 벡터DB 관리 고도화, 검색 정확도 향상(Chunking, Reranking), 관리자 기반 RAG 데이터 관리 기능 고도화'),
    ('답변 품질 향상', 100, 'Prompt 최적화 및 응답 정확도 개선, 출처 기반 답변 및 근거 제시 고도화, 답변 일관성 확보 및 품질 개선'),
    ('업무지원 및 Agent 기능', 200, '기존 VOC 분류 기능 고도화, 보고서 초안 생성 기능 확장, 규정/코드 기반 분석 지원 기능 강화(자동화 및 RPA 기능 제외)'),
    ('전자결재 연계 (API 기반)', 60, '전자결재 시스템 연계 인터페이스 구현, 결재 문서 초안 생성 기능 제공(시스템 커스터마이징 제외)'),
    ('데이터 관리 및 운영 기능', 120, '관리자 기능 고도화 및 데이터 관리 체계 개선, 데이터 갱신 및 운영 기능 강화, 개인정보 및 민감정보 필터링 적용'),
    ('DRM 보안소프트웨어', 16, '파일 복·암복화 솔루션'),
    ('sso 인증서버소프트웨어', 22, '기존 어드바이저 통합인증 기능(로그인 기능) 적용'),
]
K_MOB, K_RAG, K_ANS, K_AGT, K_APR, K_DAT, K_DRM, K_SSO = [b[0] for b in B27]
K_NEW = '신규(기획안에만 있음)'
K_PM = '공통(사업관리·품질)'

# 매칭 구분
M_EQ = '● 일치'
M_PT = '◐ 부분일치(범위확대)'
M_NEW = '★ 기획안에만 있음'
M_27 = '▲ 27예산에만 있음'

# 세부기능 유형: L=인건비(M/M), S=SW구매(식), E=직접경비(식)
# (세부기능, 유형, 수량, 단가(백만원, S/E만), 27예산 키, 산출근거/비고)
TASKS = [
 ('과제1', '가스안전 지식DB 고도화 (데이터 정비)', [
   ('1-1', '비정형 지식자산 전수조사 및 데이터 컨설팅', '부서·개인 보유 비정형 자료 전수조사, 수집→분류→정제→표준화→검증→최신화 절차 수립',
    '-', '-', M_NEW, [
     ('부서·개인 보유 자료 현황조사 및 자료목록(인벤토리) 작성', 'L', 1.0, None, K_NEW, '전 부서 대상 설문·인터뷰, 자료 유형/보안등급 분류'),
     ('지식자산 분류체계 및 데이터 정비 절차(수집~최신화) 수립', 'L', 1.0, None, K_NEW, '데이터 컨설팅 산출물: 분류체계·정비 가이드'),
   ]),
   ('1-2', '법령·KGS Code·규정·매뉴얼·FAQ 데이터 정비', '법령·KGS Code·규정·매뉴얼·FAQ 등 데이터 정비',
    K_RAG, '내부 문서 수집·전처리 체계 고도화', M_EQ, [
     ('문서 수집·변환 파이프라인 고도화(HWP/PDF/표·이미지 파싱)', 'L', 1.0, None, K_RAG, 'DRM 문서 복호화 연동 포함'),
     ('데이터 정제·전처리(중복·오류 제거, 구조화)', 'L', 1.5, None, K_RAG, ''),
     ('DRM 보안소프트웨어(파일 복·암호화) 1식', 'S', 1, 14.5, K_DRM, '27예산 16백만원(VAT포함) → 공급가 14.5 적용'),
   ]),
   ('1-3', '데이터 표준화 및 메타데이터 구축', '데이터 표준화 및 메타데이터 구축',
    K_RAG, '메타데이터 및 벡터DB 관리 고도화', M_EQ, [
     ('데이터 표준 및 메타데이터 스키마 설계', 'L', 0.5, None, K_RAG, '문서유형·시행일·개정이력·보안등급 등'),
     ('메타데이터 태깅 및 벡터DB 재적재·관리 기능', 'L', 1.5, None, K_RAG, ''),
   ]),
   ('1-4', '자료 최신화 및 버전관리 체계 구축', '자료 최신화 및 버전관리 체계 구축',
    K_DAT, '데이터 갱신 및 운영 기능 강화 (사업개요: 법령·사규·코드 자동 업로드)', M_EQ, [
     ('법령·KGS Code·사규 변경분 자동 수집·업로드', 'L', 1.0, None, K_DAT, '27 사업내용 "자동 업로드 시스템" 대응'),
     ('문서 버전·이력 관리 및 폐지자료 자동 제외', 'L', 0.5, None, K_DAT, '현행자료 반영률 KPI(≥98%) 대응'),
   ]),
   ('1-5', '지식DB 관리자 기능 개선', '지식DB 관리자 기능 개선',
    K_DAT, '관리자 기능 고도화 / 관리자 기반 RAG 데이터 관리 기능 고도화', M_EQ, [
     ('문서 등록·검수·승인 워크플로', 'L', 0.5, None, K_DAT, ''),
     ('지식DB 품질 모니터링 화면(적재현황·오류·최신성)', 'L', 0.5, None, K_RAG, ''),
   ]),
   ('1-6', '1차 사업 제외자료 확대 적재', '1차 사업 제외자료 확대 (고도화 추진방향)',
    '-', '-', M_NEW, [
     ('1차 제외자료 선별·정비·적재', 'L', 0.5, None, K_NEW, ''),
   ]),
   ('1-7', '개인정보·민감정보 필터링', '(기획안 KPI: 개인정보 침해 0건, 내부정보 외부노출 0건)',
    K_DAT, '개인정보 및 민감정보 필터링 적용', M_PT, [
     ('개인정보·민감정보 탐지 및 마스킹(적재·답변 단계)', 'L', 0.5, None, K_DAT, '과제3 대국민 입력필터와 공통 모듈화'),
   ]),
 ]),
 ('과제2', '스마트온 2.0 연계를 통한 모바일 서비스 제공', [
   ('2-1', '스마트온 2.0 연계', '스마트온 2.0 연계 (검사지원처 협업)',
    K_MOB, '기존 서비스 기반 모바일 환경 확장', M_PT, [
     ('스마트온 2.0 연계 인터페이스(API/웹뷰) 개발', 'L', 1.0, None, K_MOB, '27예산은 "모바일 환경 확장"만 명시 → 스마트온 연계로 범위 구체화'),
   ]),
   ('2-2', '계정·인증 연계', '스마트온 2.0 계정·인증 연계',
    K_SSO, '기존 어드바이저 통합인증(로그인) 적용', M_PT, [
     ('SSO 연동 및 스마트온 계정 매핑', 'L', 0.5, None, K_SSO, ''),
     ('SSO 인증서버소프트웨어 1식', 'S', 1, 20.0, K_SSO, '27예산 22백만원(VAT포함) → 공급가 20.0 적용. 스마트온 계정연계 지원 여부 확인 필요'),
   ]),
   ('2-3', '모바일 UI/UX 개발', '모바일 UI/UX 개발',
    K_MOB, '하이브리드 UI/UX 고도화', M_EQ, [
     ('현장 검사원용 모바일 화면 설계(UI/UX)', 'L', 0.5, None, K_MOB, ''),
     ('하이브리드(반응형) 모바일 UI 구현', 'L', 1.0, None, K_MOB, ''),
   ]),
   ('2-4', '모바일 검색 및 답변 근거 확인', '모바일 검색 및 답변 근거 확인',
    K_MOB, '질의응답 및 조회 기능 모바일 최적화', M_EQ, [
     ('질의응답·조회 기능 모바일 최적화', 'L', 0.5, None, K_MOB, ''),
     ('모바일 근거문서(KGS Code 조항) 뷰어', 'L', 0.5, None, K_MOB, '현장활용 예: 기준검색→Code제시→근거제공→문서연결'),
   ]),
   ('2-5', '음성입력(STT)', '음성입력(STT) 기능',
    '-', '-', M_NEW, [
     ('STT 엔진 연동 및 음성질의 처리', 'L', 0.5, None, K_NEW, ''),
     ('가스안전 전문용어 사전 구축·인식률 튜닝', 'L', 0.5, None, K_NEW, ''),
     ('STT 엔진 소프트웨어 1식(온프레미스)', 'S', 1, 30.0, K_NEW, '가정치 — 업체 견적 확인 필요(폐쇄망 설치형 기준)'),
   ]),
   ('2-6', '현장 활용성 기능 개선', '현장 활용성을 고려한 기능 개선 / 향후 현장자료·첨부파일 활용 확대',
    '-', '-', M_NEW, [
     ('현장 사진·첨부파일 첨부 질의(기초기능)', 'L', 0.5, None, K_NEW, '기획안상 "향후" 과제 — 27년 기초기능만 반영'),
   ]),
 ]),
 ('과제3', '사이버지사 VOC 질의 챗봇 구축 (대국민 서비스)', [
   ('3-1', '사이버지사 AI 상담챗봇 구축', '사이버지사 AI 상담챗봇 구축, 24시간 일반 상담',
    '-', '-', M_NEW, [
     ('대국민 웹 챗봇 UI(위젯) 개발', 'L', 1.0, None, K_NEW, '웹접근성 인증 기준 준수'),
     ('사이버지사 홈페이지 연동', 'L', 0.5, None, K_NEW, '홍보소통실 협업'),
   ]),
   ('3-2', '공개자료 기반 별도 지식DB 구축', '공개자료 기반 별도 지식DB, 내부자료와 대외자료 분리',
    '-', '-', M_NEW, [
     ('공개 가능자료(법령·KGS Code 등) 선별·적재', 'L', 1.0, None, K_NEW, '사규·인사·개인정보 원천 배제'),
     ('내부/대외 지식DB 분리(별도 인덱스·접근통제)', 'L', 0.5, None, K_NEW, ''),
   ]),
   ('3-3', 'AI 상담체계(단계별 응대) 구현', '단순→AI 자동응답 / 전문→근거포함 답변 / 복합→담당부서 이관',
    '-', '-', M_NEW, [
     ('질의 유형 분류 및 3단계 라우팅', 'L', 1.0, None, K_NEW, ''),
     ('근거자료 포함 대국민 답변 생성', 'L', 0.5, None, K_NEW, ''),
   ]),
   ('3-4', '담당부서 이관체계 및 VOC 연계', '복잡한 질의 담당부서 이관, VOC 연계',
    K_AGT, '기존 VOC 분류 기능 고도화 (내부 VOC 대상)', M_PT, [
     ('전문상담 담당부서 이관 워크플로', 'L', 0.5, None, K_NEW, ''),
     ('VOC 시스템 연계 및 기존 VOC 분류모델 재활용', 'L', 1.0, None, K_AGT, '27 Agent 항목 "VOC 분류 고도화"를 대국민 이관에 활용'),
   ]),
   ('3-5', '대국민 서비스 보안·가드레일', '(핵심원칙) 공개자료만 활용, 내부자료 원천배제',
    '-', '-', M_NEW, [
     ('입력 필터(개인정보·비속어·프롬프트 인젝션)', 'L', 0.5, None, K_NEW, ''),
     ('답변 범위 제한 및 면책·이관 안내', 'L', 0.5, None, K_NEW, ''),
   ]),
   ('3-6', '상담 운영 모니터링', '(KPI) 챗봇 이용건수, 자동응답률, 이관률, 국민 만족도',
    '-', '-', M_NEW, [
     ('상담 로그·통계 대시보드(자동응답률·이관률·만족도)', 'L', 0.5, None, K_NEW, ''),
   ]),
 ]),
 ('과제4', '"정확도"를 넘어 "신뢰성" 확보 (AI 신뢰성 고도화)', [
   ('4-1', 'RAG 검색체계·검색 알고리즘 고도화', 'RAG 검색체계 개선, 검색 알고리즘 고도화',
    K_RAG, '검색 정확도 향상(Reranking)', M_EQ, [
     ('하이브리드 검색(키워드+벡터) 적용', 'L', 0.5, None, K_RAG, ''),
     ('Reranking 모델 적용·튜닝', 'L', 0.5, None, K_RAG, ''),
   ]),
   ('4-2', '문서 전처리·청킹 개선', '문서 전처리·청킹 개선',
    K_RAG, '검색 정확도 향상(Chunking)', M_EQ, [
     ('문서구조(조·항·호/표) 기반 청킹 개선', 'L', 1.0, None, K_RAG, ''),
   ]),
   ('4-3', '답변 근거자료 제시', '답변 근거자료 제시',
    K_ANS, '출처 기반 답변 및 근거 제시 고도화', M_EQ, [
     ('출처 인용·근거 조항 하이라이트', 'L', 0.5, None, K_ANS, 'KPI 근거자료 제시율 ≥95%'),
   ]),
   ('4-4', '할루시네이션 및 치명적 오답 관리', '할루시네이션 및 치명적 오답 관리',
    K_ANS, 'Prompt 최적화 및 응답 정확도 개선, 답변 일관성 확보', M_EQ, [
     ('Prompt 최적화 및 답변 일관성 개선', 'L', 0.5, None, K_ANS, ''),
     ('답변-근거 일치 검증 및 불확실 시 응답유보(가드레일)', 'L', 1.5, None, K_ANS, 'KPI 치명적 오답 0건'),
   ]),
   ('4-5', '가스안전 전문 평가셋 구축', '가스안전 전문 평가셋 구축',
    '-', '-', M_NEW, [
     ('대표질의·정답·근거 평가셋 구축', 'L', 1.0, None, K_NEW, 'Baseline 측정 포함'),
     ('가스안전 분야 전문가 검수 자문', 'E', 1, 8.0, K_NEW, '직접경비 가정치(자문 20인·일 내외)'),
   ]),
   ('4-6', '정기 성능평가 체계 구축', '정기 성능평가 체계 구축 (정확성·근거성·최신성·안전성·완전성)',
    '-', '-', M_NEW, [
     ('자동 평가 파이프라인(회귀평가)', 'L', 1.0, None, K_NEW, '오픈소스 RAG 평가도구 활용'),
     ('5대 평가지표·KPI 대시보드', 'L', 0.5, None, K_NEW, ''),
   ]),
   ('4-7', '사용자 피드백 및 오류관리', '사용자 피드백 및 오류관리',
    K_ANS, '품질 개선', M_PT, [
     ('피드백·오류신고 수집 및 개선 이력관리', 'L', 0.5, None, K_ANS, ''),
   ]),
 ]),
 ('공통', '사업관리 및 품질보증', [
   ('0-1', '사업관리·품질보증', '(정보화사업 공통) 착수~종료 관리, 보안성 검토 대응',
    '-', '-', '공통', [
     ('사업관리(PM)', 'L', 2.0, None, K_PM, '부분투입 기준'),
     ('통합테스트·보안점검·이행 및 교육', 'L', 1.0, None, K_PM, ''),
   ]),
 ]),
]

# 27예산에만 있는 과업 (기획안 3단계 '28년 범위)
ONLY27 = [
    ('5-1', '기존 VOC 분류 기능 고도화', K_AGT, '기획안 3단계(AI 에이전트) — 단, 일부는 과제3-4 VOC 연계에 활용'),
    ('5-2', '보고서 초안 생성 기능 확장', K_AGT, '기획안 3단계(AI 에이전트, \'28년)'),
    ('5-3', '규정/코드 기반 분석 지원 기능 강화', K_AGT, '기획안 3단계(AI 에이전트, \'28년)'),
    ('5-4', '전자결재 시스템 연계 인터페이스 구현', K_APR, '기획안 3단계(전자결재 연동, \'28년)'),
    ('5-5', '결재 문서 초안 생성 기능', K_APR, '기획안 3단계(전자결재 연동, \'28년)'),
]

# ---------- styles ----------
thin = Side(style='thin', color='999999')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F = lambda **k: Font(name=FONT, size=k.pop('size', 10), **k)
HDR_FILL = PatternFill('solid', fgColor='1F3864')
SUB_FILL = PatternFill('solid', fgColor='D9E1F2')
TOT_FILL = PatternFill('solid', fgColor='FFF2CC')
INPUT_FILL = PatternFill('solid', fgColor='FFFF00')
FILLS = {M_EQ: 'E2EFDA', M_PT: 'FFF2CC', M_NEW: 'FCE4D6', M_27: 'EDEDED', '공통': 'FFFFFF'}
BLUE = '0000FF'
GREEN = '008000'
WRAP = Alignment(wrap_text=True, vertical='center')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
NUM1 = '#,##0.0;(#,##0.0);-'
NUM0 = '#,##0;(#,##0);-'

def hdr(ws, row, headers, widths=None):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = F(bold=True, color='FFFFFF'); c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

def put(ws, row, col, val, bold=False, color=None, fill=None, align=WRAP, fmt=None, border=True):
    c = ws.cell(row=row, column=col, value=val)
    kw = {'bold': bold}
    if color: kw['color'] = color
    c.font = F(**kw); c.alignment = align
    if fill: c.fill = PatternFill('solid', fgColor=fill) if isinstance(fill, str) else fill
    if fmt: c.number_format = fmt
    if border: c.border = BORDER
    return c

def title(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = F(size=14, bold=True)
    if sub: ws['A2'] = sub; ws['A2'].font = F(size=9, color='595959')

wb = Workbook()

# ================= Sheet: 개요 =================
ov = wb.active; ov.title = '0.개요'
title(ov, '가스안전 AI 어드바이저 2.0 고도화 — 과제·과업별 SW 예산 견적', '근거: ①27년 신규사업 정부·국회 심의자료(SW 758백만원) ②가스안전 AI 어드바이저 2.0 고도화 추진계획(2026.9, 제도정책부)')
ov.column_dimensions['A'].width = 34; ov.column_dimensions['B'].width = 16; ov.column_dimensions['C'].width = 70

r = 4
put(ov, r, 1, '■ 견적 산정 가정 (노란 셀 = 수정 가능 입력값)', bold=True, border=False); r += 1
hdr(ov, r, ['항목', '값', '근거/설명']); r += 1
ASSUME = [
    ('2025년 적용 응용SW개발자 월평균임금(백만원)', 6.943457, 'KOSA 2025년 적용 SW기술자 평균임금 공표(응용SW개발자, 인공지능·빅데이터개발자 포함) 월 6,943,457원', '0.000000'),
    ('2026년 임금 인상률(가정)', 0.047, '2026년 SW기술자 노임단가 전년대비 +4.7%(재게시 자료). KOSA 2026년 적용 원문 공표치로 교체 권장', '0.0%'),
    ('제경비율', 1.10, 'SW사업 대가산정 가이드: 직접인건비의 110~120%', '0%'),
    ('기술료율', 0.20, 'SW사업 대가산정 가이드: (직접인건비+제경비)의 20~40%', '0%'),
    ('부가가치세율', 0.10, '', '0%'),
    ('27년 SW 예산(백만원, VAT포함)', 758, '27년 신규사업 정부·국회 심의자료 S/W 예산 합계', '#,##0'),
]
A = {}
for name, val, src, fmt in ASSUME:
    put(ov, r, 1, name)
    c = put(ov, r, 2, val, color=BLUE, fill=INPUT_FILL, fmt=fmt, align=CENTER)
    put(ov, r, 3, src)
    A[name] = f"'0.개요'!$B${r}"
    r += 1
put(ov, r, 1, '직접인건비 단가(백만원/MM, 2027 적용 추정)', bold=True)
put(ov, r, 2, f"={A[ASSUME[0][0]]}*(1+{A[ASSUME[1][0]]})", fmt='0.00', align=CENTER)
put(ov, r, 3, '2025 월평균임금 × (1+인상률). 27년 착수 시 2027년 적용 공표치 반영 필요')
DL = f"'0.개요'!$B${r}"; r += 1
put(ov, r, 1, '투입 M/M 단가(백만원/MM, VAT별도)', bold=True)
put(ov, r, 2, f"={DL}*(1+{A['제경비율']})*(1+{A['기술료율']})", fmt='0.00', align=CENTER, fill=TOT_FILL)
put(ov, r, 3, '직접인건비 × (1+제경비율) × (1+기술료율)')
RATE = f"'0.개요'!$B${r}"; r += 1
VAT = A['부가가치세율']; BUD = A['27년 SW 예산(백만원, VAT포함)']
OV_SUM_ROW = r + 2

# ================= Sheet: 과제-과업 매칭 =================
ws = wb.create_sheet('1.과제-과업 매칭')
title(ws, '① 과제-과업 정리 및 ② 27년 예산 매칭', '매칭구분: ● 일치(27예산 반영) / ◐ 부분일치(범위 확대·구체화 필요) / ★ 기획안에만 있음(신규) / ▲ 27예산에만 있음(기획안 3단계)')
hdr(ws, 4, ['과제', '과제명', '과업ID', '과업명', '고도화 기획안 근거', '대응 27예산 SW항목', '27예산 비고상 대응내용', '매칭구분'],
    [8, 26, 7, 30, 42, 22, 38, 18])
r = 5
for tcode, tname, works in TASKS:
    start = r
    for wid, wname, src, k27, d27, m, _ in works:
        fill = FILLS.get(m)
        put(ws, r, 3, wid, align=CENTER, fill=fill); put(ws, r, 4, wname, fill=fill, bold=(m == M_NEW))
        put(ws, r, 5, src, fill=fill); put(ws, r, 6, k27, fill=fill); put(ws, r, 7, d27, fill=fill)
        put(ws, r, 8, m, align=CENTER, fill=fill, bold=True)
        r += 1
    for col, v in ((1, tcode), (2, tname)):
        ws.merge_cells(start_row=start, start_column=col, end_row=r - 1, end_column=col)
        put(ws, start, col, v, bold=True, align=CENTER)
        for rr in range(start + 1, r):
            ws.cell(row=rr, column=col).border = BORDER
start = r
for wid, wname, k27, note in ONLY27:
    fill = FILLS[M_27]
    put(ws, r, 3, wid, align=CENTER, fill=fill); put(ws, r, 4, wname, fill=fill)
    put(ws, r, 5, note, fill=fill); put(ws, r, 6, k27, fill=fill)
    put(ws, r, 7, next(b[2] for b in B27 if b[0] == k27), fill=fill)
    put(ws, r, 8, M_27, align=CENTER, fill=fill, bold=True)
    r += 1
ws.merge_cells(start_row=start, start_column=1, end_row=r - 1, end_column=1)
ws.merge_cells(start_row=start, start_column=2, end_row=r - 1, end_column=2)
put(ws, start, 1, '(참고)', bold=True, align=CENTER); put(ws, start, 2, '업무지원 Agent·전자결재 연계\n(27예산 반영, 기획안 3단계 \'28년)', bold=True, align=CENTER)
for rr in range(start + 1, r):
    ws.cell(row=rr, column=1).border = BORDER; ws.cell(row=rr, column=2).border = BORDER
r += 1
put(ws, r, 1, '매칭 요약', bold=True, border=False); r += 1
hdr_row = r
for i, h in enumerate(['구분', '과업 수'], 1):
    c = ws.cell(row=r, column=i + 1, value=h); c.font = F(bold=True, color='FFFFFF'); c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
r += 1
last = r - 4
for m in (M_EQ, M_PT, M_NEW, M_27):
    put(ws, r, 2, m, fill=FILLS[m]); put(ws, r, 3, f'=COUNTIF($H$5:$H${start + len(ONLY27) - 1},B{r})', align=CENTER)
    r += 1
ws.freeze_panes = 'C5'

# ================= Sheet: SW 구성 =================
sw = wb.create_sheet('2.SW 구성')
title(sw, '③ 고도화 기획안에 맞는 SW 표시', '27예산 SW 8개 항목의 기획안 부합 여부 + 기획안 수행에 추가로 필요한 SW(상용/오픈소스)')
put(sw, 4, 1, 'A. 27년 SW 예산 항목별 고도화 기획안 부합 여부', bold=True, border=False)
hdr(sw, 5, ['No', '27예산 SW 항목', '27예산(백만원)', '기획안 부합', '대응 기획안 과제·과업', '판단 근거/조치'],
    [5, 30, 13, 16, 36, 60])
FIT = {
    K_MOB: ('● 부합', '과제2 (2-1,2-3,2-4)', '모바일 UI/UX·조회 최적화는 일치. 기획안은 "스마트온 2.0 연계"로 구체화 → 연계 범위 명시 필요'),
    K_RAG: ('● 부합', '과제1 (1-2,1-3,1-5), 과제4 (4-1,4-2)', '수집·전처리·메타데이터·청킹·리랭킹 모두 기획안 과제1·4와 일치'),
    K_ANS: ('● 부합', '과제4 (4-3,4-4,4-7)', '근거제시·Prompt 최적화 일치. 기획안의 평가셋·정기 성능평가는 미포함 → 신규 반영'),
    K_AGT: ('▲ 미부합(3단계)', '과제3-4 일부(VOC 분류)', '기획안 로드맵상 AI 에이전트는 3단계(\'28년). VOC 분류 고도화만 대국민 VOC 이관에 활용, 나머지는 \'28년 이연 권장'),
    K_APR: ('▲ 미부합(3단계)', '-', '기획안 로드맵상 전자결재 연동은 3단계(\'28년) → 27년 예산은 기획안 신규 과업으로 재배분 권장'),
    K_DAT: ('● 부합', '과제1 (1-4,1-5,1-7)', '최신화·버전관리·관리자 기능·민감정보 필터링 일치'),
    K_DRM: ('● 부합', '과제1-2, 과제3-2', '내부 DRM 문서 수집 시 복호화 필수, 대국민 DB 분리 시 반출통제'),
    K_SSO: ('◐ 부분 부합', '과제2-2', '27예산은 "기존 어드바이저 로그인" 범위 → 스마트온 2.0 계정·인증 연계까지 확대 필요'),
}
r = 6
for i, (k, amt, note) in enumerate(B27, 1):
    fit, wk, why = FIT[k]
    fill = 'E2EFDA' if fit.startswith('●') else ('FFF2CC' if fit.startswith('◐') else 'EDEDED')
    put(sw, r, 1, i, align=CENTER); put(sw, r, 2, k); put(sw, r, 3, amt, color=BLUE, fmt=NUM0, align=CENTER)
    put(sw, r, 4, fit, bold=True, fill=fill, align=CENTER); put(sw, r, 5, wk); put(sw, r, 6, why)
    r += 1
put(sw, r, 2, '합계', bold=True, fill=TOT_FILL); put(sw, r, 3, f'=SUM(C6:C{r-1})', bold=True, fmt=NUM0, fill=TOT_FILL, align=CENTER)
put(sw, r, 4, '부합 합계', bold=True, fill=TOT_FILL, align=CENTER)
put(sw, r, 5, f'=SUMIF(D6:D{r-1},"●*",C6:C{r-1})+SUMIF(D6:D{r-1},"◐*",C6:C{r-1})', bold=True, fmt=NUM0, fill=TOT_FILL)
put(sw, r, 6, '← 기획안 부합(●+◐) 금액 / 나머지는 3단계 이연 대상', fill=TOT_FILL)
r += 2
put(sw, r, 1, 'B. 고도화 기획안 수행에 필요한 SW(솔루션) 구성', bold=True, border=False); r += 1
hdr(sw, r, ['No', 'SW(솔루션)', '도입 방식', '기획안 필요성', '관련 과업', '27예산 반영 / 비고']); r += 1
SWLIST = [
    ('DRM 보안소프트웨어', '상용 구매', '● 필수', '1-2, 3-2', '반영(16백만원). 내부 DRM 문서 수집·복호화'),
    ('SSO 인증서버소프트웨어', '상용 구매', '● 필수', '2-2', '반영(22백만원). 스마트온 2.0 계정연계 지원 여부 확인'),
    ('STT(음성인식) 엔진', '상용 구매(설치형)', '★ 신규 필요', '2-5', '미반영 → 견적 30백만원(가정) 신규 반영'),
    ('대국민 웹 챗봇(UI 위젯)', '개발', '★ 신규 필요', '3-1, 3-3', '미반영 → 인건비로 개발 반영(별도 챗봇 솔루션 구매 불요)'),
    ('개인정보·민감정보 탐지/마스킹', '개발(공통모듈)', '● 필요', '1-7, 3-5', '27 데이터관리 항목에 기능개발로 반영. 상용 도입 시 별도 산정'),
    ('문서 파싱(HWP/PDF/표) 엔진', '오픈소스/기존 활용', '◐ 권장', '1-2', '1차 구축 파서 고도화. 스캔문서 비중 높을 경우 OCR 상용 검토'),
    ('Reranker·하이브리드 검색', '오픈소스', '● 필요', '4-1', '라이선스 비용 없음(1차 구축 벡터DB 활용)'),
    ('RAG 평가 도구', '오픈소스', '★ 신규 필요', '4-5, 4-6', '라이선스 비용 없음(평가 파이프라인 개발비에 포함)'),
    ('외부망(DMZ) 서비스 인프라·웹방화벽', 'H/W·보안 장비', '★ 신규 검토', '과제3', '27 H/W(476)는 내부망 기준 → 대국민 서비스용 망분리 인프라 별도 검토 필요(본 SW 견적 제외)'),
    ('Agent 프레임워크 / 전자결재 연계 모듈', '개발', '▲ 기획안 3단계', '-', '27예산 반영(260)이나 기획안상 \'28년 → 27년 범위 제외'),
]
for i, row in enumerate(SWLIST, 1):
    fill = {'●': 'E2EFDA', '◐': 'FFF2CC', '★': 'FCE4D6', '▲': 'EDEDED'}[row[2][0]]
    put(sw, r, 1, i, align=CENTER)
    for j, v in enumerate(row, 2):
        put(sw, r, j, v, fill=fill if j == 4 else None, bold=(j == 4), align=CENTER if j in (3, 4, 5) else WRAP)
    r += 1

# ================= Sheet: 견적서 =================
es = wb.create_sheet('3.견적서')
title(es, '④ 과업 및 세부기능별 예산 견적서', '단위: 백만원 / 인건비 = 투입M/M × M/M단가(0.개요 시트) / SW구매·직접경비 = 수량 × 단가(파란색 입력값)')
hdr(es, 4, ['과제', '과업ID', '과업명', '세부기능', '매칭구분', '비목', '수량', '단위', '단가', '금액(VAT별도)', '27예산 대응항목', '산출근거/비고'],
    [8, 7, 26, 44, 16, 9, 7, 6, 9, 12, 22, 44])
r = 5
TYPE = {'L': ('인건비', 'M/M'), 'S': ('SW구매', '식'), 'E': ('직접경비', '식')}
task_total_cells = []
FIRST = r
for tcode, tname, works in TASKS:
    # 과제 header
    for col in range(1, 13):
        put(es, r, col, None, fill=SUB_FILL)
    put(es, r, 1, tcode, bold=True, fill=SUB_FILL, align=CENTER); put(es, r, 3, tname, bold=True, fill=SUB_FILL)
    trow = r; r += 1
    for wid, wname, src, k27, d27, m, items in works:
        for (fname, typ, qty, price, key, note) in items:
            nm, unit = TYPE[typ]
            put(es, r, 1, tcode, align=CENTER); put(es, r, 2, wid, align=CENTER); put(es, r, 3, wname)
            put(es, r, 4, fname); put(es, r, 5, m, fill=FILLS.get(m), align=CENTER)
            put(es, r, 6, nm, align=CENTER)
            put(es, r, 7, qty, color=BLUE, fmt='0.0', align=CENTER); put(es, r, 8, unit, align=CENTER)
            if typ == 'L':
                put(es, r, 9, f'={RATE}', color=GREEN, fmt='0.00', align=CENTER)
            else:
                put(es, r, 9, price, color=BLUE, fmt='0.0', align=CENTER)
            put(es, r, 10, f'=G{r}*I{r}', fmt=NUM1)
            put(es, r, 11, key); put(es, r, 12, note)
            r += 1
    put(es, trow, 7, f'=SUMIFS(G{trow+1}:G{r-1},F{trow+1}:F{r-1},"인건비")', bold=True, fmt='0.0', fill=SUB_FILL, align=CENTER)
    put(es, trow, 8, 'M/M', fill=SUB_FILL, align=CENTER)
    put(es, trow, 10, f'=SUM(J{trow+1}:J{r-1})', bold=True, fmt=NUM1, fill=SUB_FILL)
    task_total_cells.append((tcode, tname, trow))
LAST = r - 1
r += 1
# 합계부
def totrow(label, formula, fmt=NUM1, note=''):
    global r
    for col in range(1, 13): put(es, r, col, None, fill=TOT_FILL)
    put(es, r, 4, label, bold=True, fill=TOT_FILL); put(es, r, 10, formula, bold=True, fmt=fmt, fill=TOT_FILL)
    put(es, r, 12, note, fill=TOT_FILL)
    r += 1
    return r - 1
rows = {}
rows['L'] = totrow('인건비 소계', f'=SUMIFS(J{FIRST}:J{LAST},F{FIRST}:F{LAST},"인건비")')
rows['MM'] = totrow('총 투입 M/M', f'=SUMIFS(G{FIRST}:G{LAST},F{FIRST}:F{LAST},"인건비")', fmt='0.0')
rows['S'] = totrow('SW구매 소계', f'=SUMIFS(J{FIRST}:J{LAST},F{FIRST}:F{LAST},"SW구매")')
rows['E'] = totrow('직접경비 소계', f'=SUMIFS(J{FIRST}:J{LAST},F{FIRST}:F{LAST},"직접경비")')
rows['SUP'] = totrow('공급가액 합계', f"=J{rows['L']}+J{rows['S']}+J{rows['E']}")
rows['VAT'] = totrow('부가가치세', f"=J{rows['SUP']}*{VAT}")
rows['TOT'] = totrow('총 견적액 (VAT포함)', f"=J{rows['SUP']}+J{rows['VAT']}")
rows['BUD'] = totrow('27년 SW 예산', f"={BUD}", fmt=NUM0, note='정부·국회 심의자료')
rows['DIFF'] = totrow('예산 대비 차액 (예산 − 견적)', f"=J{rows['BUD']}-J{rows['TOT']}", note='(+) 잔액 / (−) 초과')
es.freeze_panes = 'E5'
ES_FIRST, ES_LAST = FIRST, LAST

# ================= Sheet: 27예산 대비 =================
cp = wb.create_sheet('4.27예산 대비')
title(cp, '27년 SW 예산 항목 ↔ 고도화 견적 대비표', '단위: 백만원(VAT포함). 견적액은 3.견적서의 "27예산 대응항목"별 합계 × (1+VAT)')
hdr(cp, 4, ['27예산 SW 항목', '27예산', '고도화 견적', '증감(견적−예산)', '조치'], [32, 12, 14, 16, 64])
ACT = {
    K_MOB: '스마트온 2.0 연계로 구체화(STT는 신규 항목으로 별도)',
    K_RAG: '유지',
    K_ANS: '유지(평가셋·정기평가는 신규 항목)',
    K_AGT: 'VOC 분류 재활용분만 반영, 보고서·규정분석 Agent는 \'28년(3단계) 이연',
    K_APR: '\'28년(3단계) 이연 → 기획안 신규 과업 재원으로 전환',
    K_DAT: '유지',
    K_DRM: '유지',
    K_SSO: '스마트온 계정연계까지 범위 확대',
    K_NEW: '기획안에만 있는 과업(전수조사·STT·대국민 챗봇·평가셋 등)',
    K_PM: '사업관리·통합테스트(27예산 각 항목에 내재되어 있던 공통비 명시)',
}
r = 5
keys = [(b[0], b[1]) for b in B27] + [(K_NEW, 0), (K_PM, 0)]
for k, amt in keys:
    put(cp, r, 1, k, bold=(k in (K_NEW, K_PM)))
    put(cp, r, 2, amt, color=BLUE, fmt=NUM0, align=CENTER)
    put(cp, r, 3, f"=SUMIFS('3.견적서'!$J${ES_FIRST}:$J${ES_LAST},'3.견적서'!$K${ES_FIRST}:$K${ES_LAST},A{r})*(1+{VAT})", color=GREEN, fmt=NUM1)
    put(cp, r, 4, f'=C{r}-B{r}', fmt=NUM1)
    put(cp, r, 5, ACT[k])
    r += 1
put(cp, r, 1, '합계', bold=True, fill=TOT_FILL)
for col, L in ((2, 'B'), (3, 'C'), (4, 'D')):
    put(cp, r, col, f'=SUM({L}5:{L}{r-1})', bold=True, fmt=NUM1, fill=TOT_FILL)
put(cp, r, 5, f"=IF(ABS(C{r}-'3.견적서'!J{rows['TOT']})<0.01,\"검증OK: 견적서 총액과 일치\",\"검증 오류: 견적서 총액과 불일치\")", fill=TOT_FILL)
CP_TOT = r

# ================= 개요 summary =================
r = OV_SUM_ROW
put(ov, r, 1, '■ 과제별 견적 요약 (단위: 백만원)', bold=True, border=False); r += 1
for i, h in enumerate(['과제', '투입 M/M', '금액(VAT별도)'], 1):
    c = ov.cell(row=r, column=i, value=h); c.font = F(bold=True, color='FFFFFF'); c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
r += 1
s0 = r
for tcode, tname, trow in task_total_cells:
    put(ov, r, 1, f'{tcode} {tname}')
    put(ov, r, 2, f"='3.견적서'!G{trow}", color=GREEN, fmt='0.0', align=CENTER)
    put(ov, r, 3, f"='3.견적서'!J{trow}", color=GREEN, fmt=NUM1)
    r += 1
for label, key in (('공급가액 합계', 'SUP'), ('부가가치세', 'VAT'), ('총 견적액(VAT포함)', 'TOT'), ('27년 SW 예산', 'BUD'), ('예산 대비 차액', 'DIFF')):
    put(ov, r, 1, label, bold=True, fill=TOT_FILL); put(ov, r, 2, None, fill=TOT_FILL)
    put(ov, r, 3, f"='3.견적서'!J{rows[key]}", color=GREEN, bold=True, fmt=NUM1, fill=TOT_FILL)
    r += 1
r += 1
put(ov, r, 1, '■ 범례', bold=True, border=False); r += 1
for m, desc in ((M_EQ, '27예산 SW항목과 고도화 기획안 과업이 일치'), (M_PT, '27예산에 있으나 기획안 기준으로 범위 확대·구체화 필요'),
                (M_NEW, '고도화 기획안에만 있는 과업(27예산 미반영, 신규)'), (M_27, '27예산에만 있는 과업(기획안 로드맵 3단계 \'28년)')):
    put(ov, r, 1, m, fill=FILLS[m], bold=True); put(ov, r, 2, None); put(ov, r, 3, desc); r += 1
put(ov, r, 1, '파란 글씨 / 노란 셀', color=BLUE, fill=INPUT_FILL); put(ov, r, 2, None); put(ov, r, 3, '수정 가능한 입력값(M/M, 단가, 요율). 수정 시 전 시트 자동 재계산'); r += 1
put(ov, r, 1, '초록 글씨', color=GREEN); put(ov, r, 2, None); put(ov, r, 3, '다른 시트 참조 수식'); r += 1
r += 1
put(ov, r, 1, '■ 유의사항', bold=True, border=False); r += 1
for t in ['1) M/M 단가는 KOSA 2025년 적용 평균임금에 2026년 인상률(4.7%)을 적용한 추정치이며, 착수 시점 공표치로 갱신 필요',
          '2) STT 엔진(30)·전문가 자문(8)은 가정치로 업체 견적 확인 필요. DRM·SSO는 27예산 금액(VAT포함)의 공급가 기준',
          '3) 업무지원 Agent(200)·전자결재 연계(60)는 기획안 로드맵상 3단계(\'28년)로 본 견적에서 제외하고 그 재원을 기획안 신규 과업에 재배분',
          '4) 대국민 챗봇(과제3) 운영을 위한 외부망(DMZ)·웹방화벽 등 H/W·보안 인프라는 27 H/W 예산(476, 내부망 기준)에 없어 별도 검토 필요']:
    ov.cell(row=r, column=1, value=t).font = F(); ov.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3); r += 1

for wsx in wb.worksheets:
    wsx.sheet_view.showGridLines = False
    wsx.page_setup.orientation = 'landscape'; wsx.page_setup.fitToWidth = 1; wsx.page_setup.fitToHeight = 0
    wsx.sheet_properties.pageSetUpPr.fitToPage = True
wb.save(OUT)
print('saved', OUT)
