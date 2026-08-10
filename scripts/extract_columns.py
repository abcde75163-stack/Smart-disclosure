"""
양식 파일 없이 추가 요청사항 기반으로 마스터 DB 열을 그대로 추출하는 스크립트

동작:
  - AI가 notes를 읽고 필요한 열(db_index)과 필터 조건을 결정
  - 마스터 DB에서 해당 열을 그대로 추출 → Excel로 저장
"""
import json, os, sys, datetime
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.insert(0, os.path.dirname(__file__))
from common import load_master_db, filter_by_year, to_yyyymmdd, get_year, safe_int


# ── DB 전체 컬럼 목록 (AI 참조용) ──────────────────────────────────
DB_COLUMNS_SUMMARY = """
마스터 DB 컬럼 목록 (db_index → 헤더명):
0=연번, 1=기술이전계약일, 2=기술보유기관명, 3=기술도입업체명,
4=기관유형(코드), 5=업종유형, 6=국내외(1=국내,2=국외), 7=국가명,
8=국내지역구분(지역코드), 9=사업자등록번호, 10=대표주소, 11=대표전화,
12=대표자성명, 13=홈페이지, 14=우편주소, 15=기술이전담당자명,
16=담당부서, 17=직급, 18=핸드폰, 19=팩스, 20=이메일,
21=산단담당자명, 22=산단담당부서, 23=산단전화번호, 24=산단이메일,
25=종업원수(상시), 26=연매출액(천원), 27=기술명, 28=기술명(27.기술명과 동일계열),
29=주발명자, 30=교직원번호, 31=주발명자소속, 32=전공,
33=공동발명자, 34=공동발명자교직원번호, 35=공동발명자소속, 36=공동발명자전공,
37=기술유형(코드), 38=지식재산권번호, 39=출원등록상태(1=출원,2=등록),
40=포함기술수, 41=특허비용, 42=기술분야(6T), 43=기술분류,
44=거래유형(코드), 45=계약기간, 46=계약시작일, 47=계약종료일,
48=제한사항, 49=기술료수취유형코드, 50=선급기술료(원), 51=경상기술료조건,
52=정액기술료(원), 53=총기술료(원), 54=계약상태(0=유지,1=해지,2=종료),
55=계약해지사유, 56=연구과제명, 57=부처명, 58=지원기관, 59=지원사업,
60=총연구비(천원), 61=총연구기간, 62=협약일, 63=대사업명, 64=중사업명,
65=지원기관과제번호, 66=ERP과제번호, 67=연구책임자, 68=공동연구자,
69=참여기업명, 70=정부출연금, 71=기업부담금, 72=기타,
73=입금일, 74=입금통장명의, 75=기술료수취유형,
76=현금입금액(원), 77=주식(원), 78=현물(원), 79=기타(원), 80=입금종류,
81=선급기술료분배액, 82=경상기술료분배액, 83=정액기술료분배액,
84=분배일, 85=제반비용_특허비용, 86=제반비용_중개수수료,
87=전문기관분배액, 88=발명자분배액, 89=발명자지정기관분배액,
90=산학협력단분배액, 91=기술이전사업화경비, 92=성과활용기여자보상금,
93=연구개발재투자, 94=기타특허지분액, 95=해당사업단명,
96=수납상황, 97=학진승인여부, 98=NTB등록여부, 99=기술가치평가여부,
100=계약변경일, 101=계약해지일, 102=납부기한, 103=담당자, 104=담당자(연구원)
"""

# ── 헤더명 매핑 (출력 시 사용) ─────────────────────────────────────
HEADER_MAP = {
    0: "연번", 1: "기술이전계약일", 2: "기술보유기관명", 3: "기술도입업체명",
    4: "기관유형", 5: "업종유형", 6: "국내/국외", 7: "국가명",
    8: "국내지역구분", 9: "사업자등록번호", 10: "대표주소", 11: "대표전화",
    12: "대표자성명", 13: "홈페이지", 14: "우편주소", 15: "기술이전담당자명",
    16: "담당부서", 17: "직급", 18: "핸드폰", 19: "팩스", 20: "이메일",
    21: "산단담당자명", 22: "산단담당부서", 23: "산단전화번호", 24: "산단이메일",
    25: "종업원수(상시)", 26: "연매출액(천원)", 27: "기술명", 28: "기술명",
    29: "주발명자", 30: "교직원번호", 31: "주발명자소속", 32: "전공",
    33: "공동발명자", 34: "공동발명자교직원번호", 35: "공동발명자소속", 36: "공동발명자전공",
    37: "기술유형", 38: "지식재산권번호", 39: "출원/등록상태",
    40: "포함기술수", 41: "특허비용", 42: "기술분야(6T)", 43: "기술분류",
    44: "거래유형", 45: "계약기간", 46: "계약시작일", 47: "계약종료일",
    48: "제한사항", 49: "기술료수취유형", 50: "선급기술료(원)", 51: "경상기술료조건",
    52: "정액기술료(원)", 53: "총기술료(원)", 54: "계약상태",
    55: "계약해지사유", 56: "연구과제명", 57: "부처명", 58: "지원기관", 59: "지원사업",
    60: "총연구비(천원)", 61: "총연구기간", 62: "협약일", 63: "대사업명", 64: "중사업명",
    65: "지원기관과제번호", 66: "ERP과제번호", 67: "연구책임자", 68: "공동연구자",
    69: "참여기업명", 70: "정부출연금", 71: "기업부담금", 72: "기타",
    73: "입금일", 74: "입금통장명의", 75: "기술료수취유형",
    76: "현금입금액(원)", 77: "주식(원)", 78: "현물(원)", 79: "기타(원)", 80: "입금종류",
    81: "선급기술료분배액", 82: "경상기술료분배액", 83: "정액기술료분배액",
    84: "분배일", 85: "제반비용(특허비용)", 86: "제반비용(중개수수료)",
    87: "전문기관분배액", 88: "발명자분배액", 89: "발명자지정기관분배액",
    90: "산학협력단분배액", 91: "기술이전사업화경비", 92: "성과활용기여자보상금",
    93: "연구개발재투자", 94: "기타(특허지분액)", 95: "해당사업단명",
    96: "수납상황", 97: "학진승인여부", 98: "NTB등록여부", 99: "기술가치평가여부",
    100: "계약변경일", 101: "계약해지일", 102: "납부기한", 103: "담당자", 104: "담당자(연구원)",
}

# 기본 추출 열 (요청사항이 없거나 불명확할 때)
DEFAULT_COLUMNS = [0, 1, 3, 28, 29, 37, 44, 50, 52, 73, 76]


def ask_claude(notes: str, year: int, api_key: str) -> dict:
    """
    추가 요청사항을 분석해서 추출할 열과 필터 조건을 반환.
    반환 형식:
    {
      "columns": [0, 1, 3, ...],          # db_index 목록 (순서 유지)
      "filter_mode": "contract"|"payment"|"both"|"none",
      "filter_year": 2025,                # null이면 year 파라미터 사용
      "notes": "분석 설명"
    }
    """
    client = anthropic.Anthropic(api_key=api_key)

    prompt = f"""당신은 부산대학교 산학협력단의 기술이전 담당자입니다.
아래 추가 요청사항을 분석해서 마스터 DB에서 어떤 열을 추출할지 결정해주세요.

## 기준 연도: {year}년

## 마스터 DB 컬럼 목록:
{DB_COLUMNS_SUMMARY}

## 사용자 추가 요청사항:
{notes if notes.strip() else "(없음 - 기본 주요 항목을 추출)"}

## 응답 형식 (JSON만 출력, 다른 텍스트 없이):
{{
  "columns": [정수 db_index 목록],
  "filter_mode": "contract" 또는 "payment" 또는 "both" 또는 "none",
  "filter_year": 연도_정수_또는_null,
  "notes": "선택 이유 한 줄"
}}

## 판단 기준:
- 요청사항에 특정 열이 언급되면 해당 열 포함
- 연도 필터: 계약일 기준이면 "contract", 입금일 기준이면 "payment", 둘 다면 "both", 연도 무관이면 "none"
- filter_year: 요청에 다른 연도가 있으면 그 연도, 없으면 null (기준 연도 {year} 사용)
- 요청이 없거나 불명확하면 기본 주요 열 선택: [0,1,3,28,29,37,44,50,52,73,76]
- 열 순서는 db_index 오름차순으로 정렬
- columns에는 중복 없이 정수만"""

    resp = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=512,
        messages=[{"role": "user", "content": prompt}],
    )
    text = resp.content[0].text.strip()

    # JSON 블록 추출
    import re
    m = re.search(r'\{[\s\S]*\}', text)
    if m:
        return json.loads(m.group())
    raise ValueError(f"AI 응답에서 JSON을 파싱할 수 없습니다:\n{text}")


def format_cell_value(value, db_index: int):
    """셀 값 표시용 포맷."""
    if value is None:
        return ""
    # 날짜 컬럼
    if db_index in (1, 46, 62, 73, 84, 100, 101, 102):
        return to_yyyymmdd(value)
    # 숫자 컬럼 (금액 등)
    if db_index in (50, 52, 53, 60, 70, 71, 76, 77, 78, 79,
                    81, 82, 83, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94):
        try:
            n = int(float(str(value)))
            return n
        except:
            pass
    return value


def run(master_path: str, year: int, output_path: str,
        notes: str, api_key: str) -> dict:
    """
    메인 실행 함수.
    반환: {"count": N, "columns": [...], "total_income": N, "ai_notes": "..."}
    """
    print(f"[extract_columns] 마스터 DB 로딩: {master_path}")
    all_rows = load_master_db(master_path)
    print(f"  전체 행수: {len(all_rows)}")

    # ── AI에게 열 및 필터 결정 요청 ────────────────────────────────
    print("[extract_columns] AI 분석 중...")
    try:
        ai_result = ask_claude(notes, year, api_key)
        columns    = ai_result.get("columns") or DEFAULT_COLUMNS
        filter_mode = ai_result.get("filter_mode", "both")
        filter_year = ai_result.get("filter_year") or year
        ai_notes   = ai_result.get("notes", "")
        print(f"  AI 결정 - 열: {columns}, 필터: {filter_mode}, 연도: {filter_year}")
        print(f"  AI 설명: {ai_notes}")
    except Exception as e:
        print(f"  AI 분석 실패 ({e}), 기본값 사용")
        columns     = DEFAULT_COLUMNS
        filter_mode = "both"
        filter_year = year
        ai_notes    = "AI 분석 실패 → 기본 열 사용"

    # 유효 열 인덱스만 (0~104)
    columns = sorted(set(c for c in columns if 0 <= c <= 104))

    # ── 필터링 ──────────────────────────────────────────────────────
    if filter_mode == "none":
        rows = all_rows
        print(f"  필터 없음 → {len(rows)}행")
    else:
        rows = filter_by_year(all_rows, filter_year,
                              contract_col=1, payment_col=73, mode=filter_mode)
        print(f"  {filter_year}년 필터({filter_mode}) → {len(rows)}행")

    # ── Excel 출력 ──────────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = f"{filter_year}년 추출"

    # 헤더 스타일
    header_fill  = PatternFill("solid", fgColor="1F3864")
    header_font  = Font(bold=True, color="FFFFFF", size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align   = Alignment(horizontal="left",   vertical="center")
    thin_border  = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin"),
    )

    # 헤더 행
    for col_pos, db_idx in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_pos, value=HEADER_MAP.get(db_idx, f"col_{db_idx}"))
        cell.fill      = header_fill
        cell.font      = header_font
        cell.alignment = center_align
        cell.border    = thin_border

    ws.row_dimensions[1].height = 30

    # 데이터 행
    for row_pos, row in enumerate(rows, start=2):
        for col_pos, db_idx in enumerate(columns, start=1):
            raw_val = row[db_idx] if db_idx < len(row) else None
            val = format_cell_value(raw_val, db_idx)
            cell = ws.cell(row=row_pos, column=col_pos, value=val)
            cell.border    = thin_border
            cell.alignment = left_align
            cell.font      = Font(size=9)

    # 열 너비 자동 조정 (최대 40)
    for col_pos, db_idx in enumerate(columns, start=1):
        header_len = len(HEADER_MAP.get(db_idx, "")) + 2
        max_len = header_len
        for row_pos in range(2, min(len(rows) + 2, 52)):  # 최대 50행 샘플
            cell_val = ws.cell(row=row_pos, column=col_pos).value
            if cell_val:
                max_len = max(max_len, min(len(str(cell_val)), 40))
        ws.column_dimensions[ws.cell(row=1, column=col_pos).column_letter].width = max_len + 2

    # 틀 고정
    ws.freeze_panes = "A2"

    # 총 입금액 계산 (76번 열이 포함된 경우)
    total_income = 0
    if 76 in columns:
        for row in rows:
            if 76 < len(row):
                total_income += safe_int(row[76])

    wb.save(output_path)
    print(f"[extract_columns] 저장 완료: {output_path} ({len(rows)}건)")

    return {
        "count": len(rows),
        "columns": columns,
        "total_income": total_income,
        "filter_mode": filter_mode,
        "filter_year": filter_year,
        "ai_notes": ai_notes,
    }