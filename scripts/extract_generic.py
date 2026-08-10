"""
양식 기반 AI 자동 추출기 — 모든 양식에 대해 작동

핵심 원칙:
  양식 파일 안에 있는 가이드 지침을 AI가 직접 읽고 그에 따라 추출한다.
  하드코딩된 매핑 없음. 양식 파일이 곧 명세서.
"""
import json, sys, os, re, datetime
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from feedback import write_feedback_sheet
from analyze_template import analyze
from openpyxl import load_workbook


def parse_notes_filters(notes: str) -> dict:
    """
    추가 요청사항 텍스트에서 날짜 범위·필터 기준·연구자 목록을 파싱.
    반환: {"date_start": "YYYYMMDD"|None, "date_end": "YYYYMMDD"|None,
           "date_mode": "contract"|"payment"|"both",
           "researchers": [이름,...] | None}
    """
    result = {"date_start": None, "date_end": None,
              "date_mode": "both", "researchers": None}
    if not notes:
        return result

    # 날짜 범위: 20260101~20260801 / 20260101-20260801 / 20260101 ~ 20260801
    m = re.search(r'(20\d{6})\s*[~\-~]\s*(20\d{6})', notes)
    if m:
        result["date_start"] = m.group(1)
        result["date_end"]   = m.group(2)

    # 필터 기준
    if "계약일" in notes:
        result["date_mode"] = "contract"
    elif "입금일" in notes:
        result["date_mode"] = "payment"

    # 연구자 목록: "대상 연구자: 홍길동, 김철수" 또는 "연구자: ..."
    rm = re.search(r'(?:대상\s*)?연구자\s*[:：]\s*(.+)', notes)
    if rm:
        names_raw = rm.group(1).strip()
        names = [n.strip() for n in re.split(r'[,，/\n]', names_raw) if n.strip()]
        if names:
            result["researchers"] = names

    return result

# ── 매핑 캐시 ──────────────────────────────────────────────────────
_CACHE_PATH = os.path.join(os.path.dirname(__file__), "..", "references", "mapping_cache.json")

def _normalize_key(filename: str) -> str:
    """파일명 → 캐시 조회용 정규화 키."""
    name = os.path.splitext(os.path.basename(filename))[0].lower()
    for w in ["자동추출", "수정안내", "복사본", "최종", "수정", "제출용", "붙임", "별첨"]:
        name = name.replace(w, "")
    name = re.sub(r'20\d{2}', '', name)          # 연도 제거
    name = re.sub(r'\d{4,}', '', name)            # 긴 숫자 제거
    name = re.sub(r'[\s_\-\(\)\[\]\.]+', ' ', name).strip()
    return name

def load_cache() -> dict:
    try:
        with open(_CACHE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def find_cached_mapping(template_filename: str, cache: dict) -> tuple:
    """파일명으로 캐시에서 가장 유사한 매핑 탐색. (key, mapping) 또는 (None, None)."""
    key = _normalize_key(template_filename)
    best_key, best_mapping, best_score = None, None, 0
    for cache_key, entry in cache.items():
        if cache_key.startswith("_"):
            continue
        mapping = entry.get("mapping") if isinstance(entry, dict) else None
        if not mapping:
            continue
        # 부분 문자열 매칭 점수
        score = 0
        for part in cache_key.split():
            if part and part in key:
                score += len(part)
        for part in key.split():
            if part and part in cache_key:
                score += len(part)
        if score > best_score:
            best_score, best_key, best_mapping = score, cache_key, mapping
    if best_score >= 4:   # 최소 4자 이상 겹쳐야 매칭으로 인정
        return best_key, best_mapping
    return None, None

def save_mapping_to_cache(template_filename: str, mapping: dict, cache: dict,
                          example_rows: list = None) -> dict:
    """추출 성공 후 매핑(+수동 검증 예시 행)을 캐시에 추가."""
    key = _normalize_key(template_filename)
    entry = cache.get(key, {})
    entry.update({
        "saved_at": datetime.datetime.now().strftime("%Y-%m-%d"),
        "template_hint": template_filename,
        "mapping": mapping,
    })
    if example_rows is not None:
        entry["example_rows"] = example_rows   # 수동 검증된 올바른 출력 행
    cache[key] = entry
    return cache


def extract_example_rows(corrected_file_bytes: bytes, target_sheet: str = None,
                          data_start: int = 4, n_rows: int = 5) -> list:
    """수정된 결과 파일에서 헤더 + 데이터 예시 행 추출."""
    import io
    wb = load_workbook(io.BytesIO(corrected_file_bytes), data_only=True)
    ws = None
    if target_sheet and target_sheet in wb.sheetnames:
        ws = wb[target_sheet]
    else:
        ws = wb.active

    header_row = data_start - 1
    rows_out = []
    for r in range(header_row, min(header_row + n_rows + 1, (ws.max_row or 0) + 1)):
        cells = {}
        for c in range(1, (ws.max_column or 0) + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip():
                cells[str(c)] = str(v).strip()
        if cells:
            rows_out.append({"row": r, "cells": cells, "is_header": r == header_row})
    return rows_out

# ── 지원 transform 함수 ────────────────────────────────────────────
TRANSFORMS = {
    "str":               lambda row, idx: str(row[idx]) if row[idx] is not None else "",
    "date_yyyymmdd":     lambda row, idx: to_yyyymmdd(row[idx]),
    "date_yyyymm":       lambda row, idx: to_yyyymm(row[idx]),
    "date_year":         lambda row, idx: str(get_year(row[idx])) if get_year(row[idx]) else "",
    "date_obj":          lambda row, idx: row[idx],
    "safe_int":          lambda row, idx: safe_int(row[idx]),
    "won_to_million":    lambda row, idx: round(safe_int(row[idx]) / 1_000_000, 1) if safe_int(row[idx]) else None,
    "won_to_thousand":   lambda row, idx: safe_int(row[idx]) // 1_000 if safe_int(row[idx]) else None,
    "region":            lambda row, idx: get_region_name(row[8], row[6]),
    "org_type":          lambda row, idx: ORG_TYPE_MAP.get(row[idx], str(row[idx]) if row[idx] else ""),
    "trade_type":        lambda row, idx: TRADE_TYPE_MAP.get(row[idx], str(row[idx]) if row[idx] else ""),
    "domestic_foreign":  lambda row, idx: "국외" if str(row[6]).strip() in ("2", "국외") else "국내",
    "bridge_type":       lambda row, idx: get_bridge_type(row[37], row[44]) or "",
}


def run_free_format(all_rows, template_path, mapping, year, output_path):
    """
    연구자/교수 명단에서 이름 추출 → 해당 연구자의 기술이전 실적을 자유 양식으로 생성.
    form_type='researcher_list'일 때 호출됨.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    # ── 연구자 명단 읽기 ──────────────────────────────────────────────
    wb_form   = load_workbook(template_path, data_only=True)
    res_sheet = mapping.get("researcher_sheet")
    ws_form   = wb_form[res_sheet] if res_sheet and res_sheet in wb_form.sheetnames else wb_form.active

    res_col   = mapping.get("researcher_col", 2)
    res_start = mapping.get("researcher_data_start", 2)

    researcher_names = []
    for r in range(res_start, (ws_form.max_row or 0) + 1):
        v = ws_form.cell(r, res_col).value
        if v and str(v).strip():
            name = str(v).strip()
            if name and name not in researcher_names:
                researcher_names.append(name)

    print(f"  👥 명단에서 연구자 {len(researcher_names)}명 추출: {researcher_names[:5]}")

    # ── 연구자 이름으로 마스터 DB 필터링 ────────────────────────────
    def researcher_match(row):
        investigator = str(row[29]).strip() if row[29] else ""
        if not investigator:
            return False
        return any(
            (len(name) >= 2) and (name in investigator or investigator in name)
            for name in researcher_names
        )

    matched_all = [r for r in all_rows if researcher_match(r)]

    # 연도 필터 시도 → 없으면 전체 기간
    year_filtered = filter_by_year(matched_all, year, mode="both")
    if year_filtered:
        rows_to_use = year_filtered
        year_label  = f"{year}년"
    else:
        rows_to_use = matched_all
        year_label  = "전체 기간"
        print(f"  ⚠️ {year}년 기준 결과 없음 → 전체 기간으로 확장")

    print(f"  → 필터링 결과: {len(rows_to_use)}건 ({year_label})")

    # ── 계약 단위로 집계 (동일 연번 입금 회차 합산) ─────────────────
    contracts = {}
    for row in rows_to_use:
        seq = str(row[0]) if row[0] else ""
        if not seq:
            continue
        payment = safe_int(row[76])
        if seq not in contracts:
            contracts[seq] = {"row": row, "total_pay": 0}
        contracts[seq]["total_pay"] += payment

    sorted_contracts = sorted(
        contracts.values(),
        key=lambda x: x["row"][1]
        if isinstance(x["row"][1], (datetime.datetime, datetime.date))
        else datetime.datetime(1900, 1, 1),
    )

    # ── 새 워크북 생성 ────────────────────────────────────────────────
    wb_out = Workbook()
    ws = wb_out.active
    ws.title = "기술이전 실적"

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill  = PatternFill("solid", fgColor="1F3864")
    header_font  = Font(bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    data_center = Alignment(horizontal="center", vertical="center")
    data_left   = Alignment(horizontal="left", vertical="center", wrap_text=True)

    headers = [
        ("순번",                   6),
        ("연구자명",               12),
        ("기술명",                 30),
        ("기술도입업체(지역)",     22),
        ("계약일",                 12),
        ("총기술료계약액\n(백만원)", 14),
        (f"{year_label}\n입금액(백만원)", 14),
        ("계약형태",               10),
        ("중대형여부\n(1억 이상)", 10),
    ]

    # 제목 행 (1행)
    title_cell = ws.cell(1, 1, f"기술이전 실적 현황 ({year_label})")
    title_cell.font = Font(bold=True, size=12, color="1F3864")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # 헤더 행 (2행)
    ws.row_dimensions[2].height = 36
    for c, (h, w) in enumerate(headers, 1):
        cell = ws.cell(2, c, h)
        cell.font  = header_font
        cell.fill  = header_fill
        cell.alignment = header_align
        cell.border = border
        ws.column_dimensions[cell.column_letter].width = w

    # 데이터 행 (3행~)
    total_income = 0
    for idx, info in enumerate(sorted_contracts, 1):
        row      = info["row"]
        pay_total = info["total_pay"]
        total_income += pay_total

        company    = str(row[3]).strip() if row[3] else ""
        is_foreign = str(row[6]).strip() in ("2", "국외")
        region     = "해외" if is_foreign else get_region_name(row[8])
        company_str = f"{company}({region})" if region else company

        contract_m = round((safe_int(row[52]) + safe_int(row[50])) / 1_000_000, 1)
        pay_m      = round(pay_total / 1_000_000, 1) if pay_total else None

        row_num = idx + 2
        ws.row_dimensions[row_num].height = 20

        data_cells = [
            (idx,                                        data_center),
            (str(row[29]).strip() if row[29] else "",   data_center),
            (str(row[28]).strip() if row[28] else "",   data_left),
            (company_str,                                data_left),
            (to_yyyymmdd(row[1]),                        data_center),
            (contract_m if contract_m else None,         data_center),
            (pay_m,                                      data_center),
            (get_contract_type(row),                     data_center),
            ("Y" if pay_total >= 100_000_000 else "N",  data_center),
        ]
        for c, (val, align) in enumerate(data_cells, 1):
            cell = ws.cell(row_num, c, val)
            cell.alignment = align
            cell.border    = border
            if c in (6, 7) and isinstance(val, float):
                cell.number_format = "0.0"

    # 합계 행
    if sorted_contracts:
        sum_row  = len(sorted_contracts) + 3
        sum_fill = PatternFill("solid", fgColor="D6E4FF")
        sum_font = Font(bold=True)
        ws.row_dimensions[sum_row].height = 22
        for c in range(1, len(headers) + 1):
            cell = ws.cell(sum_row, c)
            cell.fill      = sum_fill
            cell.font      = sum_font
            cell.border    = border
            cell.alignment = data_center
        ws.cell(sum_row, 1, "합계")
        ws.cell(sum_row, 6).number_format = "0.0"
        ws.cell(sum_row, 7).number_format = "0.0"
        ws.cell(sum_row, 6, round(
            sum(safe_int(i["row"][52]) + safe_int(i["row"][50]) for i in sorted_contracts) / 1_000_000, 1
        ))
        ws.cell(sum_row, 7, round(total_income / 1_000_000, 1) if total_income else None)

    wb_out.save(output_path)
    print(f"\n✅ 자유 양식 저장 완료: {output_path}")
    print(f"   연구자 {len(researcher_names)}명 | 실적 {len(sorted_contracts)}건 | 총입금 {total_income:,}원")

    return {
        "count":            len(sorted_contracts),
        "total_income":     total_income,
        "mapping_notes":    f"연구자 명단 {len(researcher_names)}명 기준 자유 양식 추출 ({year_label})",
        "manual_inputs":    [],
        "filter_mode":      year_label,
        "mapping":          mapping,
        "cache_key":        None,
        "form_type":        "researcher_list",
        "researcher_count": len(researcher_names),
    }


def get_mapping_from_claude(template_analysis, db_columns_text, transform_rules_text, notes, year, hint, api_key=None, cached_mapping=None):
    """
    양식 파일 전체 내용을 AI에 전달하고 컬럼 매핑 JSON을 받아온다.
    cached_mapping: 유사 양식에서 성공한 이전 매핑 (참고용으로 프롬프트에 주입)
    """
    try:
        import anthropic
    except ImportError:
        raise ImportError("pip install anthropic 필요")

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    # 양식 내용을 읽기 쉽게 정리
    # 헤더 행을 찾아 강조 표시 (가장 많은 셀이 채워진 행 = 헤더)
    form_text_parts = []
    for sheet in template_analysis.get("sheets", []):
        rows = sheet.get("rows", [])
        form_text_parts.append(f"\n### 시트: {sheet['name']} (최대행:{sheet.get('max_row','?')}, 최대열:{sheet.get('max_col','?')})")

        # 헤더 행 탐지: 셀이 가장 많은 행 (보통 헤더)
        header_row_num = None
        max_cells = 0
        for row_info in rows[:10]:  # 앞쪽 10행에서 탐색
            if len(row_info["cells"]) > max_cells:
                max_cells = len(row_info["cells"])
                header_row_num = row_info["row"]

        for row_info in rows:
            cells_str = ", ".join(f"[열{k}]{v}" for k, v in row_info["cells"].items())
            if row_info["row"] == header_row_num:
                form_text_parts.append(f"  ★헤더행{row_info['row']}(이 행의 모든 열을 반드시 매핑): {cells_str}")
            elif row_info["row"] > (header_row_num or 0):
                form_text_parts.append(f"  데이터행{row_info['row']}(형식참고/비어있을수있음): {cells_str}")
            else:
                form_text_parts.append(f"  행{row_info['row']}: {cells_str}")
        if sheet.get("truncated"):
            form_text_parts.append("  (이하 행 생략)")
    form_full_text = "\n".join(form_text_parts)

    if cached_mapping:
        cached_section = (
            "이전에 유사한 양식에서 성공적으로 사용된 매핑입니다. "
            "현재 양식의 헤더와 지침을 우선하되, 구조가 비슷하면 참고하세요.\n"
            f"```json\n{json.dumps(cached_mapping, ensure_ascii=False, indent=2)}\n```"
        )
        # 수동 검증된 예시 행이 있으면 추가
        example_rows = None
        if cached_mapping:
            cache = load_cache()
            key = _normalize_key(template_analysis.get("_filename", ""))
            _, found_mapping = find_cached_mapping(template_analysis.get("_filename", ""), cache)
            for ck, ce in cache.items():
                if isinstance(ce, dict) and ce.get("mapping") == cached_mapping:
                    example_rows = ce.get("example_rows")
                    break
        if example_rows:
            rows_text = []
            for row_info in example_rows:
                tag = "헤더" if row_info.get("is_header") else "데이터"
                cells_str = ", ".join(f"[열{k}]{v}" for k, v in row_info["cells"].items())
                rows_text.append(f"  {tag}행{row_info['row']}: {cells_str}")
            cached_section += (
                "\n\n### 수동 검증된 올바른 출력 예시 (이 형식을 정확히 따르세요)\n"
                + "\n".join(rows_text)
            )
    else:
        cached_section = "(없음 — 처음 보는 양식이므로 양식 파일 지침만 따르세요)"

    prompt = f"""당신은 기술이전 데이터 추출 전문가입니다.

아래에 제공된 양식 파일의 전체 내용을 읽고, 그 안에 있는 작성 지침과 예시를 따라 컬럼 매핑 JSON을 생성해주세요.

## 중요 원칙
- 양식 파일 내부의 가이드 지침, 예시 데이터, 컬럼 헤더에 명시된 형식을 최우선으로 따르세요.
- 헤더에 "(백만원)"이라고 쓰여 있으면 백만원 단위로, "(YYYYMMDD)"이면 8자리 날짜로, "(Yes/No)"이면 Y/N으로 변환하세요.
- 예시 행(XX대학교, 홍길동 등)의 형식을 정확히 따르세요. 예) "XX대학교(홍길동)" 형식이면 "부산대학교(연구자명)" 형식으로.
- 작성 지침에 인정기준이 있으면 그에 맞게 필터 조건을 설정하세요.

## 추출 연도
{year}년

## 이전 유사 양식의 성공 매핑 (참고용)
{cached_section}

## 양식 파일 전체 내용
{form_full_text}

## 마스터 DB 컬럼 목록
{db_columns_text}

## 변환 규칙
{transform_rules_text}

## 사용자 추가 요청사항
{notes if notes else "(없음)"}

## 키워드 힌트 (참고용)
{hint if hint else "(없음)"}

## 마스터 DB 주요 컬럼 참조 (db_index 혼동 방지용)
마스터 DB(기술이전총정리.xlsx)는 항상 동일한 구조입니다.
양식 파일에 해당 항목이 있을 때만 아래 인덱스를 사용하세요.
- db_index 1  → 계약일 (기술이전계약일)
- db_index 3  → 기술도입업체명 (기관/업체명)
- db_index 8  → 지역코드 (국내지역구분 코드)
- db_index 28 → 기술명
- db_index 29 → 주발명자/연구자명
- db_index 37 → 기술유형
- db_index 44 → 거래유형
- db_index 51 → 경상기술료 조건 텍스트
- db_index 52 → 정액기술료 계약금(원)
- db_index 73 → 입금일
- db_index 76 → 현금입금액(원)
- db_index 82 → 경상기술료 입금액(원)
- db_index 83 → 정액기술료 입금액(원)
양식 파일에 없는 항목은 절대 포함하지 마세요.

---

## 특수 케이스: 연구자/교수 명단 파일
양식 파일이 기술이전 성과를 기록하는 출력 양식이 아니라, **연구자·교수·인원 명단**(예: 교수님 명단, BK참여 명단, 담당자 리스트 등)으로 판단되면, 아래 형식으로만 응답하세요:

```json
{{
  "form_type": "researcher_list",
  "researcher_sheet": "연구자명이 있는 시트명",
  "researcher_col": 2,
  "researcher_data_start": 2,
  "notes": "연구자 명단 파일로 감지됨. 해당 연구자들의 기술이전 실적을 자유 양식으로 추출."
}}
```
- `researcher_col`: 연구자/교수명이 있는 열 번호(1-based)
- `researcher_data_start`: 데이터가 시작되는 행 번호(헤더 다음 행)
- 이 경우 columns 배열 불필요

판단 기준: 시트에 BK참여, 교번, 임용일, 과학인등록번호 같은 인사 정보가 있고, 기술이전 실적 기재 열이 없으면 연구자 명단.

---

반드시 아래 JSON 형식으로만 응답하세요. 설명 없이 JSON 코드블록만 출력하세요.

```json
{{
  "form_type": "output_form",
  "target_sheet": "데이터를 채울 시트명",
  "data_start_row": 4,
  "filter_mode": "contract",
  "exclude_consulting": true,
  "group_by_contract": false,
  "columns": [
    {{
      "col": 1,
      "type": "sequence",
      "label": "순번"
    }},
    {{
      "col": 2,
      "type": "db_value",
      "db_index": 3,
      "transform": "str",
      "label": "도입업체명"
    }},
    {{
      "col": 3,
      "type": "constant",
      "value": "부산대학교산학협력단",
      "label": "기술공급기관"
    }},
    {{
      "col": 4,
      "type": "researcher_concat",
      "label": "기술제공기관명(연구자명)"
    }},
    {{
      "col": 5,
      "type": "company_region",
      "label": "기술도입기업명(소재지)"
    }},
    {{
      "col": 6,
      "type": "sum",
      "db_indices": [82, 83],
      "unit": "million",
      "label": "기술료합계(백만원)"
    }}
  ],
  "manual_inputs": ["성과구분(K열) — ①②직접 선택 필요"],
  "notes": "매핑 특이사항 또는 주의사항"
}}
```

## column type 설명
- "sequence": 행 순번(1,2,3...)
- "db_value": DB의 특정 컬럼값 (db_index + transform 필수)
- "constant": 고정 문자열 (value 필수)
- "sum": 여러 DB 컬럼 합산 (db_indices + unit 필수). ※group_by_contract=true일 때 입금액 합산에는 사용 금지 — payment_amount 사용할 것
- "payment_amount": 집계된 당해연도 입금액(unit 필수). group_by_contract=true일 때 "당해연도 기술료 입금액" 같은 컬럼에 반드시 사용
- "researcher_concat": "부산대학교(연구자명)" 형식 자동 생성
- "company_region": "업체명(지역)" 형식 자동 생성
- "payment_year": 입금일에서 연도 추출
- "large_transfer": 입금액 1억 이상 여부 (Y/N)
- "contract_type": 계약형태 자동 판단 (①정액/②경상/③노하우)

## 중요: 헤더가 있는 모든 열을 빠짐없이 매핑하세요
양식 파일의 데이터 행이 비어있어도 무시하세요. 헤더 행 기준으로 모든 열을 매핑합니다.
예: 헤더에 "기술명"이 있으면 반드시 db_value(db_index=28)로 매핑해야 합니다.

## transform 목록
str, date_yyyymmdd, date_yyyymm, date_year, date_obj, safe_int,
won_to_million, won_to_thousand, region, org_type, trade_type,
domestic_foreign, bridge_type

## filter_mode
- "both": 계약일 또는 입금일이 해당 연도 (BRIDGE 방식)
- "contract": 계약일 기준
- "payment": 입금일 기준

## group_by_contract
- true: 동일 연번(계약)의 여러 입금 회차를 1행으로 집계 (TLO 방식)
- false: DB 행 그대로 (BRIDGE, 세부현황 방식)

## unit (sum 타입에서)
- "won": 원 그대로
- "million": 백만원 (1/1,000,000)
- "thousand": 천원 (1/1,000)

## data_start_row
실제 데이터를 쓸 첫 행 번호 (헤더/제목 행 다음)
"""

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text
    m = re.search(r"```json\s*([\s\S]*?)\s*```", text)
    if m:
        result = json.loads(m.group(1))
    else:
        # 순수 JSON 시도, 실패 시 {...} 블록 추출
        raw = text.strip()
        try:
            result = json.loads(raw)
        except json.JSONDecodeError:
            m2 = re.search(r'\{[\s\S]*\}', raw)
            if m2:
                result = json.loads(m2.group())
            else:
                raise ValueError(f"AI 응답에서 JSON을 파싱할 수 없습니다. 응답 내용:\n{raw[:500]}")

    # 매핑 결과 로그 출력 (디버깅용)
    print(f"  📋 AI 매핑 결과:")
    print(f"     시트: {result.get('target_sheet')} | 데이터시작행: {result.get('data_start_row')} | 필터: {result.get('filter_mode')} | 집계: {result.get('group_by_contract')}")
    for col in result.get('columns', []):
        if col.get('type') == 'db_value':
            print(f"     열{col['col']} [{col.get('label','')}] ← DB[{col['db_index']}] ({col.get('transform','')})")
        else:
            print(f"     열{col['col']} [{col.get('label','')}] ← {col['type']} {col.get('value', col.get('db_indices', ''))}")
    print(f"  📄 매핑 JSON 전체:\n{json.dumps(result, ensure_ascii=False, indent=2)}")
    return result


def get_contract_type(row):
    """계약형태 자동 판단"""
    trade     = str(row[44]).strip() if row[44] is not None else ""
    has_fixed = safe_int(row[52]) > 0
    has_cac   = bool(str(row[51]).strip()) if row[51] else False
    if trade in ("3", "8", "노하우", "단위연구노하우"):
        return "③"
    if has_fixed and has_cac:
        return "①+②"
    if has_fixed:
        return "①"
    if has_cac:
        return "②"
    return "①"


def aggregate_by_contract(rows, year):
    """같은 연번의 여러 입금 회차를 1행으로 집계."""
    contracts = {}
    for row in rows:
        seq = str(row[0]) if row[0] else None
        if not seq:
            continue
        payment = safe_int(row[76])
        p_year  = get_year(row[73])
        if seq not in contracts:
            contracts[seq] = {"row": row, "total_pay": 0, "pay_years": set()}
        contracts[seq]["total_pay"] += payment
        if p_year:
            contracts[seq]["pay_years"].add(p_year)

    result = []
    for seq, info in contracts.items():
        r     = info["row"]
        pay   = info["total_pay"]
        years = sorted(info["pay_years"])
        py    = year if year in years else (years[-1] if years else year)
        # aggregate row에 total_pay 저장 (col 76 override)
        row_list = list(r)
        # 입금 집계값을 별도 속성으로 전달할 수 없으므로 tuple에 붙이기
        result.append((r, pay, py))

    result.sort(key=lambda x: x[0][1]
                if isinstance(x[0][1], (datetime.datetime, datetime.date))
                else datetime.datetime(1900, 1, 1))
    return result


def patch_missing_columns(mapping, ws, db_columns_text, data_start):
    """
    AI가 헤더 열을 누락했을 때 db_columns.md 키워드로 자동 보완.
    헤더 행(data_start - 1)을 읽어 매핑된 열 번호와 비교한다.
    """
    # 기관마다 표현이 다를 수 있는 별칭 → db_index 직접 매핑
    ALIASES = {
        "계약명": 28, "이전기술명": 28, "기술이름": 28, "기술 명칭": 28,
        "발명자": 29, "연구책임자": 29, "발명인": 29,
        "계약업체": 3,  "도입기업": 3, "기업명": 3,
        "계약일": 1, "이전계약일": 1,
        "입금일": 73, "납입일": 73,
    }

    # db_columns.md 파싱: | db_index | label | desc |
    db_lookup = {}  # 키워드 → db_index
    for line in db_columns_text.splitlines():
        parts = [p.strip() for p in line.split("|") if p.strip()]
        if len(parts) >= 2:
            try:
                idx = int(parts[0])
            except ValueError:
                continue
            label = parts[1]
            # 숫자 접두어 제거: "27.기술명" → "기술명"
            keyword = label.split(".")[-1].strip() if "." in label else label
            if keyword:
                db_lookup[keyword] = idx
    # 별칭 병합
    db_lookup.update(ALIASES)

    header_row = data_start - 1
    mapped_cols = {c["col"] for c in mapping.get("columns", [])}

    added = []
    for col_idx in range(1, (ws.max_column or 0) + 1):
        if col_idx in mapped_cols:
            continue
        header_val = ws.cell(header_row, col_idx).value
        if not header_val:
            continue
        # 헤더 텍스트에서 키워드 매칭
        header_clean = str(header_val).replace("\n", " ").strip()
        matched_idx = None
        for keyword, db_idx in db_lookup.items():
            if keyword and keyword in header_clean:
                matched_idx = db_idx
                break
        if matched_idx is not None:
            mapping.setdefault("columns", []).append({
                "col": col_idx,
                "type": "db_value",
                "db_index": matched_idx,
                "transform": "str",
                "label": header_clean[:30],
                "_auto_patched": True,
            })
            added.append(f"열{col_idx}[{header_clean[:20]}]→DB[{matched_idx}]")
        else:
            print(f"  ⚠️ 열{col_idx} [{header_clean[:20]}] — 매핑 없음, 수동 확인 필요")

    if added:
        print(f"  🔧 자동 보완된 열: {', '.join(added)}")
    return mapping


def apply_mapping(ws, rows_data, mapping, data_start, group_by_contract, year):
    """매핑 JSON에 따라 데이터를 채운다."""
    columns = mapping.get("columns", [])
    total_income = 0

    for row_idx, row_entry in enumerate(rows_data, 1):
        r = data_start + row_idx - 1

        # group_by_contract 모드면 (row, pay_amt, pay_year) 튜플
        if group_by_contract:
            row, pay_amt, pay_year = row_entry
        else:
            row = row_entry
            pay_amt  = safe_int(row[76])
            pay_year = get_year(row[73]) or year

        for col_def in columns:
            col      = col_def.get("col")
            col_type = col_def.get("type", "db_value")

            try:
                value = None

                if col_type == "sequence":
                    value = row_idx

                elif col_type == "constant":
                    value = col_def.get("value", "")

                elif col_type == "db_value":
                    db_idx    = col_def.get("db_index", 0)
                    transform = col_def.get("transform", "str")
                    if transform in TRANSFORMS:
                        value = TRANSFORMS[transform](row, db_idx)
                    else:
                        value = row[db_idx] if db_idx < len(row) else ""
                    if transform == "date_obj" and isinstance(value, (datetime.datetime, datetime.date)):
                        ws.cell(r, col).number_format = "YYYY-MM-DD"

                elif col_type == "sum":
                    indices = col_def.get("db_indices", [])
                    unit    = col_def.get("unit", "won")
                    total   = sum(safe_int(row[i]) for i in indices if i < len(row))
                    if unit == "million":
                        value = round(total / 1_000_000, 1) if total else None
                    elif unit == "thousand":
                        value = total // 1_000 if total else None
                    else:
                        value = total if total else None

                elif col_type == "researcher_concat":
                    researcher = str(row[29]).strip() if row[29] else ""
                    value = f"부산대학교({researcher})" if researcher else "부산대학교산학협력단"

                elif col_type == "company_region":
                    company    = str(row[3]).strip() if row[3] else ""
                    is_foreign = str(row[6]).strip() in ("2", "국외")
                    region     = "해외" if is_foreign else get_region_name(row[8])
                    value      = f"{company}({region})" if region else company

                elif col_type == "payment_amount":
                    unit = col_def.get("unit", "won")
                    if unit == "million":
                        value = round(pay_amt / 1_000_000, 1) if pay_amt else None
                    elif unit == "thousand":
                        value = pay_amt // 1_000 if pay_amt else None
                    else:
                        value = pay_amt if pay_amt else None

                elif col_type == "payment_year":
                    value = str(pay_year) if pay_year else ""

                elif col_type == "large_transfer":
                    value = "Y" if pay_amt >= 100_000_000 else "N"

                elif col_type == "contract_type":
                    value = get_contract_type(row)

                if value == "":
                    value = None
                ws.cell(r, col, value)

            except Exception as e:
                print(f"  ⚠️ 행{row_idx} 열{col} ({col_def.get('label','')}) 처리 오류: {e}")

        total_income += pay_amt

    return total_income


def run(master_path, template_path, year, output_path, notes="", hint="", api_key=None):
    print(f"📂 마스터 DB 로딩: {master_path}")
    all_rows = load_master_db(master_path)

    print(f"🔍 양식 파일 전체 분석 중...")
    template_analysis = analyze(template_path)
    sheets_summary = [f"{s['name']}({len(s.get('rows',[]))}행)" for s in template_analysis["sheets"]]
    print(f"  → 시트: {sheets_summary}")

    ref_dir = os.path.join(os.path.dirname(__file__), "..", "references")
    with open(os.path.join(ref_dir, "db_columns.md"), encoding="utf-8") as f:
        db_columns_text = f.read()
    with open(os.path.join(ref_dir, "transform_rules.md"), encoding="utf-8") as f:
        transform_rules_text = f.read()

    # 캐시 조회
    cache = load_cache()
    cache_key, cached_mapping = find_cached_mapping(os.path.basename(template_path), cache)
    if cache_key:
        print(f"  📦 캐시 히트: '{cache_key}' 매핑을 참고 예시로 사용")
    else:
        print(f"  📦 캐시 없음: AI가 양식을 처음부터 분석합니다")

    print(f"🤖 AI가 양식 지침을 분석하고 매핑 중... (10~30초 소요)")
    mapping = get_mapping_from_claude(
        template_analysis, db_columns_text, transform_rules_text,
        notes, year, hint, api_key, cached_mapping=cached_mapping
    )
    if mapping.get("notes"):
        print(f"  ℹ️  {mapping['notes']}")

    # ── 연구자 명단 특수 처리 ──────────────────────────────────────
    if mapping.get("form_type") == "researcher_list":
        print(f"  → 연구자 명단 파일 감지: 자유 양식 추출로 전환")
        return run_free_format(all_rows, template_path, mapping, year, output_path)

    print(f"  → 매핑 완료: {len(mapping.get('columns', []))}개 컬럼")

    # ── notes에서 날짜 범위·연구자 파싱 ─────────────────────────────
    notes_filters = parse_notes_filters(notes)

    # 필터링 (날짜 범위 우선, 없으면 연도 전체)
    filter_mode = notes_filters["date_mode"] if notes_filters["date_start"] \
                  else mapping.get("filter_mode", "both")

    if notes_filters["date_start"] and notes_filters["date_end"]:
        from extract_columns import filter_by_date_range
        filtered = filter_by_date_range(
            all_rows,
            notes_filters["date_start"], notes_filters["date_end"],
            contract_col=1, payment_col=73, mode=filter_mode
        )
        print(f"  → 날짜범위 필터 {notes_filters['date_start']}~{notes_filters['date_end']} ({filter_mode}): {len(filtered)}건")
    else:
        filtered = filter_by_year(all_rows, year, mode=filter_mode)
        print(f"  → {year}년 해당 행: {len(filtered)}건 (기준: {filter_mode})")

    # 연구자 필터
    if notes_filters["researchers"]:
        from extract_columns import filter_by_researchers
        before = len(filtered)
        filtered = filter_by_researchers(filtered, notes_filters["researchers"],
                                         inventor_col=29, co_inventor_col=33)
        print(f"  → 연구자 필터 {notes_filters['researchers']}: {before}건 → {len(filtered)}건")

    if mapping.get("exclude_consulting", True):
        filtered = [r for r in filtered if get_bridge_type(r[37], r[44]) is not None]
        print(f"  → 기술자문 제외 후: {len(filtered)}건")

    # 집계 방식 결정
    group_by = mapping.get("group_by_contract", False)
    if group_by:
        rows_data = aggregate_by_contract(filtered, year)
        print(f"  → 계약 집계 후: {len(rows_data)}건")
    else:
        # 계약일 기준 정렬
        filtered.sort(key=lambda r: r[1] if isinstance(r[1], (datetime.datetime, datetime.date))
                      else datetime.datetime(1900, 1, 1))
        rows_data = filtered

    # 워크북 로드 및 시트 선택
    wb = load_workbook(template_path)
    target_sheet = mapping.get("target_sheet", wb.sheetnames[0])
    ws = wb[target_sheet] if target_sheet in wb.sheetnames else wb.active

    data_start = mapping.get("data_start_row", 4)
    if ws.max_row >= data_start:
        rows_to_del = ws.max_row - data_start + 1
        if rows_to_del > 0:
            ws.delete_rows(data_start, rows_to_del)

    # AI 매핑 후처리: 헤더가 있는데 매핑 누락된 열 자동 보완
    mapping = patch_missing_columns(mapping, ws, db_columns_text, data_start)

    print(f"📝 데이터 채우기 중... (시트: {ws.title}, {data_start}행부터)")
    total_income = apply_mapping(ws, rows_data, mapping, data_start, group_by, year)

    # 수동 입력 필요 안내
    manual = mapping.get("manual_inputs", [])
    if manual:
        print(f"\n⚠️  수동 입력 필요:")
        for m in manual:
            print(f"   • {m}")

    print(f"\n📊 검토 결과 분석 중...")
    # group_by=True면 (row, pay_amt, pay_year) 3-tuple에서 row만 추출
    # group_by=False면 rows_data 자체가 row 리스트
    raw_rows = [entry[0] for entry in rows_data] if group_by else list(rows_data)
    write_feedback_sheet(wb, raw_rows, year, "AI 자동 추출", total_income)

    wb.save(output_path)
    count = len(rows_data)
    print(f"\n✅ 저장 완료: {output_path}")
    print(f"   총 {count}건 | 총수입: {total_income:,}원")
    print(f"📋 검토 결과 시트에서 이상 데이터를 확인하세요.")

    return {
        "count": count,
        "total_income": total_income,
        "mapping_notes": mapping.get("notes", ""),
        "manual_inputs": manual,
        "filter_mode": filter_mode,
        "mapping": mapping,          # 캐시 저장용
        "cache_key": _normalize_key(os.path.basename(template_path)),
    }
