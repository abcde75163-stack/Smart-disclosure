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


def get_mapping_from_claude(template_analysis, db_columns_text, transform_rules_text, notes, year, hint, api_key=None):
    """
    양식 파일 전체 내용을 AI에 전달하고 컬럼 매핑 JSON을 받아온다.
    """
    try:
        import anthropic
    except ImportError:
        raise ImportError("pip install anthropic 필요")

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    # 양식 내용을 읽기 쉽게 정리
    form_text_parts = []
    for sheet in template_analysis.get("sheets", []):
        form_text_parts.append(f"\n### 시트: {sheet['name']} (최대행:{sheet.get('max_row','?')}, 최대열:{sheet.get('max_col','?')})")
        for row_info in sheet.get("rows", []):
            cells_str = ", ".join(f"[열{k}]{v}" for k, v in row_info["cells"].items())
            form_text_parts.append(f"  행{row_info['row']}: {cells_str}")
        if sheet.get("truncated"):
            form_text_parts.append("  (이하 행 생략)")
    form_full_text = "\n".join(form_text_parts)

    prompt = f"""당신은 기술이전 데이터 추출 전문가입니다.

아래에 제공된 양식 파일의 전체 내용을 읽고, 그 안에 있는 작성 지침과 예시를 따라 컬럼 매핑 JSON을 생성해주세요.

## 중요 원칙
- 양식 파일 내부의 가이드 지침, 예시 데이터, 컬럼 헤더에 명시된 형식을 최우선으로 따르세요.
- 헤더에 "(백만원)"이라고 쓰여 있으면 백만원 단위로, "(YYYYMMDD)"이면 8자리 날짜로, "(Yes/No)"이면 Y/N으로 변환하세요.
- 예시 행(XX대학교, 홍길동 등)의 형식을 정확히 따르세요. 예) "XX대학교(홍길동)" 형식이면 "부산대학교(연구자명)" 형식으로.
- 작성 지침에 인정기준이 있으면 그에 맞게 필터 조건을 설정하세요.

## 추출 연도
{year}년

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

---

반드시 아래 JSON 형식으로만 응답하세요. 설명 없이 JSON 코드블록만 출력하세요.

```json
{{
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
- "sum": 여러 DB 컬럼 합산 (db_indices + unit 필수)
- "researcher_concat": "부산대학교(연구자명)" 형식 자동 생성
- "company_region": "업체명(지역)" 형식 자동 생성
- "payment_year": 입금일에서 연도 추출
- "large_transfer": 입금액 1억 이상 여부 (Y/N)
- "contract_type": 계약형태 자동 판단 (①정액/②경상/③노하우)

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
        return json.loads(m.group(1))
    return json.loads(text.strip())


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

    print(f"🤖 AI가 양식 지침을 분석하고 매핑 중... (10~30초 소요)")
    mapping = get_mapping_from_claude(
        template_analysis, db_columns_text, transform_rules_text,
        notes, year, hint, api_key
    )
    print(f"  → 매핑 완료: {len(mapping.get('columns', []))}개 컬럼")
    if mapping.get("notes"):
        print(f"  ℹ️  {mapping['notes']}")

    # 필터링
    filter_mode = mapping.get("filter_mode", "both")
    filtered = filter_by_year(all_rows, year, mode=filter_mode)
    print(f"  → {year}년 해당 행: {len(filtered)}건 (기준: {filter_mode})")

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

    print(f"📝 데이터 채우기 중... (시트: {ws.title}, {data_start}행부터)")
    total_income = apply_mapping(ws, rows_data, mapping, data_start, group_by, year)

    # 수동 입력 필요 안내
    manual = mapping.get("manual_inputs", [])
    if manual:
        print(f"\n⚠️  수동 입력 필요:")
        for m in manual:
            print(f"   • {m}")

    print(f"\n📊 검토 결과 분석 중...")
    raw_rows = [entry[0] if isinstance(entry, tuple) else entry for entry in rows_data]
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
    }
