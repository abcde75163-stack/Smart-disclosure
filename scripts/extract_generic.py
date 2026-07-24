"""
새 양식 자동 추출기 — Claude AI 기반 컬럼 매핑
처음 보는 양식도 AI가 헤더를 분석하여 마스터 DB와 자동 매핑한다.
"""
import json, sys, os, re, datetime
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from feedback import write_feedback_sheet
from analyze_template import analyze
from openpyxl import load_workbook

# ── 지원 transform 함수 ──────────────────────────────────────────
def _make_transforms():
    return {
        "str":              lambda row, idx: str(row[idx]) if row[idx] is not None else '',
        "date_yyyymmdd":    lambda row, idx: to_yyyymmdd(row[idx]),
        "date_yyyymm":      lambda row, idx: to_yyyymm(row[idx]),
        "date_obj":         lambda row, idx: row[idx],
        "safe_int":         lambda row, idx: safe_int(row[idx]),
        "won_to_million":   lambda row, idx: won_to_million(row[idx]),
        "won_to_thousand":  lambda row, idx: won_to_thousand(row[idx]),
        "region":           lambda row, idx: get_region_name(row[8], row[6]),
        "org_type":         lambda row, idx: ORG_TYPE_MAP.get(row[idx], str(row[idx]) if row[idx] else ''),
        "trade_type":       lambda row, idx: TRADE_TYPE_MAP.get(row[idx], str(row[idx]) if row[idx] else ''),
        "domestic_foreign": lambda row, idx: '국외' if str(row[6]).strip() in ('2', '국외') else '국내',
        "bridge_type":      lambda row, idx: get_bridge_type(row[37], row[44]) or '',
    }

TRANSFORMS = _make_transforms()


# ── Claude API 매핑 요청 ────────────────────────────────────────
def get_mapping_from_claude(template_analysis, db_columns_text, transform_rules_text, notes, year, api_key=None):
    """
    Claude API를 호출하여 컬럼 매핑 JSON을 받아온다.
    Returns: dict (mapping JSON)
    """
    try:
        import anthropic
    except ImportError:
        raise ImportError("anthropic 패키지가 설치되지 않았습니다. pip install anthropic")

    if api_key:
        client = anthropic.Anthropic(api_key=api_key)
    else:
        client = anthropic.Anthropic()

    prompt = f"""당신은 기술이전 데이터 추출 전문가입니다.

아래 정보를 바탕으로, 새 양식 파일에 기술이전 마스터 DB 데이터를 채울 수 있는 컬럼 매핑 JSON을 생성해주세요.

## 추출 연도
{year}년

## 새 양식 파일 구조 분석 결과
{json.dumps(template_analysis, ensure_ascii=False, indent=2)}

## 마스터 DB 컬럼 목록
{db_columns_text}

## 변환 규칙
{transform_rules_text}

## 사용자 추가 요청사항
{notes if notes else "(없음)"}

---

반드시 아래 JSON 형식으로만 응답하세요. 설명 없이 JSON 코드블록만 출력하세요.

```json
{{
  "sheet_name": "데이터를 채울 시트명 (분석된 시트명 중 하나)",
  "data_start_row": 2,
  "filter_mode": "both",
  "exclude_consulting": true,
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
      "type": "sum",
      "db_indices": [82, 83],
      "unit": "won",
      "label": "기술료 합계(원)"
    }}
  ],
  "notes": "매핑 특이사항 기록 (수동 확인 필요 항목 등)"
}}
```

## 가이드

**filter_mode**
- "both": 계약일 또는 입금일이 해당 연도인 건 (BRIDGE 방식)
- "contract": 계약일 기준
- "payment": 입금일 기준

**column type**
- "sequence": 행 순번 (1, 2, 3 ...)
- "db_value": DB의 특정 컬럼값 (db_index + transform 필수)
- "constant": 고정 문자열 (value 필수)
- "sum": 여러 DB 컬럼 합산 (db_indices + unit 필수)

**transform 목록**
- "str": 문자열 그대로
- "date_yyyymmdd": 날짜 → YYYYMMDD (예: 20250527)
- "date_yyyymm": 날짜 → YYYY.MM
- "date_obj": datetime 객체 그대로 (엑셀 날짜 서식 적용됨)
- "safe_int": 정수 변환
- "won_to_million": 원 → 백만원 (정수)
- "won_to_thousand": 원 → 천원
- "region": 지역코드 → 광역시도명
- "org_type": 기관유형 코드 → 텍스트
- "trade_type": 거래유형 코드 → 텍스트
- "domestic_foreign": 국내/국외 텍스트
- "bridge_type": BRIDGE 기술이전유형 변환

**sum unit**
- "won": 원 그대로
- "million": 백만원
- "thousand": 천원

중요: data_start_row는 실제 데이터가 들어갈 첫 번째 행 번호입니다 (헤더 다음 행).
기존 헤더 행 위에 데이터를 쓰지 않도록 주의하세요.
"""

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text

    # JSON 코드블록 추출
    json_match = re.search(r'```json\s*([\s\S]*?)\s*```', text)
    if json_match:
        return json.loads(json_match.group(1))

    # 순수 JSON인 경우
    return json.loads(text.strip())


# ── 데이터 채우기 ───────────────────────────────────────────────
def apply_mapping(ws, rows, mapping, data_start):
    """매핑 JSON을 기반으로 워크시트에 데이터를 채운다."""
    columns = mapping.get('columns', [])
    total_income = 0

    for row_idx, row in enumerate(rows, 1):
        r = data_start + row_idx - 1

        for col_def in columns:
            col = col_def.get('col')
            col_type = col_def.get('type', 'db_value')

            try:
                if col_type == 'sequence':
                    value = row_idx

                elif col_type == 'constant':
                    value = col_def.get('value', '')

                elif col_type == 'db_value':
                    db_idx = col_def.get('db_index', 0)
                    transform = col_def.get('transform', 'str')

                    if transform in TRANSFORMS:
                        value = TRANSFORMS[transform](row, db_idx)
                    else:
                        value = row[db_idx] if db_idx < len(row) else ''

                    # 날짜 객체 서식
                    if transform == 'date_obj' and isinstance(value, (datetime.datetime, datetime.date)):
                        ws.cell(r, col).number_format = 'YYYY-MM-DD'

                elif col_type == 'sum':
                    indices = col_def.get('db_indices', [])
                    unit = col_def.get('unit', 'won')
                    total = sum(safe_int(row[i]) for i in indices if i < len(row))
                    if unit == 'million':
                        value = total // 1_000_000 if total else None
                    elif unit == 'thousand':
                        value = total // 1_000 if total else None
                    else:
                        value = total if total else None

                else:
                    value = None

                # 빈 문자열은 None으로
                if value == '':
                    value = None

                ws.cell(r, col, value)

            except Exception as e:
                print(f"  ⚠️ 행{row_idx} 열{col} ({col_def.get('label','')}) 처리 오류: {e}")
                ws.cell(r, col, None)

        # 총 수입 누적 (검토 시트용)
        total_income += safe_int(row[76]) if 76 < len(row) else 0

    return total_income


# ── 메인 실행 ───────────────────────────────────────────────────
def run(master_path, template_path, year, output_path, notes="", api_key=None):
    print(f"📂 마스터 DB 로딩: {master_path}")
    all_rows = load_master_db(master_path)

    print(f"🔍 양식 파일 분석 중...")
    template_analysis = analyze(template_path)

    # 참조 파일 로드
    ref_dir = os.path.join(os.path.dirname(__file__), '..', 'references')
    with open(os.path.join(ref_dir, 'db_columns.md'), encoding='utf-8') as f:
        db_columns_text = f.read()
    with open(os.path.join(ref_dir, 'transform_rules.md'), encoding='utf-8') as f:
        transform_rules_text = f.read()

    print(f"🤖 AI 컬럼 매핑 분석 중... (10~30초 소요)")
    mapping = get_mapping_from_claude(
        template_analysis, db_columns_text, transform_rules_text, notes, year, api_key
    )

    print(f"  → 매핑 완료: {len(mapping.get('columns', []))}개 컬럼")
    if mapping.get('notes'):
        print(f"  ℹ️  {mapping['notes']}")

    # 필터링
    filter_mode = mapping.get('filter_mode', 'both')
    filtered = filter_by_year(all_rows, year, mode=filter_mode)
    print(f"  → {year}년 해당 행: {len(filtered)}건 (기준: {filter_mode})")

    # 기술자문 제외
    if mapping.get('exclude_consulting', True):
        filtered = [r for r in filtered if get_bridge_type(r[37], r[44]) is not None]
        print(f"  → 기술자문 제외 후: {len(filtered)}건")

    # 계약일 기준 정렬
    filtered.sort(key=lambda r: r[1] if isinstance(r[1], (datetime.datetime, datetime.date))
                  else datetime.datetime(1900, 1, 1))

    # 템플릿 워크북 열기
    wb = load_workbook(template_path)
    sheet_name = mapping.get('sheet_name', wb.sheetnames[0])
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active

    data_start = mapping.get('data_start_row', 2)

    # 기존 데이터 행 제거 (헤더 보존)
    if ws.max_row >= data_start:
        rows_to_delete = ws.max_row - data_start + 1
        if rows_to_delete > 0:
            ws.delete_rows(data_start, rows_to_delete)

    # 데이터 채우기
    print(f"📝 데이터 채우기 중...")
    total_income = apply_mapping(ws, filtered, mapping, data_start)

    # 검토 결과 시트 추가
    print(f"\n📊 검토 결과 분석 중...")
    write_feedback_sheet(wb, filtered, year, '새 양식', total_income)

    wb.save(output_path)
    print(f"\n✅ 저장 완료: {output_path}")
    print(f"   총 {len(filtered)}건 | 총수입: {total_income:,}원")
    print(f"📋 '검토 결과' 시트에서 이상 데이터 및 확인 필요 항목을 확인하세요.")

    return {
        "count": len(filtered),
        "total_income": total_income,
        "mapping_notes": mapping.get('notes', ''),
        "filter_mode": filter_mode,
    }
