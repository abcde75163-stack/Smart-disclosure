"""
새 양식 파일의 헤더를 자동 분석하는 유틸리티

사용법:
  python3 analyze_template.py --template "새양식.xlsx"

출력: 각 시트의 헤더 목록을 JSON 형태로 출력
      → Claude가 이를 읽고 마스터 DB 컬럼과 자동 매핑
"""
import argparse, json, sys
from openpyxl import load_workbook

def analyze(template_path):
    wb = load_workbook(template_path, read_only=True, data_only=True)
    result = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_info = {
            'rows': ws.max_row,
            'cols': ws.max_column,
            'headers': []
        }

        # 상위 5행에서 비어있지 않은 셀 수집 (헤더 추출용)
        header_rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 5: break
            non_empty = [(j+1, str(v)[:80]) for j, v in enumerate(row) if v is not None]
            if non_empty:
                header_rows.append({'row': i+1, 'cells': non_empty})

        sheet_info['header_rows'] = header_rows
        result[sheet_name] = sheet_info

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--template', required=True)
    args = p.parse_args()
    analyze(args.template)
