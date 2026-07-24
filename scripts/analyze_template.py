"""
양식 파일 전체 내용 추출기 — 상위 5행이 아닌 전체 시트 내용 추출
지침/예시/헤더 등 양식 내 모든 정보를 캡처하여 AI에 전달한다.
"""
import json, sys
from openpyxl import load_workbook

MAX_CELL_CHARS = 300
MAX_ROWS_PER_SHEET = 200


def analyze(template_path):
    wb = load_workbook(template_path, read_only=True, data_only=True)
    result = {
        "sheets": [],
        "total_sheets": len(wb.sheetnames),
    }

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_info = {
            "name": sheet_name,
            "max_row": ws.max_row,
            "max_col": ws.max_column,
            "rows": [],
        }

        row_count = 0
        for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row_count >= MAX_ROWS_PER_SHEET:
                sheet_info["truncated"] = True
                break

            non_empty = {}
            for col_idx, val in enumerate(row, 1):
                if val is not None:
                    s = str(val).strip()
                    if s:
                        non_empty[col_idx] = s[:MAX_CELL_CHARS]

            if non_empty:
                sheet_info["rows"].append({"row": row_idx, "cells": non_empty})
                row_count += 1

        result["sheets"].append(sheet_info)

    wb.close()
    return result


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--template", required=True)
    args = p.parse_args()
    print(json.dumps(analyze(args.template), ensure_ascii=False, indent=2))
