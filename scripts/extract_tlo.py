"""
대학기술경영촉진사업 기술이전 증빙 양식 자동 생성
+ 검토 결과 시트 자동 추가
"""
import argparse, sys, os, datetime
sys.path.insert(0, os.path.dirname(__file__))
from common import *
from feedback import write_feedback_sheet
from openpyxl import Workbook

def run(master_path, year, output_path, tmc_name="과학기술사업화진흥원", period_str=None):
    print(f"📂 DB 로딩: {master_path}")
    all_rows = load_master_db(master_path)

    filtered = filter_by_year(all_rows, year, mode='contract')
    print(f"  → {year}년 계약 건: {len(filtered)}건")

    data = [r for r in filtered if get_bridge_type(r[37], r[44]) is not None]
    print(f"  → 기술자문 제외 후: {len(data)}건")

    data.sort(key=lambda r: r[1] if isinstance(r[1], (datetime.datetime, datetime.date))
              else datetime.datetime(1900, 1, 1))

    if period_str is None:
        period_str = f"{year}.1.1.~{year}.12.31."

    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}년 기술이전"

    ws.cell(1, 1, f'{year}년 대학기술경영촉진사업 기술이전 실적 내역(기간: {period_str})')
    headers = ['순번','주관기관명(TMC)','기술제공기관명','기술도입기업명','기술명',
               f'계약일자\n(8자리, 연도+월+일)','총 기술료\n계약액(백만원)','당해연도 기술료\n입금액(백만원)']
    for ci, h in enumerate(headers, 1):
        ws.cell(3, ci, h)

    total_contract = total_payment = 0

    for idx, row in enumerate(data, 1):
        r = 3 + idx
        contract_amt = safe_int(row[52]) + safe_int(row[50])
        payment_amt  = safe_int(row[82]) + safe_int(row[83])
        total_contract += contract_amt
        total_payment  += payment_amt

        ws.cell(r, 1, idx)
        ws.cell(r, 2, tmc_name)
        ws.cell(r, 3, '부산대학교산학협력단')
        ws.cell(r, 4, str(row[3]) if row[3] else '')
        ws.cell(r, 5, str(row[28]) if row[28] else '')
        ws.cell(r, 6, row[1])
        ws.cell(r, 7, contract_amt or None)
        ws.cell(r, 8, payment_amt or None)
        ws.cell(r, 6).number_format = 'YYYYMMDD'

    print("\n📊 검토 결과 분석 중...")
    write_feedback_sheet(wb, data, year, 'TLO혁신형', total_payment)

    wb.save(output_path)
    print(f"\n✅ 저장 완료: {output_path}")
    print(f"   총 {len(data)}건 | 계약액: {total_contract:,}원 | 입금액: {total_payment:,}원")
    print(f"📋 '검토 결과' 시트에서 이상 데이터 및 확인 필요 항목을 확인하세요.")

if __name__ == '__main__':
    p = argparse.ArgumentParser()
    p.add_argument('--master',  required=True)
    p.add_argument('--year',    type=int, required=True)
    p.add_argument('--output',  required=True)
    p.add_argument('--period',  default=None)
    p.add_argument('--tmc',     default='과학기술사업화진흥원')
    args = p.parse_args()
    run(args.master, args.year, args.output, args.tmc, args.period)
