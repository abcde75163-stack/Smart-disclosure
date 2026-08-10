"""
기술이전 성과 자동 추출 시스템
부산대학교 산학협력단 마스터 DB → 상위기관 보고 양식 자동 변환
"""
import streamlit as st
import tempfile, os, io, sys, traceback, re, json
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

st.set_page_config(
    page_title="기술이전 성과 자동 추출",
    page_icon="🔬",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
  /* 사이드바 완전 숨김 */
  [data-testid="stSidebar"] { display: none; }
  [data-testid="collapsedControl"] { display: none; }

  /* 전체 여백 */
  .block-container { padding-top: 2.5rem; max-width: 780px; }

  /* 타이틀 */
  .app-title {
    font-size: 1.75rem; font-weight: 800;
    color: #1F3864; margin-bottom: 0.2rem;
    text-align: center;
  }
  .app-sub {
    color: #6B7280; font-size: 0.9rem;
    text-align: center; margin-bottom: 2rem;
  }

  /* 업로드 카드 */
  .upload-card {
    border: 2px dashed #CBD5E0; border-radius: 12px;
    padding: 1.2rem 1rem; background: #F8FAFC;
    margin-bottom: 0.5rem;
  }
  .upload-label {
    font-size: 0.85rem; font-weight: 700;
    color: #374151; margin-bottom: 0.3rem;
  }
  .upload-hint { font-size: 0.78rem; color: #9CA3AF; }

  /* 업로드 완료 뱃지 */
  .badge-ok {
    display: inline-block;
    background: #D1FAE5; color: #065F46;
    border-radius: 6px; padding: 2px 8px;
    font-size: 0.78rem; font-weight: 600;
  }

  /* 실행 버튼 */
  div[data-testid="stButton"] > button[kind="primary"] {
    width: 100%; padding: 0.75rem;
    font-size: 1.05rem; font-weight: 700;
    border-radius: 10px;
    background: #1F3864 !important;
    color: white !important;
  }

  /* 다운로드 버튼 */
  .stDownloadButton > button {
    background: #1F3864 !important; color: white !important;
    font-size: 1rem !important; border-radius: 8px !important;
    padding: 0.6rem 2rem !important; width: 100%;
  }

  /* 결과 카드 */
  .result-box {
    background: #F0FDF4; border: 1.5px solid #86EFAC;
    border-radius: 10px; padding: 1rem 1.2rem; margin: 1rem 0;
  }
</style>
""", unsafe_allow_html=True)


def get_api_key():
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return os.environ.get("ANTHROPIC_API_KEY", "")


def detect_hint(filename: str) -> str:
    name = filename.lower()
    hints = []
    if any(k in name for k in ["bridge", "브릿지", "bridge3"]):
        hints.append("BRIDGE 3.0 사업 양식. 기술이전유형은 거래유형+기술유형 조합으로 결정.")
    if any(k in name for k in ["tlo", "기술경영촉진", "대학기술경영"]):
        hints.append("TLO혁신형 양식. 계약 단위 집계(group_by_contract=true). 연구자명 포함 필요.")
    if any(k in name for k in ["세부현황", "29번", "실적세부"]):
        hints.append("기술이전 실적 세부 현황 양식. 입금일 기준 필터(payment). 입금 회차별 행.")
    if any(k in name for k in ["산자부", "산업부", "현황조사"]):
        hints.append("산업부 현황조사 양식. 계약 단위 집계 가능성 높음.")
    return " / ".join(hints) if hints else ""


def detect_year(master_name: str, template_name: str) -> int:
    """파일명에서 연도 자동 감지. 없으면 현재 연도."""
    import datetime
    for name in [template_name, master_name]:
        m = re.search(r'20(\d{2})', name)
        if m:
            return int(m.group())
    return datetime.datetime.now().year


# ─── 헤더 ────────────────────────────────────────────────────────
st.markdown('<p class="app-title">🔬 기술이전 성과 자동 추출</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="app-sub">마스터 DB와 양식 파일을 올리면 AI가 자동으로 채워드립니다</p>',
    unsafe_allow_html=True
)

# API 키 체크
api_key = get_api_key()
if not api_key:
    st.error("❌ Anthropic API 키가 설정되지 않았습니다. Streamlit Cloud → Settings → Secrets에 `ANTHROPIC_API_KEY`를 추가해주세요.")
    st.stop()

# ─── 파일 업로드 ──────────────────────────────────────────────────
col_l, col_r = st.columns(2)

with col_l:
    st.markdown('<div class="upload-label">📂 마스터 DB</div>', unsafe_allow_html=True)
    master_file = st.file_uploader(
        "master", type=["xlsx"], key="master",
        label_visibility="collapsed",
        help="기술이전총정리.xlsx 형식의 마스터 데이터베이스 파일"
    )
    if master_file:
        st.markdown(
            f'<span class="badge-ok">✅ {master_file.name} ({len(master_file.getvalue())//1024} KB)</span>',
            unsafe_allow_html=True
        )
    else:
        st.markdown('<div class="upload-hint">기술이전총정리_날짜.xlsx</div>', unsafe_allow_html=True)

with col_r:
    st.markdown('<div class="upload-label">📋 양식 파일 <span style="color:#9CA3AF;font-weight:400">(선택)</span></div>', unsafe_allow_html=True)
    template_file = st.file_uploader(
        "template", type=["xlsx", "xls"], key="template",
        label_visibility="collapsed",
        help="기관에서 받은 양식 파일. 없으면 추가 요청사항을 기반으로 마스터 DB 열을 그대로 추출합니다."
    )
    if template_file:
        hint = detect_hint(template_file.name)
        st.markdown(
            f'<span class="badge-ok">✅ {template_file.name} ({len(template_file.getvalue())//1024} KB)</span>',
            unsafe_allow_html=True
        )
    else:
        hint = ""
        st.markdown('<div class="upload-hint">없으면 요청사항 기반으로 열 추출</div>', unsafe_allow_html=True)

st.markdown("---")

# ─── 추가 요청사항 ────────────────────────────────────────────────
notes = st.text_area(
    "💬 추가 요청사항 (선택)",
    placeholder=(
        "예시)\n"
        "기간: 20260101 ~ 20260801 (계약일 기준)\n"
        "대상 연구자: 홍길동, 김철수, 이영희\n"
        "추출 항목: 연번, 계약일, 기술도입업체명, 기술명, 주발명자, 총기술료, 입금일, 현금입금액"
    ),
    height=110,
)

st.markdown("<br>", unsafe_allow_html=True)

# ─── 모드 안내 ───────────────────────────────────────────────────
if master_file and not template_file:
    st.info("📝 양식 파일 없음 — 추가 요청사항을 바탕으로 마스터 DB 열을 그대로 추출합니다.")

# ─── 실행 버튼 ───────────────────────────────────────────────────
run_btn = st.button(
    "📊 추출 실행",
    type="primary",
    disabled=(not master_file),
    use_container_width=True,
)

if not master_file:
    st.caption("⬆️ 마스터 DB 파일을 올려주세요.")

# ─── 추출 실행 ───────────────────────────────────────────────────
if run_btn and master_file:
    template_name = template_file.name if template_file else ""
    year = detect_year(master_file.name, template_name)
    no_template_mode = (template_file is None)

    spinner_msg = (
        f"⏳ 추가 요청사항을 분석하고 {year}년 데이터를 추출하고 있습니다... (10~20초)"
        if no_template_mode else
        f"⏳ AI가 양식을 분석하고 {year}년 데이터를 추출하고 있습니다... (10~30초)"
    )

    with st.spinner(spinner_msg):
        with tempfile.TemporaryDirectory() as tmpdir:
            master_path   = os.path.join(tmpdir, "master.xlsx")
            output_path   = os.path.join(tmpdir, "output.xlsx")

            with open(master_path, "wb") as f:
                f.write(master_file.getvalue())

            if not no_template_mode:
                ext           = os.path.splitext(template_file.name)[1]
                template_path = os.path.join(tmpdir, f"template{ext}")
                with open(template_path, "wb") as f:
                    f.write(template_file.getvalue())

            captured   = io.StringIO()
            old_stdout = sys.stdout
            sys.stdout = captured

            success     = False
            result_info = {}
            error_msg   = ""
            tb_str      = ""

            try:
                if no_template_mode:
                    from extract_columns import run as run_columns
                    result_info = run_columns(
                        master_path, year, output_path, notes, api_key
                    )
                else:
                    from extract_generic import run as run_generic
                    result_info = run_generic(
                        master_path, template_path, year,
                        output_path, notes, hint, api_key
                    )
                success = True
            except Exception as e:
                error_msg = str(e)
                tb_str    = traceback.format_exc()
            finally:
                sys.stdout = old_stdout

            log_output = captured.getvalue()

            if success and os.path.exists(output_path):
                with open(output_path, "rb") as f:
                    result_bytes = f.read()

                is_researcher_list = result_info.get("form_type") == "researcher_list"

                if no_template_mode:
                    col_count = len(result_info.get("columns", []))
                    st.success(f"✅ 추출 완료! ({result_info.get('filter_year', year)}년 / {col_count}개 열)")
                elif is_researcher_list:
                    researcher_count = result_info.get("researcher_count", 0)
                    st.success(f"✅ 연구자 {researcher_count}명 기술이전 실적 추출 완료!")
                else:
                    st.success(f"✅ 추출 완료! ({year}년 기준)")
                st.balloons()

                # 결과 요약
                count     = result_info.get("count", 0)
                total_won = result_info.get("total_income", 0)
                total_m   = round(total_won / 1_000_000, 1) if total_won else 0
                filter_mode = result_info.get("filter_mode", "")

                c1, c2, c3 = st.columns(3)
                c1.metric("추출 건수", f"{count}건")
                c2.metric("총 수입기술료", f"{total_m}백만원")
                if no_template_mode:
                    c3.metric("추출 열 수", f"{len(result_info.get('columns', []))}개")
                elif is_researcher_list:
                    c3.metric("연구자 수", f"{result_info.get('researcher_count', 0)}명")
                else:
                    c3.metric("필터 기준", filter_mode or "-")

                # no_template_mode: AI 분석 노트
                if no_template_mode and result_info.get("ai_notes"):
                    with st.expander("🤖 AI 분석 노트"):
                        st.write(result_info["ai_notes"])

                # 수동 입력 필요 안내 (양식 모드 전용)
                if not no_template_mode:
                    manual = result_info.get("manual_inputs", [])
                    if manual:
                        st.warning("⚠️ 수동 입력이 필요한 항목이 있습니다")
                        for item in manual:
                            st.markdown(f"• {item}")

                    if result_info.get("mapping_notes"):
                        with st.expander("🤖 AI 분석 노트"):
                            st.write(result_info["mapping_notes"])

                st.divider()

                if no_template_mode:
                    master_stem = os.path.splitext(master_file.name)[0]
                    output_filename = f"{master_stem}_{year}_열추출.xlsx"
                else:
                    tname = os.path.splitext(template_file.name)[0]
                    if is_researcher_list:
                        output_filename = f"{tname}_{filter_mode}_기술이전실적.xlsx"
                    else:
                        output_filename = f"{tname}_{year}_자동추출.xlsx"

                st.download_button(
                    label="⬇️ 결과 파일 다운로드",
                    data=result_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )
                if no_template_mode:
                    st.caption(f"`{output_filename}` · 마스터 DB 원본 열 그대로 추출된 파일입니다.")
                elif is_researcher_list:
                    st.caption(f"`{output_filename}` · 연구자별 기술이전 실적이 자유 양식으로 생성되었습니다.")
                else:
                    st.caption(f"`{output_filename}` · '검토 결과' 시트에서 이상 데이터를 확인하세요.")

                with st.expander("📋 처리 상세 로그"):
                    st.code(log_output, language="text")

                # ─── 캐시 저장 (양식 모드 + 연구자 명단 제외) ─────────
                if not no_template_mode and not is_researcher_list:
                    st.markdown("---")
                    st.markdown("#### 💾 캐시 업데이트")
                    st.caption("결과를 검토 후 수정한 파일을 올리면 다음번 동일 양식 추출 정확도가 올라갑니다.")

                    corrected_file = st.file_uploader(
                        "✏️ 수동 수정된 결과 파일 업로드 (선택)",
                        type=["xlsx"],
                        key="corrected",
                        help="다운로드한 파일에서 틀린 부분을 고친 뒤 여기에 올려주세요."
                    )

                    if corrected_file and result_info.get("mapping"):
                        from extract_generic import (
                            load_cache, save_mapping_to_cache, extract_example_rows
                        )
                        target_sheet = result_info["mapping"].get("target_sheet")
                        data_start   = result_info["mapping"].get("data_start_row", 4)
                        example_rows = extract_example_rows(
                            corrected_file.getvalue(), target_sheet, data_start, n_rows=5
                        )
                        updated_cache = save_mapping_to_cache(
                            template_file.name,
                            result_info["mapping"],
                            load_cache(),
                            example_rows=example_rows,
                        )
                        cache_json = json.dumps(updated_cache, ensure_ascii=False, indent=2)

                        st.success(f"✅ 수정 파일에서 예시 {len(example_rows)}행 추출 완료!")
                        st.caption("아래 파일을 다운로드하고 GitHub의 `references/mapping_cache.json`에 덮어쓰기 하세요.")
                        st.download_button(
                            label="📥 mapping_cache.json 다운로드",
                            data=cache_json.encode("utf-8"),
                            file_name="mapping_cache.json",
                            mime="application/json",
                            use_container_width=True,
                        )
                    elif result_info.get("mapping"):
                        from extract_generic import load_cache, save_mapping_to_cache
                        updated_cache = save_mapping_to_cache(
                            template_file.name, result_info["mapping"], load_cache()
                        )
                        cache_json = json.dumps(updated_cache, ensure_ascii=False, indent=2)
                        st.download_button(
                            label="📥 수정 없이 현재 매핑만 캐시 저장",
                            data=cache_json.encode("utf-8"),
                            file_name="mapping_cache.json",
                            mime="application/json",
                            use_container_width=True,
                        )

            else:
                st.error("❌ 추출 중 오류가 발생했습니다.")
                st.markdown(f"**오류:** `{error_msg}`")
                if log_output:
                    with st.expander("처리 로그"):
                        st.code(log_output, language="text")
                with st.expander("상세 오류 (개발자용)"):
                    st.code(tb_str, language="text")

st.markdown("---")
st.caption("© 부산대학교 산학협력단  ·  Powered by Claude AI")
