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
        return st.secrets["OPENAI_API_KEY"]
    except Exception:
        return os.environ.get("OPENAI_API_KEY", "")


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


def detect_year(master_name: str, request_names=None) -> int:
    """파일명에서 연도 자동 감지. 없으면 현재 연도."""
    import datetime
    for name in list(request_names or []) + [master_name]:
        m = re.search(r'20(\d{2})', name)
        if m:
            return int(m.group())
    return datetime.datetime.now().year


def wants_column_extract_mode(notes: str, request_files) -> bool:
    """True이면 요청파일은 서식이 아니라 참고/대상 파일로 보고 요청 항목만 추출."""
    text = (notes or "").replace(" ", "")
    if not request_files:
        return True
    no_form_keywords = ["서식없", "양식없", "요청한항목", "적어놓은항목", "아래항목", "항목만", "열만"]
    return any(keyword in text for keyword in no_form_keywords)


def _format_filter_mode(mode: str) -> str:
    return {
        "contract": "계약일 기준",
        "payment": "입금일 기준",
        "both": "계약일 또는 입금일 기준",
        "none": "날짜 필터 없음",
    }.get(mode or "", mode or "-")


def render_understanding(result_info: dict, no_template_mode: bool):
    understanding = result_info.get("understanding") or {}
    if not understanding:
        return

    with st.expander("🔎 AI가 이해한 요청 내용", expanded=True):
        st.write({
            "처리 모드": understanding.get("mode", "서식 없음 / 요청 항목 직접 추출" if no_template_mode else "서식 파일 채우기"),
            "요청파일": understanding.get("request_files", []),
            "필터 기준": _format_filter_mode(understanding.get("filter_mode")),
            "기간": understanding.get("date_range") or understanding.get("filter_year") or "-",
            "대상자/참고값 필터": len(understanding.get("reference_filters") or []),
            "추가 DB 조건": len(understanding.get("db_filters") or []),
        })
        if not no_template_mode:
            st.write({
                "최종 작성양식": understanding.get("output_file_name", "-"),
                "대상 시트": understanding.get("target_sheet", "-"),
                "입력 시작 행": understanding.get("data_start_row", "-"),
                "계약 단위 집계": bool(understanding.get("group_by_contract")),
                "기술자문 제외": bool(understanding.get("exclude_consulting")),
            })


def render_column_validation(result_info: dict, no_template_mode: bool):
    rows = result_info.get("column_validation") or []
    if not rows:
        return

    title = "✅ 요청 항목 매칭 검증" if no_template_mode else "✅ 양식 컬럼 매핑 검증"
    with st.expander(title, expanded=no_template_mode):
        if no_template_mode:
            st.caption("요청사항에서 추출하기로 판단한 항목과 마스터 DB 컬럼입니다.")
        else:
            st.caption("AI가 양식의 출력 열을 어떤 DB 값으로 채우려 했는지 확인합니다.")
        st.dataframe(rows, use_container_width=True, hide_index=True)


def render_zero_result_report(result_info: dict):
    if result_info.get("count", 0) != 0:
        return

    st.warning("⚠️ 추출 결과가 0건입니다. 아래 단계별 건수 변화를 확인해주세요.")
    reasons = result_info.get("zero_result_reasons") or []
    if reasons:
        for reason in reasons:
            st.markdown(f"• {reason}")

    diagnostics = result_info.get("diagnostics") or []
    if diagnostics:
        with st.expander("📉 0건 원인 상세"):
            st.dataframe(diagnostics, use_container_width=True, hide_index=True)


# ─── 헤더 ────────────────────────────────────────────────────────
st.markdown('<p class="app-title">🔬 기술이전 성과 자동 추출</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="app-sub">마스터 DB와 양식 파일을 올리면 AI가 자동으로 채워드립니다</p>',
    unsafe_allow_html=True
)

# API 키 체크
api_key = get_api_key()
if not api_key:
    st.error("❌ OpenAI API 키가 설정되지 않았습니다. Streamlit Cloud → Settings → Secrets에 `OPENAI_API_KEY`를 추가해주세요.")
    st.stop()

# ─── 파일 업로드 ──────────────────────────────────────────────────
col_l, col_r = st.columns([1, 1.4])

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
    st.markdown('<div class="upload-label">📋 요청파일 <span style="color:#9CA3AF;font-weight:400">(선택, 여러 개 가능)</span></div>', unsafe_allow_html=True)
    request_files = st.file_uploader(
        "request_files", type=["xlsx", "xls", "pdf", "docx"], key="request_files",
        accept_multiple_files=True,
        label_visibility="collapsed",
        help="작성양식 Excel, 작성요령 PDF/Word, 추출대상자 명단 등을 함께 올릴 수 있습니다. 어떤 파일이 어떤 역할인지는 아래 요청사항에 자연어로 적어주세요."
    )
    if request_files:
        hint = " / ".join(filter(None, [detect_hint(f.name) for f in request_files]))
        badges = "<br>".join(
            f'<span class="badge-ok">✅ {idx}. {f.name} ({len(f.getvalue())//1024} KB)</span>'
            for idx, f in enumerate(request_files, 1)
        )
        st.markdown(badges, unsafe_allow_html=True)
    else:
        hint = ""
        st.markdown('<div class="upload-hint">작성양식 Excel · 작성요령 PDF/Word · 대상자 명단 등</div>', unsafe_allow_html=True)

st.markdown("---")

# ─── 추가 요청사항 ────────────────────────────────────────────────
st.markdown(
    """
    <div style="font-size:0.82rem; color:#4B5563; line-height:1.55; margin-bottom:0.45rem;">
      <b>작성 예시</b><br>
      • <b>서식이 있는 경우:</b> 요청파일 1은 대상자 명단이고, 요청파일 2는 작성요령, 요청파일 3은 최종 작성양식이야. 요청파일 1의 교직원번호와 동일한 대상만 요청파일 3 양식에 맞게 추출해줘.<br>
      • <b>서식이 없는 경우:</b> 서식 없음. 요청파일 1의 교직원 명단에 대해서 기술이전계약일 2024~2026년 성과를 계약일, 기술명, 발명자명, 학과, 정액기술료, 계약입금일만 추출해줘.
    </div>
    """,
    unsafe_allow_html=True,
)
notes = st.text_area(
    "💬 추가 요청사항 (선택)",
    placeholder=(
        "여기에 자연어로 요청사항을 입력하세요.\n"
        "예: 서식 없음. 요청파일 1의 교직원 명단 기준으로 계약일, 기술명, 발명자명만 추출해줘."
    ),
    height=130,
)

st.markdown("<br>", unsafe_allow_html=True)

# ─── 모드 안내 ───────────────────────────────────────────────────
if master_file and not request_files:
    st.info("📝 양식 파일 없음 — 추가 요청사항을 바탕으로 마스터 DB 열을 그대로 추출합니다.")
elif master_file and request_files:
    st.info("🧾 요청파일 포함 — 파일 역할과 요청사항을 자연어로 판단해 추출합니다.")

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
    request_names = [f.name for f in request_files] if request_files else []
    year = detect_year(master_file.name, request_names)
    no_template_mode = wants_column_extract_mode(notes, request_files)

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

            request_paths = []
            for idx, request_file in enumerate(request_files or [], 1):
                ext = os.path.splitext(request_file.name)[1]
                request_path = os.path.join(tmpdir, f"request_{idx}{ext}")
                with open(request_path, "wb") as f:
                    f.write(request_file.getvalue())
                request_paths.append(request_path)

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
                        master_path, year, output_path, notes, api_key,
                        request_file_paths=request_paths,
                        request_file_names=request_names,
                    )
                else:
                    from extract_generic import run as run_generic
                    result_info = run_generic(
                        master_path, request_paths[0], year,
                        output_path, notes, hint, api_key,
                        request_file_paths=request_paths,
                        request_file_names=request_names,
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

                render_understanding(result_info, no_template_mode)
                render_column_validation(result_info, no_template_mode)
                render_zero_result_report(result_info)

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

                    if result_info.get("mapping_quality"):
                        quality = result_info["mapping_quality"]
                        with st.expander("🎯 AI 매핑 신뢰도 / 자동 검증"):
                            st.write({
                                "전체 confidence": quality.get("overall_confidence"),
                                "헤더 열 수": quality.get("header_columns"),
                                "매핑 열 수": quality.get("mapped_columns"),
                                "검증 이슈 수": len(quality.get("issues", [])),
                                "낮은 confidence 열 수": len(quality.get("low_confidence", [])),
                            })
                            if quality.get("issues"):
                                st.warning("자동 검증에서 확인이 필요한 항목이 있습니다.")
                                st.json(quality.get("issues"))
                            if quality.get("low_confidence"):
                                st.info("AI가 낮은 confidence로 표시한 열입니다.")
                                st.json(quality.get("low_confidence"))

                st.divider()

                if no_template_mode:
                    master_stem = os.path.splitext(master_file.name)[0]
                    output_filename = f"{master_stem}_{year}_열추출.xlsx"
                else:
                    output_index = int(result_info.get("output_file_index") or 1)
                    output_index = max(1, min(output_index, len(request_files or [None])))
                    tname = os.path.splitext(request_files[output_index - 1].name)[0]
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
                        output_index = int(result_info.get("output_file_index") or 1)
                        output_index = max(1, min(output_index, len(request_files or [None])))
                        cache_template_name = request_files[output_index - 1].name
                        target_sheet = result_info["mapping"].get("target_sheet")
                        data_start   = result_info["mapping"].get("data_start_row", 4)
                        example_rows = extract_example_rows(
                            corrected_file.getvalue(), target_sheet, data_start, n_rows=5
                        )
                        updated_cache = save_mapping_to_cache(
                            cache_template_name,
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
                        output_index = int(result_info.get("output_file_index") or 1)
                        output_index = max(1, min(output_index, len(request_files or [None])))
                        cache_template_name = request_files[output_index - 1].name
                        updated_cache = save_mapping_to_cache(
                            cache_template_name, result_info["mapping"], load_cache()
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
st.caption("© 부산대학교 산학협력단  ·  Powered by OpenAI")
