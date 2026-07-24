"""
기술이전 성과 자동 추출 시스템
부산대학교 산학협력단 마스터 DB → 상위기관 보고 양식 자동 변환

핵심 원칙:
  양식 파일 내부 지침을 AI가 읽고 따른다.
  하드코딩 없음. 양식이 달라져도 동작한다.
"""
import streamlit as st
import tempfile, os, io, sys, traceback
from pathlib import Path

SCRIPTS_DIR = Path(__file__).parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))

st.set_page_config(
    page_title="기술이전 성과 자동 추출",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .main-title { font-size:2rem; font-weight:800; color:#1F3864; margin-bottom:0.2rem; }
  .sub-caption { color:#6B7280; font-size:0.95rem; margin-bottom:1.5rem; }
  .stDownloadButton > button {
    background-color:#1F3864 !important; color:white !important;
    font-size:1rem !important; padding:0.6rem 2rem !important; border-radius:8px !important;
  }
</style>
""", unsafe_allow_html=True)


def get_api_key():
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return os.environ.get("ANTHROPIC_API_KEY", "")


def detect_hint(filename: str) -> str:
    """파일명에서 힌트 텍스트 생성 (AI에게 참고용으로 전달)."""
    name = filename.lower()
    hints = []
    if any(k in name for k in ["bridge", "브릿지", "bridge3"]):
        hints.append("BRIDGE 3.0 사업 양식으로 보임. 기술이전유형은 거래유형+기술유형 조합으로 결정.")
    if any(k in name for k in ["tlo", "기술경영촉진", "대학기술경영"]):
        hints.append("TLO혁신형 양식으로 보임. 계약 단위로 집계(group_by_contract=true). 연구자명 포함 필요.")
    if any(k in name for k in ["세부현황", "29번", "실적세부"]):
        hints.append("기술이전 실적 세부 현황 양식으로 보임. 입금일 기준 필터(payment). 입금 회차별 행.")
    if any(k in name for k in ["산자부", "산업부", "현황조사"]):
        hints.append("산업부 현황조사 양식. 계약 단위 집계 가능성 높음.")
    return " / ".join(hints) if hints else ""


# ─── 사이드바 ────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ 추출 설정")
    st.divider()

    st.markdown("**📅 추출 연도**")
    year = st.number_input("year", min_value=2015, max_value=2035, value=2025,
                           label_visibility="collapsed")
    st.divider()

    st.markdown("**📂 마스터 DB**")
    st.caption("기술이전총정리.xlsx 형식")
    master_file = st.file_uploader("master_db", type=["xlsx"], key="master",
                                   label_visibility="collapsed")
    if master_file:
        st.success(f"✅ {master_file.name}  ({len(master_file.getvalue())//1024} KB)")
    st.divider()

    st.markdown("**📋 양식 파일** (필수)")
    st.caption("기관에서 받은 양식 파일. 내부 지침을 AI가 읽고 따릅니다.")
    template_file = st.file_uploader("template", type=["xlsx", "xls"], key="template",
                                     label_visibility="collapsed")
    if template_file:
        hint = detect_hint(template_file.name)
        st.info(f"📄 {template_file.name}  ({len(template_file.getvalue())//1024} KB)")
        if hint:
            st.caption(f"🔍 힌트: {hint}")
    else:
        hint = ""
    st.divider()

    st.markdown("**💬 추가 요청사항**")
    st.caption("특이사항, 필터 조건, 컬럼 설명 등 자유롭게 입력")
    notes = st.text_area(
        "notes",
        placeholder=(
            "예시:\n"
            "• 입금일 기준으로 필터링해주세요\n"
            "• 국외 건은 제외해주세요\n"
            "• 금액 단위가 천원입니다\n"
            "• 3번 컬럼은 정액+경상 합산 금액입니다\n"
            "• 기술자문 포함해주세요"
        ),
        height=160,
        label_visibility="collapsed"
    )
    st.divider()
    st.caption("© 부산대학교 산학협력단")


# ─── 메인 ───────────────────────────────────────────────────────
st.markdown('<p class="main-title">🔬 기술이전 성과 자동 추출 시스템</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-caption">양식 파일 내부 지침을 AI가 읽고 자동으로 데이터를 채웁니다. 양식이 달라져도 동작합니다.</p>',
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("추출 연도", f"{year}년")
with col2:
    st.metric("마스터 DB", "✅ 업로드됨" if master_file else "⏳ 미업로드")
with col3:
    st.metric("양식 파일", "✅ 업로드됨" if template_file else "⏳ 미업로드")

st.divider()

if not master_file or not template_file:
    missing = []
    if not master_file:
        missing.append("마스터 DB")
    if not template_file:
        missing.append("양식 파일")
    st.info(f"👈 왼쪽 사이드바에서 **{'과 '.join(missing)}**을 업로드해주세요.")

    st.markdown("""
    ### 📌 사용 방법
    1. **마스터 DB** 업로드 — `기술이전총정리_YYYYMMDD.xlsx`
    2. **양식 파일** 업로드 — 기관에서 받은 Excel 파일 (어떤 양식이든 가능)
    3. **추출 연도** 설정
    4. **추가 요청사항** 입력 *(선택)* — 특이사항이나 기관별 요청
    5. **추출 실행** 클릭
    6. 결과 파일 **다운로드**

    > 💡 AI가 양식 파일 내부의 작성 지침, 예시 데이터, 컬럼 헤더를 모두 읽고  
    > 그에 맞게 마스터 DB에서 데이터를 추출합니다.  
    > **양식이 바뀌어도 별도 설정 없이 동작합니다.**
    """)
    st.stop()

api_key = get_api_key()
if not api_key:
    st.error(
        "❌ Anthropic API 키가 설정되지 않았습니다.\n\n"
        "Streamlit Cloud → Settings → Secrets에 `ANTHROPIC_API_KEY`를 추가해주세요."
    )
    with st.expander("secrets.toml 예시"):
        st.code('ANTHROPIC_API_KEY = "sk-ant-api03-..."', language="toml")
    st.stop()

col_btn, col_hint = st.columns([1, 4])
with col_btn:
    run_btn = st.button("📊 추출 실행", type="primary", use_container_width=True)
with col_hint:
    st.info("🤖 AI가 양식 파일 내부 지침을 분석합니다. 10~30초 소요될 수 있습니다.")

if run_btn:
    with st.spinner(f"⏳ {year}년 데이터 추출 중... AI가 양식 지침을 분석하고 있습니다."):
        with tempfile.TemporaryDirectory() as tmpdir:
            master_path = os.path.join(tmpdir, "master.xlsx")
            ext = os.path.splitext(template_file.name)[1]
            template_path = os.path.join(tmpdir, f"template{ext}")
            output_path   = os.path.join(tmpdir, "output.xlsx")

            with open(master_path, "wb") as f:
                f.write(master_file.getvalue())
            with open(template_path, "wb") as f:
                f.write(template_file.getvalue())

            captured  = io.StringIO()
            old_stdout = sys.stdout
            sys.stdout = captured

            success    = False
            result_info = {}
            error_msg  = ""
            tb_str     = ""

            try:
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

                st.success("✅ 추출 완료!")
                st.balloons()

                summary_lines = []
                warn_lines    = []
                for line in log_output.strip().split("\n"):
                    if any(t in line for t in ["→", "✅", "총", "건", "수입"]):
                        summary_lines.append(line.strip())
                    if "⚠️" in line or "수동" in line:
                        warn_lines.append(line.strip())

                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    st.markdown("**📈 추출 결과 요약**")
                    for line in summary_lines[-5:]:
                        st.markdown(f"• {line.lstrip('→').strip()}")

                with col_r2:
                    manual = result_info.get("manual_inputs", [])
                    if manual:
                        st.markdown("**⚠️ 수동 입력 필요**")
                        for item in manual:
                            st.warning(item)

                if result_info.get("mapping_notes"):
                    st.info(f"🤖 AI 분석 노트: {result_info['mapping_notes']}")

                st.divider()

                tname = os.path.splitext(template_file.name)[0]
                output_filename = f"{tname}_{year}_자동추출.xlsx"

                col_dl, col_info = st.columns([1, 3])
                with col_dl:
                    st.download_button(
                        label="⬇️ 결과 파일 다운로드",
                        data=result_bytes,
                        file_name=output_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                with col_info:
                    st.caption(
                        f"파일명: `{output_filename}`  ·  "
                        f"'검토 결과' 시트에서 이상 데이터 및 수동 입력 필요 항목을 확인하세요."
                    )

                with st.expander("📋 처리 상세 로그"):
                    st.code(log_output, language="text")

            else:
                st.error("❌ 추출 중 오류가 발생했습니다.")
                st.markdown(f"**오류 내용:** `{error_msg}`")
                if log_output:
                    with st.expander("처리 로그"):
                        st.code(log_output, language="text")
                with st.expander("상세 오류 (개발자용)"):
                    st.code(tb_str, language="text")
                st.markdown("""
                **해결 방법:**
                - 마스터 DB에 `내역` 시트가 있는지 확인하세요.
                - 양식 파일이 손상되지 않았는지 확인하세요.
                - 추가 요청사항에 컬럼 설명을 추가하면 정확도가 높아집니다.
                """)
