
"""
기술이전 성과 자동 추출 시스템
부산대학교 산학협력단 마스터 DB → 상위기관 보고 양식 자동 변환
 
Streamlit Cloud 배포용
"""
import streamlit as st
import tempfile
import os
import io
import sys
import json
import traceback
from pathlib import Path
 
# 스크립트 경로 추가
SCRIPTS_DIR = Path(__file__).parent / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
 
# ───────────────────────────────────────────────────────────────
# 페이지 설정
# ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="기술이전 성과 자동 추출",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded",
)
 
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 800; color: #1F3864; margin-bottom: 0.2rem; }
    .sub-caption { color: #6B7280; font-size: 0.95rem; margin-bottom: 1.5rem; }
    .form-badge {
        display: inline-block;
        background: #DBEAFE; color: #1E40AF;
        padding: 2px 10px; border-radius: 12px;
        font-size: 0.82rem; font-weight: 600;
    }
    .ai-badge {
        display: inline-block;
        background: #EDE9FE; color: #5B21B6;
        padding: 2px 10px; border-radius: 12px;
        font-size: 0.82rem; font-weight: 600;
    }
    .stDownloadButton > button {
        background-color: #1F3864 !important;
        color: white !important;
        font-size: 1rem !important;
        padding: 0.6rem 2rem !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)
 
# ───────────────────────────────────────────────────────────────
# 상수 / 설정
# ───────────────────────────────────────────────────────────────
 
KNOWN_FORMS = {
    "BRIDGE 3.0": {
        "keywords": ["bridge", "브릿지", "bridge3", "bridge 3", "브리지"],
        "desc": "BRIDGE 3.0 사업 기술이전 세부 목록",
        "needs_template": False,
    },
    "TLO혁신형": {
        "keywords": ["tlo", "대학기술경영", "기술경영촉진", "tlo혁신"],
        "desc": "대학기술경영촉진사업 기술이전 증빙",
        "needs_template": False,
    },
    "기술이전 실적 세부 현황": {
        "keywords": ["세부현황", "실적세부", "29번", "29호", "세부 현황", "실적세부현황"],
        "desc": "기술이전 실적 세부 현황 (29번 양식)",
        "needs_template": False,
    },
}
 
FORM_DISPLAY_NAMES = {
    "BRIDGE 3.0": "📊 BRIDGE 3.0",
    "TLO혁신형": "🎓 TLO혁신형",
    "기술이전 실적 세부 현황": "📋 실적 세부 현황",
    "새 양식 (AI 자동 분석)": "🤖 새 양식 (AI)",
}
 
# ───────────────────────────────────────────────────────────────
# 헬퍼 함수
# ───────────────────────────────────────────────────────────────
 
def detect_form_type(filename: str) -> str:
    """파일명에서 양식 유형 자동 감지."""
    name = filename.lower()
    for form, info in KNOWN_FORMS.items():
        if any(kw in name for kw in info["keywords"]):
            return form
    return "새 양식 (AI 자동 분석)"
 
 
def get_api_key() -> str:
    """Streamlit secrets 또는 환경변수에서 API 키 가져오기."""
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        return os.environ.get("ANTHROPIC_API_KEY", "")
 
 
def run_extraction(master_path, template_path, year, form_type, notes, output_path, api_key):
    """양식 유형에 따라 적절한 추출 스크립트 실행."""
    from extract_bridge30 import run as run_bridge
    from extract_tlo import run as run_tlo
    from extract_detail import run as run_detail
    from extract_generic import run as run_generic
 
    if form_type == "BRIDGE 3.0":
        run_bridge(master_path, year, output_path, template_path)
        return {"form": "BRIDGE 3.0", "count": None}
 
    elif form_type == "TLO혁신형":
        run_tlo(master_path, year, output_path, template_path=template_path)
        return {"form": "TLO혁신형", "count": None}
 
    elif form_type == "기술이전 실적 세부 현황":
        run_detail(master_path, year, output_path)
        return {"form": "기술이전 실적 세부 현황", "count": None}
 
    else:
        # 새 양식: AI 자동 분석
        if not template_path:
            raise ValueError("새 양식 분석을 위해 양식 파일이 필요합니다.")
        result = run_generic(master_path, template_path, year, output_path, notes, api_key)
        result["form"] = "새 양식 (AI 분석)"
        return result
 
 
# ───────────────────────────────────────────────────────────────
# 사이드바
# ───────────────────────────────────────────────────────────────
 
with st.sidebar:
    st.markdown("## ⚙️ 추출 설정")
    st.divider()
 
    # 1. 연도
    st.markdown("**📅 추출 연도**")
    year = st.number_input(
        "year", min_value=2015, max_value=2035, value=2025,
        label_visibility="collapsed"
    )
 
    st.divider()
 
    # 2. 마스터 DB
    st.markdown("**📂 마스터 DB**")
    st.caption("기술이전총정리.xlsx 형식")
    master_file = st.file_uploader(
        "master_db", type=["xlsx"], key="master",
        label_visibility="collapsed"
    )
    if master_file:
        size_kb = len(master_file.getvalue()) / 1024
        st.success(f"✅ {master_file.name}  ({size_kb:.0f} KB)")
 
    st.divider()
 
    # 3. 양식 파일
    st.markdown("**📋 출력 양식 파일** (선택)")
    st.caption("없으면 기본 형식으로 생성. 새 양식은 필수.")
    template_file = st.file_uploader(
        "template", type=["xlsx", "xls"], key="template",
        label_visibility="collapsed"
    )
 
    auto_detected = ""
    if template_file:
        size_kb = len(template_file.getvalue()) / 1024
        auto_detected = detect_form_type(template_file.name)
        st.info(f"📄 {template_file.name}  ({size_kb:.0f} KB)")
 
    st.divider()
 
    # 4. 양식 유형
    st.markdown("**🏷️ 양식 유형 선택**")
    if auto_detected:
        st.caption(f"자동 감지: **{auto_detected}**")
 
    form_options = ["자동 감지"] + list(KNOWN_FORMS.keys()) + ["새 양식 (AI 자동 분석)"]
    form_type_sel = st.selectbox(
        "form_type", form_options, label_visibility="collapsed"
    )
 
    # 실제 사용할 양식 유형 결정
    if form_type_sel == "자동 감지":
        form_type = auto_detected if auto_detected else "새 양식 (AI 자동 분석)"
    else:
        form_type = form_type_sel
 
    # 선택된 양식 표시
    if form_type in FORM_DISPLAY_NAMES:
        badge = "ai-badge" if "AI" in form_type else "form-badge"
        st.markdown(
            f'→ <span class="{badge}">{FORM_DISPLAY_NAMES.get(form_type, form_type)}</span>',
            unsafe_allow_html=True
        )
 
    st.divider()
 
    # 5. 추가 요청사항
    st.markdown("**💬 추가 요청사항**")
    st.caption("특이사항, 새 양식 컬럼 설명, 필터 조건 등")
    notes = st.text_area(
        "notes",
        placeholder=(
            "예시:\n"
            "• 입금일 기준으로만 필터링해주세요\n"
            "• 국외 건은 제외해주세요\n"
            "• 금액 단위가 천원입니다\n"
            "• 3번 컬럼은 총 기술료(정액+경상)입니다\n"
            "• 기술자문 포함해주세요"
        ),
        height=150,
        label_visibility="collapsed"
    )
 
    st.divider()
    st.caption("© 부산대학교 산학협력단")
 
 
# ───────────────────────────────────────────────────────────────
# 메인 화면
# ───────────────────────────────────────────────────────────────
 
st.markdown('<p class="main-title">🔬 기술이전 성과 자동 추출 시스템</p>', unsafe_allow_html=True)
st.markdown(
    '<p class="sub-caption">부산대학교 산학협력단 마스터 DB → 상위기관 보고 양식 자동 변환 · 알려진 양식 및 새 양식 모두 지원</p>',
    unsafe_allow_html=True
)
 
# ── 상태 요약 카드 ──
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("추출 연도", f"{year}년")
with col2:
    st.metric("마스터 DB", "✅ 업로드됨" if master_file else "⏳ 미업로드")
with col3:
    st.metric("양식 파일", "✅ 업로드됨" if template_file else "미업로드 (선택)")
with col4:
    st.metric("선택 양식", FORM_DISPLAY_NAMES.get(form_type, form_type))
 
st.divider()
 
# ── 마스터 DB 없을 때 안내 ──
if not master_file:
    tab1, tab2 = st.tabs(["📌 사용 방법", "📋 지원 양식 목록"])
 
    with tab1:
        st.markdown("""
        ### 사용 순서
 
        1. **마스터 DB** 업로드 — `기술이전총정리_YYYYMMDD.xlsx`
        2. **양식 파일** 업로드 *(선택)* — 제출할 기관의 양식 파일
        3. **추출 연도** 설정
        4. **양식 유형** 확인 — 파일명으로 자동 감지, 또는 직접 선택
        5. **추가 요청사항** 입력 — 새 양식이거나 특이사항이 있는 경우
        6. **추출 실행** 버튼 클릭
        7. 결과 파일 **다운로드**
 
        > 💡 **새 양식**은 양식 파일을 업로드하면 AI가 컬럼을 자동 분석합니다.
        > 추가 요청사항 란에 컬럼 설명이나 필터 조건을 자유롭게 입력해주세요.
        """)
 
    with tab2:
        st.markdown("""
        | 양식명 | 자동 인식 키워드 | 양식 파일 필요 여부 |
        |---|---|---|
        | BRIDGE 3.0 기술이전 세부 목록 | bridge, 브릿지 | 선택 (있으면 기존 양식 활용) |
        | 대학기술경영촉진사업 기술이전 증빙 | TLO, 대학기술경영 | 불필요 |
        | 기술이전 실적 세부 현황 (29번) | 세부현황, 29번, 실적세부 | 불필요 |
        | **새 양식** (AI 자동 분석) | *그 외 모든 파일명* | **필수** |
 
        > 새 양식은 Anthropic Claude API를 사용하여 컬럼을 분석합니다.
        > API 키가 설정되어 있어야 합니다.
        """)
 
    st.stop()
 
# ── 새 양식 검증 ──
if form_type == "새 양식 (AI 자동 분석)" and not template_file:
    st.warning("⚠️ 새 양식 AI 분석을 위해 왼쪽에서 **양식 파일**을 업로드해주세요.")
    st.stop()
 
# ── API 키 검증 (새 양식만) ──
api_key = get_api_key()
if form_type == "새 양식 (AI 자동 분석)" and not api_key:
    st.error(
        "❌ Anthropic API 키가 설정되지 않았습니다.\n\n"
        "Streamlit Cloud에서 **Settings → Secrets**에 `ANTHROPIC_API_KEY`를 추가해주세요."
    )
    with st.expander("secrets.toml 예시"):
        st.code('ANTHROPIC_API_KEY = "sk-ant-api03-..."', language="toml")
    st.stop()
 
# ── 실행 버튼 ──
col_btn, col_hint = st.columns([1, 4])
with col_btn:
    run_btn = st.button("📊 추출 실행", type="primary", use_container_width=True)
with col_hint:
    if form_type == "새 양식 (AI 자동 분석)":
        st.info("🤖 AI가 양식을 분석합니다. 10~30초 소요될 수 있습니다.")
    elif notes:
        st.info(f"💬 추가 요청사항 적용됨: {notes[:60]}{'...' if len(notes) > 60 else ''}")
 
# ── 추출 실행 ──
if run_btn:
    progress_placeholder = st.empty()
    log_placeholder = st.empty()
 
    with st.spinner(f"⏳ {year}년 {form_type} 데이터 추출 중..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            # 파일 저장
            master_path = os.path.join(tmpdir, "master.xlsx")
            output_path = os.path.join(tmpdir, "output.xlsx")
 
            with open(master_path, "wb") as f:
                f.write(master_file.getvalue())
 
            template_path = None
            if template_file:
                ext = os.path.splitext(template_file.name)[1]
                template_path = os.path.join(tmpdir, f"template{ext}")
                with open(template_path, "wb") as f:
                    f.write(template_file.getvalue())
 
            # stdout 캡처
            captured = io.StringIO()
            old_stdout = sys.stdout
            sys.stdout = captured
 
            success = False
            result_info = {}
            error_msg = ""
            tb_str = ""
 
            try:
                result_info = run_extraction(
                    master_path, template_path, year,
                    form_type, notes, output_path, api_key
                )
                success = True
            except Exception as e:
                error_msg = str(e)
                tb_str = traceback.format_exc()
            finally:
                sys.stdout = old_stdout
 
            log_output = captured.getvalue()
 
            if success and os.path.exists(output_path):
                with open(output_path, "rb") as f:
                    result_bytes = f.read()
 
                # ── 성공 UI ──────────────────────────────────────
                st.success(f"✅ 추출 완료!")
                st.balloons()
 
                # 로그에서 요약 정보 파싱
                summary_lines = []
                manual_lines = []
                for line in log_output.strip().split("\n"):
                    if any(tag in line for tag in ["→", "✅", "총", "건", "수입"]):
                        summary_lines.append(line.strip())
                    if "⚠️" in line or "수동" in line:
                        manual_lines.append(line.strip())
 
                col_r1, col_r2 = st.columns(2)
                with col_r1:
                    st.markdown("**📈 추출 결과 요약**")
                    for line in summary_lines[-5:]:
                        st.markdown(f"• {line.lstrip('→').strip()}")
 
                with col_r2:
                    if manual_lines:
                        st.markdown("**⚠️ 수동 확인 필요**")
                        for line in manual_lines:
                            st.warning(line.lstrip("⚠️").strip())
 
                # AI 매핑 노트
                if result_info.get("mapping_notes"):
                    st.info(f"🤖 AI 분석 노트: {result_info['mapping_notes']}")
 
                st.divider()
 
                # ── 다운로드 버튼 ────────────────────────────────
                output_filename = f"{result_info.get('form', form_type)}_{year}_자동생성.xlsx"
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
                        f"파일명: `{output_filename}` · "
                        f"'검토 결과' 시트에서 이상 데이터 및 수동 입력 필요 항목을 확인하세요."
                    )
 
                # ── 처리 로그 (접기) ─────────────────────────────
                with st.expander("📋 처리 상세 로그"):
                    st.code(log_output, language="text")
 
            else:
                # ── 오류 UI ──────────────────────────────────────
                st.error(f"❌ 추출 중 오류가 발생했습니다.")
                st.markdown(f"**오류 내용:** `{error_msg}`")
 
                if log_output:
                    with st.expander("처리 로그"):
                        st.code(log_output, language="text")
 
                with st.expander("상세 오류 (개발자용)"):
                    st.code(tb_str, language="text")
 
                st.markdown("""
                **해결 방법:**
                - 마스터 DB 파일이 `내역` 시트를 포함하는지 확인해주세요.
                - 양식 파일이 손상되지 않았는지 확인해주세요.
                - 추가 요청사항에 구체적인 설명을 추가하면 새 양식 분석 정확도가 높아집니다.
                """)
