# Streamlit Cloud 배포 가이드

## 1. GitHub 저장소 준비

이 `streamlit_app/` 폴더 전체를 GitHub private 저장소에 업로드합니다.

```
my-repo/
├── app.py
├── requirements.txt
├── scripts/
│   ├── common.py
│   ├── feedback.py
│   ├── extract_bridge30.py
│   ├── extract_tlo.py
│   ├── extract_detail.py
│   ├── analyze_template.py
│   └── extract_generic.py
├── references/
│   ├── db_columns.md
│   └── transform_rules.md
└── .streamlit/
    └── config.toml
```

## 2. Streamlit Cloud 배포

1. [share.streamlit.io](https://share.streamlit.io) 접속
2. **New app** 클릭
3. GitHub 저장소 연결
4. Main file path: `app.py`
5. **Deploy** 클릭

## 3. API 키 설정

배포 후 **Settings → Secrets** 에서 아래 추가:

```toml
ANTHROPIC_API_KEY = "sk-ant-api03-여기에실제키입력"
```

## 4. 이메일 접근 제어 설정

Streamlit Cloud Community 플랜에서 이메일 기반 접근 제어:

1. 앱 설정 → **Sharing** 탭
2. **"Invite viewers only"** 선택
3. 허용할 이메일 주소 입력 (팀원 이메일)
4. Save

접근 허용된 이메일만 로그인 후 사용 가능.

## 5. 업데이트 방법

GitHub 저장소에 파일을 push하면 Streamlit Cloud가 자동으로 재배포합니다.
