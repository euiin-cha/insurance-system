# 한화손보 AI기반 상품 비교분석시스템

손해보험사 간편건강보험 상품 요약서 PDF를 업로드하면  
담보를 자동 추출·비교하고 경쟁력 강화방안 보고서를 생성하는 AI 시스템입니다.

## 주요 기능

| 메뉴 | 기능 |
|------|------|
| 📁 파일 업로드 | 회사별 상품 요약서 PDF 업로드 (한화·삼성·현대·메리츠·KB·DB) |
| 🗂 상품 구조 비교 | 상품명·종·형·보험기간·납입기간·납입면제 등 11개 항목 자동 추출·비교 |
| 🔍 담보 내용 비교 | 담보명+지급사유 유사도 기반 1:1 자동 매핑 + 비교표 생성 |
| 📋 경쟁력 강화방안 | SUMMARY / 상품구조비교 / 담보운영비교 / SWOT / 액션플랜 보고서 자동 생성 |

## 로컬 실행

```bash
# 1. 패키지 설치
pip install -r requirements.txt

# 2. API 키 설정 (경쟁력 강화방안 메뉴 사용 시)
set ANTHROPIC_API_KEY=sk-ant-여기에키입력   # Windows
export ANTHROPIC_API_KEY=sk-ant-여기에키입력 # Mac/Linux

# 3. 실행
python -m streamlit run app.py
```

## Render 배포 방법

### 1단계: GitHub에 올리기

```bash
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/아이디/레포이름.git
git push -u origin main
```

### 2단계: Render 연결

1. [render.com](https://render.com) → **New → Web Service**
2. GitHub 레포 선택
3. 설정 확인 (자동 입력됨):
   - Build: `pip install -r requirements.txt`
   - Start: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`

### 3단계: 환경변수 설정 (**필수**)

Render Dashboard → **Environment 탭** 에서 추가:

| Key | Value |
|-----|-------|
| `ANTHROPIC_API_KEY` | `sk-ant-api03-...` |

→ **Save Changes** 후 **Manual Deploy**

## 파일 구조

```
/
├── app.py                  # 메인 앱
├── requirements.txt        # 의존성
├── render.yaml             # Render 설정
├── .streamlit/
│   └── config.toml         # Streamlit 서버 설정
├── .gitignore
└── README.md
```

## 주의사항

- PDF는 텍스트 추출 가능한 파일이어야 합니다 (스캔 이미지 PDF 불가)
- `경쟁력 강화방안` 메뉴는 `ANTHROPIC_API_KEY` 환경변수 필요
- Render 무료 플랜은 15분 비활성 시 슬립 → 첫 요청 시 약 30초 대기
