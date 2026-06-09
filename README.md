# 한화손보 AI기반 상품 비교분석시스템

손해보험사 3N5 간편건강보험 상품 요약서 PDF를 업로드하면
담보 데이터를 자동 추출하여 비교분석 및 경쟁력 강화방안 보고서를 생성하는 AI 시스템입니다.

## 주요 기능

| 메뉴 | 기능 |
|------|------|
| 📁 파일 업로드 | 회사별 상품 요약서 PDF 업로드 (한화·삼성·현대·메리츠·KB·DB) |
| 🗂 상품 구조 비교 | 상품명·종·형·보험기간·납입기간·납입면제 등 11개 항목 자동 추출·비교 |
| 🔍 담보 내용 비교 | 담보명+지급사유 유사도 기반 1:1 자동 매핑 + 비교표 생성 |
| 📋 경쟁력 강화방안 | SUMMARY / 상품구조비교 / 담보운영비교 / SWOT / 액션플랜 보고서 자동 생성 |

## 로컬 실행

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. API 키 설정 (경쟁력 강화방안 메뉴 사용 시 필요)
# Windows
set ANTHROPIC_API_KEY=sk-ant-여기에키입력
# Mac/Linux
export ANTHROPIC_API_KEY=sk-ant-여기에키입력

# 3. 실행
streamlit run app.py
# 또는
python -m streamlit run app.py
```

## Render 배포

### 1. GitHub 레포지토리 생성 후 Push
```bash
git init
git add .
git commit -m "first commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

### 2. Render 연결
1. [render.com](https://render.com) 접속 → New → Web Service
2. GitHub 레포지토리 연결
3. 아래 설정 확인:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`
   - **Runtime**: Python 3

### 3. 환경변수 설정 (필수)
Render Dashboard → Environment 탭:

| Key | Value |
|-----|-------|
| `ANTHROPIC_API_KEY` | `sk-ant-api03-...` (Anthropic Console에서 발급) |

### 4. 배포 완료
- 첫 배포 약 3~5분 소요
- 이후 GitHub push 시 자동 재배포

## 파일 구조

```
/
├── app.py                    # 메인 애플리케이션
├── requirements.txt          # Python 의존성
├── render.yaml               # Render 배포 설정
├── .streamlit/
│   └── config.toml           # Streamlit 서버 설정
├── .gitignore
└── README.md
```

## 기술 스택

- **Frontend**: Streamlit
- **PDF 파싱**: pdfplumber
- **AI 보고서**: Anthropic Claude API (claude-sonnet-4-5)
- **엑셀 생성**: openpyxl
- **Word 생성**: python-docx
- **배포**: Render

## 주의사항

- PDF는 텍스트 추출 가능한 파일이어야 합니다 (스캔 이미지 PDF 불가)
- `경쟁력 강화방안` 메뉴는 `ANTHROPIC_API_KEY` 환경변수 필요
- Render 무료 플랜은 비활성 시 슬립 모드 진입 (첫 요청 시 약 30초 대기)
