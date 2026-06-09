# InsureCompare

손해보험 6개사 상품요약서 PDF를 AI로 자동 비교분석하는 웹 애플리케이션입니다.

## 주요 기능

- **상품 구조 비교**: 상품명·종·형·보통약관·보험기간·납입기간·가입연령·납입면제·부가서비스·할인 11개 항목 자동 추출
- **담보 내용 비교**: Bigram Dice 유사도 기반 담보 자동 매핑 (담보명→지급사유 2단계)
- **장단점 보고서**: 기준사 대비 비교사 경쟁력 자동 분석
- **엑셀 저장**: 분석 결과 .xlsx 파일 다운로드

## 지원 회사

한화손해보험 / 삼성화재 / 현대해상 / 메리츠화재 / KB손해보험 / DB손해보험

## 사용 방법

1. [Anthropic Console](https://console.anthropic.com)에서 API 키 발급
2. 앱 상단 API 키 입력란에 입력
3. 각 회사 상품요약서 PDF 업로드
4. 원하는 분석 탭에서 기준사/비교사 선택 후 실행

## 기술 스택

- Frontend: Vanilla HTML/CSS/JavaScript (단일 파일)
- AI: Claude Sonnet 4.6 API
- PDF 파싱: PDF.js 3.11.174
- 엑셀 출력: SheetJS (xlsx)
- 폰트/아이콘: Google Fonts (Noto Sans KR), Tabler Icons

## 로컬 실행

별도 서버 불필요. `index.html`을 브라우저에서 직접 열거나:

```bash
# Python 간이 서버 사용 시
python3 -m http.server 8080
# → http://localhost:8080 접속
```

## Render 배포

1. 이 저장소를 GitHub에 push
2. [Render](https://render.com) → New → Static Site
3. Repository 연결
4. 설정:
   - **Build Command**: (비워두기)
   - **Publish Directory**: `.`
5. Deploy

## 주의사항

- Claude API 키는 브라우저에서 직접 호출됩니다 (CORS 허용 헤더 포함)
- API 키는 사용자 브라우저에만 존재하며 서버에 저장되지 않습니다
- PDF 파싱은 클라이언트에서 처리되며 서버로 파일이 전송되지 않습니다

## 한화손해보험 AI 경진대회 출품작

2026년 한화손해보험 AI 경진대회 출품 프로젝트입니다.
