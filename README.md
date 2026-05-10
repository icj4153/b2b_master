# B2B Excel Sourcing Pipeline

B2B 도매처 상품 엑셀을 자동 수집하고, 상품명/가격 정보를 표준화한 뒤 제철 키워드와 네이버 트렌드 분석 화면으로 확인하는 Python 프로젝트입니다.

## 주요 파일

- `b2b_excel.py`: 도매처 엑셀 다운로드 및 통합
- `season_analyzer.py`: 통합 데이터에 제철 키워드/월 정보 추가
- `web_page.py`: Streamlit 분석 대시보드
- `env_utils.py`: `.env` 환경변수 로더
- `.env.example`: 필요한 환경변수 예시

## 설치

```powershell
pip install -r requirements.txt
playwright install
```

## 환경변수

`.env.example`을 참고해서 `.env` 파일을 만듭니다.

```text
B2B_USER_ID=your_b2b_login_id
B2B_USER_PW=your_b2b_login_password
NAVER_CLIENT_ID=your_naver_client_id
NAVER_CLIENT_SECRET=your_naver_client_secret
```

`.env`는 비밀번호와 API 키가 들어가므로 GitHub에 올리지 않습니다.

## 실행

도매처 엑셀 수집 및 통합:

```powershell
python b2b_excel.py
```

제철 분석 파일 생성:

```powershell
python season_analyzer.py
```

Streamlit 대시보드 실행:

```powershell
streamlit run web_page.py
```
