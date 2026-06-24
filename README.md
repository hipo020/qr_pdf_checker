# PDF QR 링크 자동 검수기

CSV의 `페이지`, `카드번호`, `제목`, `유튜브링크`를 기준으로 PDF 안의 영상 카드 제목, QR URL, PDF 하이퍼링크 URL이 맞는지 검수하는 Streamlit 앱입니다.

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
streamlit run app.py
```

또는 Windows에서 `streamlit` 명령이 안 잡힐 때:

```bash
python -m streamlit run app.py
```

## 사용 순서

1. PDF 파일을 업로드합니다.
2. 기존 CSV 파일을 업로드합니다.
3. `검수 실행` 버튼을 누릅니다.
4. 화면 표에서 OK / CHECK / ISSUE를 확인합니다.
5. `엑셀 검수표 다운로드` 버튼으로 결과를 저장합니다.

## 주요 개선 기능

- 문제 항목 필터: 전체 / 문제만 / ISSUE만 / CHECK만 / 제목 문제 / QR 문제 / 하이퍼링크 문제
- 검색 기능: 제목, URL, 문제유형, 비고 검색
- 문제유형별 요약
- 페이지 미리보기: QR과 하이퍼링크 영역을 박스로 표시
- 제목 위치 자동 보정: 디자인이 조금 바뀌어도 카드 영역 안에서 CSV 제목과 가장 비슷한 텍스트를 탐색
- 엑셀 리포트 개선: 요약, 전체 검수, 문제 항목만, 기타 QR/링크 시트 구성

## 상태값

- OK: 제목, QR, 하이퍼링크가 기준 CSV와 일치하거나 실질적으로 같은 영상/플레이리스트로 판단됨
- CHECK: 영상 ID는 맞지만 URL 세부값이 다르거나 제목 위치 확인이 필요한 항목
- ISSUE: 제목, QR, 하이퍼링크 중 누락 또는 불일치가 있는 항목
- EXTRA_QR / EXTRA_LINK / INTERNAL_LINK: CSV 영상 카드 기준에 포함되지 않은 기타 QR/링크

## 문제유형

- TITLE_MISMATCH: PDF 카드 제목과 CSV 제목이 다름
- TITLE_CHECK: 제목은 페이지에 있으나 카드 영역에서 정확히 잡히지 않아 확인 필요
- MISSING_QR: QR을 찾지 못함
- QR_MISMATCH: QR URL이 CSV URL과 다름
- QR_URL_DETAIL_DIFF: 영상 ID는 같지만 URL 세부값이 다름
- MISSING_LINK: 하이퍼링크를 찾지 못함
- LINK_MISMATCH: 하이퍼링크 URL이 CSV URL과 다름
- LINK_URL_DETAIL_DIFF: 영상 ID는 같지만 URL 세부값이 다름

## 옵션

- QR 인식 해상도: QR 인식 실패가 있으면 2.5~3.0으로 올려보세요.
- 제목 탐색 Y 시작/끝: 제목의 기본 탐색 범위입니다.
- 제목 위치 자동 보정: 켜두는 것을 권장합니다.
- 영상 ID가 같으면 OK 처리: 켜두면 `index` 값만 다른 URL은 OK로 처리됩니다.
- 미리보기에서 QR/링크 박스 표시: 페이지 미리보기에서 QR/링크 위치를 확인할 수 있습니다.

## GitHub / Streamlit Cloud 업로드 파일

GitHub에는 아래 파일만 올리면 됩니다.

```text
app.py
requirements.txt
README.md
```

올리지 않아도 되는 파일:

```text
__pycache__/
*.pyc
```
