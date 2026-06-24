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
3. 필요한 검수 옵션을 선택합니다.
4. `검수 실행` 버튼을 누릅니다.
5. 화면 표에서 OK / CHECK / ISSUE를 확인합니다.
6. `엑셀 검수표 다운로드` 버튼으로 결과를 저장합니다.

## 검수하는 항목

영상 카드 기준으로 아래 항목을 비교합니다.

- CSV 제목 = PDF 카드 안의 영상 제목
- CSV 유튜브링크 = QR 링크
- CSV 유튜브링크 = 카드 하이퍼링크
- QR 링크 = 카드 하이퍼링크
- 연결된 유튜브 실제 영상 제목 ≈ CSV 제목
- 연결된 유튜브 실제 영상 제목 ≈ PDF 카드 제목

유튜브 실제 제목 검수는 YouTube oEmbed를 사용하므로 인터넷 연결이 필요합니다. Streamlit Cloud에 배포한 앱에서는 사용할 수 있습니다.

## 주요 기능

- 문제 항목 필터: 전체 / 문제만 / ISSUE만 / CHECK만 / 제목 문제 / 유튜브 제목 문제 / QR 문제 / 하이퍼링크 문제
- 검색 기능: 제목, URL, 문제유형, 비고 검색
- 문제유형별 요약
- 페이지 미리보기: QR과 하이퍼링크 영역을 박스로 표시
- 제목 위치 자동 보정: 디자인이 조금 바뀌어도 카드 영역 안에서 CSV 제목과 가장 비슷한 텍스트를 탐색
- 유튜브 실제 제목 검수: 연결된 영상의 실제 제목을 가져와 CSV/PDF 제목과 유사도 비교
- 엑셀 리포트: 요약, 전체 검수, 문제 항목만, 기타 QR/링크 시트 구성

## 상태값

- OK: 제목, QR, 하이퍼링크가 기준 CSV와 일치하거나 실질적으로 같은 영상/플레이리스트로 판단됨
- CHECK: 영상 ID는 맞지만 URL 세부값이 다르거나, 제목 유사도가 애매해 확인이 필요한 항목
- ISSUE: 제목, QR, 하이퍼링크, 유튜브 실제 제목 중 누락 또는 불일치 가능성이 높은 항목
- EXTRA_QR / EXTRA_LINK / INTERNAL_LINK: CSV 영상 카드 기준에 포함되지 않은 기타 QR/링크

## 문제유형

- TITLE_MISMATCH: PDF 카드 제목과 CSV 제목이 다름
- TITLE_CHECK: 제목은 페이지에 있으나 카드 영역에서 정확히 잡히지 않아 확인 필요
- YOUTUBE_TITLE_MISMATCH: 연결된 유튜브 실제 제목과 CSV/PDF 제목 차이 큼
- YOUTUBE_TITLE_CHECK: 연결된 유튜브 실제 제목과 CSV/PDF 제목 유사도 확인 권장
- YOUTUBE_TITLE_FETCH_FAIL: 유튜브 실제 제목을 가져오지 못함
- MISSING_QR: QR을 찾지 못함
- QR_MISMATCH: QR URL이 CSV URL과 다름
- QR_URL_DETAIL_DIFF: 영상 ID는 같지만 URL 세부값이 다름
- MISSING_LINK: 하이퍼링크를 찾지 못함
- LINK_MISMATCH: 하이퍼링크 URL이 CSV URL과 다름
- LINK_URL_DETAIL_DIFF: 영상 ID는 같지만 URL 세부값이 다름
- QR_LINK_MISMATCH: QR URL과 카드 하이퍼링크 URL이 서로 다름

## 옵션

- QR 인식 해상도: QR 인식 실패가 있으면 2.5~3.0으로 올려보세요.
- 제목 탐색 Y 시작/끝: 제목의 기본 탐색 범위입니다.
- 제목 위치 자동 보정: 켜두는 것을 권장합니다.
- 영상 ID가 같으면 OK 처리: 켜두면 `index` 값만 다른 URL은 OK로 처리됩니다.
- 유튜브 실제 영상 제목까지 검수: QR/하이퍼링크의 영상 URL로 실제 YouTube 제목을 가져와 비교합니다.
- 유튜브 제목 OK 기준: 이 값 이상이면 OK로 처리합니다.
- 유튜브 제목 CHECK 기준: 이 값 이상이면 CHECK, 미만이면 ISSUE로 처리합니다.
- 미리보기에서 QR/링크 박스 표시: 페이지 미리보기에서 QR/링크 위치를 확인할 수 있습니다.

## GitHub / Streamlit Cloud 업로드 파일

GitHub에는 아래 파일만 올리면 됩니다.

```text
app.py
requirements.txt
README.md
.gitignore
```

올리지 않아도 되는 파일:

```text
__pycache__/
*.pyc
```
