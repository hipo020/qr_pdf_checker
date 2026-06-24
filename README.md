# PDF QR 링크 자동 검수기

CSV의 `페이지`, `카드번호`, `제목`, `유튜브링크`를 기준으로 PDF 안의 영상 카드 제목, QR URL, PDF 하이퍼링크 URL, 연결된 YouTube 실제 제목을 검수하는 Streamlit 앱입니다.

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
streamlit run app.py
```

Windows에서 `streamlit` 명령이 안 잡힐 때:

```bash
python -m streamlit run app.py
```

## 사용 순서

1. PDF 파일을 업로드합니다.
2. 기존 CSV 파일을 업로드합니다.
3. 필요한 검수 옵션을 선택합니다.
4. `검수 실행` 버튼을 누릅니다.
5. `보기용 요약` 탭에서 문제 항목을 먼저 확인합니다.
6. URL 전체 확인이 필요하면 `상세 원자료` 탭을 봅니다.
7. 위치 확인이 필요하면 `페이지 미리보기` 탭을 봅니다.
8. `엑셀 검수표 다운로드` 버튼으로 결과를 저장합니다.

## 검수하는 항목

영상 카드 기준으로 아래 항목을 비교합니다.

- CSV 제목 = PDF 카드 안의 영상 제목
- CSV 유튜브링크 = QR 링크
- CSV 유튜브링크 = 카드 하이퍼링크
- QR 링크 = 카드 하이퍼링크
- 연결된 YouTube 실제 영상 제목 ≈ CSV 제목
- 연결된 YouTube 실제 영상 제목 ≈ PDF 카드 제목

YouTube 실제 제목 검수는 YouTube oEmbed를 사용하므로 인터넷 연결이 필요합니다. Streamlit Cloud에 배포한 앱에서는 사용할 수 있습니다.

## 화면 구성

- `보기용 요약`: 긴 URL을 숨기고 상태, 페이지, 카드번호, 제목, 링크 검수, 조치 메모 중심으로 표시합니다.
- `상세 원자료`: CSV URL, QR URL, 하이퍼링크 URL, YouTube 실제 제목 등 모든 데이터를 표시합니다.
- `기타 QR/링크`: CSV 영상 카드 기준에 매칭되지 않은 QR, 외부 링크, 내부 페이지 이동 링크를 표시합니다.
- `페이지 미리보기`: PDF 페이지 위에 QR/하이퍼링크 위치를 박스로 표시합니다.

## 상태값

- `OK`: 제목, QR, 하이퍼링크가 기준 CSV와 일치하거나 실질적으로 같은 영상/플레이리스트로 판단됨
- `CHECK`: 영상 ID는 맞지만 URL 세부값이 다르거나, 제목 유사도가 애매해 확인이 필요한 항목
- `ISSUE`: 제목, QR, 하이퍼링크, YouTube 실제 제목 중 누락 또는 불일치 가능성이 높은 항목
- `EXTRA_QR` / `EXTRA_LINK` / `INTERNAL_LINK`: CSV 영상 카드 기준에 포함되지 않은 기타 QR/링크

## 주요 문제유형

- `TITLE_MISMATCH`: PDF 카드 제목과 CSV 제목이 다름
- `TITLE_CHECK`: 제목은 페이지에 있으나 카드 영역에서 정확히 잡히지 않아 확인 필요
- `MISSING_QR`: QR을 찾지 못함
- `QR_MISMATCH`: QR URL이 CSV URL과 다름
- `MISSING_LINK`: 하이퍼링크를 찾지 못함
- `LINK_MISMATCH`: 하이퍼링크 URL이 CSV URL과 다름
- `QR_LINK_MISMATCH`: QR URL과 카드 하이퍼링크 URL이 서로 다름
- `YOUTUBE_TITLE_MISMATCH`: 연결된 YouTube 실제 제목과 CSV/PDF 제목 차이 큼
- `YOUTUBE_TITLE_CHECK`: 연결된 YouTube 실제 제목과 CSV/PDF 제목 유사도 확인 권장
- `YOUTUBE_TITLE_FETCH_FAIL`: YouTube 실제 제목을 가져오지 못함

## 옵션

- `QR 인식 해상도`: QR 인식 실패가 있으면 2.5~3.0으로 올려보세요.
- `제목 탐색 Y 시작/끝`: 제목의 기본 탐색 범위입니다.
- `제목 위치 자동 보정`: 카드 영역 안에서 CSV 제목과 가장 비슷한 텍스트를 찾습니다. 켜두는 것을 권장합니다.
- `영상 ID가 같으면 OK 처리`: 켜두면 `index` 값만 다른 URL은 OK로 처리합니다.
- `YouTube 실제 영상 제목까지 검수`: QR/하이퍼링크의 영상 URL로 실제 YouTube 제목을 가져와 비교합니다.
- `유튜브 제목 검수 기준`: `보통 (추천)`, `엄격`, `느슨` 중 선택합니다. 유튜브 제목 앞의 `Kia How To |` 접두어만 제외하고 나머지 문구는 모두 비교합니다.

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


## 유튜브 제목 검수 기준

앱 화면에서는 숫자 기준 대신 `보통 (추천)`, `엄격`, `느슨` 중 하나를 선택합니다.

- `보통 (추천)`: 일반적인 검수용 기본값입니다.
- `엄격`: 제목 문구 차이도 더 민감하게 잡습니다. 최종 납품 전 확인용으로 적합합니다.
- `느슨`: 유튜브 제목에 부가 문구가 많아도 핵심 제목이 비슷하면 통과시키는 기준입니다.

유튜브 제목 앞에 붙는 `Kia How To |` 접두어만 제목 비교에서 제외합니다. `How to use digital key 2`, `Guide`, `Video` 같은 나머지 문구는 실제 제목의 일부로 보고 비교합니다.
