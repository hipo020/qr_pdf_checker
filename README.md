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

## 사용 순서

1. 브라우저가 열리면 PDF 파일을 업로드합니다.
2. 기존 CSV 파일을 업로드합니다.
3. `검수 실행` 버튼을 누릅니다.
4. 화면 표에서 OK / CHECK / ISSUE를 확인합니다.
5. `엑셀 검수표 다운로드` 버튼으로 결과를 저장합니다.

## 상태값

- OK: 제목, QR, 하이퍼링크가 모두 기준 CSV와 일치
- CHECK: 영상 ID는 맞지만 index 등 URL 세부값이 다르거나 제목 위치 확인이 필요한 항목
- ISSUE: 제목, QR, 하이퍼링크 중 누락 또는 불일치가 있는 항목
- EXTRA_QR / EXTRA_LINK / INTERNAL_LINK: CSV 영상 카드 기준에 포함되지 않은 기타 QR/링크

## 옵션

- QR 인식 해상도: QR 인식 실패가 있으면 2.5~3.0으로 올려보세요.
- 제목 탐색 Y 시작/끝: PDF 디자인에서 제목 위치가 바뀌면 조정하세요. 현재 샘플 PDF 기준 기본값은 520~600입니다.
