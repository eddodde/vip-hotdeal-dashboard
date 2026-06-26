# 🔥 VIP 핫딜 트렌드 대시보드

VIP라운지 **핫딜 콘텐츠**의 매출·트래픽·전환을 **기간별 추세**로 보는 Streamlit 대시보드.
데일리 실적은 따로 확인하므로, 이 대시보드는 "흐름(트렌드)" 파악에 집중합니다.

## 분석 섹션
- **핵심 요약** — 거래액·주문·UV·전환율·객단가 + 직전 동기간 대비 증감
- **매출 추세** — 일/주/월 거래액·주문 추이, 이동평균, 연도(YoY) 비교
- **트래픽 추세** — 영역별(VIP핫딜·VIP라운지·기념일·기획전) UV/PV, 핫딜 점유율
- **전환·효율** — 전환율(주문/UV), 객단가 추이
- **베스트** — 상품·브랜드·MD·카테고리 거래액 랭킹
- **브랜드·카테고리 믹스** — 매출 구성 변화 (stacked area)
- **오전·오후 슬롯** — 슬롯별 매출 비교·점유

## 데이터
repo에 커밋된 깨끗한 CSV에서 로드합니다.
- `data/hotdeal.csv` — 슬롯·상품 단위 일별 매출 (date, slot, brand, category, md, UV/PV/주문/거래액 …)
- `data/table_trend.csv` — 콘텐츠 영역별 일별 UV/PV (date, metric, area, value)

기간: 2023년 + 2024-07 ~ 현재. **2024년 상반기는 원본에 데이터 공백**(두 export를 이어붙인 형태)이라 차트에서 선이 끊겨 표시됩니다.

## 데이터 갱신 (중요)
원본 `핫딜.xlsx` / `Table.xlsx`는 회사 **DRM(Softcamp)으로 암호화**돼 있어 pandas/openpyxl로 직접 못 읽고, Excel "CSV로 저장"도 다시 암호화돼 깨집니다. 그래서 변환은 **Excel COM 기반 PowerShell 스크립트**로 합니다.

```powershell
# 두 원본 파일을 Downloads에 받아둔 뒤:
./convert.ps1
# (다른 위치면)  ./convert.ps1 -Hotdeal "경로\핫딜.xlsx" -Table "경로\Table.xlsx"
```

`data/*.csv`가 새로 생성됩니다. 이후 반영 방법은 **둘 중 하나**:

**① git push (영구 반영)**
```powershell
git add data; git commit -m "데이터 갱신"; git push
```

**② 브라우저 업로드 (그때그때, git 불필요)**
대시보드 사이드바 **⚙️ 설정 → 📤 데이터 올리기**에서 방금 만든
`hotdeal.csv` / `table_trend.csv` 를 올리면 그 세션에 바로 반영됩니다.

> ⚠️ 원본 `.xlsx` 는 DRM 암호화라 업로드해도 못 읽습니다. **반드시 `convert.ps1` 로 만든 CSV** 를 올리세요(잘못 올리면 안내 메시지가 뜹니다).

## 로컬 실행 (선택)
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 지표 메모
- **거래액**은 VAT 제외 금액이며 원 단위 분수값일 수 있음(가격/1.1 등).
- **UV/PV**는 페이지 총계(중복 제거), **매출·주문**은 슬롯(오전+오후) 합산.
- **전환율** = 주문건수 / 페이지 UV (슬롯 필터 영향 없음).
