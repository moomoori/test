# 정산 대시보드 (Settlement Report)

## 설치 및 실행

```bash
pip install -r requirements.txt
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

## 테스트 방법

`samples/` 폴더의 파일 5개를 한 번에 업로드하면 쇼핑몰별 컬럼 인식이 어떻게 되는지 확인할 수 있습니다.

| 파일 | 형식 특징 |
|---|---|
| coupang_settlement.csv | 한글 표준 컬럼 (매출액, 판매수수료, 광고비) |
| naver_smartstore.csv | 첫 줄 제목행, 결제금액·서비스이용료 |
| 11st_settlement.csv | 판매금액·택배비, 마지막 합계행 자동 제외 |
| gmarket_settlement.csv | 주문금액·마켓수수료, 쿠폰비는 기타비용 |
| global_shopify_export.csv | 영문 헤더 (Gross Sales, Fees, Cost), 금액 단위는 원화로 환산된 예시 |

샘플은 실제 정산서가 아닌 가상 데이터입니다. 실제 파일의 컬럼명이 인식되지 않으면 화면의 드롭다운에서 직접 지정하세요.
