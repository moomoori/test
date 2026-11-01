# Settlement Report (정산 대시보드)

## 설치 및 실행

```bash
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

접속: http://서버IP:8501

## 파일 구성

- app.py: 메인 앱 (단일 파일로 실행 가능)
- .streamlit/config.toml: 다크 모드 테마 설정 (app.py와 같은 위치에 두세요)
- requirements.txt: 의존성 목록

## 컬럼 인식

컬럼명이 다르면 app.py의 COLUMN_ALIASES에 표기를 추가하세요. 업로드 후 '컬럼 자동 인식 결과 확인'에서 매핑을 볼 수 있습니다.
