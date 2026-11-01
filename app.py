import io
import re

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 페이지 설정 및 스타일
# ----------------------------------------------------------------------
st.set_page_config(page_title="셀러 정산 대시보드", page_icon="📊", layout="wide")

CUSTOM_CSS = """
<style>
    .block-container {padding-top: 2rem; padding-bottom: 3rem;}
    .hero {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 55%, #EC4899 100%);
        padding: 28px 32px; border-radius: 18px; color: white; margin-bottom: 24px;
        box-shadow: 0 10px 30px rgba(79, 70, 229, 0.25);
    }
    .hero h1 {margin: 0; font-size: 2rem; color: white;}
    .hero p {margin: 6px 0 0 0; opacity: 0.9;}
    div[data-testid="stMetric"] {
        background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 14px;
        padding: 16px 18px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    div[data-testid="stMetricLabel"] p {font-weight: 600; color: #6B7280;}
    .guide-card {
        background: #F9FAFB; border: 1px dashed #C7D2FE; border-radius: 14px;
        padding: 20px 24px; color: #374151;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 표준 컬럼 매핑 사전 (플랫폼마다 다른 컬럼명을 표준 이름으로 통일)
# ----------------------------------------------------------------------
COLUMN_ALIASES = {
    "platform": ["쇼핑몰", "플랫폼", "채널", "판매처", "몰이름", "platform", "mall", "channel"],
    "date":     ["정산일", "정산기준일", "주문일", "주문일자", "결제일", "날짜", "date", "orderdate"],
    "product":  ["상품명", "상품", "품목명", "제품명", "품목", "product", "productname", "item"],
    "quantity": ["수량", "주문수량", "판매수량", "qty", "quantity"],
    "sales":    ["매출액", "총매출", "매출", "판매금액", "결제금액", "상품금액", "판매가",
                 "정산대상금액", "sales", "revenue", "amount"],
    "commission": ["수수료", "판매수수료", "서비스이용료", "중개수수료", "수수료합계",
                   "판매자수수료", "commission", "fee"],
    "shipping": ["배송비", "배송비용", "택배비", "배송료", "shipping", "delivery", "deliveryfee"],
    "other":    ["기타비용", "기타수수료", "광고비", "포장비", "기타", "other", "adcost"],
    "cost":     ["원가", "매입원가", "매입가", "공급가", "cost", "purchasecost"],
}

STANDARD_LABEL = {
    "platform": "플랫폼", "date": "정산일", "product": "상품명", "quantity": "수량",
    "sales": "매출액", "commission": "수수료", "shipping": "배송비",
    "other": "기타비용", "cost": "원가",
}

EXCLUDE_KEYWORDS = ["합계", "소계", "총계", "total", "sum"]


def normalize(text) -> str:
    return re.sub(r"[\s_\-()/]", "", str(text)).lower()


def map_columns(df: pd.DataFrame) -> dict:
    """데이터프레임 컬럼을 표준 키에 매핑. {표준키: 실제컬럼명}"""
    norm_cols = {normalize(c): c for c in df.columns}
    mapping = {}
    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            a = normalize(alias)
            if a in norm_cols:
                mapping[key] = norm_cols[a]
                break
    return mapping


def read_raw(raw: bytes, name: str, header: int) -> pd.DataFrame:
    if name.lower().endswith(".csv"):
        last_err = None
        for enc in ("utf-8-sig", "cp949", "euc-kr"):
            try:
                return pd.read_csv(io.BytesIO(raw), header=header, encoding=enc,
                                   dtype=str, on_bad_lines="skip")
            except UnicodeDecodeError as e:
                last_err = e
        raise last_err
    return pd.read_excel(io.BytesIO(raw), header=header, dtype=str)


def load_settlement(raw: bytes, name: str):
    """제목 행이 있는 정산서도 대응하기 위해 0~9행을 헤더 후보로 탐색."""
    for header in range(10):
        try:
            df = read_raw(raw, name, header)
        except Exception:
            continue
        mapping = map_columns(df)
        if "sales" in mapping:
            return df, mapping
    raise ValueError("'매출액'에 해당하는 컬럼을 찾지 못했습니다.")


def to_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)      # (1,000) -> -1000
    s = s.str.replace(r"[,\s원₩$]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def process_files(uploaded_files, default_cost_rate: float):
    frames, errors = [], []

    for f in uploaded_files:
        try:
            raw = f.getvalue()
            df, m = load_settlement(raw, f.name)
        except Exception as e:
            errors.append((f.name, str(e)))
            continue

        stem = f.name.rsplit(".", 1)[0]
        out = pd.DataFrame(index=df.index)

        if "platform" in m:
            out["플랫폼"] = df[m["platform"]].fillna(stem).astype(str)
        else:
            out["플랫폼"] = stem
        out["정산일"] = df[m["date"]].fillna("") if "date" in m else ""
        out["상품명"] = df[m["product"]].fillna("") if "product" in m else ""
        out["수량"] = to_number(df[m["quantity"]]) if "quantity" in m else 0.0
        out["매출액"] = to_number(df[m["sales"]])
        out["수수료"] = to_number(df[m["commission"]]) if "commission" in m else 0.0
        out["배송비"] = to_number(df[m["shipping"]]) if "shipping" in m else 0.0
        out["기타비용"] = to_number(df[m["other"]]) if "other" in m else 0.0
        if "cost" in m:
            out["원가"] = to_number(df[m["cost"]])
        else:
            out["원가"] = out["매출액"] * default_cost_rate / 100

        # 합계/소계 행 제외 (중복 집계 방지)
        text_cols = out[["플랫폼", "정산일", "상품명"]].astype(str).agg(" ".join, axis=1)
        is_total = text_cols.str.lower().apply(lambda t: any(k in t for k in EXCLUDE_KEYWORDS))
        out = out[~is_total]
        out = out[(out["매출액"] != 0) | (out["수수료"] != 0)]  # 빈 행 제거

        out["순이익"] = (out["매출액"] - out["수수료"] - out["배송비"]
                         - out["기타비용"] - out["원가"])
        out["마진율(%)"] = np.where(out["매출액"] > 0,
                                    out["순이익"] / out["매출액"] * 100, 0).round(2)
        out["원본파일"] = f.name
        frames.append(out)

    detail = (pd.concat(frames, ignore_index=True) if frames
              else pd.DataFrame(columns=["플랫폼", "정산일", "상품명", "수량", "매출액",
                                         "수수료", "배송비", "기타비용", "원가",
                                         "순이익", "마진율(%)", "원본파일"]))
    return detail, errors


def make_summary(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame()
    g = detail.groupby("플랫폼", as_index=False).agg(
        주문건수=("매출액", "size"),
        매출액=("매출액", "sum"),
        수수료=("수수료", "sum"),
        배송비=("배송비", "sum"),
        기타비용=("기타비용", "sum"),
        원가=("원가", "sum"),
        순이익=("순이익", "sum"),
    )
    g["마진율(%)"] = np.where(g["매출액"] > 0, g["순이익"] / g["매출액"] * 100, 0).round(2)
    return g.sort_values("매출액", ascending=False).reset_index(drop=True)


def make_totals(detail: pd.DataFrame) -> dict:
    sales = float(detail["매출액"].sum())
    cost_total = float(detail[["수수료", "배송비", "기타비용", "원가"]].sum().sum())
    profit = float(detail["순이익"].sum())
    return {
        "주문건수": int(len(detail)),
        "총매출": sales,
        "총비용": cost_total,
        "순이익": profit,
        "마진율(%)": round(profit / sales * 100, 2) if sales > 0 else 0.0,
    }


# ----------------------------------------------------------------------
# 엑셀 생성
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="4F46E5")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws, percent_label_col: bool = False):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

    headers = [c.value for c in ws[1]]
    for col_idx, header in enumerate(headers, start=1):
        letter = get_column_letter(col_idx)
        values = [str(ws.cell(r, col_idx).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[letter].width = min(max(len(v) for v in values) * 1.6 + 4, 40)

        is_rate = header and "율" in str(header)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row, col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.00" if is_rate else "#,##0"

    if percent_label_col:
        # 종합요약 시트: 항목별 서식 분리 (마진율만 소수점 2자리)
        for row in range(2, ws.max_row + 1):
            label = str(ws.cell(row, 1).value or "")
            cell = ws.cell(row, 2)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.00" if "율" in label else "#,##0"


def build_excel(detail: pd.DataFrame, summary: pd.DataFrame, totals: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(list(totals.items()), columns=["항목", "값"]).to_excel(
            writer, sheet_name="종합요약", index=False)
        summary.to_excel(writer, sheet_name="플랫폼별요약", index=False)
        detail.to_excel(writer, sheet_name="마진계산내역", index=False)

        for ws in writer.book.worksheets:
            style_sheet(ws, percent_label_col=(ws.title == "종합요약"))
    return buf.getvalue()


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ 설정")
    cost_rate = st.slider(
        "기본 원가율 (%)", min_value=0, max_value=100, value=0, step=1,
        help="파일에 '원가/매입가' 컬럼이 없을 때 매출액 대비 원가율로 적용됩니다.",
    )
    st.caption("원가 컬럼이 있는 파일은 해당 값을 우선 사용합니다.")
    st.divider()
    st.markdown("**지원 플랫폼 예시**")
    st.caption("쿠팡 윙, 네이버 스마트스토어, 11번가, G마켓, 카카오 선물하기 등 "
               "컬럼명이 유사하면 자동 인식합니다.")

# ----------------------------------------------------------------------
# 메인 화면
# ----------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <h1>📊 셀러 정산 대시보드</h1>
        <p>여러 쇼핑몰의 정산 파일을 한 번에 합쳐 매출과 순이익을 자동으로 계산합니다.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "📂 정산 파일을 여기에 드래그앤드롭 하세요 (CSV / XLSX, 여러 개 가능)",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
)

if not uploaded:
    st.markdown(
        """
        <div class="guide-card">
            <h4 style="margin-top:0">🚀 사용 방법</h4>
            <ol>
                <li>쿠팡, 네이버 스마트스토어 등에서 받은 정산 파일(CSV/XLSX)을 업로드합니다.</li>
                <li>플랫폼마다 달라도 <b>매출액 · 수수료 · 배송비 · 기타비용</b> 컬럼을 자동으로 찾아 합칩니다.</li>
                <li>순이익과 마진율을 확인하고, 엑셀 보고서를 다운로드하세요.</li>
            </ol>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

with st.spinner("정산 데이터를 분석하는 중입니다..."):
    detail, errors = process_files(uploaded, cost_rate)

for name, msg in errors:
    st.error(f"❌ **{name}** 처리 실패: {msg}")

if detail.empty:
    st.warning("처리 가능한 데이터가 없습니다. 파일의 컬럼명(매출액 등)을 확인해 주세요.")
    st.stop()

summary = make_summary(detail)
totals = make_totals(detail)

# ----------------------------------------------------------------------
# 핵심 지표
# ----------------------------------------------------------------------
st.subheader("📌 핵심 지표")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("총 매출", f"{totals['총매출']:,.0f}원")
c2.metric("총 비용", f"{totals['총비용']:,.0f}원")
c3.metric("순수익", f"{totals['순이익']:,.0f}원")
c4.metric("마진율", f"{totals['마진율(%)']:.2f}%")
c5.metric("주문 건수", f"{totals['주문건수']:,}건")

st.divider()

# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
left, right = st.columns([3, 2])
with left:
    st.markdown("**플랫폼별 매출 vs 순이익**")
    chart_df = summary.set_index("플랫폼")[["매출액", "순이익"]]
    st.bar_chart(chart_df, color=["#4F46E5", "#EC4899"])
with right:
    st.markdown("**플랫폼별 마진율(%)**")
    st.bar_chart(summary.set_index("플랫폼")[["마진율(%)"]], color=["#10B981"])

# ----------------------------------------------------------------------
# 탭: 상세 데이터
# ----------------------------------------------------------------------
money_cfg = {c: st.column_config.NumberColumn(c, format="%d") for c in
             ["매출액", "수수료", "배송비", "기타비용", "원가", "순이익", "총매출"]}
rate_cfg = {"마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f%%")}

tab1, tab2, tab3 = st.tabs(["🏪 플랫폼별 요약", "🧾 종합 정산 내역", "🧮 마진 계산 내역"])

with tab1:
    st.dataframe(summary, use_container_width=True, hide_index=True,
                 column_config={**money_cfg, **rate_cfg})

with tab2:
    st.caption(f"총 {len(detail):,}건 / 업로드 파일 {len(uploaded)}개")
    show_cols = ["플랫폼", "정산일", "상품명", "수량", "매출액", "수수료",
                 "배송비", "기타비용", "원가", "순이익", "마진율(%)", "원본파일"]
    st.dataframe(detail[show_cols], use_container_width=True, hide_index=True,
                 column_config={**money_cfg, **rate_cfg})

with tab3:
    st.markdown(
        "**계산식**  \n"
        "- 순이익 = 매출액 − 수수료 − 배송비 − 기타비용 − 원가  \n"
        "- 마진율(%) = 순이익 ÷ 매출액 × 100"
    )
    st.dataframe(
        detail[["플랫폼", "상품명", "매출액", "수수료", "배송비", "기타비용", "원가",
                "순이익", "마진율(%)"]],
        use_container_width=True, hide_index=True,
        column_config={**money_cfg, **rate_cfg},
    )

# ----------------------------------------------------------------------
# 다운로드
# ----------------------------------------------------------------------
st.divider()
st.subheader("⬇️ 보고서 다운로드")
excel_bytes = build_excel(detail, summary, totals)
st.download_button(
    label="📥 종합 정산 보고서 (.xlsx) 다운로드",
    data=excel_bytes,
    file_name=f"정산보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
st.caption("시트 구성: 종합요약 / 플랫폼별요약 / 마진계산내역")
