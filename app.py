import io
import re

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 페이지 & 스타일 (라이트 · 미니멀 · 단일 포인트 컬러)
# ----------------------------------------------------------------------
st.set_page_config(page_title="Settlement Dashboard", page_icon="📊", layout="wide")

ACCENT = "#3182F6"
ACCENT_2 = "#A0C4FF"
POSITIVE = "#12B76A"
NEGATIVE = "#F04438"

st.markdown(
    f"""
    <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"], .stApp {{
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI',
                     'Noto Sans KR', sans-serif;
    }}
    .stApp {{ background: #F7F8FA; }}
    .block-container {{ padding-top: 2.5rem; padding-bottom: 4rem; max-width: 1200px; }}
    header[data-testid="stHeader"] {{ background: transparent; }}

    /* 헤더 */
    .page-title {{ font-size: 1.9rem; font-weight: 800; color: #191F28; margin: 0; letter-spacing: -0.02em; }}
    .page-sub {{ color: #8B95A1; margin-top: 6px; font-size: 0.95rem; }}

    /* 카드 */
    .card {{
        background: #FFFFFF; border: 1px solid #ECEEF2; border-radius: 16px;
        padding: 22px 24px; margin-bottom: 16px;
    }}
    .card-title {{ font-size: 0.95rem; font-weight: 700; color: #191F28; margin-bottom: 14px; }}

    /* 지표 카드 */
    .kpi {{
        background: #FFFFFF; border: 1px solid #ECEEF2; border-radius: 16px; padding: 20px 22px;
        height: 100%;
    }}
    .kpi-label {{ font-size: 0.82rem; color: #8B95A1; font-weight: 600; }}
    .kpi-value {{ font-size: 1.6rem; font-weight: 800; color: #191F28; margin-top: 8px;
                  letter-spacing: -0.02em; }}
    .kpi-delta {{ font-size: 0.82rem; margin-top: 6px; font-weight: 600; }}
    .kpi-accent {{ background: {ACCENT}; border-color: {ACCENT}; }}
    .kpi-accent .kpi-label, .kpi-accent .kpi-value, .kpi-accent .kpi-delta {{ color: #FFFFFF; }}

    /* 업로더 */
    div[data-testid="stFileUploader"] section {{
        background: #FFFFFF; border: 1.5px dashed #D1D6DD; border-radius: 16px;
    }}
    div[data-testid="stFileUploader"] section:hover {{ border-color: {ACCENT}; }}

    /* 탭 */
    button[data-baseweb="tab"] {{ font-weight: 600; }}

    /* 버튼 */
    .stButton > button, .stDownloadButton > button {{
        border-radius: 10px; font-weight: 700; border: none; padding: 0.55rem 1.1rem;
    }}
    .stDownloadButton > button[kind="primary"], .stButton > button[kind="primary"] {{
        background: {ACCENT}; color: #FFFFFF;
    }}

    /* 사이드바 */
    section[data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #ECEEF2; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 컬럼 매핑 사전
# ----------------------------------------------------------------------
COLUMN_ALIASES = {
    "platform": ["쇼핑몰", "플랫폼", "채널", "판매처", "몰이름", "platform", "mall", "channel"],
    "date": ["정산일", "정산기준일", "주문일", "주문일자", "결제일", "날짜", "date", "orderdate"],
    "product": ["상품명", "상품", "품목명", "제품명", "품목", "product", "productname", "item"],
    "quantity": ["수량", "주문수량", "판매수량", "qty", "quantity"],
    "sales": ["매출액", "총매출", "매출", "판매금액", "결제금액", "상품금액", "판매가",
              "정산대상금액", "sales", "revenue", "amount"],
    "commission": ["수수료", "판매수수료", "서비스이용료", "중개수수료", "수수료합계",
                   "판매자수수료", "commission", "fee"],
    "shipping": ["배송비", "배송비용", "택배비", "배송료", "shipping", "delivery", "deliveryfee"],
    "other": ["기타비용", "기타수수료", "광고비", "포장비", "기타", "other", "adcost"],
    "cost": ["원가", "매입원가", "매입가", "공급가", "cost", "purchasecost"],
}
EXCLUDE_KEYWORDS = ["합계", "소계", "총계", "total", "sum"]
DETAIL_COLS = ["플랫폼", "정산일", "상품명", "수량", "매출액", "수수료", "배송비",
               "기타비용", "원가", "순이익", "마진율(%)", "원본파일"]


# ----------------------------------------------------------------------
# 샘플 데이터 생성 (플랫폼마다 컬럼명이 다른 3종)
# ----------------------------------------------------------------------
class MemoryFile:
    """업로드 파일과 같은 인터페이스(name, getvalue)를 흉내 내는 래퍼."""
    def __init__(self, name: str, data: bytes):
        self.name = name
        self._data = data

    def getvalue(self) -> bytes:
        return self._data


@st.cache_data(show_spinner=False)
def make_sample_files(seed: int = 42):
    rng = np.random.default_rng(seed)
    products = ["무선 이어폰", "텀블러 500ml", "요가매트", "캠핑 의자", "가습기",
                "에코백", "보온병", "무드등", "마우스패드", "비타민C 세트"]
    prices = dict(zip(products, [39000, 22000, 28000, 54000, 45000,
                                 12000, 31000, 19000, 9000, 26000]))
    cost_ratio = dict(zip(products, [0.55, 0.45, 0.50, 0.58, 0.52,
                                     0.40, 0.48, 0.42, 0.35, 0.50]))

    def rows(n, mall):
        out = []
        for _ in range(n):
            p = rng.choice(products)
            qty = int(rng.integers(1, 5))
            sales = prices[p] * qty
            date = f"2026-{rng.integers(9, 11):02d}-{rng.integers(1, 29):02d}"
            out.append(dict(product=p, qty=qty, sales=sales,
                            cost=round(sales * cost_ratio[p]), date=date, mall=mall))
        return out

    # 쿠팡: 매출액 / 판매수수료 / 배송비 / 광고비
    cp = pd.DataFrame(rows(40, "쿠팡"))
    coupang = pd.DataFrame({
        "정산일": cp["date"], "상품명": cp["product"], "주문수량": cp["qty"],
        "매출액": cp["sales"], "판매수수료": (cp["sales"] * 0.108).round().astype(int),
        "배송비": rng.choice([0, 3000], size=len(cp)),
        "광고비": rng.integers(0, 2000, size=len(cp)),
    })

    # 네이버 스마트스토어: 결제금액 / 서비스이용료 / 배송비 (제목행 1줄 포함)
    nv = pd.DataFrame(rows(35, "네이버"))
    naver = pd.DataFrame({
        "주문일자": nv["date"], "상품명": nv["product"], "수량": nv["qty"],
        "결제금액": nv["sales"], "서비스이용료": (nv["sales"] * 0.035).round().astype(int),
        "배송비": rng.choice([0, 3500], size=len(nv)),
    })
    naver_title = pd.DataFrame([["스마트스토어 정산내역"] + [""] * (len(naver.columns) - 1)],
                               columns=naver.columns)
    naver = pd.concat([naver_title, naver], ignore_index=True)

    # 11번가: 판매금액 / 수수료 / 택배비 + 합계행 (자동 제외 확인용)
    st_rows = pd.DataFrame(rows(30, "11번가"))
    elev = pd.DataFrame({
        "판매일": st_rows["date"], "품목명": st_rows["product"], "판매수량": st_rows["qty"],
        "판매금액": st_rows["sales"], "수수료": (st_rows["sales"] * 0.12).round().astype(int),
        "택배비": rng.choice([0, 2500], size=len(st_rows)),
    })
    total = pd.DataFrame([{
        "판매일": "합계", "품목명": "", "판매수량": elev["판매수량"].sum(),
        "판매금액": elev["판매금액"].sum(), "수수료": elev["수수료"].sum(),
        "택배비": elev["택배비"].sum(),
    }])
    elev = pd.concat([elev, total], ignore_index=True)

    files = []
    buf = io.BytesIO()
    coupang.to_excel(buf, index=False)
    files.append(MemoryFile("sample_쿠팡_정산.xlsx", buf.getvalue()))

    files.append(MemoryFile("sample_네이버_정산.csv",
                            naver.to_csv(index=False).encode("utf-8-sig")))

    buf = io.BytesIO()
    elev.to_excel(buf, index=False)
    files.append(MemoryFile("sample_11번가_정산.xlsx", buf.getvalue()))
    return files


# ----------------------------------------------------------------------
# 파싱 / 매핑 / 계산
# ----------------------------------------------------------------------
def normalize(text) -> str:
    return re.sub(r"[\s_\-()/]", "", str(text)).lower()


def map_columns(df: pd.DataFrame) -> dict:
    norm_cols = {normalize(c): c for c in df.columns}
    mapping = {}
    for key, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if normalize(alias) in norm_cols:
                mapping[key] = norm_cols[normalize(alias)]
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
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(r"[,\s원₩$]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def process_files(files, default_cost_rate: float):
    frames, errors = [], []
    for f in files:
        try:
            df, m = load_settlement(f.getvalue(), f.name)
        except Exception as e:
            errors.append((f.name, str(e)))
            continue

        stem = f.name.rsplit(".", 1)[0]
        out = pd.DataFrame(index=df.index)
        out["플랫폼"] = df[m["platform"]].fillna(stem).astype(str) if "platform" in m else stem
        out["정산일"] = df[m["date"]].fillna("") if "date" in m else ""
        out["상품명"] = df[m["product"]].fillna("") if "product" in m else ""
        out["수량"] = to_number(df[m["quantity"]]) if "quantity" in m else 0.0
        out["매출액"] = to_number(df[m["sales"]])
        out["수수료"] = to_number(df[m["commission"]]) if "commission" in m else 0.0
        out["배송비"] = to_number(df[m["shipping"]]) if "shipping" in m else 0.0
        out["기타비용"] = to_number(df[m["other"]]) if "other" in m else 0.0
        out["원가"] = to_number(df[m["cost"]]) if "cost" in m else out["매출액"] * default_cost_rate / 100

        text = out[["플랫폼", "정산일", "상품명"]].astype(str).agg(" ".join, axis=1).str.lower()
        is_total = text.apply(lambda t: any(k in t for k in EXCLUDE_KEYWORDS))
        out = out[~is_total]
        out = out[(out["매출액"] != 0) | (out["수수료"] != 0)]

        out["순이익"] = out["매출액"] - out["수수료"] - out["배송비"] - out["기타비용"] - out["원가"]
        out["마진율(%)"] = np.where(out["매출액"] > 0, out["순이익"] / out["매출액"] * 100, 0).round(2)
        out["원본파일"] = f.name
        frames.append(out)

    detail = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=DETAIL_COLS)
    return detail, errors


def make_summary(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame()
    g = detail.groupby("플랫폼", as_index=False).agg(
        주문건수=("매출액", "size"), 매출액=("매출액", "sum"), 수수료=("수수료", "sum"),
        배송비=("배송비", "sum"), 기타비용=("기타비용", "sum"), 원가=("원가", "sum"),
        순이익=("순이익", "sum"),
    )
    g["마진율(%)"] = np.where(g["매출액"] > 0, g["순이익"] / g["매출액"] * 100, 0).round(2)
    return g.sort_values("매출액", ascending=False).reset_index(drop=True)


def make_totals(detail: pd.DataFrame) -> dict:
    sales = float(detail["매출액"].sum())
    cost = float(detail[["수수료", "배송비", "기타비용", "원가"]].sum().sum())
    profit = float(detail["순이익"].sum())
    return {
        "주문건수": int(len(detail)), "총매출": sales, "총비용": cost, "순이익": profit,
        "마진율(%)": round(profit / sales * 100, 2) if sales > 0 else 0.0,
    }


# ----------------------------------------------------------------------
# 엑셀 생성
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor=ACCENT.lstrip("#"))
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    for col_idx, cell0 in enumerate(ws[1], start=1):
        header = str(cell0.value or "")
        letter = get_column_letter(col_idx)
        values = [str(ws.cell(r, col_idx).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[letter].width = min(max(len(v) for v in values) * 1.6 + 4, 40)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row, col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.00" if "율" in header else "#,##0"


def build_excel(detail, summary, totals) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(list(totals.items()), columns=["항목", "값"]).to_excel(
            writer, sheet_name="종합요약", index=False)
        summary.to_excel(writer, sheet_name="플랫폼별요약", index=False)
        detail.to_excel(writer, sheet_name="마진계산내역", index=False)
        for ws in writer.book.worksheets:
            style_sheet(ws)
    return buf.getvalue()


# ----------------------------------------------------------------------
# UI 컴포넌트
# ----------------------------------------------------------------------
def kpi(label: str, value: str, sub: str = "", accent: bool = False):
    cls = "kpi kpi-accent" if accent else "kpi"
    sub_html = f'<div class="kpi-delta">{sub}</div>' if sub else ""
    return f'<div class="{cls}"><div class="kpi-label">{label}</div>' \
           f'<div class="kpi-value">{value}</div>{sub_html}</div>'


def card_open(title: str):
    st.markdown(f'<div class="card"><div class="card-title">{title}</div>', unsafe_allow_html=True)


def card_close():
    st.markdown("</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ 설정")
    cost_rate = st.slider("기본 원가율 (%)", 0, 100, 0, 1,
                          help="파일에 원가 컬럼이 없을 때 매출액 대비로 적용됩니다.")
    st.caption("원가 컬럼이 있는 파일은 해당 값을 우선 사용합니다.")
    st.divider()
    use_demo = st.toggle("샘플 데이터로 체험하기", value=False,
                         help="쿠팡·네이버·11번가 형식의 가상 정산 데이터 3개를 불러옵니다.")
    st.caption("실제 정산 파일 업로드 시 이 옵션을 끄세요.")

# ----------------------------------------------------------------------
# 헤더
# ----------------------------------------------------------------------
st.markdown(
    """
    <h1 class="page-title">정산 대시보드</h1>
    <p class="page-sub">쇼핑몰별 정산 파일을 올리면 매출과 순이익을 한눈에 정리해 드려요.</p>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "정산 파일 업로드",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if use_demo:
    source_files = make_sample_files()
    st.info("샘플 데이터 3종(쿠팡·네이버·11번가)으로 분석 중입니다.")
elif uploaded:
    source_files = uploaded
else:
    st.markdown(
        """
        <div class="card" style="text-align:center; padding:48px 24px;">
            <div style="font-size:2.2rem;">📄</div>
            <div style="font-weight:700; font-size:1.05rem; margin-top:8px; color:#191F28;">
                정산 파일을 끌어다 놓으세요</div>
            <div style="color:#8B95A1; margin-top:6px; font-size:0.9rem;">
                CSV, XLSX 여러 개를 한 번에 올릴 수 있어요.<br>
                플랫폼마다 컬럼명이 달라도 자동으로 매핑합니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

with st.spinner("정산 데이터를 분석하는 중..."):
    detail, errors = process_files(source_files, cost_rate)

for name, msg in errors:
    st.error(f"**{name}** 처리 실패: {msg}")

if detail.empty:
    st.warning("처리 가능한 데이터가 없습니다. 매출액 컬럼이 포함되어 있는지 확인해 주세요.")
    st.stop()

summary = make_summary(detail)
totals = make_totals(detail)

# ----------------------------------------------------------------------
# 핵심 지표
# ----------------------------------------------------------------------
cols = st.columns(4)
cols[0].markdown(kpi("총 매출", f"{totals['총매출']:,.0f}원",
                     f"{totals['주문건수']:,}건 주문"), unsafe_allow_html=True)
cols[1].markdown(kpi("총 비용", f"{totals['총비용']:,.0f}원",
                     "수수료·배송·기타·원가"), unsafe_allow_html=True)
cols[2].markdown(kpi("순수익", f"{totals['순이익']:,.0f}원",
                     "매출 − 전체 비용", accent=True), unsafe_allow_html=True)
margin_color = POSITIVE if totals["마진율(%)"] >= 0 else NEGATIVE
cols[3].markdown(kpi("마진율", f"{totals['마진율(%)']:.2f}%",
                     f'<span style="color:{margin_color}">순이익 ÷ 매출</span>'),
                 unsafe_allow_html=True)

st.write("")

# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
left, right = st.columns([3, 2])
with left:
    card_open("플랫폼별 매출 · 순이익")
    st.bar_chart(summary.set_index("플랫폼")[["매출액", "순이익"]],
                 color=[ACCENT, ACCENT_2], height=300)
    card_close()
with right:
    card_open("플랫폼별 마진율")
    st.bar_chart(summary.set_index("플랫폼")[["마진율(%)"]], color=[POSITIVE], height=300)
    card_close()

# ----------------------------------------------------------------------
# 상세 테이블
# ----------------------------------------------------------------------
money_cfg = {c: st.column_config.NumberColumn(c, format="%d") for c in
             ["매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]}
rate_cfg = {"마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f%%")}

tab1, tab2, tab3 = st.tabs(["플랫폼별 요약", "종합 정산 내역", "마진 계산 내역"])

with tab1:
    st.dataframe(summary, hide_index=True, column_config={**money_cfg, **rate_cfg})

with tab2:
    st.caption(f"총 {len(detail):,}건 · 파일 {len(source_files)}개")
    st.dataframe(detail[DETAIL_COLS], hide_index=True, column_config={**money_cfg, **rate_cfg})

with tab3:
    st.caption("순이익 = 매출액 − 수수료 − 배송비 − 기타비용 − 원가  /  마진율 = 순이익 ÷ 매출액 × 100")
    st.dataframe(
        detail[["플랫폼", "상품명", "매출액", "수수료", "배송비", "기타비용", "원가",
                "순이익", "마진율(%)"]],
        hide_index=True, column_config={**money_cfg, **rate_cfg},
    )

# ----------------------------------------------------------------------
# 다운로드
# ----------------------------------------------------------------------
st.write("")
card_open("보고서 다운로드")
st.caption("종합요약 · 플랫폼별요약 · 마진계산내역 3개 시트로 구성된 엑셀 파일입니다.")
st.download_button(
    label="엑셀 보고서 다운로드 (.xlsx)",
    data=build_excel(detail, summary, totals),
    file_name=f"정산보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
card_close()
