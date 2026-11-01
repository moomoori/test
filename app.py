import io
import re
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 페이지 & 스타일 (다크 모드 기본 · 얇은 선 · 높은 대비)
# ----------------------------------------------------------------------
st.set_page_config(page_title="Settlement Report", page_icon="▫️", layout="wide",
                   initial_sidebar_state="expanded")

BG = "#0E0F12"
SURFACE = "#16181D"
LINE = "#3A3F4B"
TEXT = "#F5F5F5"
TEXT_SUB = "#B8BDC7"
ACCENT_TEXT = "#0E0F12"
SERIES_2 = "#8A8F9A"

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&display=swap');
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"], .stApp {{
        font-family: 'Pretendard', 'Noto Sans KR', -apple-system, 'Segoe UI', sans-serif;
    }}
    .stApp {{ background: {BG}; color: {TEXT}; }}
    .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3 {{ color: {TEXT}; }}
    .block-container {{ padding-top: 2rem; padding-bottom: 4rem; max-width: 1200px; }}
    header[data-testid="stHeader"] {{ background: transparent; }}

    .topbar {{
        display: flex; justify-content: space-between; align-items: center;
        border-top: 1px solid {LINE}; border-bottom: 1px solid {LINE};
        padding: 12px 4px; font-size: 0.8rem; letter-spacing: 0.12em;
        text-transform: uppercase; color: {TEXT_SUB};
    }}
    .brand {{ font-family: 'Playfair Display', 'Nanum Myeongjo', serif; font-size: 1.1rem;
              letter-spacing: 0.2em; font-weight: 600; color: {TEXT}; }}
    .hero {{ text-align: center; padding: 48px 0 40px 0; border-bottom: 1px solid {LINE};
             margin-bottom: 36px; }}
    .hero h1 {{ font-family: 'Playfair Display', 'Nanum Myeongjo', serif; font-weight: 400;
                font-size: 2.6rem; letter-spacing: 0.08em; text-transform: uppercase;
                margin: 0; color: {TEXT}; }}
    .hero p {{ color: {TEXT_SUB}; margin: 14px 0 0 0; font-size: 1rem; }}

    .section-label {{
        font-size: 0.78rem; letter-spacing: 0.18em; text-transform: uppercase;
        color: {TEXT}; font-weight: 700; margin: 36px 0 14px 0;
        padding-bottom: 8px; border-bottom: 1px solid {LINE};
    }}
    .kpi {{
        background: {SURFACE}; border: 1px solid {LINE}; padding: 22px;
        height: 100%; min-height: 136px;
    }}
    .kpi-label {{ font-size: 0.74rem; letter-spacing: 0.14em; text-transform: uppercase;
                  color: {TEXT_SUB}; font-weight: 600; }}
    .kpi-value {{ font-family: 'Playfair Display', 'Nanum Myeongjo', serif; font-size: 2rem;
                  color: {TEXT}; margin-top: 12px; line-height: 1.1; }}
    .kpi-sub {{ font-size: 0.84rem; color: {TEXT_SUB}; margin-top: 10px; }}
    .kpi-accent {{ background: {TEXT}; border-color: {TEXT}; }}
    .kpi-accent .kpi-label, .kpi-accent .kpi-sub {{ color: #3A3F4B; }}
    .kpi-accent .kpi-value {{ color: {ACCENT_TEXT}; }}

    .card-title {{ font-size: 0.8rem; letter-spacing: 0.14em; text-transform: uppercase;
                   font-weight: 700; color: {TEXT}; margin-bottom: 12px; }}
    .card {{ border: 1px solid {LINE}; padding: 22px 24px; background: {SURFACE}; }}

    div[data-testid="stFileUploader"] section {{
        background: {SURFACE}; border: 1px dashed {TEXT_SUB}; border-radius: 0;
    }}
    div[data-testid="stFileUploader"] section:hover {{ border-color: {TEXT}; }}
    div[data-testid="stFileUploader"] button {{
        border-radius: 0; border: 1px solid {TEXT}; background: transparent; color: {TEXT};
    }}

    button[data-baseweb="tab"] {{ font-weight: 600; letter-spacing: 0.04em;
                                  color: {TEXT_SUB} !important; }}
    button[data-baseweb="tab"][aria-selected="true"] {{ color: {TEXT} !important; }}
    div[data-baseweb="tab-highlight"] {{ background-color: {TEXT} !important; }}
    div[data-baseweb="tab-border"] {{ background-color: {LINE} !important; }}

    [data-testid="stDownloadButton"] button, .stButton > button {{
        border-radius: 0; font-weight: 700; letter-spacing: 0.06em;
        background: {TEXT}; color: {BG} !important; border: 1px solid {TEXT};
        padding: 0.7rem 1.4rem;
    }}
    [data-testid="stDownloadButton"] button:hover {{ background: #D6D6D6; border-color: #D6D6D6; }}

    section[data-testid="stSidebar"] {{ background: {SURFACE}; border-right: 1px solid {LINE}; }}
    div[data-testid="stDataFrame"] {{ border: 1px solid {LINE}; }}
    div[data-testid="stAlert"] {{ border-radius: 0; border: 1px solid {LINE}; }}
    .empty {{ border: 1px solid {LINE}; padding: 64px 24px; text-align: center;
              background: {SURFACE}; }}
    .empty-eyebrow {{ font-size: 0.75rem; letter-spacing: 0.18em; text-transform: uppercase;
                      color: {TEXT_SUB}; }}
    .empty-title {{ font-family: 'Playfair Display', 'Nanum Myeongjo', serif; font-size: 1.6rem;
                    margin-top: 10px; color: {TEXT}; }}
    .empty-desc {{ color: {TEXT_SUB}; margin-top: 10px; font-size: 0.95rem; line-height: 1.7; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 컬럼 매핑 사전
# ----------------------------------------------------------------------
COLUMN_ALIASES = {
    "platform": ["쇼핑몰", "플랫폼", "채널", "판매처", "몰이름", "platform", "mall", "channel"],
    "date": ["정산일", "정산기준일", "주문일", "주문일자", "결제일", "날짜", "판매일",
             "date", "orderdate"],
    "product": ["상품명", "상품", "품목명", "제품명", "품목", "product", "productname", "item"],
    "quantity": ["수량", "주문수량", "판매수량", "qty", "quantity"],
    "sales": ["매출액", "총매출", "매출", "매출금액", "판매금액", "판매가", "판매가격",
              "판매액", "결제금액", "결제가", "결제액", "총결제금액", "주문금액", "구매금액",
              "상품금액", "정산대상금액", "금액", "sales", "revenue", "amount"],
    "commission": ["수수료", "판매수수료", "서비스이용료", "중개수수료", "수수료합계",
                   "판매자수수료", "commission", "fee"],
    "shipping": ["배송비", "배송비용", "택배비", "배송료", "shipping", "delivery", "deliveryfee"],
    "other": ["기타비용", "기타수수료", "광고비", "포장비", "기타", "other", "adcost"],
    "cost": ["원가", "매입원가", "매입가", "공급가", "cost", "purchasecost"],
}

# 이 단어가 포함된 컬럼은 해당 항목으로 잡지 않음
NEGATIVE_WORDS = {
    "sales": ["원가", "수수료", "배송", "순이익", "이익", "할인", "쿠폰", "적립",
              "환불", "취소", "율", "비율"],
    "commission": ["순이익", "율", "비율"],
    "shipping": ["율", "비율"],
    "other": ["율", "비율"],
    "cost": ["율", "비율"],
}
MATCH_THRESHOLD = 0.7

STANDARD_LABEL = {
    "platform": "플랫폼", "date": "정산일", "product": "상품명", "quantity": "수량",
    "sales": "매출액", "commission": "수수료", "shipping": "배송비",
    "other": "기타비용", "cost": "원가",
}
EXCLUDE_KEYWORDS = ["합계", "소계", "총계", "total", "sum"]
DETAIL_COLS = ["플랫폼", "정산일", "상품명", "수량", "매출액", "수수료", "배송비",
               "기타비용", "원가", "순이익", "마진율(%)", "원본파일"]


# ----------------------------------------------------------------------
# 샘플 데이터 (플랫폼마다 컬럼명이 다른 3종)
# ----------------------------------------------------------------------
class MemoryFile:
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

    def rows(n):
        out = []
        for _ in range(n):
            p = rng.choice(products)
            qty = int(rng.integers(1, 5))
            sales = prices[p] * qty
            date = f"2026-{rng.integers(9, 11):02d}-{rng.integers(1, 29):02d}"
            out.append(dict(product=p, qty=qty, sales=sales,
                            cost=round(sales * cost_ratio[p]), date=date))
        return out

    cp = pd.DataFrame(rows(40))
    coupang = pd.DataFrame({
        "정산일": cp["date"], "상품명": cp["product"], "주문수량": cp["qty"],
        "매출액": cp["sales"], "판매수수료": (cp["sales"] * 0.108).round().astype(int),
        "배송비": rng.choice([0, 3000], size=len(cp)),
        "광고비": rng.integers(0, 2000, size=len(cp)),
    })

    nv = pd.DataFrame(rows(35))
    naver = pd.DataFrame({
        "주문일자": nv["date"], "상품명": nv["product"], "수량": nv["qty"],
        "결제금액": nv["sales"], "서비스이용료": (nv["sales"] * 0.035).round().astype(int),
        "배송비": rng.choice([0, 3500], size=len(nv)),
    })
    naver_title = pd.DataFrame([["스마트스토어 정산내역"] + [""] * (len(naver.columns) - 1)],
                               columns=naver.columns)
    naver = pd.concat([naver_title, naver], ignore_index=True)

    st_rows = pd.DataFrame(rows(30))
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
# 정규화 / 컬럼 매핑
# ----------------------------------------------------------------------
def normalize(text) -> str:
    s = str(text)
    s = re.sub(r"\(.*?\)|\[.*?\]|<.*?>", "", s)     # 결제금액(원) -> 결제금액
    s = re.sub(r"[\s_\-/·:.,]", "", s)               # 공백·구분기호 제거
    return s.lower()


def score_column(col_norm: str, alias_norm: str) -> float:
    """1.0 완전 일치 / 0.7~0.95 포함 관계 / 0.8 이상 유사도"""
    if not col_norm or not alias_norm:
        return 0.0
    if col_norm == alias_norm:
        return 1.0
    shorter, longer = sorted([col_norm, alias_norm], key=len)
    if len(shorter) >= 2 and shorter in longer:
        return 0.7 + 0.25 * len(shorter) / len(longer)
    ratio = SequenceMatcher(None, col_norm, alias_norm).ratio()
    return ratio if ratio >= 0.8 else 0.0


def is_negative(key: str, col_norm: str) -> bool:
    return any(w in col_norm for w in NEGATIVE_WORDS.get(key, []))


def map_columns(df: pd.DataFrame):
    """반환: (mapping {항목: 원본컬럼}, info {항목: (원본컬럼, 점수)})"""
    candidates = []
    for col in df.columns:
        cn = normalize(col)
        if not cn or cn.startswith("unnamed"):
            continue
        for key, aliases in COLUMN_ALIASES.items():
            if is_negative(key, cn):
                continue
            best = max(score_column(cn, normalize(a)) for a in aliases)
            if best >= MATCH_THRESHOLD:
                candidates.append((best, key, col))

    # 점수 높은 순으로 배정: 항목당 1개, 컬럼당 1개
    candidates.sort(key=lambda x: x[0], reverse=True)
    mapping, info, used_cols = {}, {}, set()
    for score, key, col in candidates:
        if key in mapping or col in used_cols:
            continue
        mapping[key] = col
        info[key] = (col, score)
        used_cols.add(col)
    return mapping, info


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
    """헤더 행을 0~9행에서 찾고, 매출액이 인식되는 첫 행을 사용한다."""
    for header in range(10):
        try:
            df = read_raw(raw, name, header)
        except Exception:
            continue
        mapping, info = map_columns(df)
        if "sales" in mapping:
            return df, mapping, info
    raise ValueError("'매출액'(판매가·결제금액·금액 등 포함)에 해당하는 컬럼을 찾지 못했습니다.")


def to_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(r"[,\s원₩$]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


# ----------------------------------------------------------------------
# 처리 / 계산
# ----------------------------------------------------------------------
def process_files(files, default_cost_rate: float):
    frames, errors, logs = [], [], []
    for f in files:
        try:
            df, m, info = load_settlement(f.getvalue(), f.name)
        except Exception as e:
            errors.append((f.name, str(e)))
            continue

        for key, (col, score) in info.items():
            logs.append({"파일": f.name, "표준항목": STANDARD_LABEL.get(key, key),
                         "인식된 컬럼": col, "일치도": f"{score:.0%}"})

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
        out["원가"] = (to_number(df[m["cost"]]) if "cost" in m
                     else out["매출액"] * default_cost_rate / 100)

        text = out[["플랫폼", "정산일", "상품명"]].astype(str).agg(" ".join, axis=1).str.lower()
        is_total = text.apply(lambda t: any(k in t for k in EXCLUDE_KEYWORDS))
        out = out[~is_total]
        out = out[(out["매출액"] != 0) | (out["수수료"] != 0)]

        out["순이익"] = out["매출액"] - out["수수료"] - out["배송비"] - out["기타비용"] - out["원가"]
        out["마진율(%)"] = np.where(out["매출액"] > 0, out["순이익"] / out["매출액"] * 100, 0).round(2)
        out["원본파일"] = f.name
        frames.append(out)

    detail = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=DETAIL_COLS)
    return detail, errors, pd.DataFrame(logs)


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
# 엑셀 생성 (열람·인쇄용 밝은 스타일)
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="111111")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    for col_idx, head in enumerate(ws[1], start=1):
        header = str(head.value or "")
        letter = get_column_letter(col_idx)
        values = [str(ws.cell(r, col_idx).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[letter].width = min(max(len(v) for v in values) * 1.6 + 4, 40)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row, col_idx)
            if isinstance(cell.value, (int, float)):
                cell.number_format = "0.00" if "율" in header else "#,##0"


def build_excel(detail, summary, totals, mapping_log) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(list(totals.items()), columns=["항목", "값"]).to_excel(
            writer, sheet_name="종합요약", index=False)
        summary.to_excel(writer, sheet_name="플랫폼별요약", index=False)
        detail.to_excel(writer, sheet_name="마진계산내역", index=False)
        if not mapping_log.empty:
            mapping_log.to_excel(writer, sheet_name="컬럼인식결과", index=False)
        for ws in writer.book.worksheets:
            style_sheet(ws)
    return buf.getvalue()


# ----------------------------------------------------------------------
# UI 컴포넌트
# ----------------------------------------------------------------------
def kpi(label: str, value: str, sub: str = "", accent: bool = False) -> str:
    cls = "kpi kpi-accent" if accent else "kpi"
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return (f'<div class="{cls}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>{sub_html}</div>')


def section(title: str):
    st.markdown(f'<div class="section-label">{title}</div>', unsafe_allow_html=True)


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Settings")
    cost_rate = st.slider("기본 원가율 (%)", 0, 100, 0, 1,
                          help="파일에 원가 컬럼이 없을 때 매출액 대비로 적용됩니다.")
    st.caption("원가 컬럼이 있는 파일은 해당 값을 우선 사용합니다.")
    st.divider()
    use_demo = st.toggle("샘플 데이터로 체험하기", value=False,
                         help="쿠팡·네이버·11번가 형식의 가상 정산 데이터 3개를 불러옵니다.")
    st.caption("실제 정산 파일을 올릴 때는 이 옵션을 끄세요.")

# ----------------------------------------------------------------------
# 타이틀
# ----------------------------------------------------------------------
st.markdown(
    """
    <div class="topbar">
        <span class="brand">SETTLEMENT</span>
        <span>Seller Dashboard</span>
    </div>
    <div class="hero">
        <h1>The Settlement Report</h1>
        <p>쇼핑몰별 정산 파일을 올리면 매출과 순이익을 한 장으로 정리해 드립니다.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader("정산 파일 업로드", type=["csv", "xlsx"],
                            accept_multiple_files=True, label_visibility="collapsed")

if use_demo:
    source_files = make_sample_files()
    st.info("샘플 데이터 3종(쿠팡·네이버·11번가)으로 분석 중입니다.")
elif uploaded:
    source_files = uploaded
else:
    st.markdown(
        """
        <div class="empty">
            <div class="empty-eyebrow">Upload</div>
            <div class="empty-title">정산 파일을 끌어다 놓으세요</div>
            <div class="empty-desc">
                CSV, XLSX 여러 개를 한 번에 올릴 수 있습니다.<br>
                플랫폼마다 컬럼명이 달라도 자동으로 맞춰서 합산합니다.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

with st.spinner("정산 데이터를 분석하는 중..."):
    detail, errors, mapping_log = process_files(source_files, cost_rate)

for name, msg in errors:
    st.error(f"{name} 처리 실패: {msg}")

if not mapping_log.empty:
    with st.expander("컬럼 자동 인식 결과 확인", expanded=False):
        st.caption("잘못 인식된 항목이 있으면 COLUMN_ALIASES 또는 NEGATIVE_WORDS를 수정하세요.")
        st.dataframe(mapping_log, hide_index=True)

if detail.empty:
    st.warning("처리 가능한 데이터가 없습니다. 매출액 컬럼이 포함되어 있는지 확인해 주세요.")
    st.stop()

summary = make_summary(detail)
totals = make_totals(detail)

# ----------------------------------------------------------------------
# 핵심 지표
# ----------------------------------------------------------------------
section("Overview")
cols = st.columns(4)
cols[0].markdown(kpi("총 매출", f"{totals['총매출']:,.0f}원",
                     f"{totals['주문건수']:,}건 주문"), unsafe_allow_html=True)
cols[1].markdown(kpi("총 비용", f"{totals['총비용']:,.0f}원",
                     "수수료 · 배송 · 기타 · 원가"), unsafe_allow_html=True)
cols[2].markdown(kpi("순수익", f"{totals['순이익']:,.0f}원",
                     "매출 − 전체 비용", accent=True), unsafe_allow_html=True)
cols[3].markdown(kpi("마진율", f"{totals['마진율(%)']:.2f}%", "순이익 ÷ 매출"),
                 unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
section("By Platform")
left, right = st.columns([3, 2])
with left:
    st.markdown('<div class="card-title">매출 · 순이익</div>', unsafe_allow_html=True)
    st.bar_chart(summary.set_index("플랫폼")[["매출액", "순이익"]],
                 color=[TEXT, SERIES_2], height=300)
with right:
    st.markdown('<div class="card-title">마진율 (%)</div>', unsafe_allow_html=True)
    st.bar_chart(summary.set_index("플랫폼")[["마진율(%)"]], color=[TEXT], height=300)

# ----------------------------------------------------------------------
# 상세 테이블
# ----------------------------------------------------------------------
section("Details")
money_cfg = {c: st.column_config.NumberColumn(c, format="%d") for c in
             ["매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]}
rate_cfg = {"마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f%%")}

tab1, tab2, tab3 = st.tabs(["플랫폼별 요약", "종합 정산 내역", "마진 계산 내역"])

with tab1:
    st.dataframe(summary, hide_index=True, column_config={**money_cfg, **rate_cfg})

with tab2:
    st.caption(f"총 {len(detail):,}건 · 파일 {len(source_files)}개")
    st.dataframe(detail[DETAIL_COLS], hide_index=True,
                 column_config={**money_cfg, **rate_cfg})

with tab3:
    st.caption("순이익 = 매출액 − 수수료 − 배송비 − 기타비용 − 원가  /  "
               "마진율 = 순이익 ÷ 매출액 × 100")
    st.dataframe(
        detail[["플랫폼", "상품명", "매출액", "수수료", "배송비", "기타비용", "원가",
                "순이익", "마진율(%)"]],
        hide_index=True, column_config={**money_cfg, **rate_cfg},
    )

# ----------------------------------------------------------------------
# 다운로드
# ----------------------------------------------------------------------
section("Export")
st.markdown(
    f'<div class="card"><div class="card-title">보고서 다운로드</div>'
    f'<div style="color:{TEXT_SUB}; font-size:0.9rem; margin-bottom:14px;">'
    f'종합요약 · 플랫폼별요약 · 마진계산내역 · 컬럼인식결과 시트로 구성된 엑셀 파일입니다.</div></div>',
    unsafe_allow_html=True,
)
st.download_button(
    label="엑셀 보고서 다운로드 (.xlsx)",
    data=build_excel(detail, summary, totals, mapping_log),
    file_name=f"정산보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
