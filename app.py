import csv
import io
import re
import warnings
from difflib import SequenceMatcher

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 페이지 & 스타일
# ----------------------------------------------------------------------
st.set_page_config(page_title="Settlement Report", page_icon="▫️", layout="wide")

BG = "#0E0F12"
SURFACE = "#16181D"
LINE = "#3A3F4B"
TEXT = "#F5F5F5"
TEXT_SUB = "#B8BDC7"
C_SALES = "#F5F5F5"
C_PROFIT = "#3B82F6"
C_BAD = "#EF4444"
COST_COLORS = {"수수료": "#F59E0B", "배송비": "#8B5CF6", "기타비용": "#EC4899", "원가": "#6B7280"}

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&display=swap');
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    html, body, [class*="css"], .stApp {{
        font-family: 'Pretendard', 'Noto Sans KR', -apple-system, 'Segoe UI', sans-serif;
    }}
    .stApp {{ background: {BG}; color: {TEXT}; }}
    .block-container {{ padding-top: 2rem; padding-bottom: 4rem; max-width: 1240px; }}
    header[data-testid="stHeader"] {{ background: transparent; }}
    .topbar {{ display:flex; justify-content:space-between; border-top:1px solid {LINE};
              border-bottom:1px solid {LINE}; padding:12px 4px; font-size:0.8rem;
              letter-spacing:0.12em; text-transform:uppercase; color:{TEXT_SUB}; }}
    .brand {{ font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.1rem;
             letter-spacing:0.2em; font-weight:600; color:{TEXT}; }}
    .hero {{ text-align:center; padding:44px 0 36px 0; border-bottom:1px solid {LINE};
            margin-bottom:28px; }}
    .hero h1 {{ font-family:'Playfair Display','Nanum Myeongjo',serif; font-weight:400;
               font-size:2.5rem; letter-spacing:0.08em; text-transform:uppercase;
               margin:0; color:{TEXT}; }}
    .hero p {{ color:{TEXT_SUB}; margin:14px 0 0 0; }}
    .section-label {{ font-size:0.78rem; letter-spacing:0.18em; text-transform:uppercase;
                     color:{TEXT}; font-weight:700; margin:40px 0 16px 0;
                     padding-bottom:8px; border-bottom:1px solid {LINE}; }}
    .kpi {{ background:{SURFACE}; border:1px solid {LINE}; padding:20px 22px; min-height:128px; }}
    .kpi-label {{ font-size:0.74rem; letter-spacing:0.14em; text-transform:uppercase;
                 color:{TEXT_SUB}; font-weight:600; }}
    .kpi-value {{ font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.9rem;
                 color:{TEXT}; margin-top:10px; line-height:1.1; }}
    .kpi-sub {{ font-size:0.84rem; color:{TEXT_SUB}; margin-top:10px; }}
    .kpi-accent {{ background:{TEXT}; border-color:{TEXT}; }}
    .kpi-accent .kpi-label, .kpi-accent .kpi-sub {{ color:#3A3F4B; }}
    .kpi-accent .kpi-value {{ color:{BG}; }}
    .card {{ border:1px solid {LINE}; background:{SURFACE}; padding:20px 24px; margin-bottom:12px; }}
    .card-title {{ font-size:0.8rem; letter-spacing:0.14em; text-transform:uppercase;
                  font-weight:700; color:{TEXT}; margin-bottom:10px; }}
    .insight {{ background:{SURFACE}; border:1px solid {LINE}; border-left:3px solid {TEXT};
               padding:16px 20px; margin-bottom:10px; }}
    .insight-title {{ font-weight:700; color:{TEXT}; margin-bottom:6px; }}
    .insight-text {{ color:{TEXT_SUB}; line-height:1.7; }}
    div[data-testid="stFileUploader"] section {{ background:{SURFACE}; border:1px dashed {TEXT_SUB};
                                               border-radius:0; }}
    button[data-baseweb="tab"] {{ font-weight:600; color:{TEXT_SUB} !important; }}
    button[data-baseweb="tab"][aria-selected="true"] {{ color:{TEXT} !important; }}
    div[data-baseweb="tab-highlight"] {{ background-color:{TEXT} !important; }}
    div[data-baseweb="tab-border"] {{ background-color:{LINE} !important; }}
    [data-testid="stDownloadButton"] button, .stButton > button {{
        border-radius:0; font-weight:700; background:{TEXT}; color:{BG} !important;
        border:1px solid {TEXT}; padding:0.7rem 1.4rem; }}
    section[data-testid="stSidebar"] {{ background:{SURFACE}; border-right:1px solid {LINE}; }}
    div[data-testid="stDataFrame"] {{ border:1px solid {LINE}; }}
    .empty {{ border:1px solid {LINE}; padding:64px 24px; text-align:center; background:{SURFACE}; }}
    .empty-title {{ font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.6rem;
                   margin-top:10px; color:{TEXT}; }}
    .empty-desc {{ color:{TEXT_SUB}; margin-top:10px; line-height:1.7; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------
# 컬럼 매핑 설정
# ----------------------------------------------------------------------
COLUMN_ALIASES = {
    "platform": ["쇼핑몰", "플랫폼", "채널", "판매처", "몰이름", "판매채널", "스토어명",
                 "platform", "mall", "channel", "store", "shop"],
    "date": ["정산일", "정산기준일", "주문일", "주문일자", "주문일시", "결제일", "결제일시",
             "판매일", "판매일시", "날짜", "정산기간", "date", "orderdate", "settledate"],
    "product": ["상품명", "상품", "품목명", "제품명", "품목", "옵션명", "상품옵션", "상품정보",
                "product", "productname", "item", "title"],
    "quantity": ["수량", "주문수량", "판매수량", "구매수량", "qty", "quantity"],
    "sales": ["매출액", "매출", "매출금액", "총매출", "판매금액", "판매가", "판매가격", "판매액",
              "판매대금", "결제금액", "결제액", "결제가", "총결제금액", "주문금액", "주문합계",
              "구매금액", "구매가", "상품금액", "총판매금액", "거래금액", "정산대상금액", "금액",
              "sales", "revenue", "amount", "price", "gmv", "paidamount", "orderamount"],
    "commission": ["수수료", "판매수수료", "서비스이용료", "중개수수료", "수수료합계", "판매자수수료",
                   "마켓수수료", "이용료", "commission", "fee", "commissionfee"],
    "shipping": ["배송비", "배송비용", "택배비", "배송료", "배송수수료", "운송비", "배송비합계",
                 "shipping", "delivery", "shippingfee", "deliveryfee"],
    "other": ["기타비용", "기타수수료", "광고비", "포장비", "기타", "쿠폰비", "프로모션비용",
              "other", "adcost", "adfee"],
    "cost": ["원가", "매입원가", "매입가", "공급가", "공급가액", "상품원가", "구매원가",
             "cost", "purchasecost"],
}
NEGATIVE_WORDS = {
    "sales": ["원가", "수수료", "배송", "순이익", "이익", "율", "비율", "할인", "쿠폰",
              "환불", "취소", "적립", "부가세", "세액", "차감"],
    "commission": ["순이익", "율", "비율"],
    "shipping": ["율", "비율"],
    "other": ["율", "비율"],
    "cost": ["율", "비율"],
}
FIELD_ORDER = ["sales", "commission", "shipping", "other", "cost", "quantity", "product", "date", "platform"]
STANDARD_LABEL = {
    "platform": "플랫폼", "date": "정산일", "product": "상품명", "quantity": "수량",
    "sales": "매출액", "commission": "수수료", "shipping": "배송비",
    "other": "기타비용", "cost": "원가",
}
MONEY_HINTS = ["금액", "가", "액", "원", "price", "amount", "sales", "매출", "결제", "판매"]
SKIP_HINTS = ["수량", "qty", "율", "비율", "일", "date", "번호", "id", "no"]
MATCH_THRESHOLD = 0.7
TOTAL_RE = re.compile(r"^(합계|소계|총계|total|sum)(\s|$)", re.IGNORECASE)
DETAIL_COLS = ["플랫폼", "정산월", "정산일", "상품명", "수량", "매출액", "수수료", "배송비",
               "기타비용", "원가", "순이익", "마진율(%)", "원본파일"]


# ----------------------------------------------------------------------
# 샘플 데이터 (쿠팡·네이버·11번가 형식)
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
    prices = dict(zip(products, [39000, 22000, 28000, 54000, 45000, 12000, 31000, 19000, 9000, 26000]))
    cost_ratio = dict(zip(products, [0.55, 0.45, 0.50, 0.58, 0.52, 0.40, 0.48, 0.42, 0.35, 0.50]))

    def rows(n):
        out = []
        for _ in range(n):
            p = rng.choice(products)
            qty = int(rng.integers(1, 5))
            sales = prices[p] * qty
            date = f"2026-{rng.integers(8, 11):02d}-{rng.integers(1, 29):02d}"
            out.append(dict(product=p, qty=qty, sales=sales, cost=round(sales * cost_ratio[p]), date=date))
        return out

    cp = pd.DataFrame(rows(40))
    coupang = pd.DataFrame({
        "정산일": cp["date"], "상품명": cp["product"], "주문수량": cp["qty"],
        "매출액": cp["sales"], "판매수수료": (cp["sales"] * 0.108).round().astype(int),
        "배송비": rng.choice([0, 3000], size=len(cp)), "광고비": rng.integers(0, 2000, size=len(cp)),
    })
    nv = pd.DataFrame(rows(35))
    naver = pd.DataFrame({
        "주문일자": nv["date"], "상품명": nv["product"], "수량": nv["qty"],
        "결제금액": nv["sales"], "서비스이용료": (nv["sales"] * 0.035).round().astype(int),
        "배송비": rng.choice([0, 3500], size=len(nv)),
    })
    title = pd.DataFrame([["스마트스토어 정산내역"] + [""] * (len(naver.columns) - 1)], columns=naver.columns)
    naver = pd.concat([title, naver], ignore_index=True)

    st_rows = pd.DataFrame(rows(30))
    elev = pd.DataFrame({
        "판매일": st_rows["date"], "품목명": st_rows["product"], "판매수량": st_rows["qty"],
        "판매금액": st_rows["sales"], "수수료": (st_rows["sales"] * 0.12).round().astype(int),
        "택배비": rng.choice([0, 2500], size=len(st_rows)),
    })
    total = pd.DataFrame([{"판매일": "합계", "품목명": "", "판매수량": elev["판매수량"].sum(),
                           "판매금액": elev["판매금액"].sum(), "수수료": elev["수수료"].sum(),
                           "택배비": elev["택배비"].sum()}])
    elev = pd.concat([elev, total], ignore_index=True)

    files = []
    buf = io.BytesIO()
    coupang.to_excel(buf, index=False)
    files.append(MemoryFile("sample_쿠팡_정산.xlsx", buf.getvalue()))
    files.append(MemoryFile("sample_네이버_정산.csv", naver.to_csv(index=False).encode("utf-8-sig")))
    buf = io.BytesIO()
    elev.to_excel(buf, index=False)
    files.append(MemoryFile("sample_11번가_정산.xlsx", buf.getvalue()))
    return files


# ----------------------------------------------------------------------
# 파싱 / 매핑
# ----------------------------------------------------------------------
def normalize(text) -> str:
    s = str(text)
    s = re.sub(r"\(.*?\)|\[.*?\]|<.*?>", "", s)
    s = re.sub(r"[\s_\-/·:.,*]", "", s)
    return s.lower()


def score_column(col_norm: str, alias_norm: str) -> float:
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


def match_columns(names):
    candidates = []
    for col in names:
        cn = normalize(col)
        if not cn or cn.startswith("unnamed") or cn.startswith("컬럼"):
            continue
        for key, aliases in COLUMN_ALIASES.items():
            if is_negative(key, cn):
                continue
            best = max(score_column(cn, normalize(a)) for a in aliases)
            if best >= MATCH_THRESHOLD:
                candidates.append((best, key, col))
    candidates.sort(key=lambda x: x[0], reverse=True)
    mapping, scores, used = {}, {}, set()
    for score, key, col in candidates:
        if key in mapping or col in used:
            continue
        mapping[key], scores[key] = col, score
        used.add(col)
    return mapping, scores


def numeric_share(series: pd.Series) -> float:
    s = series.astype(str).str.strip()
    s = s[(s != "") & (s.str.lower() != "nan")]
    if s.empty:
        return 0.0
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True).str.replace(r"[,\s원₩$]", "", regex=True)
    return float(pd.to_numeric(s, errors="coerce").notna().mean())


def to_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(r"[,\s원₩$]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def guess_sales_by_value(df: pd.DataFrame, used: set):
    best, best_sum = None, 0.0
    for col in df.columns:
        if col in used:
            continue
        cn = normalize(col)
        if is_negative("sales", cn) or any(h in cn for h in SKIP_HINTS):
            continue
        if not any(h in cn for h in MONEY_HINTS):
            continue
        if numeric_share(df[col]) < 0.9:
            continue
        total = float(to_number(df[col]).sum())
        if total > best_sum:
            best, best_sum = col, total
    return best


def read_grid(raw: bytes, name: str) -> pd.DataFrame:
    if name.lower().endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(raw), header=None, dtype=str)
    text = None
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr", "utf-16"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("인코딩을 판별하지 못했습니다. UTF-8 또는 CP949로 저장해 주세요.")
    try:
        sep = csv.Sniffer().sniff(text[:5000], delimiters=",\t;|").delimiter
    except csv.Error:
        sep = ","
    rows = list(csv.reader(io.StringIO(text), delimiter=sep))
    width = max((len(r) for r in rows), default=0)
    return pd.DataFrame([r + [""] * (width - len(r)) for r in rows], dtype=str)


def find_header_row(grid: pd.DataFrame, limit: int = 40):
    best_row, best_hits = None, 0
    for r in range(min(limit, len(grid))):
        cells = [str(v).strip() for v in grid.iloc[r].tolist()]
        cells = [c for c in cells if c and c.lower() != "nan"]
        hits = len(match_columns(cells)[0])
        if hits > best_hits:
            best_row, best_hits = r, hits
    return best_row


def dedupe(names):
    seen, out = {}, []
    for i, n in enumerate(names):
        n = n or f"컬럼{i + 1}"
        if n in seen:
            seen[n] += 1
            n = f"{n}_{seen[n]}"
        else:
            seen[n] = 0
        out.append(n)
    return out


def parse_file(f) -> dict:
    try:
        grid = read_grid(f.getvalue(), f.name)
    except Exception as e:
        return {"name": f.name, "error": f"파일을 읽지 못했습니다: {e}"}
    header_row = find_header_row(grid)
    if header_row is None:
        header_row = 0
    raw_names = [("" if pd.isna(v) else str(v).strip()) for v in grid.iloc[header_row].tolist()]
    names = dedupe(raw_names)
    df = grid.iloc[header_row + 1:].copy()
    df.columns = names
    df = df.reset_index(drop=True)
    mapping, scores = match_columns(names)
    guessed = False
    if "sales" not in mapping:
        col = guess_sales_by_value(df, used=set(mapping.values()))
        if col:
            mapping["sales"], scores["sales"], guessed = col, 0.5, True
    return {"name": f.name, "df": df, "columns": names, "header_row": header_row,
            "mapping": mapping, "scores": scores, "guessed": guessed}


def to_month(series: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        dt = pd.to_datetime(series, errors="coerce")
    return dt.dt.strftime("%Y-%m").fillna("미상")


def build_frame(p: dict, m: dict, cost_rate: float) -> pd.DataFrame:
    df = p["df"]
    idx = df.index

    def num(key):
        return to_number(df[m[key]]) if key in m else pd.Series(0.0, index=idx)

    stem = p["name"].rsplit(".", 1)[0]
    out = pd.DataFrame(index=idx)
    out["플랫폼"] = df[m["platform"]].fillna(stem).astype(str) if "platform" in m else stem
    out["정산일"] = df[m["date"]].fillna("").astype(str) if "date" in m else ""
    out["정산월"] = to_month(out["정산일"])
    out["상품명"] = df[m["product"]].fillna("").astype(str) if "product" in m else ""
    out["상품명"] = out["상품명"].replace(["", "nan"], "(상품명 없음)")
    out["수량"] = num("quantity")
    out["매출액"] = num("sales")
    out["수수료"] = num("commission")
    out["배송비"] = num("shipping")
    out["기타비용"] = num("other")
    out["원가"] = num("cost") if "cost" in m else out["매출액"] * cost_rate / 100

    cells = out[["플랫폼", "정산일", "상품명"]].astype(str).apply(
        lambda c: c.str.strip().str.lower().str.match(TOTAL_RE))
    out = out[~cells.any(axis=1)]
    out = out[(out["매출액"] != 0) | (out["수수료"] != 0)]

    out["순이익"] = out["매출액"] - out["수수료"] - out["배송비"] - out["기타비용"] - out["원가"]
    out["마진율(%)"] = np.where(out["매출액"] > 0, out["순이익"] / out["매출액"] * 100, 0).round(2)
    out["원본파일"] = p["name"]
    return out


def process_all(parsed: list, overrides: dict, cost_rate: float):
    frames, errors, logs = [], [], []
    for p in parsed:
        if p.get("error"):
            errors.append((p["name"], p["error"]))
            continue
        m = overrides.get(p["name"], p["mapping"])
        if "sales" not in m:
            errors.append((p["name"], "매출액 컬럼이 지정되지 않았습니다. 아래 매핑에서 선택해 주세요."))
            continue
        for key in FIELD_ORDER:
            if key in m:
                method = "자동" if p["mapping"].get(key) == m[key] else "수동"
                if key == "sales" and p.get("guessed") and method == "자동":
                    method = "값 기반 추정"
                logs.append({"파일": p["name"], "표준항목": STANDARD_LABEL[key],
                             "인식된 컬럼": m[key], "방식": method})
        frames.append(build_frame(p, m, cost_rate))
    detail = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=DETAIL_COLS)
    return detail, errors, pd.DataFrame(logs)


# ----------------------------------------------------------------------
# 리포트 계산
# ----------------------------------------------------------------------
def safe_div(a, b):
    return (a / b.replace(0, np.nan)).fillna(0)


def make_totals(detail: pd.DataFrame) -> dict:
    sales = float(detail["매출액"].sum())
    profit = float(detail["순이익"].sum())
    return {
        "주문건수": int(len(detail)), "총매출": sales,
        "수수료": float(detail["수수료"].sum()), "배송비": float(detail["배송비"].sum()),
        "기타비용": float(detail["기타비용"].sum()), "원가": float(detail["원가"].sum()),
        "총비용": float(detail[["수수료", "배송비", "기타비용", "원가"]].sum().sum()),
        "순이익": profit,
        "마진율(%)": round(profit / sales * 100, 2) if sales > 0 else 0.0,
    }


def make_platform_report(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame()
    g = detail.groupby("플랫폼", as_index=False).agg(
        주문건수=("매출액", "size"), 수량=("수량", "sum"), 매출액=("매출액", "sum"),
        수수료=("수수료", "sum"), 배송비=("배송비", "sum"), 기타비용=("기타비용", "sum"),
        원가=("원가", "sum"), 순이익=("순이익", "sum"),
    )
    tot_s = g["매출액"].sum()
    tot_p = g["순이익"].sum()
    g["객단가"] = safe_div(g["매출액"], g["주문건수"])
    g["수수료율(%)"] = safe_div(g["수수료"], g["매출액"]) * 100
    g["마진율(%)"] = safe_div(g["순이익"], g["매출액"]) * 100
    g["매출비중(%)"] = (g["매출액"] / (tot_s or np.nan) * 100).fillna(0)
    g["순이익비중(%)"] = (g["순이익"] / (tot_p or np.nan) * 100).fillna(0)
    g = g.sort_values("순이익", ascending=False).reset_index(drop=True)
    g.insert(0, "순위", range(1, len(g) + 1))
    return g.round(2)


def make_product_report(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame()
    g = detail.groupby(["플랫폼", "상품명"], as_index=False).agg(
        주문건수=("매출액", "size"), 수량=("수량", "sum"), 매출액=("매출액", "sum"),
        수수료=("수수료", "sum"), 원가=("원가", "sum"), 순이익=("순이익", "sum"),
    )
    g["마진율(%)"] = safe_div(g["순이익"], g["매출액"]) * 100
    return g.sort_values("순이익", ascending=False).reset_index(drop=True).round(2)


def make_monthly(detail: pd.DataFrame) -> pd.DataFrame:
    m = detail[detail["정산월"] != "미상"]
    if m.empty:
        return pd.DataFrame()
    g = m.groupby(["정산월", "플랫폼"], as_index=False).agg(매출액=("매출액", "sum"), 순이익=("순이익", "sum"))
    g["마진율(%)"] = safe_div(g["순이익"], g["매출액"]) * 100
    return g.round(2).sort_values(["정산월", "플랫폼"]).reset_index(drop=True)


def build_insights(plat, prod, monthly, totals):
    items = []
    items.append(("전체 요약",
                  f"총 매출 {totals['총매출']:,.0f}원, 순이익 {totals['순이익']:,.0f}원, "
                  f"마진율 {totals['마진율(%)']:.2f}%입니다. 총 {totals['주문건수']:,}건의 주문을 집계했습니다."))
    if plat.empty:
        return items

    top = plat.iloc[0]
    items.append(("수익 1위 플랫폼",
                  f"{top['플랫폼']}가 순이익 {top['순이익']:,.0f}원(전체의 {top['순이익비중(%)']:.1f}%)으로 가장 큽니다. "
                  f"이 플랫폼의 마진율은 {top['마진율(%)']:.2f}%입니다."))

    if len(plat) > 1:
        hi = plat.sort_values("마진율(%)", ascending=False).iloc[0]
        lo = plat.sort_values("마진율(%)").iloc[0]
        items.append(("마진율 비교",
                      f"마진율이 가장 높은 곳은 {hi['플랫폼']}({hi['마진율(%)']:.2f}%), "
                      f"가장 낮은 곳은 {lo['플랫폼']}({lo['마진율(%)']:.2f}%)입니다. "
                      f"두 곳의 차이는 {hi['마진율(%)'] - lo['마진율(%)']:.2f}%p입니다."))
        fee = plat.sort_values("수수료율(%)", ascending=False).iloc[0]
        items.append(("수수료 부담",
                      f"{fee['플랫폼']}의 수수료율이 {fee['수수료율(%)']:.2f}%로 가장 높습니다. "
                      f"카테고리별 수수료나 판매 조건을 점검해 볼 만합니다."))

    cost_items = {"수수료": totals["수수료"], "배송비": totals["배송비"],
                  "기타비용": totals["기타비용"], "원가": totals["원가"]}
    cost_sum = sum(cost_items.values())
    if cost_sum > 0:
        biggest = max(cost_items, key=cost_items.get)
        share = cost_items[biggest] / cost_sum * 100
        items.append(("비용 구조",
                      f"전체 비용({cost_sum:,.0f}원) 중 {biggest}가 {share:.1f}%로 가장 큽니다."))
        if totals["배송비"] / cost_sum * 100 > 25:
            items.append(("배송비 경고", "배송비가 비용의 25%를 넘습니다. 무료배송 기준 금액과 택배 계약을 재검토해 보세요."))
        other_share = totals["기타비용"] / cost_sum * 100
        if other_share > 20:
            items.append(("광고·쿠폰 점검",
                          f"기타비용(광고·쿠폰 등)이 비용의 {other_share:.1f}%입니다. 광고 수익률과 쿠폰 사용률을 확인해 보세요."))

    if not prod.empty:
        best = prod.iloc[0]
        items.append(("수익 1위 상품",
                      f"[{best['플랫폼']}] {best['상품명']}이 순이익 {best['순이익']:,.0f}원, "
                      f"마진율 {best['마진율(%)']:.2f}%로 가장 큽니다."))
        loss = prod[prod["순이익"] < 0]
        if not loss.empty:
            names = ", ".join(loss["상품명"].head(3).tolist())
            items.append(("적자 상품",
                          f"순이익이 마이너스인 상품이 {len(loss)}개 있습니다(예: {names}). "
                          f"판매가 또는 원가를 재검토하거나 판매 중단을 고려해 보세요."))

    if not monthly.empty:
        ms = monthly.groupby("정산월")["매출액"].sum().sort_index()
        if len(ms) >= 2:
            last, prev = ms.iloc[-1], ms.iloc[-2]
            chg = (last - prev) / prev * 100 if prev else 0
            items.append(("월별 흐름",
                          f"{ms.index[-1]} 매출은 {last:,.0f}원으로 직전월({ms.index[-2]}) 대비 {chg:+.1f}% 입니다."))

    if totals["마진율(%)"] < 10:
        items.append(("종합 제언", "전체 마진율이 10% 미만입니다. 원가율과 수수료 구조를 우선 점검하세요."))
    elif totals["마진율(%)"] >= 30:
        items.append(("종합 제언", "전체 마진율이 30% 이상으로 양호합니다. 수익성이 높은 상품의 판매 비중을 더 늘려 보세요."))
    else:
        items.append(("종합 제언", "마진율은 보통 수준입니다. 마진율이 낮은 플랫폼과 상품부터 개선하면 효과가 큽니다."))
    return items


# ----------------------------------------------------------------------
# 차트 (가로 막대 · 세로 적층)
# ----------------------------------------------------------------------
def style_chart(ch, height):
    return (ch.properties(height=height)
              .configure_view(strokeWidth=0)
              .configure(background=SURFACE)
              .configure_axis(labelColor=TEXT_SUB, titleColor=TEXT_SUB, gridColor=LINE,
                              domainColor=LINE, labelFontSize=13)
              .configure_legend(labelColor=TEXT, titleColor=TEXT, labelFontSize=13, orient="top"))


def grouped_bar(df, cat, measures, colors):
    long = df[[cat] + measures].melt(id_vars=[cat], value_vars=measures, var_name="구분", value_name="값")
    base = alt.Chart(long).encode(
        y=alt.Y(f"{cat}:N", sort=None, title=None, axis=alt.Axis(labelLimit=280, labelFontSize=14)),
        yOffset=alt.YOffset("구분:N"),
        x=alt.X("값:Q", title=None, axis=alt.Axis(format=",.0f")),
        color=alt.Color("구분:N", scale=alt.Scale(domain=measures, range=colors), legend=alt.Legend(title=None)),
        tooltip=[alt.Tooltip(f"{cat}:N"), alt.Tooltip("구분:N"), alt.Tooltip("값:Q", format=",.0f")],
    )
    bars = base.mark_bar(cornerRadiusEnd=3)
    labels = base.mark_text(align="left", dx=4, color=TEXT, fontSize=12).encode(
        text=alt.Text("값:Q", format=",.0f"))
    return style_chart(alt.layer(bars, labels), max(300, 90 * len(df) + 80))


def single_bar(df, cat, val, color, fmt=",.0f", signed_color=False):
    if signed_color:
        fill = alt.condition(f"datum['{val}'] < 0", alt.value(C_BAD), alt.value(color))
    else:
        fill = alt.value(color)
    base = alt.Chart(df).encode(
        y=alt.Y(f"{cat}:N", sort=None, title=None, axis=alt.Axis(labelLimit=280, labelFontSize=14)),
        x=alt.X(f"{val}:Q", title=None, axis=alt.Axis(format=fmt)),
        tooltip=[alt.Tooltip(f"{cat}:N"), alt.Tooltip(f"{val}:Q", format=fmt)],
    )
    bars = base.mark_bar(cornerRadiusEnd=3).encode(color=fill)
    labels = base.mark_text(align="left", dx=4, color=TEXT, fontSize=12).encode(
        text=alt.Text(f"{val}:Q", format=fmt))
    return style_chart(alt.layer(bars, labels), max(260, 64 * len(df) + 80))


def cost_stack(plat):
    long = plat[["플랫폼"] + list(COST_COLORS)].melt(id_vars="플랫폼", var_name="비용", value_name="금액")
    ch = alt.Chart(long).mark_bar(cornerRadiusEnd=0).encode(
        y=alt.Y("플랫폼:N", sort=None, title=None, axis=alt.Axis(labelLimit=280, labelFontSize=14)),
        x=alt.X("금액:Q", stack="zero", title=None, axis=alt.Axis(format=",.0f")),
        color=alt.Color("비용:N", scale=alt.Scale(domain=list(COST_COLORS), range=list(COST_COLORS.values())),
                        legend=alt.Legend(title=None)),
        tooltip=[alt.Tooltip("플랫폼:N"), alt.Tooltip("비용:N"), alt.Tooltip("금액:Q", format=",.0f")],
    )
    return style_chart(ch, max(280, 70 * len(plat) + 100))


def monthly_line(monthly):
    ch = alt.Chart(monthly).mark_line(point=alt.OverlayMarkDef(size=70), strokeWidth=3).encode(
        x=alt.X("정산월:N", title=None, axis=alt.Axis(labelAngle=0)),
        y=alt.Y("순이익:Q", title=None, axis=alt.Axis(format=",.0f")),
        color=alt.Color("플랫폼:N", legend=alt.Legend(title=None)),
        tooltip=[alt.Tooltip("정산월:N"), alt.Tooltip("플랫폼:N"),
                 alt.Tooltip("매출액:Q", format=",.0f"), alt.Tooltip("순이익:Q", format=",.0f")],
    )
    return style_chart(ch, 380)


# ----------------------------------------------------------------------
# 엑셀
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="111111")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.title == "분석코멘트":
        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 110
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        return
    for col_idx, head_cell in enumerate(ws[1], start=1):
        header = str(head_cell.value or "")
        values = [str(ws.cell(r, col_idx).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(v) for v in values) * 1.5 + 4, 40)
        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row, col_idx)
            if isinstance(cell.value, (int, float)):
                rate = ("율" in header or "비중" in header or "율" in str(ws.cell(row, 1).value or ""))
                cell.number_format = "0.00" if rate else "#,##0"


def build_excel(totals, plat, prod, monthly, detail, insights, mapping_log) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        summary_rows = [
            ("주문건수", totals["주문건수"]), ("총매출", totals["총매출"]),
            ("수수료", totals["수수료"]), ("배송비", totals["배송비"]),
            ("기타비용", totals["기타비용"]), ("원가", totals["원가"]),
            ("총비용", totals["총비용"]), ("순이익", totals["순이익"]),
            ("마진율(%)", totals["마진율(%)"]),
        ]
        pd.DataFrame(summary_rows, columns=["항목", "값"]).to_excel(w, sheet_name="종합요약", index=False)
        plat.to_excel(w, sheet_name="플랫폼별상세", index=False)
        prod.to_excel(w, sheet_name="상품별상세", index=False)
        if monthly.empty:
            pd.DataFrame({"안내": ["정산일을 날짜로 인식하지 못해 월별 추이를 만들지 않았습니다."]}).to_excel(
                w, sheet_name="월별추이", index=False)
        else:
            monthly.to_excel(w, sheet_name="월별추이", index=False)
        pd.DataFrame(insights, columns=["구분", "분석 내용"]).to_excel(w, sheet_name="분석코멘트", index=False)
        detail[DETAIL_COLS].to_excel(w, sheet_name="마진계산내역", index=False)
        if not mapping_log.empty:
            mapping_log.to_excel(w, sheet_name="컬럼인식결과", index=False)
        for ws in w.book.worksheets:
            style_sheet(ws)
    return buf.getvalue()


# ----------------------------------------------------------------------
# UI 헬퍼
# ----------------------------------------------------------------------
def kpi(label, value, sub="", accent=False):
    cls = "kpi kpi-accent" if accent else "kpi"
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return f'<div class="{cls}"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>{sub_html}</div>'


def section(title):
    st.markdown(f'<div class="section-label">{title}</div>', unsafe_allow_html=True)


MONEY_CFG = lambda cols: {c: st.column_config.NumberColumn(c, format="%d") for c in cols}
RATE_CFG = lambda cols: {c: st.column_config.NumberColumn(c, format="%.2f%%") for c in cols}


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Settings")
    cost_rate = st.slider("기본 원가율 (%)", 0, 100, 0, 1,
                          help="파일에 원가 컬럼이 없을 때 매출액 대비로 적용됩니다.")
    st.divider()
    use_demo = st.toggle("샘플 데이터로 체험하기", value=False)
    st.caption("실제 정산 파일을 올릴 때는 이 옵션을 끄세요.")

# ----------------------------------------------------------------------
# 헤더 & 업로드
# ----------------------------------------------------------------------
st.markdown(
    """
    <div class="topbar"><span class="brand">SETTLEMENT</span><span>Seller Dashboard</span></div>
    <div class="hero">
        <h1>The Settlement Report</h1>
        <p>쇼핑몰별 정산 파일을 올리면 플랫폼·상품·월별 수익과 분석 코멘트를 한 번에 보여 드립니다.</p>
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
            <div class="kpi-label">Upload</div>
            <div class="empty-title">정산 파일을 끌어다 놓으세요</div>
            <div class="empty-desc">CSV, XLSX 여러 개를 한 번에 올릴 수 있습니다.<br>
            컬럼명이 플랫폼마다 달라도 자동으로 맞춰 합산합니다.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ----------------------------------------------------------------------
# 컬럼 매핑 확인
# ----------------------------------------------------------------------
parsed = [parse_file(f) for f in source_files]

section("Column Mapping")
st.caption("파일마다 자동 인식된 컬럼을 보여 드립니다. 틀린 항목은 드롭다운에서 직접 고를 수 있습니다.")
overrides = {}
for i, p in enumerate(parsed):
    if p.get("error"):
        st.error(f"{p['name']}: {p['error']}")
        continue
    needs_review = ("sales" not in p["mapping"]) or p["guessed"]
    title = f"{p['name']} · 헤더 {p['header_row'] + 1}행 · 자동 인식 {len(p['mapping'])}개"
    with st.expander(title, expanded=needs_review):
        st.dataframe(p["df"].head(5), hide_index=True)
        chosen = {}
        gcols = st.columns(3)
        for j, key in enumerate(FIELD_ORDER):
            options = ["(없음)"] + p["columns"]
            default = p["mapping"].get(key)
            idx = options.index(default) if default in options else 0
            pick = gcols[j % 3].selectbox(STANDARD_LABEL[key], options, index=idx,
                                          key=f"map_{i}_{p['name']}_{key}")
            if pick != "(없음)":
                chosen[key] = pick
        overrides[p["name"]] = chosen

with st.spinner("정산 데이터를 분석하는 중..."):
    detail, errors, mapping_log = process_all(parsed, overrides, cost_rate)

for name, msg in errors:
    st.error(f"{name} 처리 실패: {msg}")

if detail.empty:
    st.warning("처리 가능한 데이터가 없습니다. 위 매핑에서 매출액 컬럼을 지정해 주세요.")
    st.stop()

plat = make_platform_report(detail)
prod = make_product_report(detail)
monthly = make_monthly(detail)
totals = make_totals(detail)
insights = build_insights(plat, prod, monthly, totals)

# ----------------------------------------------------------------------
# 핵심 지표
# ----------------------------------------------------------------------
section("Overview")
c = st.columns(4)
c[0].markdown(kpi("총 매출", f"{totals['총매출']:,.0f}원", f"{totals['주문건수']:,}건 주문"), unsafe_allow_html=True)
c[1].markdown(kpi("총 비용", f"{totals['총비용']:,.0f}원", "수수료 · 배송 · 기타 · 원가"), unsafe_allow_html=True)
c[2].markdown(kpi("순수익", f"{totals['순이익']:,.0f}원", "매출 − 전체 비용", accent=True), unsafe_allow_html=True)
c[3].markdown(kpi("마진율", f"{totals['마진율(%)']:.2f}%", "순이익 ÷ 매출"), unsafe_allow_html=True)

st.write("")
c = st.columns(4)
rev = totals["총매출"] or 1
for col, (label, key) in zip(c, [("수수료", "수수료"), ("배송비", "배송비"),
                                  ("기타비용", "기타비용"), ("원가", "원가")]):
    col.markdown(kpi(label, f"{totals[key]:,.0f}원", f"매출 대비 {totals[key] / rev * 100:.1f}%"),
                 unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 분석 리포트
# ----------------------------------------------------------------------
section("Analysis")
st.caption("업로드된 데이터의 수치를 규칙에 따라 해석한 분석입니다.")
for title, text in insights:
    st.markdown(f'<div class="insight"><div class="insight-title">{title}</div>'
                f'<div class="insight-text">{text}</div></div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 차트 (위에서 아래로 세로 적층, 가로 막대)
# ----------------------------------------------------------------------
section("Charts")

st.markdown('<div class="card-title">플랫폼별 매출 · 순이익</div>', unsafe_allow_html=True)
st.altair_chart(grouped_bar(plat, "플랫폼", ["매출액", "순이익"], [C_SALES, C_PROFIT]),
                use_container_width=True)

st.markdown('<div class="card-title">플랫폼별 마진율 (%)</div>', unsafe_allow_html=True)
st.altair_chart(single_bar(plat, "플랫폼", "마진율(%)", C_PROFIT, fmt=".2f", signed_color=True),
                use_container_width=True)

st.markdown('<div class="card-title">플랫폼별 비용 구조</div>', unsafe_allow_html=True)
st.altair_chart(cost_stack(plat), use_container_width=True)

if not monthly.empty:
    st.markdown('<div class="card-title">월별 순이익 추이</div>', unsafe_allow_html=True)
    st.altair_chart(monthly_line(monthly), use_container_width=True)

st.markdown('<div class="card-title">순이익 상위 상품 TOP 15</div>', unsafe_allow_html=True)
top_prod = prod.head(15).copy()
top_prod["라벨"] = top_prod.apply(lambda r: f"{r['플랫폼']} · {str(r['상품명'])[:26]}", axis=1)
st.altair_chart(single_bar(top_prod, "라벨", "순이익", C_PROFIT, signed_color=True),
                use_container_width=True)

# ----------------------------------------------------------------------
# 플랫폼별 상세 리포트
# ----------------------------------------------------------------------
section("Platform Report")
st.dataframe(
    plat,
    hide_index=True,
    use_container_width=True,
    column_config={**MONEY_CFG(["객단가", "매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]),
                   **RATE_CFG(["수수료율(%)", "마진율(%)", "매출비중(%)", "순이익비중(%)"])},
)

# ----------------------------------------------------------------------
# 플랫폼 드릴다운
# ----------------------------------------------------------------------
section("Drill Down")
choice = st.selectbox("플랫폼 선택", ["전체"] + plat["플랫폼"].tolist())
view = detail if choice == "전체" else detail[detail["플랫폼"] == choice]
v = make_totals(view)
d = st.columns(4)
d[0].markdown(kpi("매출", f"{v['총매출']:,.0f}원", f"{v['주문건수']:,}건"), unsafe_allow_html=True)
d[1].markdown(kpi("비용", f"{v['총비용']:,.0f}원"), unsafe_allow_html=True)
d[2].markdown(kpi("순수익", f"{v['순이익']:,.0f}원", accent=True), unsafe_allow_html=True)
d[3].markdown(kpi("마진율", f"{v['마진율(%)']:.2f}%"), unsafe_allow_html=True)
st.write("")

# ----------------------------------------------------------------------
# 탭: 상품 / 월별 / 원본 계산 내역
# ----------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["상품별 수익", "월별 추이", "마진 계산 내역"])
with tab1:
    view_prod = make_product_report(view)
    st.caption(f"{choice} 기준 상품 {len(view_prod):,}개 · 순이익 높은 순")
    st.dataframe(view_prod, hide_index=True, use_container_width=True,
                 column_config={**MONEY_CFG(["매출액", "수수료", "원가", "순이익"]),
                                **RATE_CFG(["마진율(%)"])})
with tab2:
    if monthly.empty:
        st.info("정산일을 날짜로 인식하지 못해 월별 추이를 만들 수 없습니다.")
    else:
        mv = monthly if choice == "전체" else monthly[monthly["플랫폼"] == choice]
        st.dataframe(mv, hide_index=True, use_container_width=True,
                     column_config={**MONEY_CFG(["매출액", "순이익"]), **RATE
