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

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&display=swap');
@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
html, body, [class*="css"], .stApp { font-family: 'Pretendard','Noto Sans KR',-apple-system,'Segoe UI',sans-serif; }
.stApp { background: __BG__; color: __TEXT__; }
.block-container { padding-top: 2rem; padding-bottom: 4rem; max-width: 1240px; }
header[data-testid="stHeader"] { background: transparent; }
.topbar { display:flex; justify-content:space-between; border-top:1px solid __LINE__;
          border-bottom:1px solid __LINE__; padding:12px 4px; font-size:0.8rem;
          letter-spacing:0.12em; text-transform:uppercase; color:__TEXT_SUB__; }
.brand { font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.1rem;
         letter-spacing:0.2em; font-weight:600; color:__TEXT__; }
.hero { text-align:center; padding:44px 0 36px 0; border-bottom:1px solid __LINE__; margin-bottom:28px; }
.hero h1 { font-family:'Playfair Display','Nanum Myeongjo',serif; font-weight:400; font-size:2.5rem;
           letter-spacing:0.08em; text-transform:uppercase; margin:0; color:__TEXT__; }
.hero p { color:__TEXT_SUB__; margin:14px 0 0 0; }
.section-label { font-size:0.78rem; letter-spacing:0.18em; text-transform:uppercase; color:__TEXT__;
                 font-weight:700; margin:40px 0 16px 0; padding-bottom:8px; border-bottom:1px solid __LINE__; }
.kpi { background:__SURFACE__; border:1px solid __LINE__; padding:20px 22px; min-height:128px; }
.kpi-label { font-size:0.74rem; letter-spacing:0.14em; text-transform:uppercase; color:__TEXT_SUB__; font-weight:600; }
.kpi-value { font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.9rem; color:__TEXT__; margin-top:10px; line-height:1.1; }
.kpi-sub { font-size:0.84rem; color:__TEXT_SUB__; margin-top:10px; }
.kpi-accent { background:__TEXT__; border-color:__TEXT__; }
.kpi-accent .kpi-label, .kpi-accent .kpi-sub { color:#3A3F4B; }
.kpi-accent .kpi-value { color:__BG__; }
.card-title { font-size:0.8rem; letter-spacing:0.14em; text-transform:uppercase; font-weight:700; color:__TEXT__; margin:18px 0 8px 0; }
.insight { background:__SURFACE__; border:1px solid __LINE__; border-left:3px solid __TEXT__; padding:16px 20px; margin-bottom:10px; }
.insight-title { font-weight:700; color:__TEXT__; margin-bottom:6px; }
.insight-text { color:__TEXT_SUB__; line-height:1.7; }
.alert-warn { border-left-color:__BAD__ !important; }
div[data-testid="stFileUploader"] section { background:__SURFACE__; border:1px dashed __TEXT_SUB__; border-radius:0; }
button[data-baseweb="tab"] { font-weight:600; color:__TEXT_SUB__ !important; }
button[data-baseweb="tab"][aria-selected="true"] { color:__TEXT__ !important; }
div[data-baseweb="tab-highlight"] { background-color:__TEXT__ !important; }
div[data-baseweb="tab-border"] { background-color:__LINE__ !important; }
[data-testid="stDownloadButton"] button, .stButton > button { border-radius:0; font-weight:700; background:__TEXT__;
        color:__BG__ !important; border:1px solid __TEXT__; padding:0.7rem 1.4rem; }
section[data-testid="stSidebar"] { background:__SURFACE__; border-right:1px solid __LINE__; }
div[data-testid="stDataFrame"] { border:1px solid __LINE__; }
.empty { border:1px solid __LINE__; padding:64px 24px; text-align:center; background:__SURFACE__; }
.empty-title { font-family:'Playfair Display','Nanum Myeongjo',serif; font-size:1.6rem; margin-top:10px; color:__TEXT__; }
.empty-desc { color:__TEXT_SUB__; margin-top:10px; line-height:1.7; }
</style>
"""
CSS = (CSS.replace("__BG__", BG).replace("__SURFACE__", SURFACE).replace("__LINE__", LINE)
          .replace("__TEXT_SUB__", TEXT_SUB).replace("__TEXT__", TEXT).replace("__BAD__", C_BAD))
st.markdown(CSS, unsafe_allow_html=True)

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
# 샘플 데이터
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


# ----------------------------------------------------------------------
# 원가표 (상품별 개당 원가)
# ----------------------------------------------------------------------
def table_to_map(df: pd.DataFrame) -> dict:
    out = {}
    if df is None or df.empty or "상품명" not in df.columns or "개당원가" not in df.columns:
        return out
    for _, r in df.iterrows():
        name = str(r["상품명"]).strip() if pd.notna(r["상품명"]) else ""
        try:
            v = float(r["개당원가"])
        except (TypeError, ValueError):
            continue
        if name and name.lower() != "nan" and not np.isnan(v):
            out[name] = v
    return out


def parse_cost_csv(raw: bytes) -> dict:
    df = None
    for enc in ("utf-8-sig", "cp949"):
        try:
            df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding=enc)
            break
        except UnicodeDecodeError:
            continue
    if df is None or not {"상품명", "개당원가"} <= set(df.columns):
        return {}
    df["개당원가"] = to_number(df["개당원가"])
    return table_to_map(df)


def build_frame(p: dict, m: dict, cost_rate: float, cost_map: dict) -> pd.DataFrame:
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

    if "cost" in m:
        out["원가"] = num("cost")
    else:
        unit = out["상품명"].map(lambda x: cost_map.get(str(x).strip(), np.nan)).astype(float)
        qty = out["수량"].where(out["수량"] > 0, 1.0)
        out["원가"] = np.where(unit.notna(), unit.fillna(0) * qty, out["매출액"] * cost_rate / 100)

    cells = out[["플랫폼", "정산일", "상품명"]].astype(str).apply(
        lambda c: c.str.strip().str.lower().str.match(TOTAL_RE))
    out = out[~cells.any(axis=1)]
    out = out[(out["매출액"] != 0) | (out["수수료"] != 0)]

    out["순이익"] = out["매출액"] - out["수수료"] - out["배송비"] - out["기타비용"] - out["원가"]
    out["마진율(%)"] = np.where(out["매출액"] > 0, out["순이익"] / out["매출액"] * 100, 0).round(2)
    out["원본파일"] = p["name"]
    return out


def process_all(parsed: list, overrides: dict, cost_rate: float, cost_map: dict):
    frames, errors, logs = [], [], []
    for p in parsed:
        if p.get("error"):
            errors.append((p["name"], p["error"]))
            continue
        m = overrides.get(p["name"], p["mapping"])
        if "sales" not in m:
            errors.append((p["name"], "매출액 컬럼이 지정되지 않았습니다. 매핑에서 선택해 주세요."))
            continue
        for key in FIELD_ORDER:
            if key in m:
                method = "자동" if p["mapping"].get(key) == m[key] else "수동"
                if key == "sales" and p.get("guessed") and method == "자동":
                    method = "값 기반 추정"
                logs.append({"파일": p["name"], "표준항목": STANDARD_LABEL[key],
                             "인식된 컬럼": m[key], "방식": method})
        frames.append(build_frame(p, m, cost_rate, cost_map))
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


def make_breakeven(prod: pd.DataFrame, fixed_cost: float, target_profit: float) -> pd.DataFrame:
    """개당 공헌이익 = 상품 순이익 / 수량 (원가·수수료·배송·기타 차감 후 값)."""
    if prod.empty:
        return pd.DataFrame()
    g = prod.copy()
    g["개당공헌이익"] = safe_div(g["순이익"], g["수량"].replace(0, np.nan))
    g["개당평균매출"] = safe_div(g["매출액"], g["수량"].replace(0, np.nan))
    def need(target):
        return np.where(g["개당공헌이익"] > 0, np.ceil(target / g["개당공헌이익"].replace(0, np.nan)), np.nan)
    g["손익분기수량"] = need(fixed_cost)
    g["목표달성수량"] = need(fixed_cost + target_profit)
    cols = ["플랫폼", "상품명", "수량", "개당평균매출", "개당공헌이익", "손익분기수량", "목표달성수량"]
    return g[cols].sort_values("개당공헌이익", ascending=False).reset_index(drop=True).round(0).astype(
        {"수량": "float"}, errors="ignore")


def compare_periods(detail: pd.DataFrame, a: str, b: str) -> pd.DataFrame:
    rows = []
    for label, month in [("A", a), ("B", b)]:
        d = detail[detail["정산월"] == month]
        t = make_totals(d) if not d.empty else {"총매출": 0, "순이익": 0, "마진율(%)": 0, "주문건수": 0}
        rows.append((month, t))
    (ma, ta), (mb, tb) = rows
    items = [("주문건수", "주문건수", "건"), ("총매출", "총매출", "원"),
             ("순이익", "순이익", "원"), ("마진율(%)", "마진율(%)", "%p")]
    out = []
    for label, key, unit in items:
        va, vb = float(ta[key]), float(tb[key])
        out.append({"지표": label, ma: va, mb: vb, "증감": vb - va,
                    "증감률(%)": round((vb - va) / va * 100, 2) if va else 0.0, "단위": unit})
    return pd.DataFrame(out)


def detect_anomalies(dv: pd.DataFrame, plat: pd.DataFrame, prod: pd.DataFrame, monthly: pd.DataFrame):
    alerts = []
    if not monthly.empty:
        for pf, g in monthly.groupby("플랫폼"):
            g = g.sort_values("정산월")
            if len(g) >= 2:
                last, prev = g.iloc[-1], g.iloc[-2]
                dm = last["마진율(%)"] - prev["마진율(%)"]
                if dm <= -5:
                    alerts.append(("경고", f"{pf} 마진율 급락",
                                   f"{prev['정산월']} {prev['마진율(%)']:.2f}% → {last['정산월']} {last['마진율(%)']:.2f}% ({dm:+.2f}%p)"))
                if prev["매출액"] > 0:
                    ds = (last["매출액"] - prev["매출액"]) / prev["매출액"] * 100
                    if ds <= -50:
                        alerts.append(("경고", f"{pf} 매출 급감",
                                       f"직전월 대비 {ds:.1f}% 감소 ({last['정산월']})"))
    if not plat.empty:
        overall = float(plat["수수료"].sum() / plat["매출액"].sum() * 100) if plat["매출액"].sum() else 0
        for _, r in plat.iterrows():
            if r["수수료율(%)"] > overall + 5:
                alerts.append(("주의", f"{r['플랫폼']} 수수료율 과다",
                               f"{r['수수료율(%)']:.2f}% (전체 평균 {overall:.2f}%)"))
    if not prod.empty:
        loss = prod[prod["순이익"] < 0]
        if not loss.empty:
            alerts.append(("주의", f"적자 상품 {len(loss)}개",
                           ", ".join(loss["상품명"].head(5).tolist())))
    if not dv.empty:
        no_cost = dv[dv["원가"] == 0]
        if len(no_cost) > 0 and len(no_cost) / len(dv) > 0.5:
            alerts.append(("주의", "원가 미입력 비율 높음",
                           f"전체 {len(dv):,}건 중 {len(no_cost):,}건의 원가가 0원입니다. 원가표를 입력해 주세요."))
    return alerts


def verify_fees(dv: pd.DataFrame, rates: dict, tolerance: float) -> pd.DataFrame:
    d = dv[dv["플랫폼"].map(lambda x: rates.get(x, 0) > 0)].copy()
    if d.empty:
        return pd.DataFrame()
    d["기대수수료"] = d.apply(lambda r: r["매출액"] * rates[r["플랫폼"]] / 100, axis=1)
    d["차이"] = d["수수료"] - d["기대수수료"]
    flagged = d[d["차이"].abs() > tolerance]
    return flagged[["플랫폼", "정산일", "상품명", "매출액", "수수료", "기대수수료", "차이"]].sort_values(
        "차이", ascending=False).reset_index(drop=True).round(0)


def build_insights(plat, prod, monthly, totals):
    items = [("전체 요약",
              f"총 매출 {totals['총매출']:,.0f}원, 순이익 {totals['순이익']:,.0f}원, "
              f"마진율 {totals['마진율(%)']:.2f}%입니다. 총 {totals['주문건수']:,}건의 주문을 집계했습니다.")]
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
        items.append(("비용 구조", f"전체 비용({cost_sum:,.0f}원) 중 {biggest}가 {share:.1f}%로 가장 큽니다."))
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


def build_share_text(totals, plat, insights, alerts) -> str:
    lines = ["[정산 요약]",
             f"총매출 {totals['총매출']:,.0f}원 / 총비용 {totals['총비용']:,.0f}원",
             f"순이익 {totals['순이익']:,.0f}원 / 마진율 {totals['마진율(%)']:.2f}%", "",
             "[플랫폼별 순이익]"]
    for _, r in plat.iterrows():
        lines.append(f"- {r['플랫폼']}: 순이익 {r['순이익']:,.0f}원 / 마진율 {r['마진율(%)']:.2f}%")
    if alerts:
        lines += ["", "[이상 징후]"] + [f"- {a[1]}: {a[2]}" for a in alerts]
    lines += ["", "[분석]"] + [f"- {t}: {x}" for t, x in insights]
    return "\n".join(lines)


# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
def style_chart(ch, height):
    return (ch.properties(height=height, width="container")
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
    labels = base.mark_text(align="left", dx=4, color=TEXT, fontSize=12).encode(text=alt.Text("값:Q", format=",.0f"))
    return style_chart(alt.layer(bars, labels), max(300, 90 * len(df) + 80))


def single_bar(df, cat, val, color, fmt=",.0f", signed_color=False):
    fill = (alt.condition(f"datum['{val}'] < 0", alt.value(C_BAD), alt.value(color))
            if signed_color else alt.value(color))
    base = alt.Chart(df).encode(
        y=alt.Y(f"{cat}:N", sort=None, title=None, axis=alt.Axis(labelLimit=280, labelFontSize=14)),
        x=alt.X(f"{val}:Q", title=None, axis=alt.Axis(format=fmt)),
        tooltip=[alt.Tooltip(f"{cat}:N"), alt.Tooltip(f"{val}:Q", format=fmt)],
    )
    bars = base.mark_bar(cornerRadiusEnd=3).encode(color=fill)
    labels = base.mark_text(align="left", dx=4, color=TEXT, fontSize=12).encode(text=alt.Text(f"{val}:Q", format=fmt))
    return style_chart(alt.layer(bars, labels), max(260, 64 * len(df) + 80))


def cost_stack(plat):
    long = plat[["플랫폼"] + list(COST_COLORS)].melt(id_vars="플랫폼", var_name="비용", value_name="금액")
    ch = alt.Chart(long).mark_bar().encode(
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
    if ws.title in ("분석코멘트", "이상징후"):
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 60 if col > 1 else 22
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
                cell.number_format = "0.00" if ("율" in header or "비중" in header) else "#,##0"


def build_excel(totals, plat, prod, monthly, detail, insights, mapping_log, alerts, be_df, cost_map) -> bytes:
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
        if be_df is not None and not be_df.empty:
            be_df.to_excel(w, sheet_name="손익분기", index=False)
        pd.DataFrame(alerts or [("정보", "이상 징후 없음", "")], columns=["등급", "항목", "내용"]).to_excel(
            w, sheet_name="이상징후", index=False)
        pd.DataFrame(insights, columns=["구분", "분석 내용"]).to_excel(w, sheet_name="분석코멘트", index=False)
        detail[DETAIL_COLS].to_excel(w, sheet_name="마진계산내역", index=False)
        if cost_map:
            pd.DataFrame(list(cost_map.items()), columns=["상품명", "개당원가"]).to_excel(
                w, sheet_name="원가표", index=False)
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
    return (f'<div class="{cls}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>{sub_html}</div>')


def section(title):
    st.markdown(f'<div class="section-label">{title}</div>', unsafe_allow_html=True)


def money_cfg(cols):
    return {c: st.column_config.NumberColumn(c, format="%d") for c in cols}


def rate_cfg(cols):
    return {c: st.column_config.NumberColumn(c, format="%.2f%%") for c in cols}


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Settings")
    cost_rate = st.slider("기본 원가율 (%)", 0, 100, 0, 1,
                          help="파일에 원가 컬럼이 없고 원가표에도 없을 때 매출액 대비로 적용됩니다.")
    st.divider()
    st.markdown("**상품별 원가표**")
    st.caption("상품명과 개당 원가를 입력하면 수량에 곱해 원가를 계산합니다.")
    if "cost_df" not in st.session_state:
        st.session_state.cost_df = pd.DataFrame({"상품명": pd.Series(dtype=str),
                                                 "개당원가": pd.Series(dtype=float)})
    cost_df = st.data_editor(st.session_state.cost_df, num_rows="dynamic",
                             hide_index=True, key="cost_editor")
    st.session_state.cost_df = cost_df
    cost_upload = st.file_uploader("원가표 불러오기 (CSV)", type=["csv"], key="cost_upload",
                                   help="열 이름: 상품명, 개당원가")
    cost_map = table_to_map(cost_df)
    if cost_upload is not None:
        cost_map.update(parse_cost_csv(cost_upload.getvalue()))
    if cost_map:
        st.download_button("원가표 CSV 내보내기",
                           data=pd.DataFrame(list(cost_map.items()), columns=["상품명", "개당원가"])
                           .to_csv(index=False).encode("utf-8-sig"),
                           file_name="원가표.csv", mime="text/csv")
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
    detail, errors, mapping_log = process_all(parsed, overrides, cost_rate, cost_map)

for name, msg in errors:
    st.error(f"{name} 처리 실패: {msg}")

if detail.empty:
    st.warning("처리 가능한 데이터가 없습니다. 위 매핑에서 매출액 컬럼을 지정해 주세요.")
    st.stop()

# ----------------------------------------------------------------------
# 기간 필터
# ----------------------------------------------------------------------
section("Period")
months = sorted([m for m in detail["정산월"].unique() if m != "미상"])
if "미상" in detail["정산월"].values:
    months.append("미상")
sel_months = st.multiselect("정산월 필터", months, default=months)
dv = detail[detail["정산월"].isin(sel_months)] if sel_months else detail.iloc[0:0]
if dv.empty:
    st.warning("선택한 정산월에 데이터가 없습니다. 정산월을 하나 이상 선택해 주세요.")
    st.stop()

plat = make_platform_report(dv)
prod = make_product_report(dv)
monthly = make_monthly(dv)
totals = make_totals(dv)
insights = build_insights(plat, prod, monthly, totals)
alerts = detect_anomalies(dv, plat, prod, monthly)

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
for col, key in zip(c, ["수수료", "배송비", "기타비용", "원가"]):
    col.markdown(kpi(key, f"{totals[key]:,.0f}원", f"매출 대비 {totals[key] / rev * 100:.1f}%"),
                 unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 이상 징후
# ----------------------------------------------------------------------
section("Alerts")
if alerts:
    for level, title, text in alerts:
        cls = "insight alert-warn" if level == "경고" else "insight"
        st.markdown(f'<div class="{cls}"><div class="insight-title">[{level}] {title}</div>'
                    f'<div class="insight-text">{text}</div></div>', unsafe_allow_html=True)
else:
    st.caption("감지된 이상 징후가 없습니다.")

# ----------------------------------------------------------------------
# 분석 리포트
# ----------------------------------------------------------------------
section("Analysis")
st.caption("업로드된 데이터의 수치를 규칙에 따라 해석한 분석입니다.")
for title, text in insights:
    st.markdown(f'<div class="insight"><div class="insight-title">{title}</div>'
                f'<div class="insight-text">{text}</div></div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 기간 비교
# ----------------------------------------------------------------------
section("Period Compare")
all_months = [m for m in sorted(detail["정산월"].unique()) if m != "미상"]
if len(all_months) >= 2:
    cc = st.columns(2)
    month_a = cc[0].selectbox("기준 월 (A)", all_months, index=len(all_months) - 2)
    month_b = cc[1].selectbox("비교 월 (B)", all_months, index=len(all_months) - 1)
    if month_a == month_b:
        st.caption("서로 다른 두 달을 선택하면 증감을 보여 드립니다.")
    else:
        cmp_df = compare_periods(detail, month_a, month_b)
        st.dataframe(cmp_df, hide_index=True, use_container_width=True,
                     column_config={**money_cfg([month_a, month_b, "증감"]),
                                    **rate_cfg(["증감률(%)"])})
else:
    st.caption("두 개 이상의 정산월이 있으면 기간 비교를 사용할 수 있습니다.")

# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
section("Charts")
st.markdown('<div class="card-title">플랫폼별 매출 · 순이익</div>', unsafe_allow_html=True)
st.altair_chart(grouped_bar(plat, "플랫폼", ["매출액", "순이익"], [C_SALES, C_PROFIT]), use_container_width=True)

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
st.altair_chart(single_bar(top_prod, "라벨", "순이익", C_PROFIT, signed_color=True), use_container_width=True)

# ----------------------------------------------------------------------
# 플랫폼 상세 리포트
# ----------------------------------------------------------------------
section("Platform Report")
st.dataframe(
    plat, hide_index=True, use_container_width=True,
    column_config={**money_cfg(["객단가", "매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]),
                   **rate_cfg(["수수료율(%)", "마진율(%)", "매출비중(%)", "순이익비중(%)"])},
)

# ----------------------------------------------------------------------
# 드릴다운
# ----------------------------------------------------------------------
section("Drill Down")
choice = st.selectbox("플랫폼 선택", ["전체"] + plat["플랫폼"].tolist())
view = dv if choice == "전체" else dv[dv["플랫폼"] == choice]
v = make_totals(view)
d = st.columns(4)
d[0].markdown(kpi("매출", f"{v['총매출']:,.0f}원", f"{v['주문건수']:,}건"), unsafe_allow_html=True)
d[1].markdown(kpi("비용", f"{v['총비용']:,.0f}원"), unsafe_allow_html=True)
d[2].markdown(kpi("순수익", f"{v['순이익']:,.0f}원", accent=True), unsafe_allow_html=True)
d[3].markdown(kpi("마진율", f"{v['마진율(%)']:.2f}%"), unsafe_allow_html=True)
st.write("")

# ----------------------------------------------------------------------
# 손익분기 / 목표
# ----------------------------------------------------------------------
section("Break-even")
bc = st.columns(2)
fixed_cost = bc[0].number_input("기간 고정비 (광고·인건비 등, 원)", min_value=0, value=0, step=10000)
target_profit = bc[1].number_input("목표 순이익 (원)", min_value=0, value=0, step=10000)
be_df = make_breakeven(make_product_report(view), float(fixed_cost), float(target_profit))
if be_df.empty:
    st.caption("표시할 상품이 없습니다.")
else:
    st.caption("개당 공헌이익 = 상품 순이익 ÷ 수량 (원가·수수료·배송·기타 차감 후). "
               "고정비와 목표를 입력하면 필요한 판매수량을 계산합니다.")
    st.dataframe(be_df.head(50), hide_index=True, use_container_width=True,
                 column_config={**money_cfg(["개당평균매출", "개당공헌이익"]),
                                **{c: st.column_config.NumberColumn(c, format="%d") for c in
                                   ["수량", "손익분기수량", "목표달성수량"]}})

# ----------------------------------------------------------------------
# 정산 검증
# ----------------------------------------------------------------------
section("Fee Verification")
st.caption("플랫폼별 약정 수수료율을 입력하면 실제 수수료와 차이가 큰 건을 찾아 보여 줍니다. 0%는 검증하지 않습니다.")
with st.expander("플랫폼별 기대 수수료율 입력", expanded=False):
    rate_inputs = {}
    rc = st.columns(3)
    for j, pf in enumerate(plat["플랫폼"].tolist()):
        rate_inputs[pf] = rc[j % 3].number_input(f"{pf} 기대 수수료율 (%)", min_value=0.0,
                                                 max_value=100.0, value=0.0, step=0.1,
                                                 key=f"rate_{pf}")
    tol = st.number_input("허용 오차 (원)", min_value=0, value=1000, step=100)
flagged = verify_fees(dv, rate_inputs, float(tol))
if not any(rate_inputs.values()):
    st.caption("기대 수수료율을 입력하면 결과가 표시됩니다.")
elif flagged.empty:
    st.caption("허용 오차를 벗어난 수수료 건이 없습니다.")
else:
    st.warning(f"허용 오차를 벗어난 건이 {len(flagged):,}건 있습니다. 차이가 큰 순서로 정렬했습니다.")
    st.dataframe(flagged, hide_index=True, use_container_width=True,
                 column_config={**money_cfg(["매출액", "수수료", "기대수수료", "차이"])})

# ----------------------------------------------------------------------
# 탭: 상품 / 월별 / 원본
# ----------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["상품별 수익", "월별 추이", "마진 계산 내역"])
with tab1:
    view_prod = make_product_report(view)
    st.caption(f"{choice} 기준 상품 {len(view_prod):,}개 · 순이익 높은 순")
    st.dataframe(view_prod, hide_index=True, use_container_width=True,
                 column_config={**money_cfg(["매출액", "수수료", "원가", "순이익"]), **rate_cfg(["마진율(%)"])})
with tab2:
    if monthly.empty:
        st.info("정산일을 날짜로 인식하지 못해 월별 추이를 만들 수 없습니다.")
    else:
        mv = monthly if choice == "전체" else monthly[monthly["플랫폼"] == choice]
        st.dataframe(mv, hide_index=True, use_container_width=True,
                     column_config={**money_cfg(["매출액", "순이익"]), **rate_cfg(["마진율(%)"])})
with tab3:
    st.dataframe(view[DETAIL_COLS], hide_index=True, use_container_width=True,
                 column_config={**money_cfg(["매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]),
                                **rate_cfg(["마진율(%)"])})

# ----------------------------------------------------------------------
# 공유 & 다운로드
# ----------------------------------------------------------------------
section("Share & Export")
share_text = build_share_text(totals, plat, insights, alerts)
st.markdown('<div class="card-title">요약 텍스트 (카카오톡·메일에 붙여넣기)</div>', unsafe_allow_html=True)
st.code(share_text, language=None)
st.download_button("요약 텍스트 다운로드 (.txt)", data=share_text.encode("utf-8"),
                   file_name=f"정산요약_{pd.Timestamp.now():%Y%m%d_%H%M}.txt", mime="text/plain")

st.markdown('<div class="card-title">종합 보고서 (엑셀)</div>', unsafe_allow_html=True)
st.caption("종합요약 · 플랫폼별상세 · 상품별상세 · 월별추이 · 손익분기 · 이상징후 · 분석코멘트 · "
           "마진계산내역 · 원가표 · 컬럼인식결과 시트로 구성됩니다.")
st.download_button(
    label="엑셀 보고서 다운로드 (.xlsx)",
    data=build_excel(totals, plat, prod, monthly, dv, insights, mapping_log, alerts, be_df, cost_map),
    file_name=f"정산보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
