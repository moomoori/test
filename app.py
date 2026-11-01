import csv
import io
import re
from difflib import SequenceMatcher

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Settlement Report", page_icon="▫️", layout="wide")

BG = "#0E0F12"
SURFACE = "#16181D"
LINE = "#3A3F4B"
TEXT = "#F5F5F5"
TEXT_SUB = "#B8BDC7"

st.markdown(f"""
<style>
.stApp {{ background:{BG}; color:{TEXT}; }}
.block-container {{ padding-top:2rem; max-width:1200px; }}
header[data-testid="stHeader"] {{ background:transparent; }}
.hero {{ text-align:center; padding:40px 0 32px; border-top:1px solid {LINE};
         border-bottom:1px solid {LINE}; margin-bottom:28px; }}
.hero h1 {{ font-weight:400; letter-spacing:.08em; text-transform:uppercase; margin:0;
            color:{TEXT}; font-size:2.2rem; }}
.hero p {{ color:{TEXT_SUB}; margin:12px 0 0; }}
.section-label {{ font-size:.78rem; letter-spacing:.18em; text-transform:uppercase; font-weight:700;
                  color:{TEXT}; margin:32px 0 14px; padding-bottom:8px; border-bottom:1px solid {LINE}; }}
.kpi {{ background:{SURFACE}; border:1px solid {LINE}; padding:20px; min-height:130px; }}
.kpi-label {{ font-size:.74rem; letter-spacing:.14em; text-transform:uppercase; color:{TEXT_SUB}; font-weight:600; }}
.kpi-value {{ font-size:1.9rem; color:{TEXT}; margin-top:12px; font-weight:700; }}
.kpi-sub {{ font-size:.84rem; color:{TEXT_SUB}; margin-top:8px; }}
.kpi-accent {{ background:{TEXT}; border-color:{TEXT}; }}
.kpi-accent .kpi-label, .kpi-accent .kpi-sub {{ color:#3A3F4B; }}
.kpi-accent .kpi-value {{ color:{BG}; }}
div[data-testid="stDataFrame"] {{ border:1px solid {LINE}; }}
[data-testid="stDownloadButton"] button {{ border-radius:0; background:{TEXT}; color:{BG} !important;
    font-weight:700; border:1px solid {TEXT}; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 설정: 컬럼 별칭 사전 (국내 · 해외 쇼핑몰 표기 변형 포함)
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
              "sales", "revenue", "amount", "price", "gmv", "grosssales", "itemprice",
              "paidamount", "orderamount"],
    "commission": ["수수료", "판매수수료", "서비스이용료", "중개수수료", "수수료합계", "판매자수수료",
                   "마켓수수료", "이용료", "commission", "fee", "fees", "commissionfee"],
    "shipping": ["배송비", "배송비용", "택배비", "배송료", "배송수수료", "운송비", "배송비합계",
                 "shipping", "delivery", "shippingfee", "deliveryfee"],
    "other": ["기타비용", "기타수수료", "광고비", "포장비", "기타", "쿠폰비", "프로모션비용",
              "other", "adcost", "adfee"],
    "cost": ["원가", "매입원가", "매입가", "공급가", "공급가액", "상품원가", "구매원가",
             "cost", "purchasecost"],
}

# 이 단어가 포함된 컬럼은 해당 항목으로 잡지 않음
NEGATIVE_WORDS = {
    "sales": ["원가", "수수료", "배송", "순이익", "이익", "율", "비율", "할인", "쿠폰",
              "환불", "취소", "적립", "부가세", "세액", "차감"],
    "commission": ["순이익", "율", "비율"],
    "shipping": ["율", "비율"],
    "other": ["율", "비율"],
    "cost": ["율", "비율"],
}

FIELD_ORDER = ["sales", "commission", "shipping", "other", "cost", "quantity",
               "product", "date", "platform"]
STANDARD_LABEL = {
    "platform": "플랫폼", "date": "정산일", "product": "상품명", "quantity": "수량",
    "sales": "매출액", "commission": "수수료", "shipping": "배송비",
    "other": "기타비용", "cost": "원가",
}
MONEY_HINTS = ["금액", "가", "액", "원", "price", "amount", "sales", "매출", "결제", "판매"]
SKIP_HINTS = ["수량", "qty", "율", "비율", "일", "date", "번호", "id", "no"]
MATCH_THRESHOLD = 0.7
TOTAL_RE = re.compile(r"^(합계|소계|총계|total|sum)(\s|$)", re.IGNORECASE)
DETAIL_COLS = ["플랫폼", "정산일", "상품명", "수량", "매출액", "수수료", "배송비",
               "기타비용", "원가", "순이익", "마진율(%)", "원본파일"]


# ----------------------------------------------------------------------
# 정규화 · 매칭
# ----------------------------------------------------------------------
def normalize(text) -> str:
    s = re.sub(r"\(.*?\)|\[.*?\]|<.*?>", "", str(text))
    return re.sub(r"[\s_\-/·:.,*]", "", s).lower()


def score_column(c: str, a: str) -> float:
    if not c or not a:
        return 0.0
    if c == a:
        return 1.0
    short, long_ = sorted([c, a], key=len)
    if len(short) >= 2 and short in long_:
        return 0.7 + 0.25 * len(short) / len(long_)
    r = SequenceMatcher(None, c, a).ratio()
    return r if r >= 0.8 else 0.0


def is_negative(key: str, cn: str) -> bool:
    return any(w in cn for w in NEGATIVE_WORDS.get(key, []))


def match_columns(names):
    """헤더 이름 목록 -> ({항목: 원본컬럼}, {항목: 일치도})"""
    cands = []
    for col in names:
        cn = normalize(col)
        if not cn or cn.startswith("unnamed") or cn.startswith("컬럼"):
            continue
        for key, aliases in COLUMN_ALIASES.items():
            if is_negative(key, cn):
                continue
            best = max(score_column(cn, normalize(a)) for a in aliases)
            if best >= MATCH_THRESHOLD:
                cands.append((best, key, col))
    cands.sort(key=lambda x: x[0], reverse=True)
    mapping, scores, used = {}, {}, set()
    for s, k, c in cands:
        if k in mapping or c in used:
            continue
        mapping[k], scores[k] = c, s
        used.add(c)
    return mapping, scores


# ----------------------------------------------------------------------
# 숫자 변환 · 값 기반 추정
# ----------------------------------------------------------------------
def to_number(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    s = s.str.replace(r"[,\s원₩$₩€¥]", "", regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def numeric_share(series: pd.Series) -> float:
    s = series.astype(str).str.strip()
    s = s[(s != "") & (s.str.lower() != "nan")]
    if s.empty:
        return 0.0
    return float(to_number(s).ne(0).mean() if False else
                 pd.to_numeric(s.str.replace(r"[,\s원₩$€¥]", "", regex=True),
                               errors="coerce").notna().mean())


def guess_sales_by_value(df: pd.DataFrame, used: set):
    """이름으로 못 찾았을 때: 금액처럼 보이는 숫자 컬럼 중 합계가 가장 큰 것을 추정."""
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


# ----------------------------------------------------------------------
# 파일 읽기 · 헤더 탐색
# ----------------------------------------------------------------------
def read_grid(raw: bytes, name: str) -> pd.DataFrame:
    """헤더 판단 없이 전체 표를 문자열 격자로 읽는다."""
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

    raw_names = [("" if pd.isna(v) else str(v).strip())
                 for v in grid.iloc[header_row].tolist()]
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


# ----------------------------------------------------------------------
# 표준 데이터 생성 · 계산
# ----------------------------------------------------------------------
def is_total_row(row) -> bool:
    return any(TOTAL_RE.match(str(v).strip()) for v in row if str(v).strip())


def build_frame(p: dict, m: dict, cost_rate: float) -> pd.DataFrame:
    df = p["df"]
    df = df[~df.apply(is_total_row, axis=1)]
    idx = df.index

    def num(key):
        return to_number(df[m[key]]) if key in m else pd.Series(0.0, index=idx)

    stem = p["name"].rsplit(".", 1)[0]
    out = pd.DataFrame(index=idx)
    out["플랫폼"] = df[m["platform"]].fillna(stem).astype(str) if "platform" in m else stem
    out["정산일"] = df[m["date"]].fillna("").astype(str) if "date" in m else ""
    out["상품명"] = df[m["product"]].fillna("").astype(str) if "product" in m else ""
    out["수량"] = num("quantity")
    out["매출액"] = num("sales")
    out["수수료"] = num("commission")
    out["배송비"] = num("shipping")
    out["기타비용"] = num("other")
    out["원가"] = num("cost") if "cost" in m else out["매출액"] * cost_rate / 100

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


def make_summary(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame()
    g = detail.groupby("플랫폼", as_index=False).agg(
        주문건수=("매출액", "size"), 매출액=("매출액", "sum"), 수수료=("수수료", "sum"),
        배송비=("배송비", "sum"), 기타비용=("기타비용", "sum"), 원가=("원가", "sum"),
        순이익=("순이익", "sum"))
    g["마진율(%)"] = np.where(g["매출액"] > 0, g["순이익"] / g["매출액"] * 100, 0).round(2)
    return g.sort_values("매출액", ascending=False).reset_index(drop=True)


def make_totals(detail: pd.DataFrame) -> dict:
    sales = float(detail["매출액"].sum())
    cost = float(detail[["수수료", "배송비", "기타비용", "원가"]].sum().sum())
    profit = float(detail["순이익"].sum())
    return {"주문건수": int(len(detail)), "총매출": sales, "총비용": cost, "순이익": profit,
            "마진율(%)": round(profit / sales * 100, 2) if sales > 0 else 0.0}


# ----------------------------------------------------------------------
# 엑셀 생성
# ----------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="111111")
HEADER_FONT = Font(bold=True, color="FFFFFF")


def style_sheet(ws):
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    for col_idx, hcell in enumerate(ws[1], start=1):
        header = str(hcell.value or "")
        values = [str(ws.cell(r, col_idx).value or "") for r in range(1, ws.max_row + 1)]
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(v) for v in values) * 1.6 + 4, 40)
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
        mapping_log.to_excel(writer, sheet_name="컬럼인식결과", index=False)
        for ws in writer.book.worksheets:
            style_sheet(ws)
    return buf.getvalue()


# ----------------------------------------------------------------------
# 사이드바
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Settings")
    cost_rate = st.slider("기본 원가율 (%)", 0, 100, 0, 1,
                          help="원가 컬럼이 없을 때 매출액 대비로 적용됩니다.")
    st.caption("원가 컬럼이 있는 파일은 그 값을 우선 사용합니다.")

# ----------------------------------------------------------------------
# 헤더 & 업로드
# ----------------------------------------------------------------------
st.markdown("""
<div class="hero">
    <h1>The Settlement Report</h1>
    <p>쇼핑몰 정산 파일을 올리면 매출과 순이익을 한 장으로 정리합니다.</p>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader("정산 파일 업로드", type=["csv", "xlsx"],
                            accept_multiple_files=True, label_visibility="collapsed")
if not uploaded:
    st.info("CSV 또는 XLSX 정산 파일을 끌어다 놓으세요. 여러 쇼핑몰 파일을 한 번에 올릴 수 있습니다.")
    st.stop()

# ----------------------------------------------------------------------
# 컬럼 매핑 (자동 인식 결과 확인 · 수동 수정)
# ----------------------------------------------------------------------
parsed = [parse_file(f) for f in uploaded]

st.markdown('<div class="section-label">Column Mapping</div>', unsafe_allow_html=True)
st.caption("파일마다 자동 인식된 컬럼을 보여 드립니다. 틀린 항목은 드롭다운에서 직접 고를 수 있습니다.")
overrides = {}
for i, p in enumerate(parsed):
    if p.get("error"):
        st.error(f"{p['name']}: {p['error']}")
        continue
    needs_review = ("sales" not in p["mapping"]) or p["guessed"]
    title = f"{p['name']}  ·  헤더 {p['header_row'] + 1}행  ·  자동 인식 {len(p['mapping'])}개"
    with st.expander(title, expanded=needs_review):
        st.caption("파일 상단 미리보기")
        st.dataframe(p["df"].head(5), hide_index=True)
        grid_cols = st.columns(3)
        chosen = {}
        for j, key in enumerate(FIELD_ORDER):
            options = ["(없음)"] + p["columns"]
            default = p["mapping"].get(key)
            idx = options.index(default) if default in options else 0
            pick = grid_cols[j % 3].selectbox(STANDARD_LABEL[key], options, index=idx,
                                              key=f"map_{i}_{key}")
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

summary = make_summary(detail)
totals = make_totals(detail)

# ----------------------------------------------------------------------
# 핵심 지표
# ----------------------------------------------------------------------
def kpi(label, value, sub="", accent=False):
    cls = "kpi kpi-accent" if accent else "kpi"
    return (f'<div class="{cls}"><div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>')


st.markdown('<div class="section-label">Overview</div>', unsafe_allow_html=True)
cols = st.columns(4)
cols[0].markdown(kpi("총 매출", f"{totals['총매출']:,.0f}원", f"{totals['주문건수']:,}건"),
                 unsafe_allow_html=True)
cols[1].markdown(kpi("총 비용", f"{totals['총비용']:,.0f}원", "수수료 · 배송 · 기타 · 원가"),
                 unsafe_allow_html=True)
cols[2].markdown(kpi("순수익", f"{totals['순이익']:,.0f}원", "매출 − 전체 비용", accent=True),
                 unsafe_allow_html=True)
cols[3].markdown(kpi("마진율", f"{totals['마진율(%)']:.2f}%", "순이익 ÷ 매출"),
                 unsafe_allow_html=True)

# ----------------------------------------------------------------------
# 차트
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">By Platform</div>', unsafe_allow_html=True)
left, right = st.columns([3, 2])
with left:
    st.markdown("**매출 · 순이익**")
    st.bar_chart(summary.set_index("플랫폼")[["매출액", "순이익"]], color=[TEXT, "#8A8F9A"], height=300)
with right:
    st.markdown("**마진율 (%)**")
    st.bar_chart(summary.set_index("플랫폼")[["마진율(%)"]], color=[TEXT], height=300)

# ----------------------------------------------------------------------
# 상세 테이블
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Details</div>', unsafe_allow_html=True)
money_cfg = {c: st.column_config.NumberColumn(c, format="%d") for c in
             ["매출액", "수수료", "배송비", "기타비용", "원가", "순이익"]}
rate_cfg = {"마진율(%)": st.column_config.NumberColumn("마진율(%)", format="%.2f%%")}

tab1, tab2, tab3, tab4 = st.tabs(["플랫폼별 요약", "종합 정산 내역", "마진 계산 내역", "컬럼 인식 결과"])
with tab1:
    st.dataframe(summary, hide_index=True, column_config={**money_cfg, **rate_cfg})
with tab2:
    st.caption(f"총 {len(detail):,}건 · 파일 {len(uploaded)}개")
    st.dataframe(detail[DETAIL_COLS], hide_index=True, column_config={**money_cfg, **rate_cfg})
with tab3:
    st.caption("순이익 = 매출액 − 수수료 − 배송비 − 기타비용 − 원가  /  마진율 = 순이익 ÷ 매출액 × 100")
    st.dataframe(detail[["플랫폼", "상품명", "매출액", "수수료", "배송비", "기타비용", "원가",
                         "순이익", "마진율(%)"]],
                 hide_index=True, column_config={**money_cfg, **rate_cfg})
with tab4:
    st.dataframe(mapping_log, hide_index=True)

# ----------------------------------------------------------------------
# 다운로드
# ----------------------------------------------------------------------
st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)
st.caption("종합요약 · 플랫폼별요약 · 마진계산내역 · 컬럼인식결과 4개 시트로 구성된 엑셀 파일입니다.")
st.download_button(
    label="엑셀 보고서 다운로드 (.xlsx)",
    data=build_excel(detail, summary, totals, mapping_log),
    file_name=f"정산보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
