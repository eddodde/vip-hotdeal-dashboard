import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components
import pathlib
import io
import re
import datetime as _dt

# ══════════════════════════════════════════════════════════════
# VIP 핫딜 트렌드 대시보드
#  - 데이터는 repo에 커밋된 깨끗한 CSV(data/*.csv)에서 로드.
#  - 원본 xlsx는 회사 DRM(Softcamp)으로 암호화돼 있어 그대로는 못 읽으므로,
#    convert.ps1(Excel COM)로 CSV를 재생성한 뒤 push 하는 워크플로.
# ══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="VIP 핫딜 트렌드 대시보드",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded",
)

DOW = ["월", "화", "수", "목", "금", "토", "일"]   # 요일 (Mon=0)

# 콘텐츠 영역(트래픽 비교용) & 색
AREAS = ["VIP핫딜", "VIP라운지", "기념일", "기획전영역"]
AREA_COLOR = {"VIP핫딜": "#E45756", "VIP라운지": "#4C72B0",
              "기념일": "#55A868", "기획전영역": "#B0B0B0"}
SLOT_COLOR = {"오전": "#4C72B0", "오후": "#DD8452"}
ACCENT = "#E45756"

# Plotly 한글 폰트 — 웹폰트를 강제 로드해 SVG 텍스트에 적용
KFONT = "'Noto Sans KR','Malgun Gothic','Apple SD Gothic Neo',sans-serif"

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');
    .metric-card {
        background: #f8f9fa; border-radius: 10px; padding: 14px 18px;
        border-left: 4px solid #E45756; margin-bottom: 10px;
    }
    .metric-label { font-size: 13px; color: #666; margin-bottom: 4px; }
    .metric-value { font-size: 24px; font-weight: 700; color: #1a1a2e; }
    .metric-sub  { font-size: 12px; color: #888; margin-top: 2px; }
    .section-title {
        font-size: 18px; font-weight: 700; color: #1a1a2e;
        margin: 26px 0 12px 0; padding-bottom: 6px;
        border-bottom: 2px solid #e9ecef; scroll-margin-top: 70px;
    }
    .hint { font-size: 12px; color: #999; margin: -4px 0 10px 0; }
    a.navlink {
        display: block; padding: 8px 12px; margin: 4px 0; border-radius: 8px;
        background: #fbf2f2; color: #C0392B; text-decoration: none;
        font-size: 14px; font-weight: 600; border: 1px solid #f2e0e0;
    }
    a.navlink:hover { background: #f8e3e3; color: #922B21; }
    .insight {
        background: #fdf3f2; border-left: 4px solid #E45756; border-radius: 8px;
        padding: 12px 16px; margin: 6px 0 14px 0; font-size: 14px; line-height: 1.6;
    }
    .insight.warn { background: #fdeeee; border-left-color: #C44E52; }
    .insight.ok   { background: #eef7f0; border-left-color: #55A868; }
    .insight b { color: #1a1a2e; }
</style>
""", unsafe_allow_html=True)


# ── 데이터 로드 ─────────────────────────────────────────────
DATA = pathlib.Path(__file__).parent / "data"
NUMCOLS = ["UV", "PV", "cust", "ord", "qty", "rev",
           "h_UV", "h_PV", "h_cust", "h_ord", "h_qty", "h_rev"]


def _isblank(x):
    return x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == ""


def _md(s):
    """'M/D' 또는 'YY/MM/DD' → (month, day, year|None)."""
    s = str(s).strip()
    m = re.match(r"^(\d+)/(\d+)/(\d+)", s)
    if m:
        return int(m.group(2)), int(m.group(3)), 2000 + int(m.group(1))
    m = re.match(r"^(\d+)/(\d+)", s)
    if m:
        return int(m.group(1)), int(m.group(2)), None
    return None


def parse_hotdeal_xlsx(raw):
    """해제된 핫딜.xlsx(그룹형) → tidy DataFrame. convert.ps1 과 동일 로직."""
    g = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=None).values
    n, ncol = g.shape
    tot = []   # [rowidx, m, d, year|None]
    for r in range(1, n):
        if _isblank(g[r, 0]):
            continue
        p = _md(g[r, 0])
        if p:
            tot.append([r, p[0], p[1], p[2]])
    if not tot:
        raise ValueError("핫딜 시트에서 일자 행을 찾지 못했습니다. 시트 구조를 확인하세요.")
    k = 0
    while k < len(tot) and tot[k][3] is None:
        k += 1
    for i in range(k - 1, -1, -1):           # 연도 역산(앞부분 M/D)
        ny = tot[i + 1][3]
        if tot[i][1] > tot[i + 1][1]:
            ny -= 1
        tot[i][3] = ny
    for i in range(k + 1, len(tot)):         # 잔여 None 순방향 보정
        if tot[i][3] is None:
            ny = tot[i - 1][3]
            if tot[i][1] < tot[i - 1][1]:
                ny += 1
            tot[i][3] = ny
    rowdate = {t[0]: f"{t[3]:04d}-{t[1]:02d}-{t[2]:02d}" for t in tot}

    def cell(r, c):
        return "" if (c >= ncol or _isblank(g[r, c])) else g[r, c]

    rows, cur = [], None
    for r in range(1, n):
        if r in rowdate:
            cur = rowdate[r]
        if _isblank(g[r, 2]):
            continue
        detail = str(g[r, 2]).strip()
        if detail == "Total":
            slot, rt = "Total", "TOTAL"
        elif "오전" in detail:
            slot, rt = "오전", "SLOT"
        elif "오후" in detail:
            slot, rt = "오후", "SLOT"
        else:
            slot, rt = detail, "SLOT"
        rows.append({
            "date": cur, "slot": slot, "row_type": rt,
            "prodcode": cell(r, 4), "prodname": cell(r, 5), "md": cell(r, 6),
            "bpu": cell(r, 7), "brand": cell(r, 8), "category": cell(r, 9),
            "UV": cell(r, 10), "PV": cell(r, 11), "cust": cell(r, 12),
            "ord": cell(r, 13), "qty": cell(r, 14), "rev": cell(r, 15),
            "h_UV": cell(r, 16), "h_PV": cell(r, 17), "h_cust": cell(r, 18),
            "h_ord": cell(r, 19), "h_qty": cell(r, 20), "h_rev": cell(r, 21),
        })
    return pd.DataFrame(rows)


def parse_table_xlsx(raw):
    """해제된 Table.xlsx(가로형) → tidy DataFrame(date,metric,area,value)."""
    g = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=None).values
    nrow, ncol = g.shape
    cols = list(range(2, ncol))

    def md_cell(x):
        if isinstance(x, (pd.Timestamp, _dt.datetime, _dt.date)):
            return x.month, x.day
        s = str(x).strip().split("/")
        return int(s[0]), int(s[1])

    md = [md_cell(g[0, c]) for c in cols]
    roll, prevm = 0, md[0][0]
    for m, _d in md:
        if m < prevm:
            roll += 1
        prevm = m
    end_year = pd.Timestamp.today().year     # 마지막 날짜 = 최근으로 가정
    years, y, prevm = [], end_year - roll, md[0][0]
    for m, _d in md:
        if m < prevm:
            y += 1
        years.append(y)
        prevm = m
    metric_of = {1: "UV", 2: "UV", 3: "UV", 4: "UV", 5: "PV", 6: "PV", 7: "PV", 8: "PV"}
    recs = []
    for r, metric in metric_of.items():
        if r >= nrow:
            continue
        area = "" if _isblank(g[r, 1]) else str(g[r, 1]).strip()
        for idx, c in enumerate(cols):
            dt = f"{years[idx]:04d}-{md[idx][0]:02d}-{md[idx][1]:02d}"
            recs.append({"date": dt, "metric": metric, "area": area, "value": g[r, c]})
    return pd.DataFrame(recs)


@st.cache_data(show_spinner=False)
def load_hotdeal(upload_bytes=None):
    """슬롯·상품 단위 일별 매출 상세. xlsx(해제본)·csv·기본 데이터 모두 지원."""
    if upload_bytes and upload_bytes[:2] == b"PK":       # 해제된 xlsx
        df = parse_hotdeal_xlsx(upload_bytes)
    else:
        src = io.BytesIO(upload_bytes) if upload_bytes else (DATA / "hotdeal.csv")
        df = pd.read_csv(src, dtype={"prodcode": str})
    if "date" not in df.columns or "row_type" not in df.columns:
        raise ValueError("핫딜 데이터 형식이 아닙니다(date, row_type 필요). "
                         "핫딜.xlsx(해제본) 또는 hotdeal.csv 인지 확인하세요.")
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df = df.dropna(subset=["date"])
    for c in NUMCOLS:
        df[c] = pd.to_numeric(df.get(c), errors="coerce").fillna(0)
    for c in ["brand", "category", "md", "prodname", "slot"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    # MD 표시용 이름(아이디 괄호 제거)
    df["md_name"] = df["md"].str.replace(r"\(.*\)", "", regex=True).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_table(upload_bytes=None):
    """콘텐츠 영역별 일별 UV/PV 장기 추세. xlsx(해제본)·csv·기본 데이터 모두 지원."""
    if upload_bytes and upload_bytes[:2] == b"PK":       # 해제된 xlsx
        df = parse_table_xlsx(upload_bytes)
    else:
        src = io.BytesIO(upload_bytes) if upload_bytes else (DATA / "table_trend.csv")
        df = pd.read_csv(src)
    if not {"date", "metric", "area", "value"}.issubset(df.columns):
        raise ValueError("트래픽 데이터 형식이 아닙니다(date, metric, area, value 필요).")
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["value"] = pd.to_numeric(df["value"], errors="coerce")
    return df.dropna(subset=["date"])


# ── 포맷·헬퍼 ───────────────────────────────────────────────
def fnum(x):
    return f"{int(round(x)):,}"


def fwon(x):
    """거래액(원) 축약 — 억/만 단위."""
    x = float(x)
    if abs(x) >= 1e8:
        return f"{x/1e8:.2f}억"
    if abs(x) >= 1e4:
        return f"{x/1e4:,.0f}만"
    return f"{x:,.0f}"


def fdelta(cur, prev):
    """직전 동일길이 기간 대비 증감률 HTML."""
    if prev in (0, None) or pd.isna(prev):
        return '<span style="color:#999">—</span>'
    chg = (cur - prev) / prev * 100
    if chg >= 0:
        return f'<span style="color:#2E7D32">▲ {chg:.1f}%</span>'
    return f'<span style="color:#C44E52">▼ {abs(chg):.1f}%</span>'


def metric_card(label, value, sub="", color=ACCENT):
    st.markdown(
        f'<div class="metric-card" style="border-left-color:{color}">'
        f'<div class="metric-label">{label}</div>'
        f'<div class="metric-value">{value}</div>'
        f'<div class="metric-sub">{sub}</div></div>',
        unsafe_allow_html=True)


def section(title, hint="", anchor=None):
    aid = f' id="{anchor}"' if anchor else ""
    st.markdown(f'<div class="section-title"{aid}>{title}</div>', unsafe_allow_html=True)
    if hint:
        st.markdown(f'<div class="hint">{hint}</div>', unsafe_allow_html=True)


def insight(html, kind=""):
    st.markdown(f'<div class="insight {kind}">{html}</div>', unsafe_allow_html=True)


def plot(fig, title=None, height=380):
    if title:
        st.markdown(f'<div style="font-weight:700;font-size:15px;margin:10px 0 -6px">{title}</div>',
                    unsafe_allow_html=True)
    fig.update_layout(font=dict(family=KFONT), height=height,
                      margin=dict(t=24, b=10, l=10, r=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0))
    st.plotly_chart(fig, use_container_width=True)


FREQ = {"일별": "D", "주별": "W-MON", "월별": "MS"}


def resample(s, freq, how="sum"):
    """일별 시계열을 집계단위로 리샘플. 데이터 공백은 NaN 유지."""
    s = s.sort_index()
    r = s.resample(FREQ[freq])
    out = r.sum() if how == "sum" else r.mean()
    if how == "sum":  # 원본에 데이터가 전혀 없는 구간은 0이 아니라 결측으로
        cnt = r.count()
        out = out.where(cnt > 0, np.nan)
    return out


def trend_word(s):
    """시계열 추세를 앞/뒤 구간 평균 비교로 판정."""
    s = s.dropna()
    if len(s) < 4:
        return "—", 0.0
    k = max(1, len(s) // 3)
    first, last = s.iloc[:k].mean(), s.iloc[-k:].mean()
    if first == 0:
        return ("증가" if last > 0 else "유지"), 0.0
    chg = (last - first) / first * 100
    word = "📈 증가" if chg > 5 else ("📉 감소" if chg < -5 else "→ 유지")
    return word, chg


def period_key(dates, gran):
    """일 인덱스 → 기간 라벨 Series. gran ∈ {연도, 월, 주차}."""
    if gran == "연도":
        return dates.year.astype(str)
    if gran == "월":
        return dates.to_period("M").astype(str)            # 2026-06
    iso = dates.isocalendar()
    return iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)


def avg_daily_rev(daily_tot, gran):
    """일별 거래액 합 → 기간별 '일평균 거래액'(운영일 평균) + 운영일수."""
    k = period_key(daily_tot.index, gran)
    g = daily_tot.groupby(k.values)
    out = pd.DataFrame({"일평균": g.mean(), "합계": g.sum(), "운영일": g.size()})
    return out.sort_index()


def daterange_stats(df, start, end):
    """[start, end] 구간의 (일평균 거래액, 운영일수, 합계)."""
    d = df[(df["date"] >= start) & (df["date"] <= end)]
    days = d["date"].nunique()
    rev = d["rev"].sum()
    return (rev / days if days else np.nan), days, rev


def pct(cur, prev):
    if prev in (0, None) or pd.isna(prev) or pd.isna(cur):
        return None
    return (cur - prev) / prev * 100


def cmp_card(title, label_cur, cur, label_prev, prev):
    """비교 카드: 현재 vs 비교 기간 일평균 + 증감%."""
    p = pct(cur, prev)
    if p is None:
        delta = '<span style="color:#999">비교 데이터 없음</span>'
    elif p >= 0:
        delta = f'<span style="color:#2E7D32">▲ {p:.1f}%</span>'
    else:
        delta = f'<span style="color:#C44E52">▼ {abs(p):.1f}%</span>'
    cv = "—" if pd.isna(cur) else fwon(cur)
    pv = "—" if (prev is None or pd.isna(prev)) else fwon(prev)
    metric_card(title, f"{cv}원 &nbsp;{delta}",
                f"{label_cur} 일평균 · {label_prev} {pv}원")


# ── 사이드바: 데이터 업로드 → 로드 ──────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    with st.expander("📤 데이터 올리기 (매일 갱신)", expanded=False):
        st.caption("**DRM 해제한** 핫딜.xlsx · Table.xlsx (또는 convert.ps1 로 만든 CSV)를 "
                   "올리면 git push 없이 바로 반영됩니다.")
        up_h = st.file_uploader("핫딜 매출 (xlsx / csv)", type=["xlsx", "csv"], key="up_h")
        up_t = st.file_uploader("트래픽 Table (xlsx / csv)", type=["xlsx", "csv"], key="up_t")


def _bytes(up):
    """업로드 검증: 아직 DRM 암호화 상태면 차단, 아니면 바이트 반환."""
    if up is None:
        return None
    raw = up.getvalue()
    if raw[:4] == b"SCDS":     # Softcamp DRM 미해제
        st.sidebar.error(f"'{up.name}' 은 아직 DRM 암호화 상태입니다. "
                         "암호화 해제(반출) 후 다시 올려 주세요.")
        st.stop()
    return raw


try:
    H = load_hotdeal(_bytes(up_h))
    T = load_table(_bytes(up_t))
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

src_note = "업로드 데이터" if (up_h or up_t) else "기본(커밋된) 데이터"
SLOTS = H[H["row_type"] == "SLOT"].copy()    # 슬롯·상품(매출 집계용, 가산 가능)
TOTD = (H[H["row_type"] == "TOTAL"]
        .set_index("date").sort_index())     # 일 단위 페이지 총계(UV/PV는 중복제거값)

# ── 헤더 ────────────────────────────────────────────────────
st.title("🔥 VIP 핫딜 트렌드 대시보드")
st.caption("VIP라운지 핫딜 콘텐츠의 매출·트래픽·전환을 기간별 추세로 진단합니다. "
           "데일리 실적은 따로 보시니, 여기선 '흐름'에 집중합니다.")

MENU = [
    ("sec-core", "🔑 핵심 요약"),
    ("sec-compare", "🔁 전년·전월·전주 비교"),
    ("sec-avg", "📊 일평균 거래액 (연/월/주)"),
    ("sec-sales", "💰 매출 추세 + 피크일"),
    ("sec-dow", "📅 요일별 분석"),
    ("sec-traffic", "🚦 트래픽 추세"),
    ("sec-conv", "🎯 전환·효율"),
    ("sec-best", "🏆 베스트 (상품·브랜드·MD)"),
    ("sec-mix", "🧩 브랜드·카테고리 믹스"),
    ("sec-slot", "⏰ 오전·오후 슬롯"),
    ("sec-table", "📋 상세 데이터"),
]

# ── 사이드바: 필터 ──────────────────────────────────────────
with st.sidebar:
    dmin = SLOTS["date"].min().date()
    dmax = SLOTS["date"].max().date()
    default_start = max(dmin, (pd.Timestamp(dmax) - pd.Timedelta(days=180)).date())
    dr = st.date_input("기간", value=(default_start, dmax),
                       min_value=dmin, max_value=dmax)
    d0, d1 = dr if isinstance(dr, tuple) and len(dr) == 2 else (default_start, dmax)
    freq = st.radio("집계 단위", list(FREQ.keys()), index=1, horizontal=True)
    sel_slots = st.multiselect("슬롯", ["오전", "오후"], default=["오전", "오후"])

    st.divider()
    st.markdown("**📂 분석 메뉴**")
    st.markdown("".join(f'<a href="#{a}" class="navlink">{lbl}</a>' for a, lbl in MENU),
                unsafe_allow_html=True)
    st.divider()
    st.caption(f"📦 {src_note} · {dmin} ~ {dmax} · 매출은 거래액(VAT 제외)")

# ── 필터 적용 ───────────────────────────────────────────────
d0 = pd.Timestamp(d0)
d1 = pd.Timestamp(d1)
mask = (SLOTS["date"] >= d0) & (SLOTS["date"] <= d1)
if sel_slots:
    mask &= SLOTS["slot"].isin(sel_slots)
FS = SLOTS[mask].copy()                       # 필터된 슬롯·상품 (매출)
td_mask = (TOTD.index >= d0) & (TOTD.index <= d1)
FT = TOTD[td_mask]                            # 필터된 일 총계 (트래픽·전환)

if FS.empty:
    st.warning("선택한 조건에 데이터가 없습니다. 기간/슬롯을 조정해 주세요.")
    st.stop()

period_days = (d1 - d0).days + 1
n_active = FS["date"].nunique()


def daily(df, col):
    return df.groupby("date")[col].sum()


# 직전 동일길이 기간(증감 비교용)
p_d1 = d0 - pd.Timedelta(days=1)
p_d0 = p_d1 - pd.Timedelta(days=period_days - 1)
pmask = (SLOTS["date"] >= p_d0) & (SLOTS["date"] <= p_d1)
if sel_slots:
    pmask &= SLOTS["slot"].isin(sel_slots)
PS = SLOTS[pmask]

# ════════════════════════════════════════════════════════════
# 0. 핵심 요약
# ════════════════════════════════════════════════════════════
section("핵심 요약",
        f"선택 기간 {d0.date()} ~ {d1.date()} ({period_days}일, 운영 {n_active}일) · "
        f"카드의 증감(▲▼)은 <b>바로 직전 같은 길이 기간</b>"
        f"({p_d0.date()} ~ {p_d1.date()}) 대비입니다",
        anchor="sec-core")

cur = dict(rev=FS["rev"].sum(), ord=FS["ord"].sum(), qty=FS["qty"].sum())
prev = dict(rev=PS["rev"].sum(), ord=PS["ord"].sum(), qty=PS["qty"].sum())
uv_sum = FT["UV"].sum()
conv = (cur["ord"] / uv_sum * 100) if uv_sum else 0
aov = (cur["rev"] / cur["ord"]) if cur["ord"] else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    metric_card("거래액 (VAT제외)", fwon(cur["rev"]),
                f"직전 대비 {fdelta(cur['rev'], prev['rev'])}")
with c2:
    metric_card("주문 건수", fnum(cur["ord"]),
                f"직전 대비 {fdelta(cur['ord'], prev['ord'])}")
with c3:
    metric_card("주문 수량", fnum(cur["qty"]),
                f"직전 대비 {fdelta(cur['qty'], prev['qty'])}")
with c4:
    metric_card("핫딜 UV", fnum(uv_sum), "페이지 방문(중복 제거 일합)")
with c5:
    metric_card("전환율 · 객단가", f"{conv:.2f}%",
                f"객단가 {fwon(aov)}원")

rev_d = resample(daily(FS, "rev"), freq)
word, chg = trend_word(rev_d)
top_brand = (FS.groupby("brand")["rev"].sum().sort_values(ascending=False))
top_brand = top_brand[top_brand.index != ""]
tb_name = top_brand.index[0] if len(top_brand) else "—"
tb_share = (top_brand.iloc[0] / FS["rev"].sum() * 100) if len(top_brand) and FS["rev"].sum() else 0
insight(
    f"선택 기간 거래액 <b>{fwon(cur['rev'])}원</b> · 주문 <b>{fnum(cur['ord'])}건</b>, "
    f"직전 동기간 대비 거래액 {fdelta(cur['rev'], prev['rev'])}. "
    f"{freq} 매출 추세는 <b>{word}</b>(기간 내 {chg:+.0f}%). "
    f"매출 1위 브랜드는 <b>{tb_name}</b>(거래액의 {tb_share:.0f}%). "
    f"전환율 {conv:.2f}%, 객단가 {fwon(aov)}원.",
    "warn" if chg < -5 else ("ok" if chg > 5 else ""))

# ════════════════════════════════════════════════════════════
#    전년·전월·전주 비교  (최신 데이터 기준, 일평균 거래액)
# ════════════════════════════════════════════════════════════
ref = SLOTS["date"].max()          # 전체 데이터 최신일 (필터 무관)
section("전년·전월·전주 비교",
        f"가장 최근 데이터({ref.date()}) 기준 · 기간 길이가 달라 <b>일평균 거래액</b>으로 비교합니다",
        anchor="sec-compare")

# 주 단위(ISO): 이번 주 vs 지난 주
wk_start = ref - pd.Timedelta(days=int(ref.weekday()))
cur_w = daterange_stats(SLOTS, wk_start, ref)
prev_w = daterange_stats(SLOTS, wk_start - pd.Timedelta(days=7), wk_start - pd.Timedelta(days=1))
# 월 단위: 이번 달 vs 지난 달 vs 작년 같은 달
m_start = ref.replace(day=1)
cur_m = daterange_stats(SLOTS, m_start, ref)
pm_end = m_start - pd.Timedelta(days=1)
prev_m = daterange_stats(SLOTS, pm_end.replace(day=1), pm_end)
ly_start = m_start.replace(year=m_start.year - 1)
ly_end = ly_start + (ref - m_start)          # 작년 같은 달의 같은 일수 구간
prev_y = daterange_stats(SLOTS, ly_start, ly_end)

cc1, cc2, cc3 = st.columns(3)
with cc1:
    cmp_card("전주 대비", f"이번 주({wk_start.date().strftime('%m/%d')}~)",
             cur_w[0], "지난 주", prev_w[0])
with cc2:
    cmp_card("전월 대비", f"{ref.month}월", cur_m[0], f"{pm_end.month}월", prev_m[0])
with cc3:
    cmp_card("전년 대비", f"{ref.year}.{ref.month:02d}", cur_m[0],
             f"{ly_start.year}.{ly_start.month:02d}", prev_y[0])

bits = []
if pct(cur_w[0], prev_w[0]) is not None:
    bits.append(f"전주 대비 일평균 {fwon(cur_w[0])}원 ({pct(cur_w[0], prev_w[0]):+.0f}%)")
if pct(cur_m[0], prev_m[0]) is not None:
    bits.append(f"전월 대비 {pct(cur_m[0], prev_m[0]):+.0f}%")
if pct(cur_m[0], prev_y[0]) is not None:
    bits.append(f"전년 동월 대비 {pct(cur_m[0], prev_y[0]):+.0f}%")
else:
    bits.append("전년 동월은 데이터 공백(2024 상반기 등)으로 비교 불가일 수 있음")
insight("최신 기준 " + " · ".join(bits) + ". "
        "이번 주/달은 <b>진행 중</b>이라 일평균(운영일 평균)으로 비교합니다.")

# ════════════════════════════════════════════════════════════
#    일평균 거래액 (연 / 월 / 주차)
# ════════════════════════════════════════════════════════════
section("일평균 거래액", "기간마다 운영일수가 다르므로 <b>운영일 1일당 평균 거래액</b>으로 봅니다 "
        "(선택 기간·슬롯 필터 적용)", anchor="sec-avg")

gran = st.radio("집계 단위", ["연도", "월", "주차"], index=1, horizontal=True, key="avg_gran")
ad = avg_daily_rev(daily(FS, "rev"), gran)
fig = go.Figure()
fig.add_bar(x=ad.index, y=ad["일평균"], marker_color=ACCENT,
            customdata=np.stack([ad["운영일"], ad["합계"]], axis=-1),
            hovertemplate="%{x}<br>일평균 %{y:,.0f}원<br>운영 %{customdata[0]}일"
                          "<br>합계 %{customdata[1]:,.0f}원<extra></extra>")
avg_ma = ad["일평균"].rolling(3, min_periods=1).mean()
fig.add_trace(go.Scatter(x=ad.index, y=avg_ma, name="추세(3기간 평균)",
                         line=dict(color="#922B21", width=1.5, dash="dash")))
fig.update_layout(yaxis_title="일평균 거래액(원)", xaxis_title=gran, showlegend=False)
plot(fig, f"{gran}별 일평균 거래액", height=380)

best_p = ad["일평균"].idxmax()
worst_p = ad["일평균"].idxmin()
insight(f"{gran} 중 일평균 거래액이 가장 높은 구간은 <b>{best_p}</b>"
        f"({fwon(ad.loc[best_p, '일평균'])}원/일), 가장 낮은 구간은 <b>{worst_p}</b>"
        f"({fwon(ad.loc[worst_p, '일평균'])}원/일)입니다.")

# ════════════════════════════════════════════════════════════
# 1. 매출 추세
# ════════════════════════════════════════════════════════════
section("매출 추세 + 피크일", f"{freq} 거래액·주문 추이 (7기간 이동평균 보조선)", anchor="sec-sales")

rev_s = resample(daily(FS, "rev"), freq)
ord_s = resample(daily(FS, "ord"), freq)
ma = rev_s.rolling(7, min_periods=2).mean()

fig = go.Figure()
fig.add_trace(go.Scatter(x=rev_s.index, y=rev_s.values, name="거래액",
                         mode="lines+markers", line=dict(color=ACCENT, width=2.4),
                         marker=dict(size=4)))
fig.add_trace(go.Scatter(x=ma.index, y=ma.values, name="거래액 이동평균",
                         line=dict(color="#922B21", width=1.5, dash="dash"), opacity=0.7))
fig.add_trace(go.Scatter(x=ord_s.index, y=ord_s.values, name="주문 건수",
                         yaxis="y2", line=dict(color="#4C72B0", width=1.5, dash="dot")))
fig.update_layout(
    yaxis=dict(title="거래액(원)"),
    yaxis2=dict(title="주문", overlaying="y", side="right", showgrid=False))
plot(fig, height=420)

# 연도(YoY) 비교 — 월별 거래액
yoy = SLOTS.copy()
if sel_slots:
    yoy = yoy[yoy["slot"].isin(sel_slots)]
yoy["year"] = yoy["date"].dt.year
ym = (yoy.groupby(["year", yoy["date"].dt.month])["rev"].sum()
      .rename_axis(["year", "month"]).reset_index())
ym["year"] = ym["year"].astype(str)   # 이산 색상
figy = px.line(ym, x="month", y="rev", color="year", markers=True,
               labels={"month": "월", "rev": "거래액(원)", "year": "연도"},
               color_discrete_sequence=px.colors.qualitative.Set2)
figy.update_xaxes(dtick=1)
plot(figy, "연도별 월간 거래액 비교 (YoY)", height=360)
insight("2023년 전체 → 2024년은 <b>7월부터</b>만 데이터가 있습니다(상반기 공백). "
        "공백 구간은 선이 끊겨 표시됩니다.", "")

# 매출 피크일 — 그날을 견인한 브랜드·상품
st.markdown('<div style="font-weight:700;font-size:15px;margin:14px 0 4px">'
            '🔝 거래액 피크일 TOP 12 — 그날 1등 상품</div>', unsafe_allow_html=True)
day_tot = daily(FS, "rev").sort_values(ascending=False)
peak_recs = []
for dt, tot in day_tot.head(12).items():
    drows = FS[FS["date"] == dt].sort_values("rev", ascending=False)
    top = drows.iloc[0]
    peak_recs.append({
        "일자": dt.date(), "요일": DOW[dt.weekday()],
        "당일 거래액": round(tot), "주요 슬롯": top["slot"],
        "브랜드": top["brand"], "상품": top["prodname"][:30],
        "상품 거래액": round(top["rev"]),
        "비중": f"{top['rev']/tot*100:.0f}%" if tot else "—",
    })
peak_df = pd.DataFrame(peak_recs)
st.dataframe(peak_df, use_container_width=True, height=320,
             column_config={
                 "당일 거래액": st.column_config.NumberColumn(format="%d"),
                 "상품 거래액": st.column_config.NumberColumn(format="%d")})
st.caption("‘당일 거래액’이 튀는 날 어떤 브랜드의 어떤 상품이 견인했는지 — "
           "비중이 높을수록 한 상품이 그날 매출을 좌우한 날입니다.")

# ════════════════════════════════════════════════════════════
#    요일별 분석 (상품 배치 전략 지원)
# ════════════════════════════════════════════════════════════
section("요일별 분석",
        "요일별 <b>일평균</b> 거래액·UV — 거래액 높은 요일에 저조 상품을 배치해 부스팅하는 전략 참고용 "
        "(선택 기간·슬롯 필터 적용)", anchor="sec-dow")

day_rev = daily(FS, "rev")                       # 일별 거래액 합
dow_rev = day_rev.groupby(day_rev.index.weekday).mean().reindex(range(7))
dow_uv = FT["UV"].groupby(FT.index.weekday).mean().reindex(range(7))
dow_ord = daily(FS, "ord").groupby(daily(FS, "ord").index.weekday).mean().reindex(range(7))

d1c, d2c = st.columns(2)
with d1c:
    fig = go.Figure(go.Bar(x=DOW, y=dow_rev.values, marker_color=ACCENT,
                           text=[fwon(v) if pd.notna(v) else "" for v in dow_rev.values],
                           textposition="outside"))
    fig.update_layout(yaxis_title="일평균 거래액(원)")
    plot(fig, "요일별 일평균 거래액", height=340)
with d2c:
    fig = go.Figure(go.Bar(x=DOW, y=dow_uv.values, marker_color="#4C72B0",
                           text=[fnum(v) if pd.notna(v) else "" for v in dow_uv.values],
                           textposition="outside"))
    fig.update_layout(yaxis_title="일평균 UV")
    plot(fig, "요일별 일평균 UV(방문)", height=340)

dow_tbl = pd.DataFrame({
    "요일": DOW,
    "일평균 거래액": dow_rev.values,
    "일평균 UV": dow_uv.values,
    "일평균 주문": dow_ord.values,
    "전환율(%)": (dow_ord.values / dow_uv.values * 100),
}).round({"일평균 거래액": 0, "일평균 UV": 0, "일평균 주문": 1, "전환율(%)": 2})
st.dataframe(dow_tbl, use_container_width=True, hide_index=True)

if dow_rev.notna().any():
    hi = int(dow_rev.idxmax())
    hu = int(dow_uv.idxmax()) if dow_uv.notna().any() else hi
    insight(f"거래액이 가장 높은 요일은 <b>{DOW[hi]}요일</b>"
            f"({fwon(dow_rev.iloc[hi])}원/일), 방문(UV)이 가장 많은 요일은 "
            f"<b>{DOW[hu]}요일</b>입니다. 전략대로라면 <b>{DOW[hi]}요일</b> 같은 고매출 요일에 "
            "평소 일평균이 낮은 상품을 배치하면 부스팅 효과를 기대할 수 있습니다.")

# ════════════════════════════════════════════════════════════
# 2. 트래픽 추세
# ════════════════════════════════════════════════════════════
section("트래픽 추세", "콘텐츠 영역별 UV/PV 흐름 — 핫딜이 VIP라운지 내에서 차지하는 위치",
        anchor="sec-traffic")

metric_t = st.radio("지표", ["UV", "PV"], index=0, horizontal=True, key="traffic_metric")
Tf = T[(T["date"] >= d0) & (T["date"] <= d1) & (T["metric"] == metric_t)]
fig = go.Figure()
for area in AREAS:
    s = Tf[Tf["area"] == area].set_index("date")["value"].sort_index()
    s = resample(s, freq, how="mean")  # 트래픽은 기간 평균(일 방문수준 유지)
    fig.add_trace(go.Scatter(x=s.index, y=s.values, name=area,
                             line=dict(color=AREA_COLOR[area], width=2)))
plot(fig, f"영역별 {metric_t} ({freq} 평균)", height=400)

# 핫딜 점유율
hd = T[(T["date"] >= d0) & (T["date"] <= d1) & (T["metric"] == "UV") &
       (T["area"] == "VIP핫딜")].set_index("date")["value"]
lng = T[(T["date"] >= d0) & (T["date"] <= d1) & (T["metric"] == "UV") &
        (T["area"] == "VIP라운지")].set_index("date")["value"]
share_series = (hd / lng.replace(0, np.nan) * 100).dropna()
if len(share_series):
    sw, sc = trend_word(share_series)
    insight(f"핫딜 UV는 VIP라운지 전체 UV의 평균 <b>{share_series.mean():.0f}%</b> "
            f"(추세 {sw}). 라운지 방문 중 핫딜 콘텐츠로 유입되는 비중입니다.")

# ════════════════════════════════════════════════════════════
# 3. 전환·효율
# ════════════════════════════════════════════════════════════
section("전환·효율 추세", "방문(UV) 대비 주문 전환율과 객단가의 흐름", anchor="sec-conv")

uv_d = FT["UV"].resample(FREQ[freq]).sum()
ord_full = resample(daily(FS, "ord"), freq)
rev_full = resample(daily(FS, "rev"), freq)
conv_s = (ord_full / uv_d.replace(0, np.nan) * 100)
aov_s = (rev_full / ord_full.replace(0, np.nan))

cc1, cc2 = st.columns(2)
with cc1:
    fig = px.area(x=conv_s.index, y=conv_s.values,
                  labels={"x": "", "y": "전환율(%)"})
    fig.update_traces(line_color="#55A868", fillcolor="rgba(85,168,104,0.2)")
    plot(fig, "전환율 (주문/UV)", height=320)
with cc2:
    fig = px.line(x=aov_s.index, y=aov_s.values, labels={"x": "", "y": "객단가(원)"})
    fig.update_traces(line_color="#8E44AD", line_width=2)
    plot(fig, "객단가 (거래액/주문)", height=320)

cw, cchg = trend_word(conv_s)
aw, achg = trend_word(aov_s)
insight(f"전환율 추세 <b>{cw}</b>({cchg:+.0f}%) · 객단가 추세 <b>{aw}</b>({achg:+.0f}%). "
        f"전환율은 페이지 총 UV 기준이라 슬롯 필터의 영향을 받지 않습니다.")

# ════════════════════════════════════════════════════════════
# 4. 베스트 (상품·브랜드·MD)
# ════════════════════════════════════════════════════════════
section("베스트 — 기간 누적 랭킹",
        "선택 기간·슬롯 기준. <b>일평균 거래액 + 하위</b>로 보면 ‘고매출 요일에 배치할 부스팅 후보(저조 상품)’를 찾을 수 있습니다",
        anchor="sec-best")

bc1, bc2, bc3 = st.columns(3)
with bc1:
    dim = st.radio("기준", ["상품", "브랜드", "MD", "카테고리"], horizontal=True, key="best_dim")
with bc2:
    basis = st.radio("정렬 지표", ["거래액 합계", "일평균 거래액"], horizontal=True, key="best_basis")
with bc3:
    order = st.radio("순서", ["상위 TOP", "하위 BOTTOM"], horizontal=True, key="best_order")
DIMCOL = {"상품": "prodname", "브랜드": "brand", "MD": "md_name", "카테고리": "category"}[dim]

agg = (FS[FS[DIMCOL] != ""].groupby(DIMCOL)
       .agg(거래액=("rev", "sum"), 주문=("ord", "sum"),
            수량=("qty", "sum"), 노출일수=("date", "nunique")))
agg["일평균"] = agg["거래액"] / agg["노출일수"].replace(0, np.nan)
sortcol = "거래액" if basis == "거래액 합계" else "일평균"
asc = order == "하위 BOTTOM"
agg = agg.sort_values(sortcol, ascending=asc)

top = agg.head(15).iloc[::-1]
fig = px.bar(top, x=sortcol, y=top.index, orientation="h",
             labels={"y": "", sortcol: f"{sortcol}(원)"},
             color_discrete_sequence=["#4C72B0" if asc else ACCENT])
fig.update_layout(yaxis=dict(tickfont=dict(size=11)))
plot(fig, f"{dim} · {sortcol} {'하위' if asc else '상위'} 15", height=480)

show = agg.head(20).copy()
for c in ["거래액", "일평균"]:
    show[c] = show[c].map(lambda v: "—" if pd.isna(v) else f"{v:,.0f}")
st.dataframe(show, use_container_width=True, height=300)
if asc and basis == "일평균 거래액":
    insight("👆 일평균 거래액 <b>하위</b> 항목들 — 평소 매출이 낮아 "
            "<b>고매출 요일(요일별 분석 참고)에 배치하면 부스팅 여지</b>가 큰 후보입니다. "
            "단, 노출일수가 너무 적은 항목은 표본이 작으니 함께 보세요.")

# ════════════════════════════════════════════════════════════
# 5. 브랜드·카테고리 믹스
# ════════════════════════════════════════════════════════════
section("브랜드·카테고리 믹스 추세", f"{freq} 매출 구성 변화 (상위 항목 + 기타)", anchor="sec-mix")

mix_dim = st.radio("구성 기준", ["브랜드", "카테고리"], horizontal=True, key="mix_dim")
MCOL = "brand" if mix_dim == "브랜드" else "category"
topn = (FS[FS[MCOL] != ""].groupby(MCOL)["rev"].sum()
        .sort_values(ascending=False).head(6).index.tolist())
mx = FS.copy()
mx["grp"] = np.where(mx[MCOL].isin(topn) & (mx[MCOL] != ""), mx[MCOL], "기타")
mx["bucket"] = mx["date"].dt.to_period(
    {"일별": "D", "주별": "W", "월별": "M"}[freq]).dt.start_time
piv = mx.groupby(["bucket", "grp"])["rev"].sum().reset_index()
fig = px.area(piv, x="bucket", y="rev", color="grp",
              labels={"bucket": "", "rev": "거래액(원)", "grp": mix_dim},
              color_discrete_sequence=px.colors.qualitative.Set2)
plot(fig, f"{mix_dim}별 거래액 구성", height=420)

# ════════════════════════════════════════════════════════════
# 6. 오전·오후 슬롯
# ════════════════════════════════════════════════════════════
section("오전·오후 슬롯 비교", "슬롯별 매출·주문 추세와 점유", anchor="sec-slot")

base = SLOTS[(SLOTS["date"] >= d0) & (SLOTS["date"] <= d1)]
sc1, sc2 = st.columns([2, 1])
with sc1:
    fig = go.Figure()
    for sl in ["오전", "오후"]:
        s = resample(base[base["slot"] == sl].groupby("date")["rev"].sum(), freq)
        fig.add_trace(go.Scatter(x=s.index, y=s.values, name=sl,
                                 line=dict(color=SLOT_COLOR[sl], width=2)))
    plot(fig, f"슬롯별 거래액 ({freq})", height=340)
with sc2:
    sl_sum = base.groupby("slot")["rev"].sum().reindex(["오전", "오후"]).fillna(0)
    fig = px.pie(values=sl_sum.values, names=sl_sum.index, hole=0.5,
                 color=sl_sum.index, color_discrete_map=SLOT_COLOR)
    plot(fig, "슬롯 매출 점유", height=340)

am, pm = base[base["slot"] == "오전"]["rev"].sum(), base[base["slot"] == "오후"]["rev"].sum()
lead = "오전" if am > pm else "오후"
insight(f"기간 합계 거래액은 <b>{lead}</b>이 우위 "
        f"(오전 {fwon(am)} · 오후 {fwon(pm)}원).")

# ════════════════════════════════════════════════════════════
# 7. 상세 데이터
# ════════════════════════════════════════════════════════════
section("상세 데이터", "필터된 슬롯·상품 단위 원본 (CSV 다운로드 가능)", anchor="sec-table")

cols = ["date", "slot", "brand", "category", "prodname", "md_name",
        "UV", "PV", "ord", "qty", "rev"]
tbl = FS[cols].sort_values("date", ascending=False)
st.dataframe(tbl, use_container_width=True, height=420)
st.download_button("⬇️ CSV 다운로드", tbl.to_csv(index=False).encode("utf-8-sig"),
                   file_name=f"vip_hotdeal_{d0.date()}_{d1.date()}.csv", mime="text/csv")

st.caption("ℹ️ 거래액은 VAT 제외 금액이며 원 단위 분수값일 수 있습니다(가격/1.1 등). "
           "UV/PV는 페이지 총계(중복 제거), 매출·주문은 슬롯 합산입니다.")

# ── 사이드바 nav 부드러운 스크롤 ────────────────────────────
components.html("""
<script>
const doc = window.parent.document;
doc.querySelectorAll('a.navlink').forEach(a=>{
  a.addEventListener('click', e=>{
    e.preventDefault();
    const el = doc.querySelector(a.getAttribute('href'));
    if(el) el.scrollIntoView({behavior:'smooth', block:'start'});
  });
});
</script>
""", height=0)
