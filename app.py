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
    details.navgrp { margin: 4px 0; }
    details.navgrp > summary {
        list-style: none; cursor: pointer; padding: 8px 12px; border-radius: 8px;
        background: #f2f5fa; color: #2E68B0; font-size: 14px; font-weight: 700;
        border: 1px solid #e3e9f2; user-select: none;
    }
    details.navgrp > summary::-webkit-details-marker { display: none; }
    details.navgrp > summary::after { content: "▸"; float: right; color: #9aa0a6; font-weight: 400; }
    details.navgrp[open] > summary { background: #e3ecf8; color: #163E78; }
    details.navgrp[open] > summary::after { content: "▾"; }
    a.navlink {
        display: block; padding: 6px 11px; margin: 3px 0 3px 10px; border-radius: 7px;
        background: #fbf2f2; color: #C0392B; text-decoration: none;
        font-size: 13px; font-weight: 600; border: 1px solid #f2e0e0;
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
    """해제된 핫딜.xlsx(그룹형) → tidy DataFrame. convert.ps1 과 동일 로직.
    일자에 연도가 없는(M/D) 행은 제외하고, 명시적 연도(YY/MM/DD)만 사용."""
    g = pd.read_excel(io.BytesIO(raw), sheet_name=0, header=None).values
    n, ncol = g.shape
    # Total 행 → 날짜. 연도 없는 행은 None 으로 표시해 이후 행을 무효화.
    rowdate = {}
    for r in range(1, n):
        if _isblank(g[r, 0]):
            continue
        p = _md(g[r, 0])
        if p and p[2] is not None:           # 연도 명시 → 사용
            rowdate[r] = f"{p[2]:04d}-{p[0]:02d}-{p[1]:02d}"
        elif p:                              # 연도 없음 → 해당 일자 그룹 제외
            rowdate[r] = None
    if not any(v for v in rowdate.values()):
        raise ValueError("핫딜 시트에서 연도가 있는 일자 행을 찾지 못했습니다.")

    def cell(r, c):
        return "" if (c >= ncol or _isblank(g[r, c])) else g[r, c]

    rows, cur = [], None
    for r in range(1, n):
        if r in rowdate:
            cur = rowdate[r]
        if cur is None:                      # 연도 없는 일자 그룹 / 첫 일자 이전 → 제외
            continue
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


def won(x):
    """원 단위 천단위 콤마 (정수)."""
    return "—" if pd.isna(x) else f"{round(float(x)):,}"


def _norm01(s):
    """Series를 0~1로 min-max 정규화(범위 0이면 0)."""
    lo, hi = s.min(), s.max()
    return (s - lo) / (hi - lo) if hi > lo else s * 0.0


def sgn(x, suffix="%", digits=1):
    """증감 표기(내부 규칙): 양수=세모 없이 녹색 숫자, 음수=빨강 △(빈 세모)."""
    if x is None or pd.isna(x):
        return '<span style="color:#999">—</span>'
    if x >= 0:
        return f'<span style="color:#2E7D32">{x:.{digits}f}{suffix}</span>'
    return f'<span style="color:#C44E52">△ {abs(x):.{digits}f}{suffix}</span>'


def fdelta(cur, prev):
    """직전 동일길이 기간 대비 증감률 HTML."""
    if prev in (0, None) or pd.isna(prev):
        return '<span style="color:#999">—</span>'
    return sgn((cur - prev) / prev * 100, "%", 1)


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
    delta = '<span style="color:#999">비교 데이터 없음</span>' if p is None else sgn(p, "%", 1)
    cv = won(cur)
    pv = "—" if prev is None else won(prev)
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

MENU_GROUPS = [
    ("개요", [
        ("sec-core", "🔑 핵심 요약"),
    ]),
    ("매출 성과", [
        ("sec-compare", "🔁 전년·전월·전주 비교"),
        ("sec-avg", "📊 일평균 거래액"),
        ("sec-sales", "💰 매출 추세·피크일"),
    ]),
    ("유입·전환", [
        ("sec-traffic", "🚦 트래픽 추세"),
        ("sec-conv", "🎯 전환·효율"),
    ]),
    ("패턴", [
        ("sec-dow", "📅 요일별 분석"),
        ("sec-slot", "⏰ 오전·오후 슬롯"),
    ]),
    ("구성·랭킹", [
        ("sec-best", "🏆 베스트"),
        ("sec-mix", "🧩 카테고리 믹스"),
    ]),
    ("실행·정리", [
        ("sec-md", "🧾 MD·브랜드 1-Pager"),
        ("sec-table", "📋 상세 데이터"),
        ("sec-insight", "🧭 인사이트 & 액션"),
    ]),
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
    rev_basis = st.radio(
        "💳 매출 기준", ["핫딜 직접 (VIP핫딜 경유)", "상품 전체"], index=0,
        help="핫딜 직접 = VIP핫딜 영역을 경유해 구매한 어트리뷰션 실적. "
             "상품 전체 = 그 핫딜 상품의 전체 거래액(다른 경로 포함). "
             "거래액·주문·수량·고객수가 이 기준으로 바뀝니다.")

    st.divider()
    st.markdown("**📂 분석 메뉴**")
    nav_html = ""
    for i, (gtitle, items) in enumerate(MENU_GROUPS):
        op = " open" if i == 0 else ""
        links = "".join(f'<a href="#{a}" class="navlink">{lbl}</a>' for a, lbl in items)
        nav_html += (f'<details class="navgrp" name="navacc"{op}>'
                     f'<summary>{gtitle}</summary>{links}</details>')
    st.markdown(nav_html, unsafe_allow_html=True)
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
        f"모든 거래액은 <b>일평균(운영일 1일당)</b> 기준 — 기간 길이 달라도 비교 가능. "
        f"증감은 직전 같은 길이 기간({p_d0.date()} ~ {p_d1.date()}) 대비 "
        f"(<span style='color:#2E7D32'>증가</span> · <span style='color:#C44E52'>△ 감소</span>)",
        anchor="sec-core")

# 기본 지표 = 일평균(운영일 평균). 누적은 기간 길이가 달라 비교 불가하므로.
nd = max(FS["date"].nunique(), 1)          # 선택 기간 운영일수
pnd = max(PS["date"].nunique(), 1)         # 직전 기간 운영일수
rev_h, rev_t = FS["h_rev"].sum(), FS["rev"].sum()
adir, atot = rev_h / nd, rev_t / nd                        # 일평균 직접·전체
adir_p, atot_p = PS["h_rev"].sum() / pnd, PS["rev"].sum() / pnd
ord_h = FS["h_ord"].sum()
aord, aord_p = ord_h / nd, PS["h_ord"].sum() / pnd          # 일평균 주문
uv_avg = FT["UV"].sum() / nd                               # 일평균 UV
attr = (rev_h / rev_t * 100) if rev_t else 0
conv_h = (ord_h / FT["UV"].sum() * 100) if FT["UV"].sum() else 0
aov_h = (rev_h / ord_h) if ord_h else 0

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    metric_card("핫딜 직접 일평균 거래액", f"{won(adir)}원",
                f"VIP핫딜 경유 · 직전 {fdelta(adir, adir_p)}", color="#E45756")
with c2:
    metric_card("상품 전체 일평균 거래액", f"{won(atot)}원",
                f"다른 경로 포함 · 직전 {fdelta(atot, atot_p)}", color="#4C72B0")
with c3:
    metric_card("어트리뷰션율", f"{attr:.0f}%", "직접 ÷ 전체 (영역 기여도)", color="#8E44AD")
with c4:
    metric_card("직접 일평균 주문", f"{aord:.1f}건",
                f"객단가 {won(aov_h)}원 · 직전 {fdelta(aord, aord_p)}", color="#E45756")
with c5:
    metric_card("일평균 UV · 전환율", fnum(uv_avg),
                f"전환율(직접) {conv_h:.2f}%")

# ── 전역 매출 기준 적용 (이후 모든 섹션) ────────────────────
BASIS_H = rev_basis.startswith("핫딜")
basis_label = "핫딜 직접 (VIP핫딜 경유)" if BASIS_H else "상품 전체"


def apply_basis(df):
    df["_rev_total"] = df["rev"].copy()      # 전체 거래액 보존(상세표·어트리뷰션용)
    df["_rev_direct"] = df["h_rev"].copy()   # 직접 거래액 보존
    if BASIS_H:
        for c in ("rev", "ord", "qty", "cust"):
            df[c] = df["h_" + c]
    return df


apply_basis(FS)
apply_basis(SLOTS)
st.markdown(
    f'<div style="background:#fff4f3;border:1px solid #f3d5d2;border-radius:8px;'
    f'padding:8px 14px;margin:4px 0 6px;font-size:13px">💳 아래 모든 매출 지표 기준: '
    f'<b>{basis_label}</b> · 거래액은 모두 <b>일평균</b> 기준 &nbsp;— 사이드바에서 변경</div>',
    unsafe_allow_html=True)

rev_d = resample(daily(FS, "rev"), freq, how="mean")
word, chg = trend_word(rev_d)
top_brand = (FS.groupby("brand")["rev"].sum().sort_values(ascending=False))
top_brand = top_brand[top_brand.index != ""]
tb_name = top_brand.index[0] if len(top_brand) else "—"
tb_share = (top_brand.iloc[0] / FS["rev"].sum() * 100) if len(top_brand) and FS["rev"].sum() else 0
insight(
    f"<b>핫딜 직접 일평균 {won(adir)}원/일</b>(영역 기여 {attr:.0f}%) · "
    f"상품 전체 일평균 {won(atot)}원/일. "
    f"현재 ‘{basis_label}’ 기준 {freq} 일평균 추세는 <b>{word}</b>(기간 내 {sgn(chg, '%', 0)}), "
    f"1위 브랜드 <b>{tb_name}</b>(거래액의 {tb_share:.0f}%).",
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
    bits.append(f"전주 대비 일평균 {won(cur_w[0])}원 ({sgn(pct(cur_w[0], prev_w[0]), '%', 0)})")
if pct(cur_m[0], prev_m[0]) is not None:
    bits.append(f"전월 대비 {sgn(pct(cur_m[0], prev_m[0]), '%', 0)}")
if pct(cur_m[0], prev_y[0]) is not None:
    bits.append(f"전년 동월 대비 {sgn(pct(cur_m[0], prev_y[0]), '%', 0)}")
else:
    bits.append("전년 동월 데이터가 없어 전년비는 비교 불가")
insight("최신 기준 " + " · ".join(bits) + ". "
        "이번 주/달은 <b>진행 중</b>이라 일평균(운영일 평균)으로 비교합니다.")


# ── 어트리뷰션율(직접÷전체) — 전월·전년 비교 + 월별 YoY 추세 ──
def attr_of(start, end):
    d = SLOTS[(SLOTS["date"] >= start) & (SLOTS["date"] <= end)]
    t = d["_rev_total"].sum()
    return (d["_rev_direct"].sum() / t * 100) if t else np.nan


a_cur, a_pm, a_py = attr_of(m_start, ref), attr_of(pm_end.replace(day=1), pm_end), attr_of(ly_start, ly_end)
st.markdown('<div style="font-weight:700;font-size:15px;margin:10px 0 2px">'
            '🎯 어트리뷰션율(직접÷전체) 비교</div>', unsafe_allow_html=True)
ac1, ac2, ac3 = st.columns(3)
with ac1:
    metric_card(f"이번 달 ({ref.month}월)", "—" if pd.isna(a_cur) else f"{a_cur:.0f}%",
                "VIP핫딜 영역 기여도", color="#8E44AD")
with ac2:
    dd = "" if (pd.isna(a_cur) or pd.isna(a_pm)) else f"({sgn(a_cur - a_pm, 'p', 0)})"
    metric_card("전월 대비", "—" if pd.isna(a_pm) else f"{a_pm:.0f}%", f"지난달 → 이번달 {dd}",
                color="#8E44AD")
with ac3:
    dd = "" if (pd.isna(a_cur) or pd.isna(a_py)) else f"({sgn(a_cur - a_py, 'p', 0)})"
    metric_card("전년 동월 대비", "—" if pd.isna(a_py) else f"{a_py:.0f}%",
                f"{ly_start.year}.{ly_start.month:02d} → 올해 {dd}", color="#8E44AD")

# 월별 어트리뷰션율 연도 비교(YoY)
at = SLOTS[SLOTS["date"].dt.year >= 2024].copy()
yy = at["date"].dt.year.rename("year")
mm = at["date"].dt.month.rename("month")
ap = at.groupby([yy, mm]).agg(d=("_rev_direct", "sum"), t=("_rev_total", "sum"))
ap["attr"] = np.where(ap["t"] > 0, ap["d"] / ap["t"] * 100, np.nan)
ap = ap.reset_index()
ap["year"] = ap["year"].astype(str)
figa = px.line(ap, x="month", y="attr", color="year", markers=True,
               labels={"month": "월", "attr": "어트리뷰션율(%)", "year": "연도"},
               color_discrete_sequence=px.colors.qualitative.Set2)
figa.update_xaxes(dtick=1)
plot(figa, "월별 어트리뷰션율 (연도 비교)", height=340)
if not pd.isna(a_cur) and not pd.isna(a_py):
    insight(f"이번 달 어트리뷰션율 <b>{a_cur:.0f}%</b> — 전년 동월({a_py:.0f}%) 대비 "
            f"{sgn(a_cur - a_py, 'p', 0)}. 이 비율이 오르면 핫딜 영역이 매출을 직접 더 많이 "
            "끌고 있다는 뜻이고, 내리면 ‘보고 나중에 구매’ 비중이 커진 것입니다.")

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
fig.update_yaxes(tickformat=",")
plot(fig, f"{gran}별 일평균 거래액", height=380)

best_p = ad["일평균"].idxmax()
worst_p = ad["일평균"].idxmin()
insight(f"{gran} 중 일평균 거래액이 가장 높은 구간은 <b>{best_p}</b>"
        f"({won(ad.loc[best_p, '일평균'])}원/일), 가장 낮은 구간은 <b>{worst_p}</b>"
        f"({won(ad.loc[worst_p, '일평균'])}원/일)입니다.")

# ════════════════════════════════════════════════════════════
# 1. 매출 추세
# ════════════════════════════════════════════════════════════
section("매출 추세 + 피크일",
        f"{freq} <b>일평균 거래액</b> 추이 — <b>핫딜 직접</b>(빨강)·<b>상품 전체</b>(파랑). "
        "기간별 일수가 달라도 비교되도록 일평균으로 표시",
        anchor="sec-sales")

tot_s = resample(daily(FS, "_rev_total"), freq, how="mean")
dir_s = resample(daily(FS, "_rev_direct"), freq, how="mean")
ord_s = resample(daily(FS, "ord"), freq, how="mean")

fig = go.Figure()
fig.add_trace(go.Scatter(x=tot_s.index, y=tot_s.values, name="상품 전체 일평균",
                         mode="lines", line=dict(color="#4C72B0", width=2)))
fig.add_trace(go.Scatter(x=dir_s.index, y=dir_s.values, name="핫딜 직접 일평균",
                         mode="lines+markers", line=dict(color=ACCENT, width=2.4),
                         marker=dict(size=4)))
fig.add_trace(go.Scatter(x=ord_s.index, y=ord_s.values, name="일평균 주문",
                         yaxis="y2", line=dict(color="#999", width=1.2, dash="dot")))
fig.update_layout(
    yaxis=dict(title="일평균 거래액(원)", tickformat=","),
    yaxis2=dict(title="일평균 주문", overlaying="y", side="right", showgrid=False))
plot(fig, height=420)
_tt = FS["_rev_total"].sum()
if _tt:
    insight(f"두 선의 간격(파랑−빨강)이 <b>핫딜 영역을 거치지 않은 매출</b>"
            f"(라운지에서 보고 나중에 재유입 구매 등)입니다. "
            f"선택 기간 평균 어트리뷰션율 {FS['_rev_direct'].sum()/_tt*100:.0f}%.")

# 거래액 피크일 — 그날을 견인한 브랜드·상품 (차트 바로 아래)
st.markdown('<div style="font-weight:700;font-size:15px;margin:14px 0 4px">'
            '🔝 거래액 피크일 TOP 12 — 그날 1등 상품</div>', unsafe_allow_html=True)
day_rev = daily(FS, "rev")
day_ord = daily(FS, "ord")
day_qty = daily(FS, "qty")
peak_recs = []
for dt in day_rev.sort_values(ascending=False).head(12).index:
    tot = day_rev[dt]
    top = FS[FS["date"] == dt].sort_values("rev", ascending=False).iloc[0]
    peak_recs.append({
        "일자": str(dt.date()), "요일": DOW[dt.weekday()],
        "당일 거래액": won(tot),
        "건수": int(day_ord.get(dt, 0)), "수량": int(day_qty.get(dt, 0)),
        "슬롯": top["slot"], "브랜드": top["brand"], "상품": top["prodname"][:28],
        "상품 거래액": won(top["rev"]),
        "비중": f"{top['rev']/tot*100:.0f}%" if tot else "—",
    })
st.dataframe(pd.DataFrame(peak_recs), use_container_width=True, height=320, hide_index=True)
st.caption("거래액 순 정렬. ‘건수·수량’을 함께 보세요 — 건수 1인데 거래액이 크면 "
           "고단가 1개로 튄 날이라 실제 수요 피크는 아닙니다. "
           "가중 우선순위(판매금액·판매량·PV)는 아래 🏆 베스트에서 봅니다.")

# 연도(YoY) 비교 — 월별 일평균 거래액 (2024년 이후만)
yoy = SLOTS.copy()
if sel_slots:
    yoy = yoy[yoy["slot"].isin(sel_slots)]
yoy = yoy[yoy["date"].dt.year >= 2024]
dly = yoy.groupby("date")["rev"].sum().reset_index()      # 일별 총
dly["year"] = dly["date"].dt.year.astype(str)
dly["month"] = dly["date"].dt.month
ym = dly.groupby(["year", "month"])["rev"].mean().reset_index()   # 월별 일평균
figy = px.line(ym, x="month", y="rev", color="year", markers=True,
               labels={"month": "월", "rev": "일평균 거래액(원)", "year": "연도"},
               color_discrete_sequence=px.colors.qualitative.Set2)
figy.update_xaxes(dtick=1)
figy.update_yaxes(tickformat=",")
plot(figy, "연도별 월간 일평균 거래액 비교 (YoY)", height=360)
insight("<b>2024년은 7월부터</b> 데이터가 있어 상반기가 비어 보입니다.", "")

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
insight(f"전환율 추세 <b>{cw}</b>({sgn(cchg, '%', 0)}) · 객단가 추세 <b>{aw}</b>({sgn(achg, '%', 0)}). "
        f"전환율은 페이지 총 UV 기준이라 슬롯 필터의 영향을 받지 않습니다.")

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
    "일평균 거래액(원)": [won(v) for v in dow_rev.values],
    "일평균 UV": [won(v) for v in dow_uv.values],
    "일평균 주문": np.round(dow_ord.values, 1),
    "전환율(%)": np.round(dow_ord.values / dow_uv.values * 100, 2),
})
st.dataframe(dow_tbl, use_container_width=True, hide_index=True)

if dow_rev.notna().any():
    hi = int(dow_rev.idxmax())
    hu = int(dow_uv.idxmax()) if dow_uv.notna().any() else hi
    insight(f"거래액이 가장 높은 요일은 <b>{DOW[hi]}요일</b>"
            f"({won(dow_rev.iloc[hi])}원/일), 방문(UV)이 가장 많은 요일은 "
            f"<b>{DOW[hu]}요일</b>입니다. 전략대로라면 <b>{DOW[hi]}요일</b> 같은 고매출 요일에 "
            "평소 일평균이 낮은 상품을 배치하면 부스팅 효과를 기대할 수 있습니다.")

# ════════════════════════════════════════════════════════════
#    오전·오후 슬롯
# ════════════════════════════════════════════════════════════
section("오전·오후 슬롯 비교", "슬롯별 매출·주문 추세와 점유", anchor="sec-slot")

base = SLOTS[(SLOTS["date"] >= d0) & (SLOTS["date"] <= d1)]
sc1, sc2 = st.columns([2, 1])
with sc1:
    fig = go.Figure()
    for sl in ["오전", "오후"]:
        s = resample(base[base["slot"] == sl].groupby("date")["rev"].sum(), freq, how="mean")
        fig.add_trace(go.Scatter(x=s.index, y=s.values, name=sl,
                                 line=dict(color=SLOT_COLOR[sl], width=2)))
    fig.update_yaxes(tickformat=",")
    plot(fig, f"슬롯별 일평균 거래액 ({freq})", height=340)
with sc2:
    sl_sum = base.groupby("slot")["rev"].sum().reindex(["오전", "오후"]).fillna(0)
    fig = px.pie(values=sl_sum.values, names=sl_sum.index, hole=0.5,
                 color=sl_sum.index, color_discrete_map=SLOT_COLOR)
    plot(fig, "슬롯 매출 점유", height=340)

am, pm = base[base["slot"] == "오전"]["rev"].sum(), base[base["slot"] == "오후"]["rev"].sum()
lead = "오전" if am > pm else "오후"
insight(f"기간 합계 거래액은 <b>{lead}</b>이 우위 "
        f"(오전 {won(am)}원 · 오후 {won(pm)}원).")

# ════════════════════════════════════════════════════════════
# 4. 베스트 (상품·브랜드·MD)
# ════════════════════════════════════════════════════════════
section("베스트 — 기간 누적 랭킹",
        "선택 기간·슬롯 기준. <b>가중 점수</b>=몰 베스트 산출식(판매금액·판매량·PV 정규화 후 가중합). "
        "<b>일평균 거래액 + 하위</b>로 보면 ‘고매출 요일에 배치할 부스팅 후보(저조 상품)’를 찾습니다",
        anchor="sec-best")

bc1, bc2, bc3 = st.columns(3)
with bc1:
    dim = st.radio("기준", ["상품", "브랜드", "MD", "카테고리"], horizontal=True, key="best_dim")
with bc2:
    basis = st.radio("정렬 지표", ["가중 점수", "일평균 거래액", "거래액 합계"],
                     horizontal=True, key="best_basis")
with bc3:
    order = st.radio("순서", ["상위 TOP", "하위 BOTTOM"], horizontal=True, key="best_order")
if basis == "가중 점수":
    st.caption("가중치 (판매금액·판매량·PV) — 각 지표를 0~1 정규화 후 가중합, 최고=100점")
    bw1, bw2, bw3 = st.columns(3)
    bw_amt = bw1.number_input("판매금액", 0.0, 1.0, 0.7, 0.05, key="bw_amt")
    bw_qty = bw2.number_input("판매량", 0.0, 1.0, 0.2, 0.05, key="bw_qty")
    bw_pv = bw3.number_input("PV", 0.0, 1.0, 0.1, 0.05, key="bw_pv")
DIMCOL = {"상품": "prodname", "브랜드": "brand", "MD": "md_name", "카테고리": "category"}[dim]

agg = (FS[FS[DIMCOL] != ""].groupby(DIMCOL)
       .agg(거래액=("rev", "sum"), 주문=("ord", "sum"), 수량=("qty", "sum"),
            PV=("PV", "sum"), 노출일수=("date", "nunique")))
agg["일평균"] = agg["거래액"] / agg["노출일수"].replace(0, np.nan)
if basis == "가중 점수":
    raw = bw_amt * _norm01(agg["거래액"]) + bw_qty * _norm01(agg["수량"]) + bw_pv * _norm01(agg["PV"])
    agg["가중점수"] = (raw / raw.max() * 100).round(1) if raw.max() else 0.0

sortcol = {"가중 점수": "가중점수", "일평균 거래액": "일평균", "거래액 합계": "거래액"}[basis]
asc = order == "하위 BOTTOM"
agg = agg.sort_values(sortcol, ascending=asc)
unit = "점" if sortcol == "가중점수" else "원"

top = agg.head(15).iloc[::-1]
fig = px.bar(top, x=sortcol, y=top.index, orientation="h",
             labels={"y": "", sortcol: f"{sortcol}({unit})"},
             color_discrete_sequence=["#4C72B0" if asc else ACCENT])
fig.update_layout(yaxis=dict(tickfont=dict(size=11)))
plot(fig, f"{dim} · {sortcol} {'하위' if asc else '상위'} 15", height=480)

show = agg.head(20).copy()
for c in ["거래액", "일평균"]:
    show[c] = show[c].map(lambda v: "—" if pd.isna(v) else f"{v:,.0f}")
st.dataframe(show, use_container_width=True, height=300)
if basis == "가중 점수":
    insight("‘가중점수’는 몰 베스트 노출 기준과 같은 방식(판매금액·판매량·PV 정규화 후 가중합)입니다. "
            "가중치를 바꾸면 우선순위가 재계산돼요. 고단가 단발 상품은 판매량·PV가 낮아 자연히 밀립니다.")
elif asc and basis == "일평균 거래액":
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
#    MD·브랜드 1-Pager (MD 핫딜 상품 요청용 성적표)
# ════════════════════════════════════════════════════════════
section("MD·브랜드 1-Pager",
        f"브랜드(또는 MD)를 골라 핫딜 성과 요약 — MD에게 핫딜 상품 요청할 때 그대로 첨부. "
        f"기간 {d0.date()} ~ {d1.date()} (원본 지표, 매출 기준 토글 무관)",
        anchor="sec-md")

# 원본(H)에서 슬롯·상품 단위로 계산 — 매출 기준 토글/스왑 영향 없음
B = H[(H["row_type"] == "SLOT") & (H["date"] >= d0) & (H["date"] <= d1)].copy()
mc1, mc2 = st.columns([1, 2])
with mc1:
    pdim = st.radio("기준", ["브랜드", "MD"], horizontal=True, key="md_dim")
PCOL = "brand" if pdim == "브랜드" else "md_name"
B = B[B[PCOL] != ""]
order = B.groupby(PCOL)["rev"].sum().sort_values(ascending=False)
with mc2:
    ent = st.selectbox(pdim, order.index.tolist(), index=0 if len(order) else None)

if ent:
    bb = B[B[PCOL] == ent]
    days = bb["date"].nunique()
    prods = bb["prodname"].nunique()
    uv = bb["UV"].sum()
    rev_t, rev_d = bb["rev"].sum(), bb["h_rev"].sum()
    halo = rev_t - rev_d
    ordc = bb["ord"].sum()
    attr_b = (rev_d / rev_t * 100) if rev_t else 0
    conv_b = (ordc / uv * 100) if uv else 0
    aov_b = (rev_t / ordc) if ordc else 0
    # 전체(모든 브랜드/MD) 평균 벤치마크
    A_attr = (B["h_rev"].sum() / B["rev"].sum() * 100) if B["rev"].sum() else 0
    A_conv = (B["ord"].sum() / B["UV"].sum() * 100) if B["UV"].sum() else 0
    A_aov = (B["rev"].sum() / B["ord"].sum()) if B["ord"].sum() else 0
    # 베스트 요일
    bd = bb.groupby("date")["rev"].sum()
    bdow = bd.groupby(bd.index.weekday).mean().reindex(range(7))
    best_dow = DOW[int(bdow.idxmax())] if bdow.notna().any() else "—"

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        metric_card("핫딜 노출", f"{days}일", f"상품 {prods}개 · 누적 UV {fnum(uv)}", color="#4C72B0")
    with k2:
        metric_card("전체 거래액", f"{won(rev_t)}원", "기간 누적(노출이 만든 총매출)", color="#E45756")
    with k3:
        metric_card("핫딜 직접", f"{won(rev_d)}원", f"어트리뷰션 {attr_b:.0f}%", color="#E45756")
    with k4:
        metric_card("헤일로 매출", f"{won(halo)}원", "할인 외 경로(노출 효과)", color="#2E7D32")

    insight(
        f"<b>{ent}</b>: 핫딜 직접 매출 {won(rev_d)}원 외에 <b>노출로 발생한 매출이 {won(halo)}원</b> "
        f"(전체의 {100-attr_b:.0f}%) — 할인 매출만 보면 과소평가됩니다. "
        f"전환율 {conv_b:.2f}%(평균 {A_conv:.2f}%, 평균 대비 {sgn(conv_b - A_conv, '%p', 2)}) · "
        f"객단가 {won(aov_b)}원(평균 대비 {sgn(aov_b - A_aov, '원', 0)}) · "
        f"어트리뷰션 {attr_b:.0f}%(평균 대비 {sgn(attr_b - A_attr, '%p', 0)}). "
        f"가장 잘 팔린 요일은 <b>{best_dow}요일</b>.",
        "ok" if conv_b >= A_conv else "")

    # MD 메일에 붙일 요약문 (복사 버튼 제공)
    pitch = (
        f"[{ent}] VIP핫딜 성과 요약 ({d0.date()} ~ {d1.date()})\n"
        f"────────────────────────────\n"
        f"• 노출: {days}일 · 상품 {prods}개 · VIP핫딜 누적 방문 UV {uv:,.0f}\n"
        f"• 매출: 전체 {rev_t:,.0f}원 = 핫딜 직접 {rev_d:,.0f}원 + 노출로 발생 {halo:,.0f}원\n"
        f"        → 할인 매출 외 '노출 효과' 매출이 {halo:,.0f}원 (전체의 {100-attr_b:.0f}%)\n"
        f"• 효율: 전환율 {conv_b:.2f}% (전체 평균 {A_conv:.2f}%) · 객단가 {aov_b:,.0f}원\n"
        f"• 추천: {best_dow}요일에 가장 잘 팔립니다 — 해당 요일 배치 제안\n"
        f"────────────────────────────\n"
        f"VIP핫딜은 광고비 없이 VIP 고객에게 브랜드를 노출하고, 할인 매출을 넘어\n"
        f"정상가·재유입 구매까지 끌어냅니다. 다음 주 핫딜 상품 제공 부탁드립니다."
    )
    st.markdown("**📋 MD 요청 메일용 요약 (우측 상단 복사 버튼)**")
    st.code(pitch, language=None)
else:
    st.info("선택한 기간에 데이터가 있는 브랜드/MD가 없습니다. 기간을 넓혀 보세요.")

# ════════════════════════════════════════════════════════════
# 7. 상세 데이터
# ════════════════════════════════════════════════════════════
section("상세 데이터", "필터된 슬롯·상품 단위 원본 — 거래액은 전체·직접 둘 다 표기 (CSV 다운로드 가능)",
        anchor="sec-table")

tbl = FS.copy().sort_values("date", ascending=False)
tbl["date"] = tbl["date"].dt.date.astype(str)
tbl["거래액_전체"] = tbl["_rev_total"].map(won)
tbl["거래액_직접"] = tbl["_rev_direct"].map(won)
cols = ["date", "slot", "brand", "category", "prodname", "md_name",
        "UV", "PV", "ord", "거래액_전체", "거래액_직접"]
st.dataframe(tbl[cols], use_container_width=True, height=420, hide_index=True)
st.download_button("⬇️ CSV 다운로드", tbl[cols].to_csv(index=False).encode("utf-8-sig"),
                   file_name=f"vip_hotdeal_{d0.date()}_{d1.date()}.csv", mime="text/csv")

st.caption("ℹ️ 거래액_전체 = 상품 전체 거래액, 거래액_직접 = VIP핫딜 영역 경유(어트리뷰션). "
           "VAT 제외·원 단위 분수값일 수 있음. UV/PV는 페이지 총계(중복 제거), 매출·주문은 슬롯 합산. "
           "주문(ord)은 현재 선택한 ‘매출 기준’을 따릅니다.")

# ════════════════════════════════════════════════════════════
#    인사이트 & 액션 (종합)
# ════════════════════════════════════════════════════════════
section("인사이트 & 액션",
        f"선택 기간·필터 기준으로 자동 요약된 현황 진단·시사점·액션 ({d0.date()} ~ {d1.date()})",
        anchor="sec-insight")

# 요약용 지표 (자체 재계산 — 섹션 독립)
i_days = max(FS["date"].nunique(), 1)
i_dir, i_tot = FS["_rev_direct"].sum(), FS["_rev_total"].sum()
i_attr = (i_dir / i_tot * 100) if i_tot else 0
i_adir = i_dir / i_days
i_drev = daily(FS, "rev")
i_dowrev = i_drev.groupby(i_drev.index.weekday).mean().reindex(range(7))
i_bestdow = DOW[int(i_dowrev.idxmax())] if i_dowrev.notna().any() else "—"
i_uv, i_ord = FT["UV"].sum(), FS["ord"].sum()
i_conv = (i_ord / i_uv * 100) if i_uv else 0
i_tbs = FS[FS["brand"] != ""].groupby("brand")["rev"].sum().sort_values(ascending=False)
i_tb = i_tbs.index[0] if len(i_tbs) else "—"
i_tw, i_tchg = trend_word(resample(daily(FS, "rev"), freq, how="mean"))


def _panel(title, color, items):
    lis = "".join(f"<li>{x}</li>" for x in items)
    st.markdown(
        f'<div style="border-left:5px solid {color};background:#fafafb;border-radius:8px;'
        f'padding:12px 18px;margin:8px 0">'
        f'<div style="font-weight:700;font-size:15px;color:#1a1a2e;margin-bottom:6px">{title}</div>'
        f'<ul style="margin:0 0 0 18px;padding:0;font-size:14px;line-height:1.75">{lis}</ul></div>',
        unsafe_allow_html=True)


_panel("📌 현황 분석 & 인사이트", "#4C72B0", [
    f"핫딜 직접 <b>일평균 {won(i_adir)}원</b> · 어트리뷰션율 <b>{i_attr:.0f}%</b> — "
    f"상품 전체 매출의 {i_attr:.0f}%만 핫딜 영역 직접 구매, 나머지 <b>{100-i_attr:.0f}%는 노출 후 다른 경로</b>(헤일로).",
    f"{freq} 매출(일평균) 추세는 <b>{i_tw}</b> ({sgn(i_tchg, '%', 0)}).",
    f"방문→구매 전환율 <b>{i_conv:.2f}%</b> · 기간 매출 1위 브랜드 <b>{i_tb}</b>.",
    f"요일 편차 존재 — <b>{i_bestdow}요일</b>의 일평균 거래액이 가장 높음.",
])
_panel("💡 시사점 (So-What)", "#E45756", [
    f"핫딜은 ‘할인 손실’이 아니라 <b>노출 채널</b> — 매출의 <b>{100-i_attr:.0f}%가 노출로 발생</b>. "
    "MD에게 핫딜 상품을 요청할 때의 핵심 근거.",
    f"고매출 요일(<b>{i_bestdow}요일</b>)에 평소 <b>일평균 저조 상품</b>을 배치하면 부스팅 여지가 큼.",
    "어트리뷰션율이 낮아지면 ‘보고 나중에 구매’ 비중↑ — 영역 기여도 저하 신호로 모니터링 필요.",
])
_panel("✅ 추후 액션 (Action)", "#2E7D32", [
    "매주 MD 요청 메일에 <b>MD·브랜드 1-Pager</b>(헤일로·성적표)를 그대로 첨부해 설득력 강화.",
    "<b>요일별 분석</b> × <b>베스트(일평균 하위)</b>를 교차해 고매출 요일에 올릴 부스팅 상품 선정.",
    "노출 우선순위는 <b>베스트 가중 점수</b>(판매금액·판매량·PV)로 정렬해 몰 기준과 정합성 유지.",
    "(데이터 확보 시) 고객 신규/재구매·마진을 연동해 <b>ROI·LTV</b> 논거로 1-Pager 보강.",
])
st.caption("※ 위 요약은 현재 선택한 기간·슬롯·매출 기준에 따라 자동 갱신됩니다.")

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
