import streamlit as st
import pandas as pd
try:
    pd.set_option("mode.dtype_backend", "numpy_nullable")
except Exception:
    pass
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
from bs4 import BeautifulSoup
import pdfplumber
import re
import os

# ─── 고정 설정 ──────────────────────────────────────────
BOK_API_KEY = st.secrets.get("BOK_API_KEY", os.environ.get("BOK_API_KEY", ""))
ANTHROPIC_API_KEY = st.secrets.get("ANTHROPIC_API_KEY", os.environ.get("ANTHROPIC_API_KEY", ""))
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "data_apr_w2")

END_DATE = "20260403"
START_DATE = "20260103"
LAST_WEEK_START = "20260330"
LAST_WEEK_END = "20260403"

# 리포트 대상 주간 (이번 주)
REPORT_WEEK_START = "20260407"  # 월요일
REPORT_WEEK_END = "20260411"    # 금요일
# 외부 기사 수집 기준: 전주 금요일 이후 발행분만
NEWS_CUTOFF_DATE = "20260404"   # 이 날짜 이후 기사만 수집

st.set_page_config(page_title="주간 환율 리포트 — 4월 1주차", page_icon="📊", layout="wide")

# ─── CSS ────────────────────────────────────────────────
st.markdown("""
<style>
div[data-testid="stMetric"] {
    background: linear-gradient(135deg, #667eea0d, #764ba20d);
    border: 1px solid #ddd; border-radius: 12px;
    padding: 16px 20px; box-shadow: 0 2px 8px rgba(0,0,0,0.03);
}
div[data-testid="stMetric"] label { font-size:.85rem!important; color:#555!important; }
div[data-testid="stMetric"] [data-testid="stMetricValue"] { font-weight:700!important; }
.section-header {
    background: linear-gradient(90deg, #2E75B6, #4a90d9);
    color: white; padding: 10px 20px; border-radius: 8px;
    font-size: 1.15rem; font-weight: 700; margin: 28px 0 16px 0;
}
.comment-box {
    background: #f0f4ff; border-left: 4px solid #2E75B6;
    padding: 14px 18px; border-radius: 8px; margin: 12px 0;
    font-size: 0.95rem; line-height: 1.7; color: #1a1a1a;
}
.insight-card {
    background: #fffbf0; border-left: 4px solid #e6a817;
    padding: 14px 18px; border-radius: 8px; margin: 10px 0;
    font-size: 0.93rem; line-height: 1.7;
}
.news-card {
    background: #f8f9fc; border: 1px solid #e0e0e0;
    padding: 16px 20px; border-radius: 10px; margin: 10px 0;
}
.fx-card {
    border: 1px solid #ddd; border-radius: 12px; padding: 18px 20px;
    background: white; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    height: 100%;
}
.fx-card-header {
    font-size: 0.8rem; color: #888; font-weight: 600;
    letter-spacing: 0.5px; margin-bottom: 6px;
}
.fx-card-dir {
    font-size: 1.6rem; font-weight: 800; margin: 4px 0 8px 0;
}
.fx-card-band {
    background: #f0f4ff; border-radius: 6px; padding: 6px 12px;
    font-size: 0.9rem; display: inline-block; margin-bottom: 8px;
}
.fx-card-detail {
    font-size: 0.85rem; color: #444; line-height: 1.6;
}
.causal-box {
    background: #f8f8ff; border-left: 3px solid #7c83db;
    padding: 12px 16px; border-radius: 6px; margin: 8px 0;
    font-size: 0.9rem; line-height: 1.7;
}
hr { border-color: #e8e8e8 !important; }
/* 슬라이더 남색 */
div[data-testid="stSlider"] [role="slider"] { background: #2E75B6 !important; }
div[data-testid="stSlider"] [data-testid="stThumbValue"] { color: #2E75B6 !important; }
div[data-testid="stSlider"] [data-testid="stTickBar"] > div { background: #2E75B6 !important; }
div[data-testid="stSlider"] div[role="slider"]::before { background: #2E75B6 !important; }
/* 슬라이더 트랙 (채워진 부분) */
div[data-testid="stSlider"] div[data-baseweb="slider"] div[role="progressbar"] { background-color: #2E75B6 !important; }
div[data-testid="stSlider"] div[data-baseweb="slider"] div:nth-child(3) { background-color: #2E75B6 !important; }
div[data-testid="stSlider"] div[data-baseweb="slider"] div:nth-child(4) div { background-color: #2E75B6 !important; }
/* 라디오 버튼 남색 */
div[role="radiogroup"] label[data-checked="true"] { background: #2E75B6 !important; color: white !important; border-color: #2E75B6 !important; }
div[data-testid="stRadio"] input:checked + div { border-color: #2E75B6 !important; }
div[data-testid="stRadio"] input:checked + div::before { background-color: #2E75B6 !important; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════
# 데이터 수집
# ═══════════════════════════════════════════════════════
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_exchange_rates() -> pd.DataFrame:
    url = (
        f"http://ecos.bok.or.kr/api/StatisticSearch/"
        f"{BOK_API_KEY}/json/kr/1/10000/731Y001/D/{START_DATE}/{END_DATE}/"
    )
    res = requests.get(url, timeout=15)
    res.raise_for_status()
    body = res.json()

    if "StatisticSearch" not in body or "row" not in body["StatisticSearch"]:
        msg = body.get("RESULT", {}).get("MESSAGE", "알 수 없는 오류")
        raise RuntimeError(f"한국은행 API 오류: {msg}")

    rows = pd.DataFrame(body["StatisticSearch"]["row"])
    rows["TIME"] = pd.to_datetime(rows["TIME"])
    rows["DATA_VALUE"] = pd.to_numeric(rows["DATA_VALUE"], errors="coerce")

    usd = (
        rows[rows["ITEM_CODE1"] == "0000001"]
        .drop_duplicates("TIME").set_index("TIME")["DATA_VALUE"].rename("USD_KRW")
    )
    cny = (
        rows[rows["ITEM_NAME1"].str.contains("위안", na=False)]
        .drop_duplicates("TIME").set_index("TIME")["DATA_VALUE"].rename("CNY_KRW")
    )
    df = pd.concat([usd, cny], axis=1).ffill().dropna()
    df["USD_CNY"] = df["USD_KRW"] / df["CNY_KRW"]
    return df


def compute_stats(df: pd.DataFrame) -> dict:
    avg_3m = df.mean()
    lw = df.loc[LAST_WEEK_START:LAST_WEEK_END]
    avg_lw = lw.mean() if not lw.empty else df.iloc[-5:].mean()
    latest = lw.iloc[-1] if not lw.empty else df.iloc[-1]
    prev_week = df.loc[:pd.Timestamp(LAST_WEEK_START) - pd.Timedelta(days=1)].iloc[-5:]
    avg_prev = prev_week.mean() if not prev_week.empty else avg_3m

    def pct(cur, base):
        return (cur - base) / base * 100 if base else 0

    return {
        "latest": latest, "avg_3m": avg_3m, "avg_lw": avg_lw, "avg_prev": avg_prev,
        # 전주 평균 vs 3개월 평균
        "usd_vs_3m": pct(avg_lw["USD_KRW"], avg_3m["USD_KRW"]),
        # 전주 평균 vs 전전주 평균
        "usd_vs_pw": pct(avg_lw["USD_KRW"], avg_prev["USD_KRW"]),
        "cny_vs_3m": pct(avg_lw["CNY_KRW"], avg_3m["CNY_KRW"]),
        "cny_vs_pw": pct(avg_lw["CNY_KRW"], avg_prev["CNY_KRW"]),
        "cross_vs_3m": pct(avg_lw["USD_CNY"], avg_3m["USD_CNY"]),
        "cross_vs_pw": pct(avg_lw["USD_CNY"], avg_prev["USD_CNY"]),
    }


# ═══════════════════════════════════════════════════════
# 그래프
# ═══════════════════════════════════════════════════════
def build_chart(df: pd.DataFrame, events: dict = None, normalized: bool = False) -> go.Figure:
    """3개월 환율 추이 차트.
    normalized=False: 절대값 (USD/KRW 좌축, CNY/KRW 우축, USD/CNY 우축 토글)
    normalized=True : 시작일=100 정규화 인덱스 (단일 좌축, 세 통화 동시 비교)
    """
    if normalized:
        return _build_chart_normalized(df, events)

    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # 전주 구간 배경 강조
    lw_dates = df.loc[LAST_WEEK_START:LAST_WEEK_END].index
    if len(lw_dates) > 0:
        x0_ms = int(lw_dates[0].timestamp() * 1000) - 43200000
        x1_ms = int(lw_dates[-1].timestamp() * 1000) + 43200000
        fig.update_layout(
            shapes=[dict(type="rect", xref="x", yref="paper",
                         x0=x0_ms, x1=x1_ms, y0=0, y1=1,
                         fillcolor="rgba(0,0,0,0.06)", line=dict(width=0), layer="below")],
        )

    # 전체 기간 선
    fig.add_trace(go.Scatter(
        x=df.index, y=df["USD_KRW"], name="USD/KRW",
        line=dict(color="#2E75B6", width=2),
        hovertemplate="%{x|%m/%d}  %{y:,.2f}원<extra>USD/KRW</extra>",
    ), secondary_y=False)
    fig.add_trace(go.Scatter(
        x=df.index, y=df["CNY_KRW"], name="CNY/KRW",
        line=dict(color="#C00000", width=2),
        hovertemplate="%{x|%m/%d}  %{y:,.2f}원<extra>CNY/KRW</extra>",
    ), secondary_y=True)
    fig.add_trace(go.Scatter(
        x=df.index, y=df["USD_CNY"], name="USD/CNY (재정)",
        line=dict(color="#548235", width=1.5, dash="dot"),
        hovertemplate="%{x|%m/%d}  %{y:.4f}<extra>USD/CNY</extra>",
        visible="legendonly",
    ), secondary_y=True)

    # 이벤트 마커 (함수 내부에서 추가)
    if events:
        ev_x, ev_y, ev_labels = [], [], []
        for d_str, label in sorted(events.items()):
            ev = pd.Timestamp(d_str)
            if ev not in df.index:
                future = df.index[df.index >= ev]
                if len(future) == 0:
                    continue
                ev = future[0]
            ev_x.append(ev)
            ev_y.append(float(df.loc[ev, "USD_KRW"]))
            ev_labels.append(f"📌 {label}")
        if ev_x:
            fig.add_trace(go.Scatter(
                x=ev_x, y=ev_y, mode="markers",
                marker=dict(size=8, color="#2E75B6", symbol="circle", opacity=0.5),
                showlegend=False, name="이벤트",
                hovertemplate="%{text}<extra></extra>",
                text=ev_labels,
            ), secondary_y=False)

    def axis_range(s, m=0.12):
        lo, hi = s.min(), s.max()
        margin = (hi - lo) * m
        return [lo - margin, hi + margin]

    fig.update_yaxes(title_text="USD/KRW (원)", secondary_y=False,
                     range=axis_range(df["USD_KRW"]), tickformat=",",
                     title_font=dict(color="#2E75B6"))
    fig.update_yaxes(title_text="CNY/KRW (원)", secondary_y=True,
                     range=axis_range(df["CNY_KRW"]), tickformat=".2f",
                     title_font=dict(color="#C00000"))
    fig.update_layout(
        height=430, template="plotly_white", hovermode="x",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=10, r=10, t=50, b=10),
        xaxis=dict(tickformat="%m/%d", dtick=7 * 86400000),
    )
    return fig


def _build_chart_normalized(df: pd.DataFrame, events: dict = None) -> go.Figure:
    """시작일 = 100 정규화 인덱스 차트 (3통화 동일 스케일 비교)."""
    base_usd = float(df["USD_KRW"].iloc[0])
    base_cny = float(df["CNY_KRW"].iloc[0])
    base_cross = float(df["USD_CNY"].iloc[0])
    norm_usd = df["USD_KRW"] / base_usd * 100
    norm_cny = df["CNY_KRW"] / base_cny * 100
    norm_cross = df["USD_CNY"] / base_cross * 100

    fig = go.Figure()

    # 전주 음영
    lw_dates = df.loc[LAST_WEEK_START:LAST_WEEK_END].index
    shapes = []
    if len(lw_dates) > 0:
        x0_ms = int(lw_dates[0].timestamp() * 1000) - 43200000
        x1_ms = int(lw_dates[-1].timestamp() * 1000) + 43200000
        shapes.append(dict(type="rect", xref="x", yref="paper",
                           x0=x0_ms, x1=x1_ms, y0=0, y1=1,
                           fillcolor="rgba(0,0,0,0.06)", line=dict(width=0), layer="below"))
    # 기준선 (100)
    shapes.append(dict(type="line", xref="paper", yref="y",
                       x0=0, x1=1, y0=100, y1=100,
                       line=dict(color="#999", width=1, dash="dash"), layer="below"))

    fig.add_trace(go.Scatter(
        x=df.index, y=norm_usd, name="USD/KRW",
        line=dict(color="#2E75B6", width=2),
        customdata=df["USD_KRW"],
        hovertemplate="%{x|%m/%d}  지수 %{y:.2f}  (절대값 %{customdata:,.2f}원)<extra>USD/KRW</extra>",
    ))
    fig.add_trace(go.Scatter(
        x=df.index, y=norm_cny, name="CNY/KRW",
        line=dict(color="#C00000", width=2),
        customdata=df["CNY_KRW"],
        hovertemplate="%{x|%m/%d}  지수 %{y:.2f}  (절대값 %{customdata:,.2f}원)<extra>CNY/KRW</extra>",
    ))
    fig.add_trace(go.Scatter(
        x=df.index, y=norm_cross, name="USD/CNY (재정)",
        line=dict(color="#548235", width=2, dash="dot"),
        customdata=df["USD_CNY"],
        hovertemplate="%{x|%m/%d}  지수 %{y:.2f}  (절대값 %{customdata:.4f})<extra>USD/CNY</extra>",
    ))

    # 이벤트 마커 (USD/KRW 정규화 라인 위에)
    if events:
        ev_x, ev_y, ev_labels = [], [], []
        for d_str, label in sorted(events.items()):
            ev = pd.Timestamp(d_str)
            if ev not in df.index:
                future = df.index[df.index >= ev]
                if len(future) == 0:
                    continue
                ev = future[0]
            ev_x.append(ev)
            ev_y.append(float(norm_usd.loc[ev]))
            ev_labels.append(f"📌 {label}")
        if ev_x:
            fig.add_trace(go.Scatter(
                x=ev_x, y=ev_y, mode="markers",
                marker=dict(size=8, color="#2E75B6", symbol="circle", opacity=0.5),
                showlegend=False, name="이벤트",
                hovertemplate="%{text}<extra></extra>",
                text=ev_labels,
            ))

    # Y축 범위: 세 시리즈 합쳐서 ±5% 마진
    all_norm = pd.concat([norm_usd, norm_cny, norm_cross])
    lo, hi = all_norm.min(), all_norm.max()
    margin = (hi - lo) * 0.15

    fig.update_layout(
        height=430, template="plotly_white", hovermode="x",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        margin=dict(l=10, r=10, t=50, b=10),
        xaxis=dict(tickformat="%m/%d", dtick=7 * 86400000),
        yaxis=dict(
            title=dict(text=f"지수 (시작일 {df.index[0].strftime('%m/%d')} = 100)", font=dict(color="#444")),
            range=[lo - margin, hi + margin],
            tickformat=".1f",
        ),
        shapes=shapes,
    )
    return fig


# ═══════════════════════════════════════════════════════
# 외부 기관 리서치 서치 (Bing RSS + 기사 본문 크롤링)
# ═══════════════════════════════════════════════════════
@st.cache_data(ttl=3600, show_spinner=False)
def search_external_research() -> list[dict]:
    """Bing 뉴스 RSS로 환율 전망 기사를 찾고, 본문에서 기관별 밴드를 추출한다.
    NEWS_CUTOFF_DATE 이후 발행 기사만 수집한다."""
    import xml.etree.ElementTree as ET
    import urllib.parse
    from datetime import datetime as _dt
    from email.utils import parsedate_to_datetime

    results = []
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    cutoff = _dt.strptime(NEWS_CUTOFF_DATE, "%Y%m%d")

    rss_queries = ["주간환율전망", "환율 전망 주간 달러원"]

    seen_urls = set()
    institution_kw = {
        "하나은행": ["하나은행", "하나금융"],
        "우리은행": ["우리은행", "우리금융"],
        "삼성증권": ["삼성증권"],
        "미래에셋증권": ["미래에셋", "미래에셋증권"],
        "NH투자증권": ["NH투자", "NH증권"],
        "KB증권": ["KB증권", "KB국민"],
        "한국투자증권": ["한국투자", "한투증권"],
    }

    for query in rss_queries:
        rss_url = f"https://www.bing.com/news/search?q={urllib.parse.quote(query)}&format=rss"
        try:
            res = requests.get(rss_url, headers=headers, timeout=10)
            root = ET.fromstring(res.text)
            items = root.findall(".//item")[:5]

            for item in items:
                link = item.findtext("link", "")
                if not link or link in seen_urls:
                    continue

                # 날짜 필터: cutoff 이후 기사만
                pub_date_str = item.findtext("pubDate", "")
                if pub_date_str:
                    try:
                        pub_date = parsedate_to_datetime(pub_date_str).replace(tzinfo=None)
                        if pub_date < cutoff:
                            continue
                    except Exception:
                        pass

                seen_urls.add(link)

                try:
                    res2 = requests.get(link, headers=headers, timeout=8, allow_redirects=True)
                    soup = BeautifulSoup(res2.text, "html.parser")
                    body_tag = soup.select_one(
                        "#article-view-content-div, [itemprop=articleBody], "
                        ".article_body, #articeBody, .news_cnt_detail_wrap, "
                        ".article-body, .newsct_article"
                    )
                    if not body_tag:
                        continue
                    body_text = body_tag.get_text(separator=" ", strip=True).replace(",", "")

                    # 기관명 + 밴드 추출
                    for inst_name, keywords in institution_kw.items():
                        if not any(kw in body_text for kw in keywords):
                            continue
                        # 해당 기관명 주변에서 밴드 탐색
                        usd_matches = re.findall(r'(\d{4})\s*[~\-]\s*(\d{4})', body_text)
                        for lo, hi in usd_matches:
                            lo_n, hi_n = sorted([int(lo), int(hi)])
                            if 1000 <= lo_n <= 2000 and (hi_n - lo_n) < 200:
                                # 중복 체크
                                exists = any(r["source"] == inst_name for r in results)
                                if not exists:
                                    results.append({
                                        "source": inst_name, "currency": "USD/KRW",
                                        "lo": lo_n, "hi": hi_n,
                                    })
                                break
                except Exception:
                    continue
        except Exception:
            continue

    return results


# ═══════════════════════════════════════════════════════
# Claude API PDF 분석
# ═══════════════════════════════════════════════════════
@st.cache_data(ttl=86400, show_spinner=False)
def analyze_pdfs_with_claude(folder: str, week_type: str = "금주") -> dict:
    """Claude API로 PDF 통화별 변동요인을 분석한다.
    week_type: '금주' 또는 '전주'
    반환: {'USD/KRW': [요인1, 요인2, 요인3], 'CNY/KRW': [...], 'USD/CNY': [...]}
    """
    if not ANTHROPIC_API_KEY or not os.path.isdir(folder):
        return {}

    # 해당 주차 PDF 텍스트 수집
    pdf_texts = []
    for fname in sorted(os.listdir(folder)):
        if not fname.endswith(".pdf") or week_type not in fname:
            continue
        path = os.path.join(folder, fname)
        try:
            with pdfplumber.open(path) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages[:6])
            src = "국민은행" if "국민" in fname else ("신한은행" if "신한" in fname else fname)
            pdf_texts.append(f"=== {src} ===\n{text[:3000]}")
        except Exception:
            continue

    if not pdf_texts:
        return {}

    combined = "\n\n".join(pdf_texts)

    prompt = f"""다음은 은행 외환 전망 리포트입니다. 이를 분석하여 각 통화별 변동 요인을 정리해주세요.

[리포트 내용]
{combined}

[요청사항]
USD/KRW, CNY/KRW, USD/CNY(재정환율) 각 통화에 대해 주요 변동 요인 3가지씩을 다음 형식으로 작성해주세요:

- 각 요인은 "원인 → 메커니즘 → 결과" 형식의 짧은 한 문장 (40자 이내)
- 반드시 → 화살표를 사용하여 인과관계를 표현 (쉼표(,)로 나열하지 말 것)
- 핵심 키워드는 **굵게** 표시 (HTML <b> 태그 사용)
- 예시: "<b>이란 전쟁</b> 불확실성 → 위험회피 심리 → <b>달러 매수</b> 자극"

JSON 형식으로 정확히 응답해주세요:
{{
  "USD/KRW": ["요인1", "요인2", "요인3"],
  "CNY/KRW": ["요인1", "요인2", "요인3"],
  "USD/CNY": ["요인1", "요인2", "요인3"]
}}
JSON만 출력하고 다른 설명은 하지 마세요."""

    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=ANTHROPIC_API_KEY, timeout=30.0)
        msg = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=1500,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # JSON 추출
        import json as _json
        # ```json ... ``` 마크다운 제거
        if "```" in raw:
            raw = re.sub(r'```(?:json)?\s*', '', raw)
            raw = raw.replace("```", "").strip()
        result = _json.loads(raw)
        return result
    except Exception as e:
        st.warning(f"Claude API 분석 실패: {e}")
        return {}


# ═══════════════════════════════════════════════════════
# 3자 자료 수집 + 통합 분석 엔진
# ═══════════════════════════════════════════════════════
def _is_clean(s: str) -> bool:
    clean_chars = len(re.findall(r'[가-힣a-zA-Z0-9\s,.%()~·\-/]', s))
    return len(s) > 15 and clean_chars / max(len(s), 1) > 0.7


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_seoulfn_weekly_outlook() -> dict:
    """서울파이낸셜에서 제목에 '[주간환율전망]'이 포함된 최신 기사 1건."""
    search_url = (
        "https://www.seoulfn.com/news/articleList.html"
        "?sc_area=A&sc_word=%EC%A3%BC%EA%B0%84%ED%99%98%EC%9C%A8%EC%A0%84%EB%A7%9D"
    )
    headers = {"User-Agent": "Mozilla/5.0"}
    try:
        res = requests.get(search_url, headers=headers, timeout=10)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, "html.parser")
        target_link = None
        for tag in soup.select(".altlist-subject a, .list-titles a"):
            if "주간환율전망" in tag.get_text(strip=True):
                target_link = tag
                break
        if not target_link:
            return {"title": "기사 없음", "body": "[주간환율전망] 기사를 찾지 못했습니다.",
                    "url": "", "sentences": []}
        href = target_link.get("href", "")
        article_url = ("https://www.seoulfn.com" + href) if href.startswith("/") else href
        title = target_link.get_text(strip=True)
        res2 = requests.get(article_url, headers=headers, timeout=10)
        res2.raise_for_status()
        soup2 = BeautifulSoup(res2.text, "html.parser")
        body_tag = soup2.select_one("#article-view-content-div, [itemprop='articleBody']")
        if body_tag:
            paragraphs = body_tag.find_all("p")
            body = "\n\n".join(
                p.get_text(strip=True) for p in paragraphs if p.get_text(strip=True)
            ) if paragraphs else body_tag.get_text(separator="\n", strip=True)
        else:
            body = "본문 추출 실패"
        sentences = [s.strip() for s in body.replace("\n", " ").split(".")
                     if len(s.strip()) > 15]
        if len(body) > 1200:
            body = body[:1200] + "…"
        return {"title": title, "body": body, "url": article_url, "sentences": sentences}
    except Exception as e:
        return {"title": "오류", "body": str(e), "url": "", "sentences": []}


def run_integrated_analysis(folder: str) -> dict:
    """3자 자료를 하나로 병합하여 통합 전망 보고서 데이터를 생성한다."""
    by_source: dict[str, list[str]] = {}
    if os.path.isdir(folder):
        for fname in sorted(os.listdir(folder)):
            if not fname.lower().endswith(".pdf"):
                continue
            # 금주 PDF만 분석 (전주 PDF 제외)
            if "전주" in fname:
                continue
            path = os.path.join(folder, fname)
            with pdfplumber.open(path) as pdf:
                text = "\n".join(page.extract_text() or "" for page in pdf.pages[:8])
            sents = [s.strip() for s in text.replace("\n", " ").split(".")
                     if _is_clean(s.strip())]
            by_source[fname.replace(".pdf", "")] = sents

    news = fetch_seoulfn_weekly_outlook()
    if news["sentences"]:
        by_source["서울파이낸셜"] = news["sentences"]

    all_sentences = []
    for sents in by_source.values():
        all_sentences.extend(sents)

    # ── 키워드 그룹 ──
    kw_groups = {
        "현상": ["하락", "상승", "약세", "강세", "급등", "급락", "변동", "횡보", "반등"],
        "원인": ["때문", "영향", "요인", "우려", "기대", "리스크", "이슈", "결과", "관세",
                 "트럼프", "FOMC", "금리", "무역", "지정학", "인플레", "경상수지"],
        "전망": ["전망", "예상", "예측", "향후", "가능성", "것으로", "될 것", "판단"],
    }
    direction_kw = {
        "상승": ["상승", "강세", "오를", "높아", "반등", "확대", "상단", "상향"],
        "하락": ["하락", "약세", "내릴", "낮아", "둔화", "축소", "하단", "하향"],
    }

    # ── 카테고리별 상위 문장 선별 (출처 태그 없이) ──
    categorized = {"현상": [], "원인": [], "전망": []}
    for s in all_sentences:
        scores = {}
        for cat, kws in kw_groups.items():
            scores[cat] = sum(1 for kw in kws if kw in s)
        best_cat = max(scores, key=scores.get)
        if scores[best_cat] > 0:
            categorized[best_cat].append((s + ".", scores[best_cat]))

    for cat in categorized:
        categorized[cat].sort(key=lambda x: x[1], reverse=True)
        seen, unique = set(), []
        for text, _ in categorized[cat]:
            short = text[:30]
            if short not in seen:
                seen.add(short)
                unique.append(text)
        categorized[cat] = unique[:3]

    # ── 자료 간 견해 차이 감지 ──
    has_divergence = False
    divergence_note = ""
    if len(by_source) >= 2:
        directions = []
        for src_name, sents in by_source.items():
            src_text = " ".join(sents)
            bull = sum(1 for kw in direction_kw["상승"] if kw in src_text)
            bear = sum(1 for kw in direction_kw["하락"] if kw in src_text)
            if bull > bear:
                directions.append((src_name, "상승"))
            elif bear > bull:
                directions.append((src_name, "하락"))
            else:
                directions.append((src_name, "보합"))
        unique_dirs = set(d[1] for d in directions)
        if len(unique_dirs) > 1:
            has_divergence = True
            parts = [f"{d[0]}은 {d[1]}" for d in directions]
            divergence_note = f"※ 시장 견해가 갈리고 있음: {', '.join(parts)} 전망."

    # ── 통합 방향성 결론 ──
    full_text = " ".join(all_sentences)
    total_bull = sum(1 for kw in direction_kw["상승"] if kw in full_text)
    total_bear = sum(1 for kw in direction_kw["하락"] if kw in full_text)
    if total_bull > total_bear * 1.3:
        conclusion_dir = "상승(원화 약세)"
    elif total_bear > total_bull * 1.3:
        conclusion_dir = "하락(원화 강세)"
    else:
        conclusion_dir = "보합(혼조)"

    # ── 경제지표 추출 ──
    indicator_map = {
        "CPI": {"name": "미국 CPI", "keywords": ["CPI", "소비자물가", "인플레"]},
        "고용": {"name": "미국 고용지표", "keywords": ["고용", "비농업", "실업률", "NFP", "일자리"]},
        "FOMC": {"name": "FOMC/연준", "keywords": ["FOMC", "연준", "파월", "금리결정", "기준금리"]},
        "PMI": {"name": "PMI(제조업)", "keywords": ["PMI", "제조업지수", "ISM"]},
        "관세": {"name": "미중 관세/무역", "keywords": ["관세", "무역전쟁", "무역갈등", "트럼프", "상호관세"]},
        "중국": {"name": "중국 경제지표", "keywords": ["중국 수출", "중국 수입", "위안화 고시", "PBOC", "중국 PMI"]},
        "수출": {"name": "한국 수출", "keywords": ["수출", "경상수지", "무역수지"]},
        "GDP": {"name": "GDP 성장률", "keywords": ["GDP", "성장률", "경제성장"]},
    }
    date_pattern = re.compile(r'(\d{1,2})[/월]\s*(\d{1,2})일?|\b(\d{1,2})일')

    indicators = []
    for key, info in indicator_map.items():
        related = [s for s in all_sentences if any(kw in s for kw in info["keywords"])]
        if not related:
            continue
        src_count = sum(1 for src_sents in by_source.values()
                        if any(any(kw in s for kw in info["keywords"]) for s in src_sents))
        date_str = ""
        for s in related:
            m = date_pattern.search(s)
            if m:
                if m.group(1) and m.group(2):
                    date_str = f"{m.group(1)}/{m.group(2)}"
                elif m.group(3):
                    date_str = f"4/{m.group(3)}"
                break
        indicators.append({
            "name": info["name"], "mentions": len(related),
            "sources": src_count, "date": date_str or "미정",
        })
    indicators.sort(key=lambda x: (x["sources"], x["mentions"]), reverse=True)

    # ── 밴드 추출 (텍스트 + 테이블 + 제목에서 종합) ──
    band_context_kw = ["전망", "예상", "밴드", "레인지", "범위", "금주", "이번주", "이번 주",
                       "예측", "타겟", "목표", "저항", "지지"]
    usd_ranges, cny_ranges = [], []

    # (A) 텍스트에서 추출
    for s in all_sentences:
        if not re.search(r'\d[\d,]*\s*[~\-]\s*\d[\d,]*', s):
            continue
        if not any(kw in s for kw in band_context_kw):
            continue
        # 콤마 제거 후 매칭 (1,495~1,525 → 1495~1525)
        s_clean = s.replace(",", "")
        matches = re.findall(r'(\d{3,4})\s*[~\-]\s*(\d{3,4})', s_clean)
        for lo, hi in matches:
            lo_n, hi_n = sorted([int(lo), int(hi)])
            if 1000 <= lo_n <= 2000:
                usd_ranges.append((lo_n, hi_n))
            elif 150 <= lo_n <= 300:
                cny_ranges.append((lo_n, hi_n))

    # (B) PDF 테이블에서 직접 추출
    if os.path.isdir(folder):
        for fname in sorted(os.listdir(folder)):
            if not fname.lower().endswith(".pdf"):
                continue
            path = os.path.join(folder, fname)
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages[:4]:
                    tables = page.extract_tables() or []
                    for tbl in tables:
                        flat = " ".join(
                            str(cell).replace(",", "") for row in tbl
                            for cell in (row or []) if cell
                        )
                        tbl_matches = re.findall(r'(\d{3,4})\s*[~\-]\s*(\d{3,4})', flat)
                        for lo, hi in tbl_matches:
                            lo_n, hi_n = sorted([int(lo), int(hi)])
                            if 1000 <= lo_n <= 2000 and (hi_n - lo_n) < 200:
                                usd_ranges.append((lo_n, hi_n))
                            elif 150 <= lo_n <= 300 and (hi_n - lo_n) < 30:
                                cny_ranges.append((lo_n, hi_n))

                    # 제목/헤더에서도 추출 (예: "금주 1,495~1,525원 전망")
                    page_text = (page.extract_text() or "").replace(",", "")
                    title_matches = re.findall(
                        r'(?:금주|이번\s*주|주간|예상|전망)[^\n]{0,20}?(\d{3,4})\s*[~\-]\s*(\d{3,4})',
                        page_text
                    )
                    for lo, hi in title_matches:
                        lo_n, hi_n = sorted([int(lo), int(hi)])
                        if 1000 <= lo_n <= 2000 and (hi_n - lo_n) < 200:
                            usd_ranges.append((lo_n, hi_n))
                        elif 150 <= lo_n <= 300 and (hi_n - lo_n) < 30:
                            cny_ranges.append((lo_n, hi_n))

    band = {}
    if usd_ranges:
        band["USD/KRW"] = (
            sum(r[0] for r in usd_ranges) // len(usd_ranges),
            sum(r[1] for r in usd_ranges) // len(usd_ranges),
        )
    if cny_ranges:
        band["CNY/KRW"] = (
            sum(r[0] for r in cny_ranges) // len(cny_ranges),
            sum(r[1] for r in cny_ranges) // len(cny_ranges),
        )
    # USD/CNY 재정환율 밴드: USD/KRW, CNY/KRW 밴드가 둘 다 있으면 역산
    if "USD/KRW" in band and "CNY/KRW" in band:
        usd_lo, usd_hi = band["USD/KRW"]
        cny_lo, cny_hi = band["CNY/KRW"]
        cross_lo = round(usd_lo / cny_hi, 2)
        cross_hi = round(usd_hi / cny_lo, 2)
        band["USD/CNY"] = (cross_lo, cross_hi)

    # ── 통화별 심층 분석 (PDF 정독 기반) ──
    currency_kw = {
        "USD_KRW": {
            "detect": ["달러", "USD", "원/달러", "달러원", "원달러"],
            # 환율 상승(원화 약세) = bull
            "bull": ["상방", "상방우세", "상방압력", "상단", "환율상승", "원화약세", "원화절하",
                     "달러강세", "상승압력", "쏠림", "상회"],
            # 환율 하락(원화 강세) = bear
            "bear": ["하방", "하방우세", "하방압력", "하단", "환율하락", "원화강세", "원화절상",
                     "달러약세", "하락압력", "하회"],
        },
        "CNY_KRW": {
            "detect": ["위안", "CNY", "원/위안", "위안화", "중국"],
            "bull": ["위안강세", "위안화강세", "위안화반등", "위안화절상", "위안화상승",
                     "중국경기회복", "중국수출호조", "내수회복"],
            "bear": ["위안약세", "위안화약세", "위안화절하", "위안화하락", "위안화둔감",
                     "내수둔화", "내수부진", "중국경기둔화", "중국리스크",
                     "안정적흐름", "안정적관리", "둔감"],
        },
        "USD_CNY": {
            "detect": ["재정환율", "USD/CNY", "달러/위안", "달러위안"],
            "bull": ["위안절하", "위안약세", "달러강세"],
            "bear": ["위안절상", "위안강세", "달러약세", "안정적관리", "안정적흐름", "둔감"],
        },
    }

    currency_analysis = {}
    for cur_key, kws in currency_kw.items():
        related = [s for s in all_sentences if any(kw in s for kw in kws["detect"])]

        # 환율 전용 키워드로 방향성 판단 (띄어쓰기 무시)
        related_joined = " ".join(related).replace(" ", "")
        bull_score = sum(1 for kw in kws["bull"] if kw.replace(" ", "") in related_joined)
        bear_score = sum(1 for kw in kws["bear"] if kw.replace(" ", "") in related_joined)

        if bull_score > bear_score:
            direction = "상승"
            arrow = "↑"
            color = "#C00000"
        elif bear_score > bull_score:
            direction = "하락"
            arrow = "↓"
            color = "#4A90D9"
        else:
            direction = "보합"
            arrow = "→"
            color = "#856404"

        # 인과관계 문장 추출 (원인 키워드 + 해당 통화 동시 언급)
        cause_kws = ["때문", "영향", "요인", "기대", "우려", "리스크", "결과", "이슈",
                     "부진", "확대", "축소", "둔화", "회복"]
        causal_sentences = []
        for s in related:
            if any(ck in s for ck in cause_kws) and len(s) > 20:
                causal_sentences.append(s + ".")
        # 중복 제거, 상위 2개
        seen_c, unique_c = set(), []
        for cs in causal_sentences:
            short = cs[:25]
            if short not in seen_c:
                seen_c.add(short)
                unique_c.append(cs)
        causal_sentences = unique_c[:2]

        # 핵심 변수 추출
        var_kws = {"관세": "관세", "트럼프": "트럼프", "금리": "금리", "FOMC": "FOMC",
                   "유가": "유가", "무역": "무역", "지정학": "지정학", "인플레": "인플레",
                   "경상수지": "경상수지", "PBOC": "PBOC", "내수": "내수"}
        key_vars = [v for k, v in var_kws.items()
                    if any(k in s for s in related)][:3]

        # USD/CNY 특수 처리: 절상/절하
        if cur_key == "USD_CNY":
            if direction == "상승":
                sub_label = "위안화 절하"
            elif direction == "하락":
                sub_label = "위안화 절상"
            else:
                sub_label = "방향성 혼조"
        else:
            sub_label = ""

        # 상방/하방 변수 분류
        upside_kws = ["리스크", "지정학", "관세", "트럼프", "유가", "인플레", "매파", "강세"]
        downside_kws = ["순매수", "유입", "완화", "회복", "둔화", "비둘기", "절상", "수출"]
        upside_vars, downside_vars = [], []
        for s in related:
            for kw in upside_kws:
                if kw in s and kw not in upside_vars:
                    upside_vars.append(kw)
            for kw in downside_kws:
                if kw in s and kw not in downside_vars:
                    downside_vars.append(kw)

        # 전문가 코멘트 생성 (PDF 논거 + 밴드 수치 반영)
        expert_comment = ""
        if cur_key == "USD_KRW":
            top_cause = causal_sentences[0].split(".")[-2] if causal_sentences else ""
            cause_phrase = top_cause.strip() if top_cause and len(top_cause) > 10 else "대외 불확실성 확대"
            if direction == "상승":
                expert_comment = f"{cause_phrase}에 따른 달러 강세 기조로 상방 압력 우위 예상."
            elif direction == "하락":
                expert_comment = f"{cause_phrase} 영향으로 원화 강세 전환 가능성, 하단 테스트 예상."
            else:
                expert_comment = f"상승·하락 요인 혼재로 박스권 흐름 전망."

        elif cur_key == "CNY_KRW":
            cny_cause = ""
            for s in related:
                if any(kw in s for kw in ["내수", "부진", "둔화", "수출", "회복"]):
                    cny_cause = s.strip()
                    break
            if not cny_cause:
                cny_cause = "중국 경기 불확실성"
            if direction == "하락":
                expert_comment = f"위안화 프록시 통화인 원화의 동반 약세 속 하방 압력 우위. 원인: {cny_cause[:60]}."
            elif direction == "상승":
                expert_comment = f"위안화 반등 기대에 따른 상방 시도 예상. 배경: {cny_cause[:60]}."
            else:
                expert_comment = f"위안화 방향성 혼조 속 박스권 흐름 전망."

        elif cur_key == "USD_CNY":
            if direction == "상승":
                expert_comment = "달러 강세와 위안화 절하 압력 병존. 인민은행의 환율 방어 의지와 달러 강세 간의 힘겨루기 구간."
            elif direction == "하락":
                expert_comment = "인민은행의 적극적 개입 기대로 위안화 절상 압력. 달러 약세 동조 시 하락폭 확대 가능."
            else:
                expert_comment = "인민은행 환율 방어 의지와 달러 강세 간 균형 구간. 방향성 제한적."

        currency_analysis[cur_key] = {
            "direction": direction,
            "arrow": arrow,
            "color": color,
            "causal": causal_sentences,
            "key_vars": key_vars,
            "sub_label": sub_label,
            "expert_comment": expert_comment,
            "upside_vars": upside_vars[:3],
            "downside_vars": downside_vars[:3],
        }

    # ── 전문가 통합 관전평 (3자 자료 핵심 관통) ──
    source_views = []
    for src_name, sents in by_source.items():
        src_text = " ".join(sents)
        bull = sum(1 for kw in direction_kw["상승"] if kw in src_text)
        bear = sum(1 for kw in direction_kw["하락"] if kw in src_text)
        if bull > bear:
            source_views.append((src_name, "변동성 확대/상승"))
        elif bear > bull:
            source_views.append((src_name, "하향 안정/하락"))
        else:
            source_views.append((src_name, "혼조"))

    # 공통 키워드 추출
    common_vars_count = {}
    important_kws = ["관세", "트럼프", "금리", "FOMC", "지정학", "유가", "무역", "인플레", "PBOC", "내수"]
    for kw in important_kws:
        cnt = sum(1 for src_sents in by_source.values() if any(kw in s for s in src_sents))
        if cnt >= 2:
            common_vars_count[kw] = cnt
    common_top = sorted(common_vars_count.keys(), key=lambda k: common_vars_count[k], reverse=True)

    if source_views and common_top:
        views_str = ", ".join([f"{sv[0]}은 {sv[1]}" for sv in source_views])
        common_str = " · ".join(common_top[:3])
        expert_insight = f"{views_str}을 전망하고 있으나, 공통적으로 {common_str}을(를) 핵심 변수로 지목함."
    elif source_views:
        views_str = ", ".join([f"{sv[0]}은 {sv[1]}" for sv in source_views])
        expert_insight = f"{views_str}을 전망."
    else:
        expert_insight = "3자 자료에서 공통 핵심 변수를 추출하지 못했습니다."

    # ── 원픽 관전 포인트 ──
    if indicators:
        top = indicators[0]
        date_p = f"{top['date']} " if top["date"] != "미정" else ""
        band_p = ""
        if "CNY/KRW" in band:
            lo, hi = band["CNY/KRW"]
            band_p = f" (CNY/KRW 예상 밴드: {lo}~{hi}원)"
        watchpoint = (
            f"{date_p}{top['name']} 발표 결과가 금주 환율 방향성의 핵심 변수이며, "
            f"결과에 따라 위안화 환율 변동폭이 확대될 수 있으므로 환전 타이밍에 유의 필요{band_p}."
        )
    else:
        watchpoint = "금주 주요 경제지표 발표에 따른 환율 방향성 변동에 유의할 것."

    # ── 기관별 밴드 테이블 구성 (3통화) ──
    inst_bands = []

    def _add_band(src, currency, lo_n, hi_n):
        """중복 없이 inst_bands에 추가."""
        exists = any(b["기관"] == src and b["통화"] == currency for b in inst_bands)
        if not exists:
            if currency == "USD/CNY":
                inst_bands.append({"기관": src, "통화": currency, "주간 예상 범위": f"{lo_n}~{hi_n}"})
            else:
                inst_bands.append({"기관": src, "통화": currency, "주간 예상 범위": f"{lo_n:,}~{hi_n:,}"})

    def _classify_range(lo_n, hi_n):
        """숫자 범위를 통화로 분류."""
        if 1000 <= lo_n <= 2000 and (hi_n - lo_n) < 200:
            return "USD/KRW"
        elif 150 <= lo_n <= 300 and (hi_n - lo_n) < 30:
            return "CNY/KRW"
        elif 5 <= lo_n <= 10 and (hi_n - lo_n) < 3:
            return "USD/CNY"
        return None

    # PDF 테이블에서 기관별 밴드
    if os.path.isdir(folder):
        for fname in sorted(os.listdir(folder)):
            if not fname.lower().endswith(".pdf"):
                continue
            path = os.path.join(folder, fname)
            src_name = "국민은행" if "국민" in fname else ("신한은행" if "신한" in fname else fname.replace(".pdf", ""))
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages[:4]:
                    # 테이블 스캔
                    tables = page.extract_tables() or []
                    for tbl in tables:
                        flat = " ".join(str(c).replace(",", "") for row in tbl for c in (row or []) if c)
                        # 정수 범위 (USD/KRW, CNY/KRW)
                        for lo, hi in re.findall(r'(\d{3,4})\s*[~\-]\s*(\d{3,4})', flat):
                            lo_n, hi_n = sorted([int(lo), int(hi)])
                            cur = _classify_range(lo_n, hi_n)
                            if cur:
                                _add_band(src_name, cur, lo_n, hi_n)
                        # 소수점 범위 (USD/CNY: 6.870~6.920)
                        for lo, hi in re.findall(r'(\d\.\d{2,3})\s*[~\-]\s*(\d\.\d{2,3})', flat):
                            lo_f, hi_f = sorted([float(lo), float(hi)])
                            if 5 <= lo_f <= 10:
                                _add_band(src_name, "USD/CNY", lo_f, hi_f)

                    # 제목/본문에서도 추출
                    page_text = (page.extract_text() or "").replace(",", "")
                    for lo, hi in re.findall(r'(?:금주|주간|예상|전망)[^\n]{0,20}?(\d{3,4})\s*[~\-]\s*(\d{3,4})', page_text):
                        lo_n, hi_n = sorted([int(lo), int(hi)])
                        cur = _classify_range(lo_n, hi_n)
                        if cur:
                            _add_band(src_name, cur, lo_n, hi_n)

    # 서울파이낸셜 뉴스 밴드
    if news.get("sentences"):
        news_text = " ".join(news["sentences"]).replace(",", "")
        for lo, hi in re.findall(r'(\d{3,4})\s*[~\-]\s*(\d{3,4})', news_text):
            lo_n, hi_n = sorted([int(lo), int(hi)])
            cur = _classify_range(lo_n, hi_n)
            if cur:
                _add_band("서울파이낸셜", cur, lo_n, hi_n)

    # 외부 기관 (Bing RSS)
    ext = search_external_research()
    for e in ext:
        _add_band(e["source"], e["currency"], e["lo"], e["hi"])

    return {
        "news": news,
        "categorized": categorized,
        "conclusion_dir": conclusion_dir,
        "has_divergence": has_divergence,
        "divergence_note": divergence_note,
        "indicators": indicators,
        "band": band,
        "watchpoint": watchpoint,
        "all_sentences": all_sentences,
        "currency_analysis": currency_analysis,
        "expert_insight": expert_insight,
        "inst_bands": inst_bands,
    }


# ═══════════════════════════════════════════════════════
#  페이지 렌더링
# ═══════════════════════════════════════════════════════

st.markdown("## 📊 주간 환율 리포트 — 2026년 4월 2주차")
st.caption("수집 기간: 2026-01-03 ~ 2026-04-03 · 한국은행 ECOS API")

# ── 데이터 로딩 ──
with st.spinner("데이터 로딩 중..."):
    try:
        df = fetch_exchange_rates()
    except Exception as e:
        st.error(f"데이터 로딩 실패: {e}")
        st.stop()

stats = compute_stats(df)
latest = stats["latest"]

with st.spinner("기관별 리서치 통합 분석 중..."):
    report = run_integrated_analysis(DATA_DIR)

band = report["band"]
ca = report["currency_analysis"]

# ── 국민은행 금주 PDF에서 주간 밴드 추출 ──
inst_bands = report.get("inst_bands", [])
recent5 = df.tail(5)
_kb_usd, _kb_cross = None, None

# 파일명에 "금주"가 포함된 국민은행 PDF 찾기
kb_pdfs = [f for f in os.listdir(DATA_DIR) if "국민" in f and "금주" in f and f.endswith(".pdf")]
if kb_pdfs:
    with pdfplumber.open(os.path.join(DATA_DIR, kb_pdfs[0])) as pdf:
        _kb_text = "\n".join(p.extract_text() or "" for p in pdf.pages[:2]).replace(",", "")
        m_u = re.search(r'USDKRW\s+(\d{4})\s*[~\-]\s*(\d{4})', _kb_text)
        if m_u:
            _kb_usd = (int(m_u.group(1)), int(m_u.group(2)))
        m_c = re.search(r'USDCNY\s+(\d\.\d{2,3})\s*[~\-]\s*(\d\.\d{2,3})', _kb_text)
        if m_c:
            _kb_cross = (float(m_c.group(1)), float(m_c.group(2)))

band["USD/KRW"] = _kb_usd or (int(round(recent5["USD_KRW"].mean()) - 5), int(round(recent5["USD_KRW"].mean()) + 5))
band["USD/CNY"] = _kb_cross or (round(recent5["USD_CNY"].mean() - 0.03, 2), round(recent5["USD_CNY"].mean() + 0.03, 2))
usd_lo, usd_hi = band["USD/KRW"]
cross_lo, cross_hi = band["USD/CNY"]
cny_lo = int(round(usd_lo / cross_hi))
cny_hi = int(round(usd_hi / cross_lo))
band["CNY/KRW"] = (cny_lo, cny_hi)

# ── 방향성 판단 (1순위: 수치 → 2순위: PDF 보정) ──
avg_lw = stats["avg_lw"]

def _decide_dir(band_mid, prev_avg, threshold, cur_key):
    pct = (band_mid - prev_avg) / prev_avg if prev_avg else 0
    if pct > threshold:
        return "상승"
    elif pct < -threshold:
        return "하락"
    # 보합 → PDF 보정
    pdf_d = ca.get(cur_key, {}).get("direction", "보합")
    return pdf_d if pdf_d != "보합" else "보합"

usd_dir = _decide_dir((usd_lo + usd_hi) / 2, avg_lw["USD_KRW"], 0.005, "USD_KRW")
cny_dir = _decide_dir((cny_lo + cny_hi) / 2, avg_lw["CNY_KRW"], 0.004, "CNY_KRW")
cross_dir = _decide_dir((cross_lo + cross_hi) / 2, avg_lw["USD_CNY"], 0.002, "USD_CNY")
cross_sub = "위안화 절하" if cross_dir == "상승" else ("위안화 절상" if cross_dir == "하락" else "")

def _dir_color(d):
    return {"상승": "#C00000", "하락": "#4A90D9"}.get(d, "#2E8B57")

def _dir_arrow(d):
    return {"상승": "↑", "하락": "↓"}.get(d, "→")

def _dir_bg(d):
    return {"상승": ("#fdf2f2", "#C00000"), "하락": ("#f2f6fd", "#4A90D9")}.get(d, ("#f2faf5", "#2E8B57"))

def _forecast_card(label, band_str, direction, sub=""):
    color = _dir_color(direction)
    arrow = _dir_arrow(direction)
    sub_html = f' <span style="font-size:0.8rem;color:#888;">{sub}</span>' if sub else ""
    return (
        f'<div style="background:linear-gradient(135deg,#667eea0d,#764ba20d);'
        f'border:1px solid #ddd;border-radius:12px;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);">'
        f'<div style="font-size:0.85rem;color:#555;">{label}</div>'
        f'<div style="font-size:2rem;font-weight:700;color:#1a1a1a;margin:4px 0;">{band_str}</div>'
        f'<div style="font-size:1rem;color:{color};font-weight:600;">{arrow} {direction}{sub_html}</div>'
        f'</div>'
    )

# ── 사이드바: 외화 데이터 업로드 ──
with st.sidebar:
    st.subheader("외화 데이터 업로드")
    uploaded_cash = st.file_uploader("보유 현금 (Excel/CSV)", type=["xlsx", "csv"], key="cash")
    uploaded_ar = st.file_uploader("이번 주 채권 AR (Excel/CSV)", type=["xlsx", "csv"], key="ar")
    uploaded_ap = st.file_uploader("이번 주 채무 AP (Excel/CSV)", type=["xlsx", "csv"], key="ap")
    st.caption("컬럼: 통화, 금액, 보유환율(현금만)")

def _load_fx_data(uploaded, default_data, has_rate=False):
    if uploaded:
        d = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
        d["금액"] = [float(str(x).replace(",", "")) for x in d["금액"].tolist()]
        if has_rate and "보유환율" in d.columns:
            d["보유환율"] = [float(str(x).replace(",", "")) for x in d["보유환율"].tolist()]
        return d
    return default_data

cash_df = _load_fx_data(uploaded_cash, pd.DataFrame({
    "통화": ["USD", "CNY"], "금액": [5000000.0, 30000000.0], "보유환율": [1450.00, 198.50],
}), has_rate=True)

ar_df = _load_fx_data(uploaded_ar, pd.DataFrame({
    "통화": ["USD", "CNY"], "금액": [0.0, 0.0],
}))

ap_df = _load_fx_data(uploaded_ap, pd.DataFrame({
    "통화": ["USD", "CNY"], "금액": [0.0, 0.0],
}))

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 1: 주간 외환 관리 가이드라인
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
latest_date = df.index[-1].strftime("%Y-%m-%d")
st.markdown(f'<div class="section-header">1. 주간 외환 관리 가이드라인 ({latest_date} 기준)</div>', unsafe_allow_html=True)

# ── 순 노출액 계산 ──
rate_map = {"USD": float(latest["USD_KRW"]), "CNY": float(latest["CNY_KRW"])}
net_exposure = {}
cash_data = {}
for _, row in cash_df.iterrows():
    cur = str(row["통화"])
    cash_data[cur] = {"금액": float(row["금액"]), "보유환율": float(row.get("보유환율", 0))}
    net_exposure[cur] = float(row["금액"])

for _, row in ar_df.iterrows():
    cur = str(row["통화"])
    net_exposure[cur] = net_exposure.get(cur, 0) + float(row["금액"])

for _, row in ap_df.iterrows():
    cur = str(row["통화"])
    net_exposure[cur] = net_exposure.get(cur, 0) - float(row["금액"])

# ── 데이터 요약 텍스트 생성 (Claude에 전달) ──
cny_cash = cash_data.get("CNY", {}).get("금액", 0)
cny_book = cash_data.get("CNY", {}).get("보유환율", 0)
cny_net = net_exposure.get("CNY", 0)
usd_cash = cash_data.get("USD", {}).get("금액", 0)
usd_book = cash_data.get("USD", {}).get("보유환율", 0)
usd_net = net_exposure.get("USD", 0)
cny_mkt = rate_map.get("CNY", 0)
usd_mkt = rate_map.get("USD", 0)
cross_rate = usd_mkt / cny_mkt if cny_mkt else 0

cny_ar_val = float(ar_df[ar_df["통화"] == "CNY"]["금액"].sum()) if "CNY" in ar_df["통화"].values else 0
cny_ap_val = float(ap_df[ap_df["통화"] == "CNY"]["금액"].sum()) if "CNY" in ap_df["통화"].values else 0
usd_ar_val = float(ar_df[ar_df["통화"] == "USD"]["금액"].sum()) if "USD" in ar_df["통화"].values else 0
usd_ap_val = float(ap_df[ap_df["통화"] == "USD"]["금액"].sum()) if "USD" in ap_df["통화"].values else 0

summary_text = f"""[당사 외화 포지션 현황 — {latest_date} 기준]

■ CNY (위안화)
- 보유 현금: {cny_cash:,.0f} CNY (장부단가: {cny_book:,.2f}원)
- 이번 주 채권(AR): {cny_ar_val:,.0f} CNY
- 이번 주 채무(AP): {cny_ap_val:,.0f} CNY
- 순 노출액: {cny_net:,.0f} CNY

■ USD (달러)
- 보유 현금: {usd_cash:,.0f} USD (장부단가: {usd_book:,.2f}원)
- 이번 주 채권(AR): {usd_ar_val:,.0f} USD
- 이번 주 채무(AP): {usd_ap_val:,.0f} USD
- 순 노출액: {usd_net:,.0f} USD

■ 현재 시장 환율
- CNY/KRW: {cny_mkt:,.2f}원 (장부 대비 {(cny_mkt-cny_book)/cny_book*100 if cny_book else 0:+.2f}%)
- USD/KRW: {usd_mkt:,.2f}원 (장부 대비 {(usd_mkt-usd_book)/usd_book*100 if usd_book else 0:+.2f}%)
- USD/CNY (재정환율): {cross_rate:.4f}
- 금주 USD/KRW 전망: {usd_dir} ({usd_lo:,}~{usd_hi:,}원)
- 금주 CNY/KRW 전망: {cny_dir} ({cny_lo:,}~{cny_hi:,}원)"""

# ── 단기/장기 채권채무 분리 ──
def _split_term(d):
    """채권/채무를 단기(이번주)/장기(이후)로 분리."""
    if "구분" in d.columns:
        short = d[d["구분"].str.contains("단기|이번주", na=False)]
        long = d[d["구분"].str.contains("장기|이후", na=False)]
    else:
        short = d
        long = pd.DataFrame(columns=d.columns)
    return short, long

usd_ar_short = usd_ar_val
usd_ap_short = usd_ap_val
usd_ar_long = 0.0
usd_ap_long = 0.0
if "구분" in ar_df.columns:
    sm = ar_df[(ar_df["통화"] == "USD") & (ar_df["구분"].str.contains("장기|이후", na=False))]
    usd_ar_long = float(sm["금액"].sum()) if not sm.empty else 0.0
    usd_ar_short = usd_ar_val - usd_ar_long
if "구분" in ap_df.columns:
    sm = ap_df[(ap_df["통화"] == "USD") & (ap_df["구분"].str.contains("장기|이후", na=False))]
    usd_ap_long = float(sm["금액"].sum()) if not sm.empty else 0.0
    usd_ap_short = usd_ap_val - usd_ap_long

# CNY: 채권은 30일 이내/1년 이내 분리, 채무는 1년 이내 단일
cny_ar_short = cny_ar_val
cny_ar_long = 0.0
if "구분" in ar_df.columns:
    sm = ar_df[(ar_df["통화"] == "CNY") & (ar_df["구분"].str.contains("장기|이후", na=False))]
    cny_ar_long = float(sm["금액"].sum()) if not sm.empty else 0.0
    cny_ar_short = cny_ar_val - cny_ar_long

# ── 통화별 의사결정 분석 (Claude 직접 분석) ──
@st.cache_data(ttl=86400, show_spinner=False)
def _ai_decision(currency: str, payload: str) -> dict:
    """포지션 데이터를 Claude가 직접 보고 [현황-리스크-실무 제안]을 JSON으로 반환."""
    result = {"current": "", "risks": [], "actions": [], "error": ""}
    if not ANTHROPIC_API_KEY:
        result["error"] = "API 키 미설정"
        return result
    try:
        from anthropic import Anthropic
        import json as _json
        client = Anthropic(api_key=ANTHROPIC_API_KEY, timeout=30.0)
        msg = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=1500,
            system=f"""너는 글로벌 의류 기업 F&F의 외환 트레저리 전략가야.
{currency} 포지션 데이터를 직접 분석해서 [현황 - 리스크 - 실무 제안] 의사결정 분석을 작성해.

[출력 포맷] 정확히 다음 JSON만 출력 (마크다운 코드블록 없이):
{{
  "current": "한 단락 현황 요약 (보유/채권/채무/평가손익/순노출 종합. <b>HTML 강조</b> 가능)",
  "risks": ["리스크 1 (한 줄)", "리스크 2", "..."],
  "actions": ["① 제안 1 (한 줄)", "② 제안 2", "③ 제안 3 (선택)"]
}}

[분석 지침]
- 결론부터 말하고 근거 숫자(금액·환율·%)를 반드시 인용
- 평가이익/손실의 강도, 순노출 방향, 환율 민감도(10원 변동 시 영향), 채무 부담을 종합 판단
- 리스크는 2~3개, 실무 제안은 2~3개로 압축
- 각 항목 한국어 한 줄 (최대 60자)
- 환전 권고 시 구체적 비중(%)과 금액을 제시
- 평가손실 구간이면 보유 유지·환전 보류 권고
- 단기 채무가 보유+채권을 초과하면 즉시 매수 경고
- 한국어. 다른 텍스트 출력 금지.""",
            messages=[{"role": "user", "content": payload}],
        )
        text = msg.content[0].text.strip()
        # ```json ... ``` 마크다운 제거
        if text.startswith("```"):
            text = text.strip("`")
            if text.startswith("json"):
                text = text[4:]
            text = text.strip()
        parsed = _json.loads(text)
        result["current"] = parsed.get("current", "")
        result["risks"] = parsed.get("risks", []) or []
        result["actions"] = parsed.get("actions", []) or []
        return result
    except Exception as e:
        result["error"] = str(e)
        return result


def _render_ai_decision(decision: dict):
    """AI 의사결정 결과를 통일된 박스로 렌더."""
    if decision.get("error"):
        st.markdown(
            f'<div style="margin-top:16px;padding:12px 16px;background:#fdf2f2;border-left:4px solid #C00000;border-radius:6px;font-size:0.9rem;">'
            f'⚠️ <b>AI 의사결정 분석 실패</b>: {decision["error"]}'
            f'</div>',
            unsafe_allow_html=True,
        )
        return
    risks_html = "".join(f"&nbsp;&nbsp;&nbsp;&nbsp;• {r}<br>" for r in decision["risks"])
    actions_html = "".join(f"&nbsp;&nbsp;&nbsp;&nbsp;{a}<br>" for a in decision["actions"])
    st.markdown(
        f'<div style="margin-top:16px;padding:16px 20px;background:#fafbff;border:1px solid #d6d9e3;border-radius:8px;font-size:0.92rem;line-height:1.7;">'
        f'<div style="font-weight:700;font-size:1.0rem;margin-bottom:10px;color:#2E75B6;">📋 의사결정 분석 <span style="font-size:0.78rem;color:#888;font-weight:400;">(Claude AI)</span></div>'
        f'<div style="margin-bottom:10px;"><b style="color:#333;">▸ 현황</b><br>&nbsp;&nbsp;&nbsp;&nbsp;{decision["current"]}</div>'
        f'<div style="margin-bottom:10px;"><b style="color:#C00000;">▸ 리스크</b><br>{risks_html}</div>'
        f'<div><b style="color:#2E8B57;">▸ 실무 제안</b><br>{actions_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── UI 렌더링 ──

# 모듈 레벨에서 USD/CNY 의사결정 분석 미리 계산 (탭 + HTML 다운로드 양쪽에서 재사용)
_g_cash_pnl = (usd_mkt - usd_book) * usd_cash if usd_book else 0
_g_pnl_pct = (usd_mkt - usd_book) / usd_book * 100 if usd_book else 0
_g_net_exposure = usd_cash + usd_ar_short + usd_ar_long - usd_ap_short - usd_ap_long
_g_net_mkt_krw = _g_net_exposure * usd_mkt
_g_usd_liquidity = usd_cash + usd_ar_short - usd_ap_short
_g_sens_per_10 = abs(_g_net_exposure) * 10
_g_usd_payload = (
    f"[USD 포지션 현황]\n"
    f"- 보유 현금: {usd_cash:,.0f} USD (장부 평균 {usd_book:,.2f}원, 당일 매매기준율 {usd_mkt:,.2f}원)\n"
    f"- 평가손익: {_g_cash_pnl:+,.0f}원 ({_g_pnl_pct:+.2f}%) → 백만원 단위 {_g_cash_pnl/1_000_000:+,.0f}\n"
    f"- 미결 채권 (AR): 단기 {usd_ar_short:,.0f} / 장기 {usd_ar_long:,.0f} USD\n"
    f"- 미결 채무 (AP): 단기 {usd_ap_short:,.0f} / 장기 {usd_ap_long:,.0f} USD\n"
    f"- 단기 유동성 (현금+단기AR-단기AP): {_g_usd_liquidity:+,.0f} USD\n"
    f"- 순 노출금액 (현금+AR-AP): {_g_net_exposure:+,.0f} USD = {_g_net_mkt_krw/1_000_000:+,.0f}백만 원\n"
    f"- 환율 민감도: 환율 10원 변동 시 약 {_g_sens_per_10/10_000:,.0f}만 원 손익 변동\n\n"
    f"[금주 USD/KRW 전망]\n"
    f"- 밴드: {usd_lo:,} ~ {usd_hi:,}원 (중간 {(usd_lo+usd_hi)/2:,.1f})\n"
    f"- 방향: {usd_dir}"
)

_g_cny_cash_pnl = (cny_mkt - cny_book) * cny_cash if cny_book else 0
_g_cny_pnl_pct = (cny_mkt - cny_book) / cny_book * 100 if cny_book else 0
_g_cny_net = cny_cash + cny_ar_val - cny_ap_val
_g_cny_net_mkt_krw = _g_cny_net * cny_mkt
_g_cny_sens_per_10 = abs(_g_cny_net) * 10
_g_cny_payload = (
    f"[CNY 포지션 현황]\n"
    f"- 보유 현금: {cny_cash:,.0f} CNY (장부 평균 {cny_book:,.2f}원, 당일 매매기준율 {cny_mkt:,.2f}원)\n"
    f"- 평가손익: {_g_cny_cash_pnl:+,.0f}원 ({_g_cny_pnl_pct:+.2f}%) → 백만원 단위 {_g_cny_cash_pnl/1_000_000:+,.0f}\n"
    f"- 미결 채권 (AR): 30일 이내 {cny_ar_short:,.0f} / 1년 이내 {cny_ar_long:,.0f} CNY (합계 {cny_ar_val:,.0f})\n"
    f"- 미결 채무 (AP): 1년 이내 {cny_ap_val:,.0f} CNY\n"
    f"- 순 노출금액 (현금+AR-AP): {_g_cny_net:+,.0f} CNY = {_g_cny_net_mkt_krw/1_000_000:+,.0f}백만 원\n"
    f"- 환율 민감도: 환율 10원 변동 시 약 {_g_cny_sens_per_10/10_000:,.0f}만 원 손익 변동\n\n"
    f"[금주 CNY/KRW 전망]\n"
    f"- CNY/KRW 밴드: {cny_lo:,} ~ {cny_hi:,}원 (중간 {(cny_lo+cny_hi)/2:,.1f})\n"
    f"- USD/CNY 재정환율 밴드: {cross_lo} ~ {cross_hi} (당일 {cross_rate:.4f})\n"
    f"- 환전 대상은 KRW(즉시 차익 확정) 또는 USD(미래 결제 대비)로 분기 판단 필요"
)

with st.spinner("Claude가 USD·CNY 의사결정 분석 중..."):
    g_usd_decision = _ai_decision("USD", _g_usd_payload)
    g_cny_decision = _ai_decision("CNY", _g_cny_payload)

# USD / CNY 탭 분리
tab_usd, tab_cny = st.tabs(["🇺🇸 USD 유동성 진단", "🇨🇳 CNY 수익 전략"])

def _position_table(currency, cash_amt, book_rate, mkt_rate, ar_short, ar_long, ap_short, ap_long, symbol="$"):
    """첨부 사진 양식의 보유현황 표."""
    book_krw = cash_amt * book_rate
    mkt_krw = cash_amt * mkt_rate
    pnl = (mkt_rate - book_rate) * cash_amt
    pnl_color = "#C00000" if pnl > 0 else "#4A90D9"
    return (
        f'<table style="width:100%;border-collapse:collapse;font-size:0.88rem;border:1px solid #ddd;">'
        f'<tr style="background:#f0f4ff;text-align:center;">'
        f'<th rowspan="2" style="padding:8px;border:1px solid #ddd;">날짜</th>'
        f'<th rowspan="2" style="padding:8px;border:1px solid #ddd;">통화</th>'
        f'<th rowspan="2" style="padding:8px;border:1px solid #ddd;">보유금액</th>'
        f'<th colspan="2" style="padding:8px;border:1px solid #ddd;">채권 (AR)</th>'
        f'<th colspan="2" style="padding:8px;border:1px solid #ddd;">채무 (AP)</th>'
        f'<th colspan="2" style="padding:8px;border:1px solid #ddd;">장부 기준</th>'
        f'<th colspan="2" style="padding:8px;border:1px solid #ddd;">당일 기준</th>'
        f'<th rowspan="2" style="padding:8px;border:1px solid #ddd;">외환차손익(원)</th>'
        f'</tr>'
        f'<tr style="background:#f0f4ff;text-align:center;">'
        f'<th style="padding:6px;border:1px solid #ddd;">단기</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">장기</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">단기</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">장기</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">보유 평균환율</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">매매기준율</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액</th>'
        f'</tr>'
        f'<tr style="text-align:center;">'
        f'<td style="padding:8px;border:1px solid #eee;">{latest_date}</td>'
        f'<td style="padding:8px;border:1px solid #eee;font-weight:700;">{currency}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{cash_amt:,.2f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{ar_short:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{ar_long:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{ap_short:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{ap_long:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{book_rate:,.2f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{book_krw:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{mkt_rate:,.2f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;">{mkt_krw:,.0f}</td>'
        f'<td style="padding:8px;border:1px solid #eee;text-align:right;font-weight:700;color:{pnl_color};">{pnl:+,.0f}</td>'
        f'</tr>'
        f'</table>'
    )

with tab_usd:
    cash_pnl = (usd_mkt - usd_book) * usd_cash if usd_book else 0
    usd_liquidity = usd_cash + usd_ar_short - usd_ap_short
    net_warn = usd_liquidity < 0
    net_exposure = usd_cash + usd_ar_short + usd_ar_long - usd_ap_short - usd_ap_long

    # 포맷터 ───────────────────────────────────────────
    def _f_amt_usd(v):
        if not v: return "-"
        return f"{v:,.0f}"
    def _f_rate(v):
        if not v: return "-"
        if abs(v - round(v)) < 1e-6:
            return f"{int(round(v)):,}"
        return f"{v:,.2f}"
    def _f_krw_mil(v):
        if not v: return "-"
        return f"{v / 1_000_000:,.0f}"
    def _f_pnl_mil(v):
        if not v: return "-"
        val = abs(v) / 1_000_000
        if v > 0:
            return f'<span style="color:#C00000;font-weight:700;">▲ {val:,.0f}</span>'
        return f'<span style="color:#4A90D9;font-weight:700;">▼ {val:,.0f}</span>'

    # 원화 환산 ──────────────────────────────────────
    cash_book_krw = usd_cash * usd_book if usd_book else 0
    cash_mkt_krw = usd_cash * usd_mkt
    ar_s_krw = usd_ar_short * usd_mkt
    ar_l_krw = usd_ar_long * usd_mkt
    ap_s_krw = usd_ap_short * usd_mkt
    ap_l_krw = usd_ap_long * usd_mkt
    net_mkt_krw = net_exposure * usd_mkt

    # 데이터 행: (항목, 구분, 외화금액, 장부환율, 장부원화, 당일환율, 당일원화, 현재기준 외환차손익, is_ap)
    rows = [
        ("보유현금", "", usd_cash, usd_book, cash_book_krw, usd_mkt, cash_mkt_krw, cash_pnl, False),
        ("미결채권", "단기 회수예정", usd_ar_short, 0, 0, usd_mkt, ar_s_krw, 0, False),
        ("미결채권", "장기 누적미결", usd_ar_long, 0, 0, usd_mkt, ar_l_krw, 0, False),
        ("미결채무", "단기 지급예정", -usd_ap_short, 0, 0, usd_mkt, -ap_s_krw, 0, True),
        ("미결채무", "장기 미결현황", -usd_ap_long, 0, 0, usd_mkt, -ap_l_krw, 0, True),
    ]

    pnl_pct = (usd_mkt - usd_book) / usd_book * 100 if usd_book else 0

    # 같은 항목명이 연속되면 rowspan으로 셀 병합
    labels = [r[0] for r in rows]
    rowspan_map = {}  # idx -> rowspan 값 (병합 시작 행)
    skip_label = set()  # 라벨 셀 출력 생략할 행
    i = 0
    while i < len(labels):
        j = i
        while j + 1 < len(labels) and labels[j + 1] == labels[i]:
            j += 1
        if j > i:
            rowspan_map[i] = j - i + 1
            for k in range(i + 1, j + 1):
                skip_label.add(k)
        i = j + 1

    rows_html = ""
    for idx, (label, sub, amt, book, book_krw, mkt, mkt_krw, pnl, is_ap) in enumerate(rows):
        amt_color = "color:#C00000;" if (is_ap and amt < 0) else ""
        if idx in skip_label:
            label_cell = ""
        else:
            rs = rowspan_map.get(idx, 1)
            label_cell = (
                f'<td rowspan="{rs}" style="padding:9px 12px;border:1px solid #eee;'
                f'font-weight:600;background:#f8f9fc;text-align:center;vertical-align:middle;">{label}</td>'
            )
        rows_html += (
            f'<tr>'
            f'{label_cell}'
            f'<td style="padding:9px 12px;border:1px solid #eee;color:#555;">{sub}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;{amt_color}">{_f_amt_usd(amt)}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_f_rate(book)}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_f_krw_mil(book_krw)}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_f_rate(mkt)}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_f_krw_mil(mkt_krw)}</td>'
            f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_f_pnl_mil(pnl)}</td>'
            f'</tr>'
        )

    # 순 노출금액 (Net Exposure) 강조 행
    net_amt_color = "color:#C00000;" if net_exposure < 0 else ""
    rows_html += (
        f'<tr style="background:#ececec;">'
        f'<td style="padding:13px 12px;border:1px solid #ccc;font-weight:700;font-size:1.05rem;">순 노출금액</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;color:#444;font-size:0.9rem;">현금 + 채권 − 채무</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;font-size:1.1rem;{net_amt_color}">{_f_amt_usd(net_exposure)}</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;color:#888;">-</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;color:#888;">-</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;">{_f_rate(usd_mkt)}</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;font-size:1.1rem;">{_f_krw_mil(net_mkt_krw)}</td>'
        f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-size:1.05rem;">{_f_pnl_mil(cash_pnl)}</td>'
        f'</tr>'
    )

    # USD 환율 정보 박스 (표 위)
    usd_fcst_mid = (usd_lo + usd_hi) / 2
    st.markdown(
        f'<div style="margin-bottom:10px;padding:10px 14px;background:#f0f4ff;border-radius:6px;font-size:0.85rem;color:#333;">'
        f'📊 <b>당일 매매기준율</b>: {usd_mkt:,.2f}원 &nbsp;|&nbsp; '
        f'<b>금주 전망 중간</b>: {usd_fcst_mid:,.1f}원 (밴드 {usd_lo:,}~{usd_hi:,})'
        f'</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem;border:1px solid #ddd;">'
        # 헤더 1행
        f'<tr style="background:#f0f4ff;text-align:center;">'
        f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">항목</th>'
        f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">구분</th>'
        f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">금액(USD)</th>'
        f'<th colspan="2" style="padding:10px;border:1px solid #ddd;">장부 기준</th>'
        f'<th colspan="2" style="padding:10px;border:1px solid #ddd;">당일 기준</th>'
        f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">현재기준 외환차손익<br>'
        f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
        f'</tr>'
        # 헤더 2행 (장부/당일 하위)
        f'<tr style="background:#f0f4ff;text-align:center;">'
        f'<th style="padding:6px;border:1px solid #ddd;">보유 평균환율</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액<br>'
        f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
        f'<th style="padding:6px;border:1px solid #ddd;">매매기준율</th>'
        f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액<br>'
        f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
        f'</tr>'
        f'{rows_html}'
        f'</table>'
        f'<div style="margin-top:6px;font-size:0.78rem;color:#888;">참고) 채권: 전기일 기준(단기 90일 내 / 장기 누적분) &nbsp;·&nbsp; 채무: 만기일 기준(단기 30일 내 / 잔여분)</div>',
        unsafe_allow_html=True
    )

    # ── 의사결정 분석 (모듈 레벨에서 미리 계산된 결과 렌더) ──
    _render_ai_decision(g_usd_decision)

with tab_cny:
    cny_pnl_total = (cny_mkt - cny_book) * cny_cash if cny_book else 0
    cny_pnl_pct = (cny_mkt - cny_book) / cny_book * 100 if cny_book else 0

    @st.fragment
    def _run_cny_tab():
        # 라디오는 시나리오 시뮬 표 바로 위로 이동 (아래쪽에서 사용)
        # 여기서는 기본값으로 초기화만
        sim_target = st.session_state.get("cny_target", "KRW")
        sim_pct = st.session_state.get("cny_pct", "100%")

        pct_val = int(sim_pct.replace("%", "")) / 100
        sim_amt = cny_cash * pct_val

        # 금주 전망 중간값 (시뮬 적용)
        cny_fcst = (cny_lo + cny_hi) / 2
        usd_fcst = (usd_lo + usd_hi) / 2
        cross_fcst = (cross_lo + cross_hi) / 2

        # 장부단가 라벨 (USD 선택 시 재정환율로 환산)
        if sim_target == "KRW":
            book_display = cny_book
            book_label = "KRW/CNY"
        else:
            # CNY 장부단가(KRW)를 USD/CNY로 변환: 보유 CNY → USD 환산 평균 단가
            # 장부 USD 단가 = 장부 CNY/KRW ÷ 현재 USD/KRW (대략)
            book_display = (1 / (cny_book / usd_mkt)) if (cny_book and usd_mkt) else 0
            book_label = "USD/CNY"

        if sim_target == "KRW":
            sim_pnl = (cny_fcst - cny_book) * sim_amt if cny_book else 0
            sim_converted = sim_amt * cny_fcst
            converted_str = f"{sim_converted:,.0f} KRW"
            sim_label = f"환전 시뮬 ({sim_pct} → KRW)"
            sim_sub = f"환전 예정액 {sim_amt:,.0f} CNY"
            sim_rate_display = cny_fcst
        else:
            sim_usd = sim_amt / cross_fcst if cross_fcst else 0
            book_krw_val = sim_amt * cny_book
            mkt_usd_to_krw = sim_usd * usd_fcst
            sim_pnl = mkt_usd_to_krw - book_krw_val
            sim_converted = sim_usd
            converted_str = f"${sim_usd:,.2f}"
            sim_label = f"환전 시뮬 ({sim_pct} → USD)"
            sim_sub = f"환전 예정액 {sim_amt:,.0f} CNY → ${sim_usd:,.2f}"
            sim_rate_display = cross_fcst

        def _f_amt(v): return f"{v:,.0f}" if v else "-"
        def _f_rate(v): return f"{v:,.2f}" if v else "-"
        def _f_pnl(v):
            if v == 0: return "-"
            color = "#C00000" if v > 0 else "#4A90D9"
            return f'<span style="color:{color};font-weight:700;">{v:+,.0f}</span>'

        # 현재 평가손익 (환전 통화에 따라 분리)
        if sim_target == "KRW":
            cur_pnl = (cny_mkt - cny_book) * cny_cash if cny_book else 0
            cur_label = "현재 평가손익 (→ KRW)"
            cur_sub = "전량 환전 시"
            cur_rate_show = cny_mkt
        else:
            usd_equiv = cny_cash / cross_rate if cross_rate else 0
            book_krw_full = cny_cash * cny_book
            mkt_usd_to_krw_full = usd_equiv * usd_mkt
            cur_pnl = mkt_usd_to_krw_full - book_krw_full
            cur_label = "현재 평가손익 (→ USD)"
            cur_sub = f"전량 환전 시 ${usd_equiv:,.2f}"
            cur_rate_show = cross_rate

        # 장부단가 표시 포맷 (KRW: 1234.56, USD: 0.1234)
        def _f_book(v):
            if not v: return "-"
            return f"{v:,.4f}" if sim_target == "USD" else f"{v:,.2f}"

        # 적용환율 표시 포맷
        def _f_rate2(v):
            if not v: return "-"
            return f"{v:,.4f}" if sim_target == "USD" else f"{v:,.2f}"

        # 메인 표 (보유 현황 + 현재 평가손익)
        # 적용 환율 정보 박스 (표 위)
        if sim_target == "KRW":
            rate_info = (
                f'<b>당일 매매기준율</b>: {cny_mkt:,.2f}원 &nbsp;|&nbsp; '
                f'<b>금주 전망 중간</b>: {cny_fcst:,.1f}원 (밴드 {cny_lo}~{cny_hi})'
            )
        else:
            rate_info = (
                f'<b>당일 재정환율</b>: {cross_rate:.4f} &nbsp;|&nbsp; '
                f'<b>금주 전망 재정환율 중간</b>: {cross_fcst:.4f} (밴드 {cross_lo}~{cross_hi})'
            )
        st.markdown(
            f'<div style="margin-bottom:10px;padding:10px 14px;background:#f0f4ff;border-radius:6px;font-size:0.85rem;color:#333;">'
            f'📊 {rate_info}</div>',
            unsafe_allow_html=True
        )

        # ── 메인 보유 현황 표 (KRW 기준 고정) ──
        cny_cash_book_krw = cny_cash * cny_book if cny_book else 0
        cny_cash_mkt_krw = cny_cash * cny_mkt
        cny_ar_short_krw = cny_ar_short * cny_mkt
        cny_ar_long_krw = cny_ar_long * cny_mkt
        cny_ap_mkt_krw = cny_ap_val * cny_mkt
        cny_net = cny_cash + cny_ar_val - cny_ap_val
        cny_net_mkt_krw = cny_net * cny_mkt
        cny_cash_pnl = (cny_mkt - cny_book) * cny_cash if cny_book else 0

        def _fc_amt(v):
            if not v: return "-"
            return f"{v:,.0f}"
        def _fc_rate(v):
            if not v: return "-"
            if abs(v - round(v, 2)) < 1e-6 and v == int(v):
                return f"{int(v):,}"
            return f"{v:,.2f}"
        def _fc_krw_mil(v):
            if not v: return "-"
            return f"{v / 1_000_000:,.0f}"
        def _fc_pnl_mil(v):
            if not v: return "-"
            val = abs(v) / 1_000_000
            if v > 0:
                return f'<span style="color:#C00000;font-weight:700;">▲ {val:,.0f}</span>'
            return f'<span style="color:#4A90D9;font-weight:700;">▼ {val:,.0f}</span>'

        # 데이터 행: (항목, 구분, 외화금액, 장부환율, 장부원화, 당일환율, 당일원화, 현재기준 외환차손익, is_ap)
        cny_rows = [
            ("보유현금", "", cny_cash, cny_book, cny_cash_book_krw, cny_mkt, cny_cash_mkt_krw, cny_cash_pnl, False),
            ("미결채권", "단기 회수예정", cny_ar_short, 0, 0, cny_mkt, cny_ar_short_krw, 0, False),
            ("미결채권", "장기 누적미결", cny_ar_long, 0, 0, cny_mkt, cny_ar_long_krw, 0, False),
            ("미결채무", "장기 미결현황", -cny_ap_val, 0, 0, cny_mkt, -cny_ap_mkt_krw, 0, True),
        ]

        # 같은 항목명이 연속되면 rowspan으로 셀 병합
        cny_labels = [r[0] for r in cny_rows]
        cny_rowspan_map = {}
        cny_skip_label = set()
        i = 0
        while i < len(cny_labels):
            j = i
            while j + 1 < len(cny_labels) and cny_labels[j + 1] == cny_labels[i]:
                j += 1
            if j > i:
                cny_rowspan_map[i] = j - i + 1
                for k in range(i + 1, j + 1):
                    cny_skip_label.add(k)
            i = j + 1

        cny_rows_html = ""
        for idx, (label, sub, amt, book, book_krw_v, mkt, mkt_krw_v, pnl, is_ap) in enumerate(cny_rows):
            amt_color = "color:#C00000;" if (is_ap and amt < 0) else ""
            if idx in cny_skip_label:
                label_cell = ""
            else:
                rs = cny_rowspan_map.get(idx, 1)
                label_cell = (
                    f'<td rowspan="{rs}" style="padding:9px 12px;border:1px solid #eee;'
                    f'font-weight:600;background:#f8f9fc;text-align:center;vertical-align:middle;">{label}</td>'
                )
            cny_rows_html += (
                f'<tr>'
                f'{label_cell}'
                f'<td style="padding:9px 12px;border:1px solid #eee;color:#555;">{sub}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;{amt_color}">{_fc_amt(amt)}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_fc_rate(book)}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_fc_krw_mil(book_krw_v)}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_fc_rate(mkt)}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_fc_krw_mil(mkt_krw_v)}</td>'
                f'<td style="padding:9px 12px;border:1px solid #eee;text-align:right;">{_fc_pnl_mil(pnl)}</td>'
                f'</tr>'
            )

        # 순 노출금액 강조 행
        cny_net_color = "color:#C00000;" if cny_net < 0 else ""
        cny_rows_html += (
            f'<tr style="background:#ececec;">'
            f'<td style="padding:13px 12px;border:1px solid #ccc;font-weight:700;font-size:1.05rem;text-align:center;">순노출금액</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;color:#444;font-size:0.9rem;">현금 + 채권 − 채무</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;font-size:1.1rem;{cny_net_color}">{_fc_amt(cny_net)}</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;color:#888;">-</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;color:#888;">-</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;">{_fc_rate(cny_mkt)}</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-weight:700;font-size:1.1rem;">{_fc_krw_mil(cny_net_mkt_krw)}</td>'
            f'<td style="padding:13px 12px;border:1px solid #ccc;text-align:right;font-size:1.05rem;">{_fc_pnl_mil(cny_cash_pnl)}</td>'
            f'</tr>'
        )

        st.markdown(
            f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem;border:1px solid #ddd;">'
            # 헤더 1행
            f'<tr style="background:#f0f4ff;text-align:center;">'
            f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">항목</th>'
            f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">구분</th>'
            f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">금액(CNY)</th>'
            f'<th colspan="2" style="padding:10px;border:1px solid #ddd;">장부 기준</th>'
            f'<th colspan="2" style="padding:10px;border:1px solid #ddd;">당일 기준</th>'
            f'<th rowspan="2" style="padding:10px;border:1px solid #ddd;">현재기준 외환차손익<br>'
            f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
            f'</tr>'
            # 헤더 2행 (장부/당일 하위)
            f'<tr style="background:#f0f4ff;text-align:center;">'
            f'<th style="padding:6px;border:1px solid #ddd;">보유 평균환율</th>'
            f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액<br>'
            f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
            f'<th style="padding:6px;border:1px solid #ddd;">매매기준율</th>'
            f'<th style="padding:6px;border:1px solid #ddd;">원화환산금액<br>'
            f'<span style="font-size:0.72rem;font-weight:400;color:#666;">(백만원)</span></th>'
            f'</tr>'
            f'{cny_rows_html}'
            f'</table>'
            f'<div style="margin-top:6px;font-size:0.78rem;color:#888;">참고) 채권: 전기일 기준(단기 30일 내 / 장기 누적분) &nbsp;·&nbsp; 채무: 만기일 기준(잔여분)</div>',
            unsafe_allow_html=True
        )

        # ── 의사결정 분석 (모듈 레벨에서 미리 계산된 결과 렌더) ──
        _render_ai_decision(g_cny_decision)

        # ── 보유현금 환전 시뮬 표 ──
        st.markdown(
            "<div style='margin-top:18px;font-weight:700;font-size:0.95rem;'>💱 보유현금 환전 시뮬레이션</div>",
            unsafe_allow_html=True,
        )
        opt1, opt2 = st.columns(2)
        with opt1:
            sim_target = st.radio("환전 통화", ["KRW", "USD"], horizontal=True, key="cny_target")
        with opt2:
            sim_pct = st.radio("환전 비중", ["30%", "50%", "70%", "100%"], horizontal=True, index=3, key="cny_pct")
        # 변경된 값 반영 — 표의 보유현금(cny_cash) × 환전 비중
        pct_val = int(sim_pct.replace("%", "")) / 100
        sim_amt = cny_cash * pct_val

        # 시나리오 계산 (3개)
        if sim_target == "KRW":
            sc_low_rate, sc_mid_rate, sc_hi_rate = float(cny_lo), cny_fcst, float(cny_hi)
        else:
            sc_low_rate, sc_mid_rate, sc_hi_rate = float(cross_lo), cross_fcst, float(cross_hi)

        def _calc_pnl(rate):
            if sim_target == "KRW":
                return (rate - cny_book) * sim_amt if cny_book else 0
            sc_usd = sim_amt / rate if rate else 0
            return (sc_usd * usd_fcst) - (sim_amt * cny_book)

        def _row_html(label, rate, bg=""):
            pnl = _calc_pnl(rate)
            return (
                f'<tr style="{bg}">'
                f'<td style="padding:8px 12px;border:1px solid #eee;font-weight:600;background:#f8f9fc;">{label}</td>'
                f'<td style="padding:8px 12px;border:1px solid #eee;color:#555;">{sim_pct} 환전</td>'
                f'<td style="padding:8px 12px;border:1px solid #eee;text-align:right;">{_f_amt(sim_amt)}</td>'
                f'<td style="padding:8px 12px;border:1px solid #eee;text-align:right;">{_f_book(book_display)}</td>'
                f'<td style="padding:8px 12px;border:1px solid #eee;text-align:right;">{_f_rate2(rate)}</td>'
                f'<td style="padding:8px 12px;border:1px solid #eee;text-align:right;">{_f_pnl(pnl)}</td>'
                f'</tr>'
            )

        def _sim_table(rows_html):
            return (
                f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem;border:1px solid #ddd;margin-top:6px;">'
                f'<tr style="background:#fff7e6;text-align:center;">'
                f'<th style="padding:10px;border:1px solid #ddd;">시나리오</th>'
                f'<th style="padding:10px;border:1px solid #ddd;">구분</th>'
                f'<th style="padding:10px;border:1px solid #ddd;">금액(CNY)</th>'
                f'<th style="padding:10px;border:1px solid #ddd;">장부단가 ({book_label})</th>'
                f'<th style="padding:10px;border:1px solid #ddd;">적용환율 ({book_label})</th>'
                f'<th style="padding:10px;border:1px solid #ddd;">외환차손익(원)</th>'
                f'</tr>{rows_html}</table>'
            )

        # 메인: 평균(중간값)만 표시
        st.markdown(
            _sim_table(_row_html("평균 (중간값)", sc_mid_rate)),
            unsafe_allow_html=True
        )

        # 토글: 최저 / 평균 / 최고 시나리오 비교
        with st.expander("▶ 환율 변동 시나리오 상세보기"):
            st.markdown(
                _sim_table(
                    _row_html("최저 (밴드 하단)", sc_low_rate) +
                    _row_html("평균 (중간값)", sc_mid_rate) +
                    _row_html("최고 (밴드 상단)", sc_hi_rate)
                ),
                unsafe_allow_html=True
            )

        # 시뮬 환차손익 (강조용 변수)
        sim_pnl = _calc_pnl(sc_mid_rate)
        if sim_target == "KRW":
            converted_str = f"{sim_amt * sc_mid_rate:,.0f} KRW"
        else:
            converted_str = f"${sim_amt / sc_mid_rate if sc_mid_rate else 0:,.2f}"

        # 시뮬 환차손익 (선택 비중 기준)
        sim_pnl_pct = (sim_pnl / (sim_amt * cny_book) * 100) if (sim_amt * cny_book) else 0
        sim_pnl_color = "#C00000" if sim_pnl > 0 else "#4A90D9"
        st.markdown(
            f'<div style="margin-top:12px;padding:14px 18px;background:#fdf8f0;border-left:4px solid #e6a817;border-radius:6px;">'
            f'<div style="font-size:0.85rem;color:#666;">📈 예상 환차손익 ({sim_pct} → {sim_target})</div>'
            f'<div style="font-size:1.4rem;font-weight:800;color:{sim_pnl_color};margin-top:4px;">'
            f'{sim_pnl:+,.0f}원 <span style="font-size:0.9rem;color:#888;">({sim_pnl_pct:+.2f}%)</span></div>'
            f'<div style="font-size:0.85rem;color:#444;margin-top:6px;">'
            f'선택: {sim_pct} → {sim_target} 환전 시 → <b>{converted_str}</b> 확보'
            f'</div></div>',
            unsafe_allow_html=True
        )

    _run_cny_tab()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 2: 금주 환율 전망
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown(f'<div class="section-header">2. 금주 환율 전망 ({REPORT_WEEK_START[4:6]}/{REPORT_WEEK_START[6:]} ~ {REPORT_WEEK_END[4:6]}/{REPORT_WEEK_END[6:]})</div>', unsafe_allow_html=True)

# 금주 전망 밴드 카드
fc1, fc2, fc3 = st.columns(3)
fc1.markdown(_forecast_card("USD / KRW", f"{usd_lo:,} ~ {usd_hi:,} 원", usd_dir), unsafe_allow_html=True)
fc2.markdown(_forecast_card("CNY / KRW", f"{cny_lo:,} ~ {cny_hi:,} 원", cny_dir), unsafe_allow_html=True)
fc3.markdown(_forecast_card("USD / CNY (재정)", f"{cross_lo} ~ {cross_hi}", cross_dir, cross_sub), unsafe_allow_html=True)

st.divider()

# 통화별 분석 (통합 표 — PDF 기반 변동요인)
st.markdown("##### 📝 통화별 분석")
st.caption("국민은행 · 신한은행 금주 전망 PDF + 서울파이낸셜 종합")

def _dir_badge(d):
    c = _dir_color(d)
    a = _dir_arrow(d)
    return f'<span style="color:{c};font-weight:700;">{a} {d}</span>'

# Claude API로 금주 PDF 분석
with st.spinner("Claude AI가 금주 PDF를 분석하는 중..."):
    claude_factors = analyze_pdfs_with_claude(DATA_DIR, week_type="금주")

usd_factors = claude_factors.get("USD/KRW", ["분석 데이터 없음"])
cny_factors = claude_factors.get("CNY/KRW", ["분석 데이터 없음"])
cross_factors = claude_factors.get("USD/CNY", ["분석 데이터 없음"])

# 표 생성
def _factor_html(factors):
    return "<br>".join(f"• {f}" for f in factors)

# 전망 셀: 방향 + 변동금액 + 변동%
avg_lw = stats["avg_lw"]

def _outlook_cell(direction, band_lo, band_hi, prev_avg, fmt="krw"):
    mid = (band_lo + band_hi) / 2
    diff = mid - prev_avg
    color = _dir_color(direction)
    arrow = _dir_arrow(direction)
    if fmt == "krw":
        diff_str = f"{diff:+,.1f}원"
    else:
        diff_str = f"{diff:+.4f}"
    return (
        f'{arrow} <b>{direction}</b><br>'
        f'<span style="font-size:0.78rem;color:#888;">전주 평균 대비 {diff_str}</span>'
    )

usd_outlook = _outlook_cell(usd_dir, usd_lo, usd_hi, avg_lw["USD_KRW"])
cny_outlook = _outlook_cell(cny_dir, cny_lo, cny_hi, avg_lw["CNY_KRW"])
cross_outlook = _outlook_cell(cross_dir, cross_lo, cross_hi, avg_lw["USD_CNY"], fmt="cross")

st.markdown(
    f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem;border:1px solid #ddd;">'
    f'<tr style="background:#f0f4ff;">'
    f'<th style="padding:10px 14px;border:1px solid #ddd;text-align:center;width:12%;">통화</th>'
    f'<th style="padding:10px 14px;border:1px solid #ddd;text-align:center;width:18%;">전망</th>'
    f'<th style="padding:10px 14px;border:1px solid #ddd;text-align:center;width:18%;">예상 밴드</th>'
    f'<th style="padding:10px 14px;border:1px solid #ddd;text-align:center;width:52%;">변동 요인</th>'
    f'</tr>'
    f'<tr>'
    f'<td style="padding:10px;border:1px solid #eee;font-weight:700;text-align:center;white-space:nowrap;">USD/KRW</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;color:{_dir_color(usd_dir)};">{usd_outlook}</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;font-weight:600;white-space:nowrap;">{usd_lo:,}~{usd_hi:,}원</td>'
    f'<td style="padding:10px;border:1px solid #eee;line-height:1.8;">{_factor_html(usd_factors)}</td>'
    f'</tr>'
    f'<tr>'
    f'<td style="padding:10px;border:1px solid #eee;font-weight:700;text-align:center;white-space:nowrap;">CNY/KRW</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;color:{_dir_color(cny_dir)};">{cny_outlook}</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;font-weight:600;white-space:nowrap;">{cny_lo:,}~{cny_hi:,}원</td>'
    f'<td style="padding:10px;border:1px solid #eee;line-height:1.8;">{_factor_html(cny_factors)}</td>'
    f'</tr>'
    f'<tr>'
    f'<td style="padding:10px;border:1px solid #eee;font-weight:700;text-align:center;white-space:nowrap;">USD/CNY</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;color:{_dir_color(cross_dir)};">{cross_outlook}</td>'
    f'<td style="padding:10px;border:1px solid #eee;text-align:center;font-weight:600;white-space:nowrap;">{cross_lo}~{cross_hi}</td>'
    f'<td style="padding:10px;border:1px solid #eee;line-height:1.8;">{_factor_html(cross_factors)}</td>'
    f'</tr>'
    f'</table>',
    unsafe_allow_html=True,
)

# 뉴스 원문 (통화별 분석 표 바로 아래)
news = report["news"]
if news.get("url"):
    with st.expander(f"📰 {news.get('title', '')[:45]}…"):
        if news.get("body"):
            st.markdown(news["body"][:600] + "…" if len(news.get("body", "")) > 600 else news.get("body", ""))
        st.caption(f"[기사 원문]({news['url']})")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 3: 환율 추이 및 전주 동향
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown('<div class="section-header">3. 환율 추이 및 전주 동향</div>', unsafe_allow_html=True)

# ── 직전 3개월 환율 추이 ──
st.markdown(f"##### 📈 직전 3개월 환율 추이 ({START_DATE[4:6]}/{START_DATE[6:]} ~ {END_DATE[4:6]}/{END_DATE[6:]})")

# 주요 이벤트 딕셔너리
_chart_events = {
    "2026-02-28": "이란 전쟁 발발",
    "2026-03-07": "미국 2월 고용지표 발표",
    "2026-03-12": "미국 2월 CPI 발표",
    "2026-03-19": "FOMC 금리 동결",
    "2026-03-20": "한국은행 금통위 기준금리 동결",
    "2026-03-25": "호르무즈 해협 봉쇄 이슈",
    "2026-04-01": "트럼프 이란 최후통첩",
    "2026-04-03": "국제유가 배럴당 130달러 돌파",
}
for ind in report.get("indicators", []):
    if ind["date"] == "미정":
        continue
    try:
        parts = ind["date"].split("/")
        d_str = f"2026-{int(parts[0]):02d}-{int(parts[1]):02d}"
        if d_str not in _chart_events:
            _chart_events[d_str] = ind["name"]
    except (ValueError, IndexError):
        continue

@st.fragment
def _render_trend_chart():
    mode = st.radio(
        "표시 모드",
        ["📊 변동률 인덱스 (시작일=100)", "💹 절대값"],
        horizontal=True,
        key="chart_mode",
        help="변동률 인덱스는 시작일을 100으로 정규화해 세 통화의 상대 변동을 한 그래프에서 비교합니다.",
    )
    normalized = mode.startswith("📊")
    fig = build_chart(df, events=_chart_events, normalized=normalized)
    st.plotly_chart(fig, use_container_width=True)

_render_trend_chart()

# 요약 테이블 (그래프 바로 아래) — 3개월 평균만 표시
summary = pd.DataFrame({
    "": ["3개월 평균"],
    "USD/KRW": [f"{stats['avg_3m']['USD_KRW']:,.2f}"],
    "CNY/KRW": [f"{stats['avg_3m']['CNY_KRW']:,.2f}"],
    "USD/CNY": [f"{stats['avg_3m']['USD_CNY']:.4f}"],
})
st.dataframe(summary, use_container_width=True, hide_index=True)

st.divider()

# ── 전주 환율 현황 ──
st.markdown(f"##### 📋 전주 환율 현황 ({LAST_WEEK_START[4:6]}/{LAST_WEEK_START[6:]} ~ {LAST_WEEK_END[4:6]}/{LAST_WEEK_END[6:]})")

# 메트릭 카드 (% + 금액 차이 병기)
usd_diff = stats["avg_lw"]["USD_KRW"] - stats["avg_prev"]["USD_KRW"]
cny_diff = stats["avg_lw"]["CNY_KRW"] - stats["avg_prev"]["CNY_KRW"]
cross_diff = stats["avg_lw"]["USD_CNY"] - stats["avg_prev"]["USD_CNY"]

c1, c2, c3 = st.columns(3)
c1.metric("USD/KRW (전주 평균)", f"{stats['avg_lw']['USD_KRW']:,.2f} 원",
          delta=f"{stats['usd_vs_pw']:+.2f}% ({usd_diff:+,.2f}원) 전전주 대비", delta_color="inverse")
c2.metric("CNY/KRW (전주 평균)", f"{stats['avg_lw']['CNY_KRW']:,.2f} 원",
          delta=f"{stats['cny_vs_pw']:+.2f}% ({cny_diff:+,.2f}원) 전전주 대비", delta_color="inverse")
c3.metric("USD/CNY 재정 (전주 평균)", f"{stats['avg_lw']['USD_CNY']:.4f}",
          delta=f"{stats['cross_vs_pw']:+.2f}% ({cross_diff:+.4f}) 전전주 대비", delta_color="inverse")

# 전주 복기: 전망 vs 실적

# 전주 실적 데이터 (분석용)
lw_data = df.loc[LAST_WEEK_START:LAST_WEEK_END]

# 전주 PDF에서 밴드 추출 (국민은행 전주 전망)
prev_forecast = {}
try:
    prev_pdfs = [f for f in os.listdir(DATA_DIR) if "전주" in f and f.endswith(".pdf")]
except Exception:
    prev_pdfs = []

# 전주 PDF 텍스트도 수집 (복기 코멘트용)
prev_all_text = ""
for pf in prev_pdfs:
    with pdfplumber.open(os.path.join(DATA_DIR, pf)) as pdf:
        pt = "\n".join(p.extract_text() or "" for p in pdf.pages[:3]).replace(",", "")
        prev_all_text += pt + " "
    mu = re.search(r'USDKRW\s+(\d{4})\s*[~\-]\s*(\d{4})', pt)
    if mu and "USD/KRW" not in prev_forecast:
        prev_forecast["USD/KRW"] = (int(mu.group(1)), int(mu.group(2)))
    mc = re.search(r'USDCNY\s+(\d\.\d{2,3})\s*[~\-]\s*(\d\.\d{2,3})', pt)
    if mc and "USD/CNY" not in prev_forecast:
        prev_forecast["USD/CNY"] = (float(mc.group(1)), float(mc.group(2)))

# CNY/KRW 역산
if "USD/KRW" in prev_forecast and "USD/CNY" in prev_forecast:
    pu_lo, pu_hi = prev_forecast["USD/KRW"]
    pc_lo, pc_hi = prev_forecast["USD/CNY"]
    prev_forecast["CNY/KRW"] = (int(round(pu_lo / pc_hi)), int(round(pu_hi / pc_lo)))

if prev_forecast and not lw_data.empty:
    # 통화별 복기 코멘트 (Claude 분석)
    st.markdown("##### 📝 통화별 분석")

    with st.spinner("Claude AI가 전주 PDF를 분석하는 중..."):
        prev_claude_factors = analyze_pdfs_with_claude(DATA_DIR, week_type="전주")

    comment_rows = []
    for cur, col, fmt in [("USD/KRW", "USD_KRW", "krw"), ("CNY/KRW", "CNY_KRW", "krw"), ("USD/CNY", "USD_CNY", "cross")]:
        if cur not in prev_forecast:
            continue
        f_lo, f_hi = prev_forecast[cur]
        a_lo = float(lw_data[col].min())
        a_hi = float(lw_data[col].max())
        a_avg = float(lw_data[col].mean())

        if a_hi > f_hi:
            result = "⬆️ 상단 상회"
        elif a_lo < f_lo:
            result = "⬇️ 하단 하회"
        else:
            result = "✅ 범위 내"

        # Claude가 분석한 전주 변동요인
        prev_factors = prev_claude_factors.get(cur, ["분석 데이터 없음"])
        cause = "<br>".join(f"• {f}" for f in prev_factors)

        if fmt == "cross":
            fc_str = f"{f_lo}~{f_hi}"
            act_str = f"{a_avg:.4f} ({a_lo:.4f}~{a_hi:.4f})"
        else:
            fc_str = f"{int(f_lo):,}~{int(f_hi):,}"
            act_str = f"{a_avg:,.2f} ({a_lo:,.2f}~{a_hi:,.2f})"

        comment_rows.append({"cur": cur, "forecast": fc_str, "actual": act_str, "result": result, "cause": cause})

    if comment_rows:
        rows_html = ""
        for r in comment_rows:
            rows_html += (
                f'<tr>'
                f'<td style="padding:10px;border:1px solid #eee;font-weight:700;text-align:center;white-space:nowrap;">{r["cur"]}</td>'
                f'<td style="padding:10px;border:1px solid #eee;text-align:center;white-space:nowrap;">{r["forecast"]}</td>'
                f'<td style="padding:10px;border:1px solid #eee;text-align:center;white-space:nowrap;">{r["actual"]}</td>'
                f'<td style="padding:10px;border:1px solid #eee;text-align:center;white-space:nowrap;">{r["result"]}</td>'
                f'<td style="padding:10px;border:1px solid #eee;line-height:1.7;">{r["cause"]}</td>'
                f'</tr>'
            )
        st.markdown(
            f'<table style="width:100%;border-collapse:collapse;font-size:0.9rem;border:1px solid #ddd;">'
            f'<tr style="background:#f0f4ff;">'
            f'<th style="padding:10px;border:1px solid #ddd;text-align:center;">통화</th>'
            f'<th style="padding:10px;border:1px solid #ddd;text-align:center;">전주 전망</th>'
            f'<th style="padding:10px;border:1px solid #ddd;text-align:center;">전주 실적(평균/범위)</th>'
            f'<th style="padding:10px;border:1px solid #ddd;text-align:center;">결과</th>'
            f'<th style="padding:10px;border:1px solid #ddd;text-align:center;">주요 원인</th>'
            f'</tr>{rows_html}</table>',
            unsafe_allow_html=True,
        )
else:
    st.caption("전주 전망 PDF가 없어 복기 분석을 수행할 수 없습니다.")



# ── HTML 리포트 다운로드 ──
def _gen_html():
    chart_html = build_chart(df, events=_chart_events).to_html(include_plotlyjs="cdn", full_html=False)

    # ── 포맷터 ──
    def f_amt(v):
        return "-" if not v else f"{v:,.0f}"
    def f_rate(v):
        if not v: return "-"
        if abs(v - round(v)) < 1e-6:
            return f"{int(round(v)):,}"
        return f"{v:,.2f}"
    def f_krw_mil(v):
        return "-" if not v else f"{v/1_000_000:,.0f}"
    def f_pnl_mil(v):
        if not v: return "-"
        val = abs(v) / 1_000_000
        if v > 0:
            return f'<span style="color:#C00000;font-weight:700;">▲ {val:,.0f}</span>'
        return f'<span style="color:#4A90D9;font-weight:700;">▼ {val:,.0f}</span>'

    # ── USD 표 ──
    cash_book_krw = usd_cash * usd_book if usd_book else 0
    cash_mkt_krw = usd_cash * usd_mkt
    ar_s_krw = usd_ar_short * usd_mkt
    ar_l_krw = usd_ar_long * usd_mkt
    ap_s_krw = usd_ap_short * usd_mkt
    ap_l_krw = usd_ap_long * usd_mkt

    usd_rows_html = (
        f'<tr><td rowspan="1" style="font-weight:600;background:#f8f9fc;text-align:center;">보유현금</td>'
        f'<td></td>'
        f'<td style="text-align:right;">{f_amt(usd_cash)}</td>'
        f'<td style="text-align:right;">{f_rate(usd_book)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cash_book_krw)}</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cash_mkt_krw)}</td>'
        f'<td style="text-align:right;">{f_pnl_mil(_g_cash_pnl)}</td></tr>'
        f'<tr><td rowspan="2" style="font-weight:600;background:#f8f9fc;text-align:center;vertical-align:middle;">미결채권</td>'
        f'<td>단기 회수예정</td>'
        f'<td style="text-align:right;">{f_amt(usd_ar_short)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(ar_s_krw)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr><td>장기 누적미결</td>'
        f'<td style="text-align:right;">{f_amt(usd_ar_long)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(ar_l_krw)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr><td rowspan="2" style="font-weight:600;background:#f8f9fc;text-align:center;vertical-align:middle;">미결채무</td>'
        f'<td>단기 지급예정</td>'
        f'<td style="text-align:right;color:#C00000;">{f_amt(-usd_ap_short)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(-ap_s_krw)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr><td>장기 미결현황</td>'
        f'<td style="text-align:right;color:#C00000;">{f_amt(-usd_ap_long)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(-ap_l_krw)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr style="background:#ececec;font-weight:700;">'
        f'<td style="text-align:center;">순노출금액</td>'
        f'<td>현금 + 채권 − 채무</td>'
        f'<td style="text-align:right;">{f_amt(_g_net_exposure)}</td>'
        f'<td style="text-align:right;color:#888;">-</td><td style="text-align:right;color:#888;">-</td>'
        f'<td style="text-align:right;">{f_rate(usd_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(_g_net_mkt_krw)}</td>'
        f'<td style="text-align:right;">{f_pnl_mil(_g_cash_pnl)}</td></tr>'
    )

    # ── CNY 표 ──
    cny_cash_book_krw_h = cny_cash * cny_book if cny_book else 0
    cny_cash_mkt_krw_h = cny_cash * cny_mkt
    cny_ar_s_krw_h = cny_ar_short * cny_mkt
    cny_ar_l_krw_h = cny_ar_long * cny_mkt
    cny_ap_krw_h = cny_ap_val * cny_mkt

    cny_rows_html_h = (
        f'<tr><td style="font-weight:600;background:#f8f9fc;text-align:center;">보유현금</td>'
        f'<td></td>'
        f'<td style="text-align:right;">{f_amt(cny_cash)}</td>'
        f'<td style="text-align:right;">{f_rate(cny_book)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cny_cash_book_krw_h)}</td>'
        f'<td style="text-align:right;">{f_rate(cny_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cny_cash_mkt_krw_h)}</td>'
        f'<td style="text-align:right;">{f_pnl_mil(_g_cny_cash_pnl)}</td></tr>'
        f'<tr><td rowspan="2" style="font-weight:600;background:#f8f9fc;text-align:center;vertical-align:middle;">미결채권</td>'
        f'<td>단기 회수예정</td>'
        f'<td style="text-align:right;">{f_amt(cny_ar_short)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(cny_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cny_ar_s_krw_h)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr><td>장기 누적미결</td>'
        f'<td style="text-align:right;">{f_amt(cny_ar_long)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(cny_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(cny_ar_l_krw_h)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr><td style="font-weight:600;background:#f8f9fc;text-align:center;">미결채무</td>'
        f'<td>장기 미결현황</td>'
        f'<td style="text-align:right;color:#C00000;">{f_amt(-cny_ap_val)}</td>'
        f'<td style="text-align:right;">-</td><td style="text-align:right;">-</td>'
        f'<td style="text-align:right;">{f_rate(cny_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(-cny_ap_krw_h)}</td>'
        f'<td style="text-align:right;">-</td></tr>'
        f'<tr style="background:#ececec;font-weight:700;">'
        f'<td style="text-align:center;">순노출금액</td>'
        f'<td>현금 + 채권 − 채무</td>'
        f'<td style="text-align:right;">{f_amt(_g_cny_net)}</td>'
        f'<td style="text-align:right;color:#888;">-</td><td style="text-align:right;color:#888;">-</td>'
        f'<td style="text-align:right;">{f_rate(cny_mkt)}</td>'
        f'<td style="text-align:right;">{f_krw_mil(_g_cny_net_mkt_krw)}</td>'
        f'<td style="text-align:right;">{f_pnl_mil(_g_cny_cash_pnl)}</td></tr>'
    )

    # ── AI 의사결정 분석 박스 (HTML) ──
    def render_ai_box(decision):
        if decision.get("error"):
            return f'<div style="background:#fdf2f2;border-left:4px solid #C00000;padding:10px 14px;margin:8px 0;font-size:0.88rem;">⚠️ AI 분석 실패: {decision["error"]}</div>'
        risks_html = "".join(f"&nbsp;&nbsp;&nbsp;&nbsp;• {r}<br>" for r in decision.get("risks", []))
        actions_html = "".join(f"&nbsp;&nbsp;&nbsp;&nbsp;{a}<br>" for a in decision.get("actions", []))
        return (
            f'<div style="background:#fafbff;border:1px solid #d6d9e3;border-radius:8px;padding:14px 18px;margin:8px 0 16px;font-size:0.9rem;line-height:1.7;">'
            f'<div style="font-weight:700;font-size:0.95rem;margin-bottom:8px;color:#2E75B6;">📋 의사결정 분석 <span style="font-size:0.78rem;color:#888;font-weight:400;">(Claude AI)</span></div>'
            f'<div style="margin-bottom:8px;"><b style="color:#333;">▸ 현황</b><br>&nbsp;&nbsp;&nbsp;&nbsp;{decision.get("current","")}</div>'
            f'<div style="margin-bottom:8px;"><b style="color:#C00000;">▸ 리스크</b><br>{risks_html}</div>'
            f'<div><b style="color:#2E8B57;">▸ 실무 제안</b><br>{actions_html}</div>'
            f'</div>'
        )

    usd_ai_html = render_ai_box(g_usd_decision)
    cny_ai_html = render_ai_box(g_cny_decision)

    # 전망 밴드 방향 표시
    def _hdir(d):
        c = {"상승":"#C00000","하락":"#4A90D9"}.get(d,"#2E8B57")
        a = {"상승":"↑","하락":"↓"}.get(d,"→")
        return f'<span style="color:{c};font-weight:700;">{a} {d}</span>'

    # 변동요인
    usd_f = "<br>".join(f"• {f}" for f in usd_factors)
    cny_f = "<br>".join(f"• {f}" for f in cny_factors)
    cross_f = "<br>".join(f"• {f}" for f in cross_factors)

    # 표 헤더 공통
    pos_header = (
        '<tr><th rowspan="2">항목</th><th rowspan="2">구분</th><th rowspan="2">금액</th>'
        '<th colspan="2">장부 기준</th><th colspan="2">당일 기준</th>'
        '<th rowspan="2">현재기준 외환차손익<br><span style="font-weight:400;font-size:0.72rem;color:#666;">(백만원)</span></th></tr>'
        '<tr><th>보유 평균환율</th><th>원화환산금액 <span style="font-weight:400;font-size:0.72rem;color:#666;">(백만원)</span></th>'
        '<th>매매기준율</th><th>원화환산금액 <span style="font-weight:400;font-size:0.72rem;color:#666;">(백만원)</span></th></tr>'
    )

    return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<title>주간 환율 리포트 — 2026년 4월 2주차</title>
<style>
body{{font-family:'Malgun Gothic',sans-serif;margin:0;padding:20px 40px;color:#222;background:#fff;}}
h2{{margin-bottom:4px;}} .cap{{font-size:0.8rem;color:#888;margin-bottom:16px;}}
.sh{{background:linear-gradient(90deg,#2E75B6,#4a90d9);color:white;padding:10px 20px;border-radius:8px;font-size:1.1rem;font-weight:700;margin:24px 0 12px;}}
.subh{{font-size:1.0rem;font-weight:700;margin:18px 0 8px;color:#2E75B6;}}
table{{width:100%;border-collapse:collapse;font-size:0.88rem;margin:8px 0;}}
th{{background:#f0f4ff;padding:8px 10px;text-align:center;border:1px solid #ddd;font-weight:700;}}
td{{padding:8px 10px;border:1px solid #eee;}}
.note{{font-size:0.78rem;color:#888;margin-top:4px;}}
.rate-info{{background:#f0f4ff;border-radius:6px;padding:8px 14px;margin:6px 0;font-size:0.85rem;}}
.mc{{display:flex;gap:16px;margin:12px 0;}}
.mc>div{{flex:1;background:linear-gradient(135deg,#667eea08,#764ba208);border:1px solid #ddd;border-radius:12px;padding:16px 20px;}}
.mc .lb{{font-size:0.82rem;color:#555;}} .mc .vl{{font-size:1.8rem;font-weight:700;margin:4px 0;}}
.ft{{font-size:0.75rem;color:#aaa;margin-top:20px;border-top:1px solid #eee;padding-top:8px;}}
</style></head><body>
<h2>📊 주간 환율 리포트 — 2026년 4월 2주차</h2>
<div class="cap">한국은행 ECOS API · 국민은행/신한은행 PDF · 서울파이낸셜 · Claude AI 의사결정 분석</div>

<div class="sh">1. 주간 외환 관리 가이드라인 ({latest_date} 기준)</div>

<div class="subh">🇺🇸 USD 유동성 진단</div>
<div class="rate-info">📊 <b>당일 매매기준율</b>: {usd_mkt:,.2f}원 &nbsp;|&nbsp; <b>금주 전망 중간</b>: {(usd_lo+usd_hi)/2:,.1f}원 (밴드 {usd_lo:,}~{usd_hi:,})</div>
<table>{pos_header}{usd_rows_html}</table>
<div class="note">참고) 채권: 전기일 기준(단기 90일 내 / 장기 누적분) · 채무: 만기일 기준(단기 30일 내 / 잔여분)</div>
{usd_ai_html}

<div class="subh">🇨🇳 CNY 수익 전략</div>
<div class="rate-info">📊 <b>당일 매매기준율</b>: {cny_mkt:,.2f}원 &nbsp;|&nbsp; <b>금주 전망 중간</b>: {(cny_lo+cny_hi)/2:,.1f}원 (밴드 {cny_lo}~{cny_hi})</div>
<table>{pos_header.replace('금액', '금액(CNY)').replace('항목</th><th rowspan="2">구분', '항목</th><th rowspan="2">구분')}{cny_rows_html_h}</table>
<div class="note">참고) 채권: 전기일 기준(단기 30일 내 / 장기 누적분) · 채무: 만기일 기준(잔여분)</div>
{cny_ai_html}

<div class="sh">2. 금주 환율 전망 ({REPORT_WEEK_START[4:6]}/{REPORT_WEEK_START[6:]} ~ {REPORT_WEEK_END[4:6]}/{REPORT_WEEK_END[6:]})</div>
<div class="mc">
<div><div class="lb">USD/KRW (금주 전망)</div><div class="vl">{usd_lo:,} ~ {usd_hi:,} 원</div>{_hdir(usd_dir)}</div>
<div><div class="lb">CNY/KRW (금주 전망)</div><div class="vl">{cny_lo:,} ~ {cny_hi:,} 원</div>{_hdir(cny_dir)}</div>
<div><div class="lb">USD/CNY (금주 전망)</div><div class="vl">{cross_lo} ~ {cross_hi}</div>{_hdir(cross_dir)}</div>
</div>
<table>
<tr><th>통화</th><th>전망</th><th>예상 밴드</th><th>변동 요인</th></tr>
<tr><td style="text-align:center;font-weight:700;">USD/KRW</td><td style="text-align:center;">{_hdir(usd_dir)}</td>
<td style="text-align:center;font-weight:600;">{usd_lo:,}~{usd_hi:,}원</td><td>{usd_f}</td></tr>
<tr><td style="text-align:center;font-weight:700;">CNY/KRW</td><td style="text-align:center;">{_hdir(cny_dir)}</td>
<td style="text-align:center;font-weight:600;">{cny_lo:,}~{cny_hi:,}원</td><td>{cny_f}</td></tr>
<tr><td style="text-align:center;font-weight:700;">USD/CNY</td><td style="text-align:center;">{_hdir(cross_dir)}</td>
<td style="text-align:center;font-weight:600;">{cross_lo}~{cross_hi}</td><td>{cross_f}</td></tr>
</table>

<div class="sh">3. 환율 추이 및 전주 동향</div>
{chart_html}

<div class="mc">
<div><div class="lb">USD/KRW (전주 평균)</div><div class="vl">{stats['avg_lw']['USD_KRW']:,.2f} 원</div></div>
<div><div class="lb">CNY/KRW (전주 평균)</div><div class="vl">{stats['avg_lw']['CNY_KRW']:,.2f} 원</div></div>
<div><div class="lb">USD/CNY (전주 평균)</div><div class="vl">{stats['avg_lw']['USD_CNY']:.4f}</div></div>
</div>

<div class="ft">ⓒ F&F 자금팀 · 한국은행 ECOS API · 서울파이낸셜 · Claude AI · 2026년 4월 2주차 리포트</div>
</body></html>"""

st.divider()
st.download_button(
    label="📥 HTML 리포트 다운로드",
    data=_gen_html().encode("utf-8"),
    file_name="주간환율리포트_2026년4월2주차.html",
    mime="text/html",
)

# ── 푸터 ──
st.caption("ⓒ F&F 자금팀 · 한국은행 ECOS API · 서울파이낸셜 · 2026년 4월 2주차 리포트")
