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
def build_chart(df: pd.DataFrame, events: dict = None) -> go.Figure:
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
            model="claude-3-5-sonnet-20241022",
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

# ── Claude API 자율 분석 ──
@st.cache_data(ttl=86400, show_spinner=False)
def _ai_fx_strategy(context: str) -> str:
    if not ANTHROPIC_API_KEY:
        return "⚠️ API 키가 설정되지 않아 AI 분석을 수행할 수 없습니다."
    try:
        from anthropic import Anthropic
        client = Anthropic(api_key=ANTHROPIC_API_KEY, timeout=30.0)
        msg = client.messages.create(
            model="claude-3-5-sonnet-latest",
            max_tokens=2000,
            system="""너는 글로벌 기업의 외환 전략 최고 책임자(CFO)야.
아래 데이터를 보고 '금주 외환 실행 전략 가이드'를 작성해.

[자율 분석 미션]
1) USD 유동성 진단:
   - 단기 미결 채무가 보유 현금+단기 채권보다 크면 '🚨 즉시 달러 매수 필요' 경고를 생성.
   - 순 노출액이 (-)면 부족분만큼 환전 권고.

2) CNY 수익 전략 (보유 잔액 기준, 채권채무 무시):
   - 환전 시점: 금주 환율 전망과 장부가 대비 현재 환율을 비교해 지금 환전이 적절한지 판단.
   - 환전 비중: 보유 CNY의 몇 %를 환전할지 구체 수치 제안 (예: 30%, 50%, 100%).
   - 환전 대상: KRW(즉시 환차익 확정) vs USD(미래 결제 대비 가성비) 중 어느 쪽이 외환차손익 극대화에 유리한지 비교 분석.

[작성 규칙]
- 두 섹션을 명확히 구분해서 작성:
  ## 🚨 USD 유동성 진단
  ## 💱 CNY 환전 전략
- 결론부터 말하고, 근거 숫자를 반드시 포함.
- 마크다운(볼드, 불렛). 한국어. 500자 이내.""",
            messages=[{"role": "user", "content": context}],
        )
        return msg.content[0].text.strip()
    except Exception as e:
        return f"⚠️ AI 분석 실패: {e}"

with st.spinner("AI가 외환 포지션을 분석하는 중..."):
    ai_strategy = _ai_fx_strategy(summary_text)

# ── UI 렌더링 ──

# 최상단: AI 전략 제안
with st.chat_message("assistant"):
    st.markdown(f"**🤖 금주 외환 실행 전략 가이드**\n\n{ai_strategy}")

st.markdown("")

# USD / CNY 탭 분리
tab_usd, tab_cny = st.tabs(["🇺🇸 USD 유동성 진단", "🇨🇳 CNY 수익 전략"])

with tab_usd:
    usd_liquidity = usd_cash + usd_ar_short - usd_ap_short
    u1, u2, u3, u4 = st.columns(4)
    u1.metric("보유 현금", f"${usd_cash:,.0f}")
    u2.metric("단기 AR/AP", f"+${usd_ar_short:,.0f} / -${usd_ap_short:,.0f}")
    u3.metric("장기 AR/AP", f"+${usd_ar_long:,.0f} / -${usd_ap_long:,.0f}")
    u4.metric("최종 순 노출액", f"${usd_net:,.0f}",
              delta="유동성 부족" if usd_liquidity < 0 else "유동성 양호",
              delta_color="inverse" if usd_liquidity < 0 else "normal")

    if usd_liquidity < 0:
        st.markdown(
            f'<div style="margin-top:12px;padding:12px 16px;background:#fdf2f2;border-left:4px solid #C00000;border-radius:6px;font-size:0.9rem;">'
            f'🚨 <b>즉시 달러 매수 필요</b>: 단기 유동성 <b>${abs(usd_liquidity):,.0f}</b> 부족'
            f'</div>', unsafe_allow_html=True)
    else:
        st.markdown(
            f'<div style="margin-top:12px;padding:12px 16px;background:#f0fdf4;border-left:4px solid #2E8B57;border-radius:6px;font-size:0.9rem;">'
            f'✅ <b>유동성 양호</b>: 단기 여유 <b>${usd_liquidity:,.0f}</b>'
            f'</div>', unsafe_allow_html=True)

with tab_cny:
    cny_pnl_per = (cny_mkt - cny_book) if cny_book else 0
    cny_pnl_total = cny_pnl_per * cny_cash
    cny_pnl_pct = (cny_mkt - cny_book) / cny_book * 100 if cny_book else 0

    cc1, cc2, cc3, cc4 = st.columns(4)
    cc1.metric("현재 보유 잔액", f"¥{cny_cash:,.0f}")
    cc2.metric("당일 매매기준율", f"{cny_mkt:,.2f}원",
               delta=f"{cny_pnl_pct:+.2f}% vs 장부", delta_color="inverse")
    cc3.metric("장부단가", f"{cny_book:,.2f}원")
    pnl_color_str = "+" if cny_pnl_total >= 0 else ""
    cc4.metric("잠재 환차익", f"{pnl_color_str}{cny_pnl_total:,.0f}원",
               delta_color="inverse")


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

st.divider()

# CNY 환전 시뮬레이터
st.markdown("##### 💱 CNY 환전 시뮬레이터")

@st.fragment
def _run_simulator():
    """시뮬레이터만 독립 리렌더링 (페이지 전체 리로드 방지)."""

    # 엑셀에서 CNY 보유 데이터
    cny_rows = [i for i, c in enumerate(cash_df["통화"].tolist()) if str(c).upper() == "CNY"]
    if cny_rows:
        sim_cny = float(str(cash_df.iloc[cny_rows[0]]["금액"]).replace(",", ""))
        sim_book = float(str(cash_df.iloc[cny_rows[0]]["보유환율"]).replace(",", ""))
    else:
        sim_cny = 30000000.0
        sim_book = 198.50

    usd_rate = float(latest["USD_KRW"])

    sim_target = st.radio("환전 통화", ["KRW", "USD"], horizontal=True)

    sl1, sl2 = st.columns(2)
    with sl1:
        sim_pct = st.selectbox("환전 비율", ["30%", "50%", "70%", "100%"], index=3)
        pct_val = int(sim_pct.replace("%", "")) / 100
        sim_cny_actual = sim_cny * pct_val
    with sl2:
        if sim_target == "KRW":
            sim_rate = st.slider("적용 환율 (CNY/KRW)",
                                 min_value=float(cny_lo), max_value=float(cny_hi),
                                 value=round(float(latest["CNY_KRW"]), 1), step=0.1, format="%.1f원")
        else:
            sim_rate = st.slider("적용 재정환율 (USD/CNY)",
                                 min_value=float(cross_lo), max_value=float(cross_hi),
                                 value=round(float(latest["USD_CNY"]), 2), step=0.01, format="%.2f")

    if sim_target == "KRW":
        sim_krw = sim_cny_actual * sim_rate
        sim_pnl = sim_cny_actual * (sim_rate - sim_book)
        conv_label = f"{sim_krw:,.0f} 원"
        rate_label = f"적용환율 {sim_rate:,.2f}원"
        pnl_formula = f"({sim_rate:,.2f} - {sim_book:,.2f}) × {sim_cny_actual:,.0f}"
    else:
        sim_usd = sim_cny_actual / sim_rate
        book_cross = float(latest["USD_CNY"])
        sim_pnl_usd = sim_cny_actual * (1/sim_rate - 1/book_cross)
        sim_pnl = sim_pnl_usd * usd_rate
        conv_label = f"${sim_usd:,.2f}"
        rate_label = f"적용 재정환율 {sim_rate:.2f}"
        pnl_formula = f"CNY÷{sim_rate:.2f} vs CNY÷{book_cross:.2f}"

    pnl_color = "#C00000" if sim_pnl > 0 else "#4A90D9"

    st.markdown(
        f'<div style="display:flex;gap:12px;margin-top:8px;">'
        f'<div style="flex:1;background:#f8f9fc;border:1px solid #ddd;border-radius:10px;padding:14px 18px;">'
        f'<div style="font-size:0.8rem;color:#888;">환전 CNY ({sim_pct})</div>'
        f'<div style="font-size:1.3rem;font-weight:700;">{sim_cny_actual:,.0f}</div>'
        f'<div style="font-size:0.75rem;color:#888;">전체 {sim_cny:,.0f} · 장부 {sim_book:,.2f}원</div>'
        f'</div>'
        f'<div style="flex:1;background:#f8f9fc;border:1px solid #ddd;border-radius:10px;padding:14px 18px;">'
        f'<div style="font-size:0.8rem;color:#888;">환전 금액 ({sim_target})</div>'
        f'<div style="font-size:1.3rem;font-weight:700;">{conv_label}</div>'
        f'<div style="font-size:0.75rem;color:#888;">{rate_label}</div>'
        f'</div>'
        f'<div style="flex:1;background:#f8f9fc;border:1px solid #ddd;border-radius:10px;padding:14px 18px;">'
        f'<div style="font-size:0.8rem;color:#888;">외환차손익 (KRW)</div>'
        f'<div style="font-size:1.3rem;font-weight:700;color:{pnl_color};">{sim_pnl:+,.0f} 원</div>'
        f'<div style="font-size:0.75rem;color:#888;">{pnl_formula}</div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

_run_simulator()


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

fig = build_chart(df, events=_chart_events)
st.plotly_chart(fig, use_container_width=True)

# 요약 테이블 (그래프 바로 아래)
lw_data = df.loc[LAST_WEEK_START:LAST_WEEK_END]
if not lw_data.empty:
    summary = pd.DataFrame({
        "": ["전주 평균", "전주 최고", "전주 최저", "3개월 평균"],
        "USD/KRW": [f"{lw_data['USD_KRW'].mean():,.2f}", f"{lw_data['USD_KRW'].max():,.2f}",
                    f"{lw_data['USD_KRW'].min():,.2f}", f"{stats['avg_3m']['USD_KRW']:,.2f}"],
        "CNY/KRW": [f"{lw_data['CNY_KRW'].mean():,.2f}", f"{lw_data['CNY_KRW'].max():,.2f}",
                    f"{lw_data['CNY_KRW'].min():,.2f}", f"{stats['avg_3m']['CNY_KRW']:,.2f}"],
        "USD/CNY": [f"{lw_data['USD_CNY'].mean():.4f}", f"{lw_data['USD_CNY'].max():.4f}",
                    f"{lw_data['USD_CNY'].min():.4f}", f"{stats['avg_3m']['USD_CNY']:.4f}"],
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

    # 보유현황 테이블
    hold_rows = ""
    for i in range(len(currencies)):
        p = pnl[i]
        pc = "#C00000" if p > 0 else "#4A90D9"
        hold_rows += (
            f'<tr><td>{latest_date}</td><td>{currencies[i]}</td>'
            f'<td style="text-align:right;">{amt[i]:,.2f}</td>'
            f'<td style="text-align:right;">{book[i]:,.2f}</td>'
            f'<td style="text-align:right;">{book_krw[i]:,.0f}</td>'
            f'<td style="text-align:right;">{mkt[i]:,.2f}</td>'
            f'<td style="text-align:right;">{mkt_krw[i]:,.0f}</td>'
            f'<td style="text-align:right;font-weight:700;color:{pc};">{p:+,.0f}</td></tr>'
        )
    tp_color = "#C00000" if total_pnl > 0 else "#4A90D9"

    # 전망 밴드
    def _hdir(d):
        c = {"상승":"#C00000","하락":"#4A90D9"}.get(d,"#2E8B57")
        a = {"상승":"↑","하락":"↓"}.get(d,"→")
        return f'<span style="color:{c};font-weight:700;">{a} {d}</span>'

    # 변동요인
    usd_f = "<br>".join(f"• {f}" for f in usd_factors)
    cny_f = "<br>".join(f"• {f}" for f in cny_factors)
    cross_f = "<br>".join(f"• {f}" for f in cross_factors)

    return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="UTF-8">
<title>주간 환율 리포트 — 2026년 4월 2주차</title>
<style>
body{{font-family:'Malgun Gothic',sans-serif;margin:0;padding:20px 40px;color:#222;background:#fff;}}
h2{{margin-bottom:4px;}} .cap{{font-size:0.8rem;color:#888;margin-bottom:16px;}}
.sh{{background:linear-gradient(90deg,#2E75B6,#4a90d9);color:white;padding:10px 20px;border-radius:8px;font-size:1.1rem;font-weight:700;margin:24px 0 12px;}}
table{{width:100%;border-collapse:collapse;font-size:0.88rem;margin:8px 0;}}
th{{background:#f0f4ff;padding:8px 10px;text-align:center;border:1px solid #ddd;font-weight:700;}}
td{{padding:8px 10px;border:1px solid #eee;}}
.mc{{display:flex;gap:16px;margin:12px 0;}}
.mc>div{{flex:1;background:linear-gradient(135deg,#667eea08,#764ba208);border:1px solid #ddd;border-radius:12px;padding:16px 20px;}}
.mc .lb{{font-size:0.82rem;color:#555;}} .mc .vl{{font-size:1.8rem;font-weight:700;margin:4px 0;}}
.ft{{font-size:0.75rem;color:#aaa;margin-top:20px;border-top:1px solid #eee;padding-top:8px;}}
</style></head><body>
<h2>📊 주간 환율 리포트 — 2026년 4월 2주차</h2>
<div class="cap">한국은행 ECOS API · 국민은행/신한은행 PDF · 서울파이낸셜</div>

<div class="sh">1. 당사 외화 보유 현황 (매매기준율: {latest_date})</div>
<table>
<tr><th rowspan="2">날짜</th><th rowspan="2">통화</th><th rowspan="2">보유금액</th>
<th colspan="2">장부 기준</th><th colspan="2">당일 기준</th><th rowspan="2">외환차손익(원)</th></tr>
<tr><th>보유 평균환율</th><th>원화환산금액</th><th>매매기준율</th><th>원화환산금액</th></tr>
{hold_rows}
</table>
<div style="text-align:right;font-size:1rem;font-weight:700;color:{tp_color};margin-top:8px;">현재기준 외환차손익: {total_pnl:+,.0f}원</div>

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

<div class="ft">ⓒ F&F 자금팀 · 한국은행 ECOS API · 서울파이낸셜 · 2026년 4월 2주차 리포트</div>
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
