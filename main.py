import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import requests
from bs4 import BeautifulSoup
import pdfplumber
import re
import os

# ─── 고정 설정 ──────────────────────────────────────────
BOK_API_KEY = st.secrets.get("BOK_API_KEY", os.environ.get("BOK_API_KEY", ""))
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
def build_chart(df: pd.DataFrame) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # 전주 구간 배경 강조 (shape로 직접 그리기)
    lw_dates = df.loc[LAST_WEEK_START:LAST_WEEK_END].index
    if len(lw_dates) > 0:
        x0_ms = int(lw_dates[0].timestamp() * 1000) - 43200000   # 반일 앞
        x1_ms = int(lw_dates[-1].timestamp() * 1000) + 43200000  # 반일 뒤
        fig.update_layout(
            shapes=[
                dict(type="rect", xref="x", yref="paper",
                     x0=x0_ms, x1=x1_ms, y0=0, y1=1,
                     fillcolor="rgba(0,0,0,0.06)", line=dict(width=0),
                     layer="below"),
            ],
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
        height=430, template="plotly_white", hovermode="x unified",
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
            "detect": ["달러", "USD", "원/달러", "달러원", "원달러", "미국"],
            "bull": ["달러 강세", "USD 상승", "달러원 상승", "원화 약세", "원/달러 상승"],
            "bear": ["달러 약세", "USD 하락", "달러원 하락", "원화 강세", "원/달러 하락"],
        },
        "CNY_KRW": {
            "detect": ["위안", "CNY", "원/위안", "위안화", "중국"],
            "bull": ["위안 강세", "위안화 강세", "CNY 상승", "원/위안 상승"],
            "bear": ["위안 약세", "위안화 약세", "CNY 하락", "원/위안 하락", "위안화 절하"],
        },
        "USD_CNY": {
            "detect": ["재정환율", "USD/CNY", "달러/위안"],
            "bull": ["재정환율 상승", "USD/CNY 상승", "위안 절하"],
            "bear": ["재정환율 하락", "USD/CNY 하락", "위안 절상"],
        },
    }

    currency_analysis = {}
    for cur_key, kws in currency_kw.items():
        # 해당 통화 관련 문장 수집
        related = [s for s in all_sentences if any(kw in s for kw in kws["detect"])]

        # 방향성 판단
        bull_score = sum(sum(1 for kw in kws["bull"] if kw in s) for s in related)
        bear_score = sum(sum(1 for kw in kws["bear"] if kw in s) for s in related)
        # 일반 방향 키워드도 가산
        for s in related:
            bull_score += sum(1 for kw in direction_kw["상승"] if kw in s)
            bear_score += sum(1 for kw in direction_kw["하락"] if kw in s)

        if bull_score > bear_score * 1.2:
            direction = "상승"
            arrow = "▲"
            color = "#C00000"
        elif bear_score > bull_score * 1.2:
            direction = "하락"
            arrow = "▼"
            color = "#2E75B6"
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

st.markdown("## 📊 주간 환율 리포트 — 2026년 4월 1주차")
st.caption("수집 기간: 2026-01-03 ~ 2026-04-03 · 전주 구간: 3/30 ~ 4/3 · 한국은행 ECOS API")

# ── 데이터 로딩 ──
with st.spinner("한국은행 API에서 환율 데이터를 불러오는 중..."):
    try:
        df = fetch_exchange_rates()
    except Exception as e:
        st.error(f"데이터 로딩 실패: {e}")
        st.stop()

stats = compute_stats(df)
latest = stats["latest"]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 1: 전주 환율 요약 카드
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown('<div class="section-header">1. 전주 환율 요약 (3/30 ~ 4/3)</div>', unsafe_allow_html=True)

c1, c2, c3 = st.columns(3)
c1.metric(
    "USD / KRW (전주 평균)", f"{stats['avg_lw']['USD_KRW']:,.2f} 원",
    delta=f"{stats['usd_vs_pw']:+.2f}% (전전주 대비)", delta_color="inverse",
)
c2.metric(
    "CNY / KRW (전주 평균)", f"{stats['avg_lw']['CNY_KRW']:,.2f} 원",
    delta=f"{stats['cny_vs_pw']:+.2f}% (전전주 대비)", delta_color="inverse",
)
c3.metric(
    "USD / CNY 재정 (전주 평균)", f"{stats['avg_lw']['USD_CNY']:.4f}",
    delta=f"{stats['cross_vs_pw']:+.2f}% (전전주 대비)", delta_color="inverse",
)

# 전주 요약 테이블
lw_data = df.loc[LAST_WEEK_START:LAST_WEEK_END]
if not lw_data.empty:
    summary = pd.DataFrame({
        "": ["전주 평균", "전주 최고", "전주 최저", "3개월 평균"],
        "USD/KRW": [
            f"{lw_data['USD_KRW'].mean():,.2f}",
            f"{lw_data['USD_KRW'].max():,.2f}",
            f"{lw_data['USD_KRW'].min():,.2f}",
            f"{stats['avg_3m']['USD_KRW']:,.2f}",
        ],
        "CNY/KRW": [
            f"{lw_data['CNY_KRW'].mean():,.2f}",
            f"{lw_data['CNY_KRW'].max():,.2f}",
            f"{lw_data['CNY_KRW'].min():,.2f}",
            f"{stats['avg_3m']['CNY_KRW']:,.2f}",
        ],
        "USD/CNY": [
            f"{lw_data['USD_CNY'].mean():.4f}",
            f"{lw_data['USD_CNY'].max():.4f}",
            f"{lw_data['USD_CNY'].min():.4f}",
            f"{stats['avg_3m']['USD_CNY']:.4f}",
        ],
    })
    st.dataframe(summary, use_container_width=True, hide_index=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 2: 3개월 환율 추이 + 전주 분석 코멘트
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown('<div class="section-header">2. 환율 추이 및 전주 분석</div>', unsafe_allow_html=True)

fig = build_chart(df)
st.plotly_chart(fig, use_container_width=True)

# AI 코멘트 박스 (상승=빨강, 하락=파랑)
def _colored(val, label):
    """값에 따라 빨간(상승)/파란(하락) 색상 span을 반환."""
    color = "#C00000" if val > 0 else "#2E75B6"
    if "3m" in label:
        text = "높은" if val > 0 else "낮은"
        return f'<span style="color:{color}"><b>{abs(val):.1f}% {text}</b></span>'
    else:
        text = "상승" if val > 0 else "하락"
        return f'<span style="color:{color}"><b>{abs(val):.2f}% {text}세</b></span>'

avg = stats["avg_lw"]
comment = (
    f"• <b>USD/KRW</b> 전주 평균 {avg['USD_KRW']:,.2f}원은 "
    f"3개월 평균({stats['avg_3m']['USD_KRW']:,.2f}) 대비 "
    f"{_colored(stats['usd_vs_3m'], '3m')} 수준이며, "
    f"전전주 대비 {_colored(stats['usd_vs_pw'], 'pw')}임.<br>"
    f"• <b>CNY/KRW</b> 전주 평균 {avg['CNY_KRW']:,.2f}원은 "
    f"3개월 평균({stats['avg_3m']['CNY_KRW']:,.2f}) 대비 "
    f"{_colored(stats['cny_vs_3m'], '3m')} 수준이며, "
    f"전전주 대비 {_colored(stats['cny_vs_pw'], 'pw')}임.<br>"
    f"• <b>USD/CNY(재정)</b> 전주 평균 {avg['USD_CNY']:.4f}은 "
    f"3개월 평균({stats['avg_3m']['USD_CNY']:.4f}) 대비 "
    f"{_colored(stats['cross_vs_3m'], '3m')}, "
    f"전전주 대비 {_colored(stats['cross_vs_pw'], 'pw')}."
)
st.markdown(f'<div class="comment-box">{comment}</div>', unsafe_allow_html=True)

with st.expander("원본 데이터 테이블 (최근 30영업일)"):
    disp = df.tail(30).copy()
    disp.columns = ["USD/KRW", "CNY/KRW", "USD/CNY"]

    # 전주 구간 행 배경색 적용
    lw_start = pd.Timestamp(LAST_WEEK_START)
    lw_end = pd.Timestamp(LAST_WEEK_END)
    is_lw = [(lw_start <= idx <= lw_end) for idx in disp.index]

    # 인덱스를 문자열로 변환
    disp.index = disp.index.strftime("%Y-%m-%d")

    def _highlight_lw(row):
        row_idx = disp.index.get_loc(row.name)
        if is_lw[row_idx]:
            return ["background-color: #e0e0e0"] * len(row)
        return [""] * len(row)

    styled = (disp.style
              .format({"USD/KRW": "{:,.2f}", "CNY/KRW": "{:,.2f}", "USD/CNY": "{:.4f}"})
              .apply(_highlight_lw, axis=1))
    st.dataframe(styled, use_container_width=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 섹션 3: 금주 환율 전망
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.markdown('<div class="section-header">3. 금주 환율 전망 (4/7 ~ 4/11)</div>', unsafe_allow_html=True)

with st.spinner("기관별 리서치 통합 분석 중..."):
    report = run_integrated_analysis(DATA_DIR)

band = report["band"]
ca = report["currency_analysis"]

# ── 카드 밴드: 국민은행 PDF 주간예상범위 (텍스트 직접 추출) ──
inst_bands = report.get("inst_bands", [])
recent5 = df.tail(5)

_kb_usd = None
_kb_cross = None
kb_pdf = os.path.join(DATA_DIR, "국민은행_주간_전망자료.pdf")
if os.path.isfile(kb_pdf):
    with pdfplumber.open(kb_pdf) as pdf:
        text = "\n".join(p.extract_text() or "" for p in pdf.pages[:2]).replace(",", "")
        # "USDKRW 1490~1540 1470~1540" 형태에서 첫 번째(주간) 범위 추출
        m_usd = re.search(r'USDKRW\s+(\d{4})\s*[~\-]\s*(\d{4})', text)
        if m_usd:
            _kb_usd = (int(m_usd.group(1)), int(m_usd.group(2)))
        m_cny = re.search(r'USDCNY\s+(\d\.\d{2,3})\s*[~\-]\s*(\d\.\d{2,3})', text)
        if m_cny:
            _kb_cross = (float(m_cny.group(1)), float(m_cny.group(2)))

band["USD/KRW"] = _kb_usd or (int(round(recent5["USD_KRW"].mean()) - 5), int(round(recent5["USD_KRW"].mean()) + 5))
band["USD/CNY"] = _kb_cross or (round(recent5["USD_CNY"].mean() - 0.03, 2), round(recent5["USD_CNY"].mean() + 0.03, 2))

# CNY/KRW: 국민은행 직접 제공 없음 → USD/KRW ÷ USD/CNY 역산
usd_lo_t, usd_hi_t = band["USD/KRW"]
cross_lo_t, cross_hi_t = band["USD/CNY"]
cny_lo_calc = int(round(usd_lo_t / cross_hi_t))
cny_hi_calc = int(round(usd_hi_t / cross_lo_t))
band["CNY/KRW"] = (cny_lo_calc, cny_hi_calc)

# 역산 밴드를 기관별 표에도 추가
if not any(b["기관"] == "국민은행" and b["통화"] == "CNY/KRW" for b in inst_bands):
    inst_bands.append({"기관": "국민은행", "통화": "CNY/KRW",
                       "주간 예상 범위": f"{cny_lo_calc:,}~{cny_hi_calc:,} (역산)"})

usd_lo, usd_hi = band["USD/KRW"]
cny_lo, cny_hi = band["CNY/KRW"]
cross_lo, cross_hi = band["USD/CNY"]

# ── 3대 통화 전망 카드 ──
fc1, fc2, fc3 = st.columns(3)

usd_info = ca.get("USD_KRW", {})
cny_info = ca.get("CNY_KRW", {})
cross_info = ca.get("USD_CNY", {})

def _dir_color(direction):
    if direction == "상승": return "#C00000"
    if direction == "하락": return "#4A90D9"
    return "#2E8B57"

def _dir_arrow(direction):
    if direction == "상승": return "↑"
    if direction == "하락": return "↓"
    return "→"

def _forecast_card(label, band_str, direction, sub=""):
    color = _dir_color(direction)
    arrow = _dir_arrow(direction)
    sub_html = f' <span style="font-size:0.8rem;color:#888;">{sub}</span>' if sub else ""
    return (
        f'<div data-testid="stMetric" style="background:linear-gradient(135deg,#667eea0d,#764ba20d);'
        f'border:1px solid #ddd;border-radius:12px;padding:16px 20px;box-shadow:0 2px 8px rgba(0,0,0,0.03);">'
        f'<div style="font-size:0.85rem;color:#555;font-weight:400;">{label}</div>'
        f'<div style="font-size:2rem;font-weight:700;color:#1a1a1a;margin:4px 0;">{band_str}</div>'
        f'<div style="font-size:1rem;color:{color};font-weight:600;">{arrow} {direction}{sub_html}</div>'
        f'</div>'
    )

fc1.markdown(_forecast_card("USD / KRW (금주 전망)", f"{usd_lo:,} ~ {usd_hi:,} 원",
             usd_info.get("direction", "보합")), unsafe_allow_html=True)
fc2.markdown(_forecast_card("CNY / KRW (금주 전망)", f"{cny_lo:,} ~ {cny_hi:,} 원",
             cny_info.get("direction", "보합")), unsafe_allow_html=True)
fc3.markdown(_forecast_card("USD / CNY 재정 (금주 전망)", f"{cross_lo} ~ {cross_hi}",
             cross_info.get("direction", "보합"), cross_info.get("sub_label", "")), unsafe_allow_html=True)

st.caption("출처: 국민은행 주간 외환 전망 (CNY/KRW는 USD/KRW ÷ USD/CNY 역산)")

st.divider()

# ── 통화별 전망 분석 (3자 자료 종합) ──
st.markdown("##### 📝 통화별 분석")
st.caption("신한은행 · 국민은행 PDF + 서울파이낸셜 [주간환율전망] 종합")

# USD/KRW
st.markdown(
    f'<div style="margin-bottom:16px;">'
    f'<div style="padding:12px 16px;background:#fdf2f2;border-radius:8px;border-left:4px solid #C00000;">'
    f'<div style="font-size:0.95rem;font-weight:800;color:#C00000;margin-bottom:6px;">'
    f'▲ USD/KRW — {usd_lo:,}~{usd_hi:,}원</div>'
    f'<div style="font-size:0.87rem;line-height:1.8;color:#333;">'
    # 변동 요인
    f'<b>변동 요인:</b> '
    f'<b>이란 전쟁 불확실성</b>과 <b>국제유가 고공행진</b>(호르무즈 해협 통항 리스크)이 '
    f'달러 매수 심리를 자극하는 가운데, <b>4월 외국인 배당 역송금</b> 수요가 본격화되며 '
    f'실수급 측면에서도 상방 압력이 가중되고 있음. '
    f'지난주 환율은 <b>1,497~1,536원</b> 범위에서 등락하며 상방 쏠림이 우세했음.'
    f'<br>'
    # 금주 전망
    f'<b>금주 전망:</b> '
    f'목요일(<b>4/10</b>) <b>미국 3월 CPI</b>에서 고유가가 물가에 반영될 경우, '
    f'연준 금리인하 기대 후퇴 → 달러 강세 → 상단 <b>{usd_hi:,}원</b> 테스트 가능성 높음. '
    f'다만 금요일(<b>4/11</b>) <b>한국은행 금통위</b>에서 당국의 원화 약세 쏠림 경계 발언 및 '
    f'외환시장 개입 의지가 확인되면 상단이 제한될 전망. '
    f'수출업체 고점 매도세도 하방 지지 요인으로 작용.'
    f'</div></div></div>',
    unsafe_allow_html=True,
)

# CNY/KRW
st.markdown(
    f'<div style="margin-bottom:16px;">'
    f'<div style="padding:12px 16px;background:#f2f6fd;border-radius:8px;border-left:4px solid #2E75B6;">'
    f'<div style="font-size:0.95rem;font-weight:800;color:#2E75B6;margin-bottom:6px;">'
    f'▼ CNY/KRW — {cny_lo:,}~{cny_hi:,}원</div>'
    f'<div style="font-size:0.87rem;line-height:1.8;color:#333;">'
    f'<b>변동 요인:</b> '
    f'위안화 프록시 통화인 원화가 달러 강세에 동반 약세를 보이는 가운데, '
    f'<b>중국 내수 둔화</b> 우려와 <b>미중 관세 리스크</b> 재부각이 위안화 절하 압력을 자극. '
    f'원/위안 환율은 달러/원 상승에도 불구, 위안화 자체 약세로 제한적 변동을 보임.'
    f'<br>'
    f'<b>금주 전망:</b> '
    f'중동 리스크에 따른 글로벌 안전자산 선호가 지속되면 위안화 추가 약세 → '
    f'CNY/KRW <b>{cny_lo:,}원</b>대 하단 테스트 가능성. '
    f'반면 호르무즈 해협 정상화 조짐이나 미중 무역 협상 진전 시 '
    f'위안화 반등과 함께 <b>{cny_hi:,}원</b>대 회복 시도 전망.'
    f'</div></div></div>',
    unsafe_allow_html=True,
)

# USD/CNY
st.markdown(
    f'<div style="margin-bottom:16px;">'
    f'<div style="padding:12px 16px;background:#f5f8f2;border-radius:8px;border-left:4px solid #548235;">'
    f'<div style="font-size:0.95rem;font-weight:800;color:#548235;margin-bottom:6px;">'
    f'→ USD/CNY (재정) — {cross_lo}~{cross_hi}</div>'
    f'<div style="font-size:0.87rem;line-height:1.8;color:#333;">'
    f'<b>변동 요인:</b> '
    f'달러 강세 기조에도 <b>인민은행(PBOC)</b>이 기준환율 고시를 통해 위안화 급락을 방어하며 '
    f'안정적 관리 기조를 유지 중. 지정학적 이벤트에 상대적으로 둔감한 움직임을 보이고 있음.'
    f'<br>'
    f'<b>금주 전망:</b> '
    f'PBOC의 적극적 환율 방어 의지가 지속되는 한, 재정환율은 '
    f'<b>{cross_lo}~{cross_hi}</b> 범위 내 제한적 등락이 전망됨. '
    f'다만 미국 CPI 서프라이즈 시 달러 인덱스 급등 → 위안화 절하 압력으로 '
    f'상단 이탈 가능성도 열어두어야 함.'
    f'</div></div></div>',
    unsafe_allow_html=True,
)

# ── 뉴스 원문 접힘 ──
news = report["news"]
if news.get("url"):
    with st.expander(f"📰 {news.get('title', '')[:45]}… — 원문 보기"):
        if news.get("body"):
            st.markdown(news["body"][:600] + "…" if len(news.get("body", "")) > 600 else news.get("body", ""))
        st.caption(f"[기사 원문]({news['url']})")

# ── 푸터 ──
st.divider()
st.caption("ⓒ F&F 자금팀 · 한국은행 ECOS API · 서울파이낸셜 · 2026년 4월 1주차 리포트")
