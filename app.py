from pathlib import Path
import html
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import streamlit.components.v1 as components
import yfinance as yf

st.set_page_config(page_title="股票指標評分看盤", page_icon="📈", layout="wide")

st.markdown(
    """
<style>
  :root {
    --app-ink: #18243a;
    --app-muted: #64748b;
    --app-line: #e4eaf2;
    --app-blue: #2563eb;
    --app-panel: #ffffff;
  }
  .stApp { background: #f5f7fb; color: var(--app-ink); }
  [data-testid="stHeader"] { background: rgba(245, 247, 251, .88); }
  [data-testid="stSidebar"] {
    background: #eef2f8; border-right: 1px solid #e1e7f0;
  }
  [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
  [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
    color: #24344e; letter-spacing: -.02em;
  }
  [data-testid="stSidebar"] [data-testid="stTextInput"],
  [data-testid="stSidebar"] [data-testid="stSelectbox"],
  [data-testid="stSidebar"] [data-testid="stExpander"] {
    border-radius: 11px;
  }
  [data-testid="stMainBlockContainer"] {
    max-width: 1580px; padding-top: 1.4rem; padding-bottom: 3rem;
  }
  .app-hero {
    display: flex; align-items: center; gap: 16px; margin: 0 0 20px;
    padding: 20px 24px; border: 1px solid #e1e9f5; border-radius: 18px;
    background: linear-gradient(115deg, #ffffff 0%, #f1f6ff 100%);
    box-shadow: 0 8px 24px rgba(27, 55, 100, .045);
  }
  .app-hero-icon {
    display: grid; place-items: center; flex: 0 0 48px; height: 48px;
    border-radius: 15px; color: #1d4ed8; background: #e7efff; font-size: 25px;
  }
  .app-hero-eyebrow { color: #54709b; font-size: 10px; font-weight: 800; letter-spacing: .14em; }
  .app-hero-title { margin-top: 2px; color: #17243b; font-size: 25px; font-weight: 800; line-height: 1.25; }
  .app-hero-copy { margin-top: 4px; color: #64748b; font-size: 13px; }
  .zone-status-card {
    box-sizing: border-box; min-height: 112px; height: 100%; padding: 16px 18px;
    border: 1px solid; border-radius: 15px; box-shadow: 0 4px 14px rgba(24, 36, 58, .035);
  }
  .zone-status-label { color: #64748b; font-size: 12px; font-weight: 650; }
  .zone-status-value { margin-top: 7px; font-size: 1.65rem; font-weight: 800; line-height: 1.2; }
  .zone-status-stage { display: inline-flex; margin-top: 8px; padding: 3px 9px; border-radius: 999px; background: rgba(255, 255, 255, .78); font-size: 12px; font-weight: 700; }
  .zone-value { color: #991b1b; background: #fff1f2; border-color: #fecdd3; }
  .zone-hot { color: #166534; background: #ecfdf3; border-color: #bbf7d0; }
  .zone-neutral { color: #92400e; background: #fffbeb; border-color: #fde68a; }
  .app-section-heading { display: flex; align-items: center; gap: 8px; margin: 1.1rem 0 .55rem; }
  .app-section-heading h3 { margin: 0; color: #24344e; font-size: 1.2rem; font-weight: 700; line-height: 1.35; }
  .app-help-mark {
    position: relative; display: inline-flex; flex: 0 0 17px; align-items: center; justify-content: center;
    width: 17px; height: 17px; border: 1.5px solid #8993a3; border-radius: 50%;
    color: #7b8494; background: transparent; font-size: 11px; font-weight: 700;
    line-height: 1; cursor: help; outline: none;
  }
  .app-help-mark:focus-visible { box-shadow: 0 0 0 3px rgba(37, 99, 235, .18); }
  .app-help-tooltip {
    position: absolute; z-index: 1000; top: calc(100% + 8px); left: 50%;
    width: max-content; max-width: min(340px, 78vw); padding: 10px 12px;
    border: 1px solid #334155; border-radius: 9px; color: #f8fafc; background: #1e293b;
    box-shadow: 0 8px 22px rgba(15, 23, 42, .2); font-size: 12px; font-weight: 450;
    line-height: 1.55; text-align: left; white-space: normal; overflow-wrap: anywhere;
    opacity: 0; visibility: hidden; pointer-events: none; transform: translate(-50%, 3px);
    transition: opacity .12s ease, transform .12s ease, visibility .12s ease;
  }
  .app-help-mark:hover .app-help-tooltip,
  .app-help-mark:focus .app-help-tooltip { opacity: 1; visibility: visible; transform: translate(-50%, 0); }
  [data-testid="stMetric"] {
    min-height: 112px; padding: 17px 18px; border: 1px solid var(--app-line);
    border-radius: 15px; background: var(--app-panel);
    box-shadow: 0 4px 14px rgba(24, 36, 58, .035);
  }
  [data-testid="stMetricLabel"] { color: var(--app-muted); font-size: 12px; font-weight: 650; }
  [data-testid="stMetricValue"] { color: var(--app-ink); font-weight: 760; }
  [data-testid="stTabs"] { margin-top: 18px; }
  [data-baseweb="tab-list"] {
    gap: 6px; padding: 6px; border: 1px solid var(--app-line); border-radius: 14px;
    background: #eef2f8;
  }
  [data-baseweb="tab"] {
    height: 42px; padding: 0 15px; border-radius: 10px; color: #53627a;
    font-size: 13px; font-weight: 650;
  }
  [data-baseweb="tab"][aria-selected="true"] {
    color: #1d4ed8; background: #ffffff; box-shadow: 0 2px 7px rgba(25, 47, 84, .09);
  }
  [data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 15px;
  }
  [data-testid="stExpander"] {
    border: 1px solid var(--app-line); border-radius: 13px; background: #fff;
  }
  [data-testid="stDataFrame"], [data-testid="stTable"] {
    overflow: hidden; border: 1px solid var(--app-line); border-radius: 12px;
  }
  .stButton > button, [data-testid="stDownloadButton"] button {
    border-radius: 10px; font-weight: 650; transition: transform .15s ease, box-shadow .15s ease;
  }
  .stButton > button:hover, [data-testid="stDownloadButton"] button:hover {
    transform: translateY(-1px); box-shadow: 0 5px 12px rgba(37, 99, 235, .13);
  }
  [data-testid="stTextInput"] input, [data-testid="stNumberInput"] input,
  [data-testid="stSelectbox"] [role="combobox"] {
    border-radius: 10px;
  }
  [data-testid="stAlert"] { border-radius: 12px; }
  [data-testid="stAlert"] [data-testid="stMarkdownContainer"],
  [data-testid="stAlert"] [data-testid="stMarkdownContainer"] p,
  [data-testid="stAlert"] [data-testid="stMarkdownContainer"] h3 {
    color: #24344e !important; opacity: 1 !important;
  }
  [data-testid="stWidgetLabel"] p, [data-testid="stRadio"] label p {
    color: #334155 !important; opacity: 1 !important;
  }
  [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] p {
    color: #526176 !important; opacity: 1 !important;
  }
  @media (max-width: 720px) {
    [data-testid="stMainBlockContainer"] { padding: .8rem .85rem 2rem; }
    .app-hero { gap: 11px; padding: 15px; margin-bottom: 13px; border-radius: 14px; }
    .app-hero-icon { flex-basis: 40px; height: 40px; border-radius: 12px; font-size: 21px; }
    .app-hero-title { font-size: 20px; }
    .app-hero-copy { font-size: 11px; line-height: 1.45; }
    [data-testid="stMetric"] { min-height: 94px; padding: 12px; }
    [data-testid="stMetricValue"] { font-size: 1.35rem; }
    [data-baseweb="tab-list"] { gap: 3px; padding: 4px; }
    [data-baseweb="tab"] { height: 38px; padding: 0 9px; font-size: 12px; }
  }
</style>
""",
    unsafe_allow_html=True,
)

# User identity is supplied by Streamlit's OIDC integration. Configure [auth]
# in Streamlit secrets before deploying; never accept an owner id from a widget.
try:
    _auth_secrets = st.secrets["auth"]
    _auth_ready = all(_auth_secrets.get(key) for key in ("redirect_uri", "cookie_secret", "client_id", "client_secret", "server_metadata_url"))
except Exception:
    _auth_ready = False
try:
    _supabase_secrets = st.secrets["supabase"]
    _supabase_ready = bool(_supabase_secrets.get("url") and _supabase_secrets.get("secret_key"))
except Exception:
    _supabase_ready = False

if not _auth_ready:
    st.error("尚未完成 Google 登入設定。請在 Streamlit Cloud 的 App settings → Secrets 填入 [auth] 設定後再重新啟動。")
    st.stop()
if not _supabase_ready:
    st.error("尚未完成資料庫設定。請在 Streamlit Cloud 的 App settings → Secrets 填入 [supabase] 的 url 與 secret_key。")
    st.stop()

if not st.user.is_logged_in:
    st.title("股票指標評分看盤")
    st.write("請使用 Google 帳戶登入，以載入個人自選股與訊號紀錄。")
    if st.button("使用 Google 登入", type="primary"):
        st.login()
    st.stop()

st.sidebar.caption(f"登入中：{st.user.get('email', 'Google 使用者')}")
if st.sidebar.button("登出", key="logout_button"):
    st.logout()
    st.stop()

YF_CACHE_DIR = Path(r"C:\Users\USER\Documents\Codex\2026-09-25\new-chat\outputs\StockAssistant\yf_cache")
YF_CACHE_DIR.mkdir(parents=True, exist_ok=True)
yf.set_tz_cache_location(str(YF_CACHE_DIR))
ZONE_STATS_COMMISSION_RATE = 0.1425 / 100
ZONE_STATS_SELL_TAX_RATE = 0.3 / 100
ZONE_STATS_SLIPPAGE_RATE = 0.05 / 100
LIVE_STOP_REFERENCE_ATR_MULTIPLE = 2.0

st.markdown(
    """
<div class="app-hero">
  <div class="app-hero-icon">📈</div>
  <div>
    <div class="app-hero-eyebrow">MARKET DASHBOARD · SHORT-TERM VIEW</div>
    <div class="app-hero-title">股票指標評分看盤</div>
    <div class="app-hero-copy">整合個股技術訊號與台美市場水位，快速掌握趨勢、風險與觀察依據。行情可能延遲，請留意資料日期。</div>
  </div>
</div>
""",
    unsafe_allow_html=True,
)


def section_heading_with_help(title: str, explanation: str) -> None:
    """Render a section title with an inline, hover/focus question-mark tooltip."""
    safe_title = html.escape(title)
    safe_explanation = html.escape(explanation, quote=True)
    st.markdown(
        f"""<div class="app-section-heading"><h3>{safe_title}</h3><span class="app-help-mark" tabindex="0" aria-label="補充說明">?<span class="app-help-tooltip" role="tooltip">{safe_explanation}</span></span></div>""",
        unsafe_allow_html=True,
    )


@st.cache_data(ttl=86400, show_spinner=False)
def load_taiwan_stock_directory() -> pd.DataFrame:
    """Download the daily official directory for TWSE listed and TPEx OTC companies."""
    sources = [
        ("上市", "TW", "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
         "公司代號", "公司簡稱", "公司名稱"),
        ("上櫃", "TWO", "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
         "SecuritiesCompanyCode", "CompanyAbbreviation", "CompanyName"),
    ]
    rows = []
    for market, suffix, url, code_key, short_key, full_key in sources:
        try:
            response = requests.get(url, timeout=20)
            response.raise_for_status()
            for item in response.json():
                code = str(item.get(code_key, "")).strip()
                short_name = str(item.get(short_key, "")).strip()
                full_name = str(item.get(full_key, "")).strip()
                if code and code.isdigit() and (short_name or full_name):
                    rows.append({
                        "代號": code, "簡稱": short_name or full_name,
                        "公司全名": full_name or short_name, "市場": market,
                        "symbol": f"{code}.{suffix}",
                    })
        except Exception:
            continue
    return pd.DataFrame(rows, columns=["代號", "簡稱", "公司全名", "市場", "symbol"])


def search_taiwan_stocks(query: str, directory: pd.DataFrame) -> pd.DataFrame:
    """Match Taiwan tickers, abbreviations, or official company names."""
    text = str(query).strip()
    if not text or directory.empty:
        return directory.iloc[0:0]
    code_query = text.upper().replace(".TWO", "").replace(".TW", "")
    exact = directory[directory["代號"] == code_query]
    if text.upper().endswith(".TWO"):
        exact = exact[exact["市場"] == "上櫃"]
    elif text.upper().endswith(".TW"):
        exact = exact[exact["市場"] == "上市"]
    if not exact.empty:
        return exact
    folded = text.casefold()
    mask = (
        directory["簡稱"].str.casefold().str.contains(folded, regex=False)
        | directory["公司全名"].str.casefold().str.contains(folded, regex=False)
        | directory["代號"].str.contains(text, regex=False)
    )
    return directory.loc[mask].head(30)


@st.cache_data(ttl=86400, show_spinner=False)
def get_yahoo_symbol_name(symbol: str) -> str:
    try:
        info = yf.Ticker(symbol).get_info()
        return str(info.get("shortName") or info.get("longName") or symbol)
    except Exception:
        return symbol


def get_symbol_name(symbol: str, directory: pd.DataFrame) -> str:
    normalized = str(symbol).strip().upper()
    if not directory.empty:
        row = directory[directory["symbol"].str.upper() == normalized]
        if not row.empty:
            return str(row.iloc[0]["簡稱"])
    return get_yahoo_symbol_name(normalized)


@st.cache_data(ttl=900, show_spinner=False)
def get_history(symbol: str, period: str = "2y") -> pd.DataFrame:
    data = yf.download(symbol, period=period, auto_adjust=False, progress=False, threads=False)
    if data.empty:
        return data
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)
    return prepare_history_prices(data.rename_axis("Date").dropna(subset=["Close"]))


@st.cache_data(ttl=3600, show_spinner=False)
def get_history_range(symbol: str, start_date: str, end_date: str) -> pd.DataFrame:
    """Fetch warm-up history before the selected backtest window."""
    start = pd.Timestamp(start_date) - pd.Timedelta(days=400)
    end = pd.Timestamp(end_date) + pd.Timedelta(days=1)
    data = yf.download(symbol, start=start.strftime("%Y-%m-%d"), end=end.strftime("%Y-%m-%d"),
                       auto_adjust=False, progress=False, threads=False)
    if data.empty:
        return data
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)
    return prepare_history_prices(data.rename_axis("Date").dropna(subset=["Close"]))


def prepare_history_prices(data: pd.DataFrame) -> pd.DataFrame:
    """Keep raw tradable OHLC beside adjusted OHLC used by scoring and backtests."""
    out = data.copy()
    price_columns = [column for column in ("Open", "High", "Low", "Close") if column in out]
    for column in price_columns:
        out[f"Raw{column}"] = out[column]
    if "Adj Close" in out and "Close" in out:
        factor = (out["Adj Close"] / out["Close"]).replace([np.inf, -np.inf], np.nan).fillna(1.0)
        for column in price_columns:
            out[column] = out[column] * factor
    return out


def add_indicators(data: pd.DataFrame) -> pd.DataFrame:
    df = data.copy()
    close, high, low = df["Close"], df["High"], df["Low"]
    df["SMA5"] = close.rolling(5).mean()
    df["SMA10"] = close.rolling(10).mean()
    df["SMA20"] = close.rolling(20).mean()
    df["SMA50"] = close.rolling(50).mean()
    df["SMA60"] = close.rolling(60).mean()
    df["SMA120"] = close.rolling(120).mean()
    df["SMA200"] = close.rolling(200).mean()
    df["STD20"] = close.rolling(20).std()
    df["BB_Upper"] = df["SMA20"] + 2 * df["STD20"]
    df["BB_Lower"] = df["SMA20"] - 2 * df["STD20"]
    df["RSI14"] = 100 - 100 / (1 + close.diff().clip(lower=0).ewm(alpha=1/14, adjust=False).mean() /
                                  (-close.diff().clip(upper=0)).ewm(alpha=1/14, adjust=False).mean())
    ema12 = close.ewm(span=12, adjust=False).mean()
    ema26 = close.ewm(span=26, adjust=False).mean()
    df["MACD"] = ema12 - ema26
    df["MACDSignal"] = df["MACD"].ewm(span=9, adjust=False).mean()
    lo9, hi9 = low.rolling(9).min(), high.rolling(9).max()
    rsv = 100 * (close - lo9) / (hi9 - lo9).replace(0, np.nan)
    df["K"] = rsv.ewm(com=2, adjust=False).mean()
    df["D"] = df["K"].ewm(com=2, adjust=False).mean()
    df["VolMA20"] = df["Volume"].rolling(20).mean()
    df["Bias20"] = (close / df["SMA20"] - 1) * 100
    df["OBV"] = (np.sign(close.diff()).fillna(0) * df["Volume"]).cumsum()
    df["OBV_MA20"] = df["OBV"].rolling(20).mean()
    prev_close = close.shift(1)
    tr = pd.concat([high - low, (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
    up_move, down_move = high.diff(), -low.diff()
    plus_dm = up_move.where((up_move > down_move) & (up_move > 0), 0.0)
    minus_dm = down_move.where((down_move > up_move) & (down_move > 0), 0.0)
    df["ATR14"] = tr.ewm(alpha=1/14, adjust=False).mean()
    plus_di = 100 * plus_dm.ewm(alpha=1/14, adjust=False).mean() / df["ATR14"]
    minus_di = 100 * minus_dm.ewm(alpha=1/14, adjust=False).mean() / df["ATR14"]
    dx = 100 * (plus_di - minus_di).abs() / (plus_di + minus_di).replace(0, np.nan)
    df["ADX14"] = dx.ewm(alpha=1/14, adjust=False).mean()
    df["PlusDI"] = plus_di
    df["MinusDI"] = minus_di
    return df


def score_stock(df: pd.DataFrame):
    row = df.iloc[-1]
    parts = []
    sell_points = 0
    def add(name, points, max_points, sell, reason):
        nonlocal sell_points
        sell_points += sell
        parts.append({"參考指標": name, "偏多分數": f"{points}/{max_points}",
                      "偏弱分數": f"{sell}/{max_points}", "判斷依據": reason})

    trend = 0
    trend += 10 if row.Close > row.SMA20 else 0
    trend += 10 if row.SMA20 > row.SMA60 else 0
    trend += 5 if row.SMA60 > row.SMA120 else 0
    trend_sell = (10 if row.Close < row.SMA20 else 0) + (10 if row.SMA20 < row.SMA60 else 0) + (5 if row.SMA60 < row.SMA120 else 0)
    add("均線趨勢", trend, 25, trend_sell, f"收盤 {row.Close:.2f}；20日線 {row.SMA20:.2f}；60日線 {row.SMA60:.2f}；120日線 {row.SMA120:.2f}")

    momentum = 0
    momentum += 10 if row.MACD > row.MACDSignal else 0
    momentum += 10 if 50 <= row.RSI14 <= 70 else (5 if 40 <= row.RSI14 < 50 else 0)
    momentum += 5 if row.K > row.D else 0
    momentum_sell = (10 if row.MACD < row.MACDSignal else 0) + (10 if row.RSI14 > 70 or row.RSI14 < 35 else 0) + (5 if row.K < row.D else 0)
    add("MACD／RSI／KD 動能", momentum, 25, momentum_sell, f"RSI {row.RSI14:.1f}；MACD {'在訊號線上方' if row.MACD > row.MACDSignal else '在訊號線下方'}；K {row.K:.1f}／D {row.D:.1f}")

    volume = 0
    volume += 10 if row.Close > row.SMA20 and row.Volume > row.VolMA20 else 0
    volume += 10 if row.Close >= df.Close.tail(20).max() * 0.98 else 0
    volume_sell = (10 if row.Close < row.SMA20 and row.Volume > row.VolMA20 else 0) + (10 if row.Close < df.Close.tail(20).min() * 1.02 else 0)
    add("量價與突破", volume, 20, volume_sell, f"今日量為20日均量 {row.Volume / row.VolMA20:.2f} 倍；布林上／下軌 {row.BB_Upper:.2f}／{row.BB_Lower:.2f}")

    risk = 0
    risk += 10 if -5 <= row.Bias20 <= 5 else (5 if -8 <= row.Bias20 <= 8 else 0)
    risk += 10 if 35 <= row.RSI14 <= 68 else 0
    risk_sell = (10 if row.Bias20 > 8 else 0) + (10 if row.RSI14 > 70 or row.RSI14 < 35 else 0)
    add("風險與乖離", risk, 20, risk_sell, f"股價對20日線乖離 {row.Bias20:.1f}%；ATR {row.ATR14:.2f}；ADX {row.ADX14:.1f}")
    total = round((trend + momentum + volume + risk) / 90 * 100)
    sell_total = round(min(sell_points, 90) / 90 * 100)
    label = "偏多觀察" if total >= 70 else ("中性觀望" if total >= 45 else "偏弱觀察")
    return total, sell_total, label, pd.DataFrame(parts)



def classify_zone(df: pd.DataFrame, buy_score: int, sell_score: int):
    """Classify continuation opportunity vs. elevated-energy/reversal risk."""
    row = df.iloc[-1]
    value_reasons, heat_reasons, weakness_reasons = [], [], []

    if row.RSI14 >= 70:
        heat_reasons.append(f"RSI {row.RSI14:.1f}，短線能量偏高")
    if row.Bias20 >= 8:
        heat_reasons.append(f"20日乖離 {row.Bias20:+.1f}%，價格明顯高於短期均線")
    if row.Close >= row.BB_Upper:
        heat_reasons.append("收盤觸及或高於布林上軌")
    if len(df) >= 21 and row.Close / df.Close.iloc[-21] - 1 >= 0.15:
        heat_reasons.append("近20個交易日漲幅達15%，短線升溫")
    if row.Close < row.SMA10:
        weakness_reasons.append("收盤跌破10日均線")
    if row.SMA5 < row.SMA10:
        weakness_reasons.append("5日均線低於10日均線")
    if row.MACD < row.MACDSignal:
        weakness_reasons.append("MACD低於訊號線")
    if row.K < row.D:
        weakness_reasons.append("KD短線動能轉弱")
    if row.Close < row.Open and row.Volume > row.VolMA20 * 1.5:
        weakness_reasons.append("放量收黑，出現賣壓")
    if len(df) >= 6 and row.SMA20 < df.SMA20.iloc[-6]:
        weakness_reasons.append("20日均線斜率轉下")

    # High energy alone is a caution, not an automatic sell signal.
    if heat_reasons:
        turning_down = len(weakness_reasons) >= 2 or row.Close < row.SMA20
        reasons = heat_reasons[:2] + (weakness_reasons[:2] if turning_down else [])
        if turning_down:
            stage = "轉弱確認"
            action = "高檔且已有轉弱確認；持股者可依風險承受度分批減碼或執行原定出場規則。"
        else:
            stage = "偏熱觀察"
            action = "目前偏熱但尚未確認轉弱；避免追價，可評估分批鎖利，不把單一過熱指標當成全數賣出訊號。"
        return "升溫區", stage, reasons, action

    trend_ok = (
        row.Close > row.SMA20 and row.SMA20 > row.SMA50
        and len(df) >= 6 and row.SMA20 > df.SMA20.iloc[-6]
    )
    momentum_ok = row.Close > df.Close.iloc[-11] and (row.MACD > row.MACDSignal or row.K > row.D) if len(df) >= 11 else False
    prior_high = df.Close.iloc[-21:-1].max() if len(df) >= 21 else np.nan
    breakout = (
        np.isfinite(prior_high) and row.Close > prior_high
        and row.Volume >= row.VolMA20 * 1.1
    )
    pullback = trend_ok and row.Close <= row.SMA20 * 1.03 and row.Close >= row.SMA20
    continuation = trend_ok and momentum_ok

    if buy_score >= 60 and sell_score < 40 and row.RSI14 < 70 and row.Bias20 < 8 and (breakout or pullback or continuation):
        if breakout:
            value_reasons.append("收盤突破前20日高點，成交量至少為20日均量的1.1倍")
        if pullback:
            value_reasons.append("上升趨勢中的股價回到20日線附近並守穩")
        if continuation:
            value_reasons.append("20／50日趨勢向上，近10日報酬為正且動能未轉弱")
        value_reasons.append(f"偏多分數 {buy_score}/100、偏弱風險 {sell_score}/100")
        return "價值區", "續漲條件成立", value_reasons[:4], "續漲條件較完整，列為候選；仍需設定部位與停損，這不是獲利保證。"

    neutral_reasons = []
    if not trend_ok:
        neutral_reasons.append("20／50日上升趨勢尚未同時成立")
    if not momentum_ok:
        neutral_reasons.append("近10日報酬或動能尚未確認")
    if buy_score < 60 or sell_score >= 40:
        neutral_reasons.append(f"偏多／偏弱分數 {buy_score}/{sell_score}，未達續漲條件")
    if not neutral_reasons:
        neutral_reasons.append("突破、回踩與續漲條件尚未形成明確優勢")
    return "空檔", "條件未齊", neutral_reasons[:4], "先觀察；等待趨勢與動能更一致。"


@st.cache_data(ttl=3600, show_spinner=False)
def classify_zone_history(df: pd.DataFrame) -> pd.Series:
    """Classify every candle using only that date and earlier data."""
    zones = []
    for end in range(len(df)):
        history = df.iloc[:end + 1]
        buy_score, sell_score, _, _ = score_stock(history)
        zone, _, _, _ = classify_zone(history, buy_score, sell_score)
        zones.append(zone)
    return pd.Series(zones, index=df.index, name="區域", dtype="object")


def current_timing_action(zone: str, stage: str, buy_score: int, sell_score: int,
                          holding: bool, market: dict | None, buy_threshold: int,
                          sell_threshold: int):
    """Turn the current technical state into a clear, non-automated timing label."""
    market_value = market.get("score") if market else None
    reasons = []
    if market_value is None:
        reasons.append("大盤資料不足，暫不給出新的買進候選判斷")
    else:
        market_date = pd.Timestamp(market["date"]).strftime("%Y-%m-%d")
        reasons.append(f"{market['benchmark_name']}環境分數 {market_value}/100，配置參考上限 {market['alloc']:.1f}%，資料日 {market_date}")

    if zone == "升溫區" and stage == "轉弱確認":
        reasons.append("個股處於升溫區且已有轉弱確認")
        if holding:
            return "偏向減碼／賣出候選", "持股者可檢查原定停損並考慮減碼或出場；未持有者不宜追買。", reasons, "error"
        return "暫不買進", "目前是高檔轉弱訊號；若未持有，先觀察，不把它當成買點。", reasons, "warning"

    if holding and sell_score >= sell_threshold:
        reasons.append(f"偏弱風險 {sell_score}/100，達到設定門檻 {sell_threshold}")
        return "偏向減碼／賣出候選", "偏弱訊號已達門檻；持股者可依預先設定的風險規則處理。", reasons, "error"

    if zone == "升溫區":
        reasons.append("短線能量偏高，但尚未確認轉弱")
        if holding:
            return "持有觀察／考慮分批鎖利", "目前沒有明確的追買條件；持股者可提高警覺並檢查停損。", reasons, "warning"
        return "等待，不追買", "短線偏熱但尚未轉弱；等待回檔或新的續漲條件，不追高。", reasons, "warning"

    if zone == "價值區":
        reasons.extend([f"偏多分數 {buy_score}/100", f"偏弱風險 {sell_score}/100"])
        if market_value is None:
            if holding:
                return "持有觀察，不加碼", "大盤資料不足；目前不新增買進部位，持股者依停損規則觀察。", reasons, "warning"
            return "等待大盤資料", "個股條件偏多，但無法確認市場水位；暫不列為新的買進候選。", reasons, "warning"
        if market_value < 40:
            if holding:
                return "持有觀察，不加碼", "個股條件偏多但大盤環境偏弱；暫不加碼，持股者檢查停損與曝險。", reasons, "warning"
            return "等待，不新增部位", "個股條件偏多，但大盤環境分數低於 40；先控制整體風險。", reasons, "warning"
        if holding:
            return "偏向續抱觀察", "個股續漲條件仍成立；這不等於適合加碼，請依停損與部位上限管理。", reasons, "info"
        if buy_score >= buy_threshold:
            return "偏向買進候選", "個股與大盤條件符合短線候選；依最新收盤資料判讀，下一交易日仍需確認價格與風險，不是盤中即時喊單。", reasons, "success"
        reasons.append(f"偏多分數未達你設定的買進門檻 {buy_threshold}")
        return "等待確認，不買進", "雖進入價值區，但偏多分數未達你設定的加嚴門檻。", reasons, "warning"

    if holding:
        reasons.append(f"偏多 {buy_score}/100、偏弱 {sell_score}/100；尚無明確買進或賣出訊號")
        return "持有觀察，不加碼", "目前區域條件不足以支持新買進；持股者依停損規則觀察。", reasons, "info"
    reasons.append(f"偏多分數 {buy_score}/100，尚未形成價值區條件")
    return "等待，不買進", "目前未達短線買進候選條件；等待趨勢、動能與市場環境更一致。", reasons, "warning"


def concise_timing_reason(action: str, zone: str, stage: str, buy_score: int,
                          sell_score: int, holding: bool, market: dict | None,
                          buy_threshold: int, sell_threshold: int) -> str:
    """Return one concise scanner reason; scores and dates are shown in other columns."""
    if action == "偏向買進候選":
        return f"價值區成立，偏多分數 {buy_score} 達門檻"
    if action == "等待確認，不買進":
        return f"偏多分數 {buy_score} 未達門檻 {buy_threshold}"
    if action == "等待，不新增部位":
        return "大盤環境偏弱，先不新增部位"
    if action == "等待大盤資料":
        return "大盤資料不足，暫不列買進候選"
    if action == "偏向減碼／賣出候選":
        if stage == "轉弱確認":
            return "高檔轉弱，出現賣壓訊號"
        return f"偏弱風險 {sell_score} 達門檻 {sell_threshold}"
    if action == "持有觀察／考慮分批鎖利":
        return "短線偏熱，尚未確認轉弱"
    if action == "偏向續抱觀察":
        return "續漲條件仍成立，尚未達減碼門檻"
    if action == "持有觀察，不加碼":
        if market is None or market.get("score", 0) < 40:
            return "大盤偏弱或資料不足，持股暫不加碼"
        return "尚無明確出場訊號，依停損規則觀察"
    if action == "暫不買進":
        return "高檔轉弱，避免進場"
    if action == "等待，不追買":
        return "短線偏熱，等待降溫或動能確認"
    if action == "等待，不買進":
        return "趨勢或動能條件尚未齊"
    if holding and sell_score >= sell_threshold:
        return "偏弱風險達減碼門檻"
    if zone == "空檔":
        return "目前趨勢訊號不明確"
    return "等待條件更明確"


@st.cache_data(ttl=3600, show_spinner=False)
def evaluate_zone_outcomes(df: pd.DataFrame, commission: float, sell_tax: float, slippage: float) -> pd.DataFrame:
    """Evaluate de-duplicated past zone signals using only subsequent closes."""
    horizons = (5, 10, 20)
    cost_buy = (1 + slippage) * (1 + commission)
    cost_sell = (1 - slippage) * (1 - commission - sell_tax)
    sampled = {"價值區": [], "升溫區｜偏熱觀察": [], "升溫區｜轉弱確認": []}
    last_sample = {key: -1000 for key in sampled}
    # Input already has warmed indicators; 20 rows provide breakout and recent-return lookback.
    last_start = 20

    for i in range(last_start, len(df) - max(horizons)):
        past = df.iloc[:i + 1]
        buy_score, sell_score, _, _ = score_stock(past)
        zone, stage, _, _ = classify_zone(past, buy_score, sell_score)
        bucket = "價值區" if zone == "價值區" else (f"升溫區｜{stage}" if zone == "升溫區" else None)
        if bucket not in sampled or i - last_sample[bucket] < 20:
            continue
        last_sample[bucket] = i
        entry = float(df.Close.iloc[i])
        future = df.iloc[i + 1:i + 1 + max(horizons)]
        if entry <= 0 or len(future) < max(horizons):
            continue
        event = {}
        for h in horizons:
            exit_px = float(df.Close.iloc[i + h])
            net_return = ((exit_px * cost_sell) / (entry * cost_buy) - 1) * 100
            adverse = (float(future.Low.iloc[:h].min()) / entry - 1) * 100
            event[h] = (net_return, adverse)
        sampled[bucket].append(event)

    rows = []
    for zone, events in sampled.items():
        for h in horizons:
            values = [e[h][0] for e in events if h in e]
            drawdowns = [e[20][1] for e in events if 20 in e]
            if not values:
                rows.append({
                    "訊號區域／狀態": zone,
                    "觀察期間(交易日)": h,
                    "去重後訊號數": 0,
                    "扣估計成本後正報酬率": np.nan,
                    "平均淨報酬": np.nan,
                    "20日內平均最大盤中跌幅": np.nan,
                })
                continue
            rows.append({
                "訊號區域／狀態": zone,
                "觀察期間(交易日)": h,
                "去重後訊號數": len(values),
                "扣估計成本後正報酬率": round(float(np.mean(np.array(values) > 0) * 100), 1),
                "平均淨報酬": round(float(np.mean(values)), 2),
                "20日內平均最大盤中跌幅": round(float(np.mean(drawdowns)), 2) if drawdowns else np.nan,
            })
    return pd.DataFrame(rows)



def add_score_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Calculate daily scores using only data available at each daily close."""
    out = df.copy()
    out["BuyRaw"] = (
        10 * (out.Close > out.SMA20) + 10 * (out.SMA20 > out.SMA60) + 5 * (out.SMA60 > out.SMA120)
        + 10 * (out.MACD > out.MACDSignal)
        + 10 * ((out.RSI14 >= 50) & (out.RSI14 <= 70))
        + 5 * ((out.RSI14 >= 40) & (out.RSI14 < 50))
        + 5 * (out.K > out.D)
        + 10 * ((out.Close > out.SMA20) & (out.Volume > out.VolMA20))
        + 10 * (out.Close >= out.Close.rolling(20).max() * 0.98)
        + 10 * ((out.Bias20 >= -5) & (out.Bias20 <= 5))
        + 5 * (((out.Bias20 >= -8) & (out.Bias20 < -5)) | ((out.Bias20 > 5) & (out.Bias20 <= 8)))
        + 10 * ((out.RSI14 >= 35) & (out.RSI14 <= 68))
    )
    out["SellRaw"] = (
        10 * (out.Close < out.SMA20) + 10 * (out.SMA20 < out.SMA60) + 5 * (out.SMA60 < out.SMA120)
        + 10 * (out.MACD < out.MACDSignal)
        + 10 * ((out.RSI14 > 70) | (out.RSI14 < 35))
        + 5 * (out.K < out.D)
        + 10 * ((out.Close < out.SMA20) & (out.Volume > out.VolMA20))
        + 10 * (out.Close < out.Close.rolling(20).min() * 1.02)
        + 10 * (out.Bias20 > 8)
        + 10 * ((out.RSI14 > 70) | (out.RSI14 < 35))
    )
    out["BuyScore"] = out.BuyRaw / 90 * 100
    out["SellScore"] = out.SellRaw / 90 * 100
    return out


def signal_reasons(row: pd.Series, buy: bool = True) -> str:
    reasons = []
    if buy:
        if row.Close > row.SMA20: reasons.append("收盤高於20日均線")
        if row.SMA20 > row.SMA60: reasons.append("20日線高於60日線")
        if row.SMA60 > row.SMA120: reasons.append("60日線高於120日線")
        if row.MACD > row.MACDSignal: reasons.append("MACD高於訊號線")
        if 50 <= row.RSI14 <= 70: reasons.append(f"RSI在強勢區({row.RSI14:.0f})")
        elif 40 <= row.RSI14 < 50: reasons.append(f"RSI由弱轉穩({row.RSI14:.0f})")
        if row.K > row.D: reasons.append("KD偏強")
        if row.Close > row.SMA20 and row.Volume > row.VolMA20: reasons.append("價量高於20日均值")
        if row.Close >= row.Close_20MAX * .98: reasons.append("接近20日收盤高點")
    else:
        if row.Close < row.SMA20: reasons.append("收盤跌破20日均線")
        if row.SMA20 < row.SMA60: reasons.append("20日線低於60日線")
        if row.MACD < row.MACDSignal: reasons.append("MACD低於訊號線")
        if row.RSI14 > 70: reasons.append(f"RSI偏熱({row.RSI14:.0f})")
        if row.RSI14 < 35: reasons.append(f"RSI偏弱({row.RSI14:.0f})")
        if row.K < row.D: reasons.append("KD偏弱")
        if row.Close < row.SMA20 and row.Volume > row.VolMA20: reasons.append("放量且收盤低於20日線")
    return "、".join(reasons) if reasons else "綜合指標分數達門檻"


def run_backtest(data: pd.DataFrame, buy_threshold: int, sell_threshold: int,
                 position_pct: float, capital: float, commission: float,
                 sell_tax: float, slippage: float, prior_buy_score: float = 0.0,
                 risk_fraction: float = 0.005, atr_stop_multiple: float = 2.0,
                 max_holding_days: int = 20):
    """Close-derived signals, next-session fills, ATR stop, risk-sized positions."""
    df = (data.copy() if {"BuyScore", "SellScore"}.issubset(data.columns) else add_score_columns(data))
    if "ATR14" not in df:
        df = add_indicators(df)
    if "MarketAlloc" not in df: df["MarketAlloc"] = 60.0
    if "MarketScore" not in df: df["MarketScore"] = 60.0
    if "MarketProxyClose" not in df: df["MarketProxyClose"] = df.Close
    df["Close_20MAX"] = df.Close.rolling(20).max()
    df = df.dropna(subset=["BuyScore", "SellScore", "Open", "High", "Low", "Close", "ATR14"]).copy()
    if len(df) < 2: return None
    cash, shares = float(capital), 0
    trades, curve, exposure = [], [], []
    entry_price = entry_cash = 0.0
    entry_date = entry_signal_date = None
    entry_index = None
    entry_reason = ""
    entry_score = entry_market_score = entry_alloc = entry_suggested_pct = 0.0
    entry_stop_price = entry_risk_budget = entry_risk_amount = 0.0
    pending = None

    def record_exit(date, fill, signal_date, reason, exit_score="—"):
        net_proceeds = shares * fill * (1 - commission - sell_tax)
        pnl = net_proceeds - entry_cash
        trades.append({
            "進場訊號日": pd.Timestamp(entry_signal_date).strftime("%Y-%m-%d"),
            "買進成交日": pd.Timestamp(entry_date).strftime("%Y-%m-%d"),
            "出場訊號日": signal_date if isinstance(signal_date, str) else pd.Timestamp(signal_date).strftime("%Y-%m-%d"),
            "賣出成交日": pd.Timestamp(date).strftime("%Y-%m-%d"),
            "買進價": round(entry_price, 2), "估算停損價": round(entry_stop_price, 2),
            "單筆風險上限": round(entry_risk_budget), "估計停損虧損": round(entry_risk_amount),
            "賣出價": round(fill, 2), "股數": shares,
            "建議投入比例": f"{entry_suggested_pct:.1f}%", "實際投入金額": round(entry_cash),
            "進場偏多分數": round(entry_score, 1), "當時市場分數": round(entry_market_score, 1),
            "進場理由": entry_reason, "出場偏弱分數": round(exit_score, 1) if isinstance(exit_score, (int, float, np.number)) else exit_score,
            "出場理由": reason, "報酬率(扣費)": round(pnl / entry_cash * 100, 2), "損益(扣費)": round(pnl),
        })
        return net_proceeds

    for i, (date, row) in enumerate(df.iterrows()):
        closed_by_stop = False
        if pending and pending["side"] == "sell" and shares:
            fill = float(row.Open) * (1 - slippage)
            cash += record_exit(date, fill, pending["signal_date"], pending["reason"], pending.get("score", "—"))
            shares, entry_price, entry_cash, entry_date, entry_signal_date = 0, 0.0, 0.0, None, None
            entry_index = None
            entry_stop_price = entry_risk_budget = entry_risk_amount = 0.0
        elif pending and pending["side"] == "buy" and not shares:
            fill = float(row.Open) * (1 + slippage)
            # Budget is constrained by risk, configured position size, market regime, and available cash.
            market_cap = float(np.clip(pending["market_alloc"], 10, 80)) / 100
            strength = float(np.clip(.5 + .5 * (pending["score"] - buy_threshold) / max(100 - buy_threshold, 1), .5, 1))
            suggested_pct = min(position_pct, market_cap) * strength
            atr_value = float(pending["atr"])
            stop_price = max(fill - atr_value * atr_stop_multiple, 0.0)
            stop_proceeds_per_share = stop_price * (1 - slippage) * (1 - commission - sell_tax)
            risk_per_share = fill * (1 + commission) - stop_proceeds_per_share
            risk_budget = max(cash, 0.0) * risk_fraction
            if atr_value > 0 and risk_per_share > 0 and stop_price > 0:
                risk_limited_shares = int(risk_budget / risk_per_share)
                allocation_budget = cash * suggested_pct
                budget = min(allocation_budget, risk_limited_shares * fill * (1 + commission))
                shares = int(budget / (fill * (1 + commission)))
            else:
                shares = 0
            if shares:
                entry_price, entry_cash, entry_date = fill, shares * fill * (1 + commission), date
                entry_signal_date, entry_reason = pending["signal_date"], pending["reason"]
                entry_score, entry_market_score = pending["score"], pending["market_score"]
                entry_alloc = entry_cash / capital * 100
                entry_suggested_pct = suggested_pct * 100
                entry_index = i
                entry_stop_price = stop_price
                entry_risk_budget = risk_budget
                entry_risk_amount = shares * risk_per_share
                cash -= entry_cash
        pending = None

        # Simulate an intraday stop with daily OHLC: gaps below the stop fill at the open;
        # otherwise a touched stop fills at the stop price, with configured slippage.
        if shares and entry_stop_price > 0 and float(row.Low) <= entry_stop_price:
            raw_exit = min(float(row.Open), entry_stop_price)
            fill = raw_exit * (1 - slippage)
            cash += record_exit(date, fill, date, "觸及ATR停損；跳空時以當日開盤估算成交", "—")
            shares, entry_price, entry_cash, entry_date, entry_signal_date = 0, 0.0, 0.0, None, None
            entry_index = None
            entry_stop_price = entry_risk_budget = entry_risk_amount = 0.0
            closed_by_stop = True

        marked_equity = cash + shares * float(row.Close)
        curve.append({"Date": date, "Equity": marked_equity})
        exposure.append(shares * float(row.Close) / marked_equity if marked_equity else 0.0)
        if i < len(df) - 1 and not closed_by_stop:
            prior_buy = float(df.BuyScore.iloc[i - 1]) if i > 0 else prior_buy_score
            if shares and float(row.SellScore) >= sell_threshold:
                custom_reason = row.get("SellReason", "")
                reason = str(custom_reason) if isinstance(custom_reason, str) and custom_reason else signal_reasons(row, False)
                pending = {"side":"sell", "signal_date":pd.Timestamp(date), "score":float(row.SellScore), "reason":reason}
            elif shares and entry_index is not None and (i - entry_index + 1) >= max_holding_days:
                pending = {"side":"sell", "signal_date":pd.Timestamp(date), "score":"—",
                           "reason":f"已持有 {i - entry_index + 1} 個交易日，達最長持有設定"}
            elif not shares and float(row.BuyScore) >= buy_threshold and prior_buy < buy_threshold:
                custom_reason = row.get("BuyReason", "")
                reason = str(custom_reason) if isinstance(custom_reason, str) and custom_reason else signal_reasons(row, True)
                pending = {"side":"buy", "signal_date":pd.Timestamp(date), "score":float(row.BuyScore),
                           "market_score":float(row.MarketScore), "market_alloc":float(row.MarketAlloc),
                           "atr":float(row.ATR14), "reason":reason}
    last_date, last_row = df.index[-1], df.iloc[-1]
    if shares:
        fill = float(last_row.Close) * (1 - slippage)
        cash += record_exit(last_date, fill, "期末結算", "回測區間結束，以期末收盤估算平倉", "—")
    curve[-1]["Equity"] = cash
    equity = pd.DataFrame(curve).set_index("Date")["Equity"]
    drawdown = (equity / equity.cummax() - 1) * 100
    years = max((equity.index[-1] - equity.index[0]).days / 365.25, 1 / 365.25)
    cagr = ((equity.iloc[-1] / capital) ** (1 / years) - 1) * 100 if equity.iloc[-1] > 0 else -100.0
    trades_df = pd.DataFrame(trades)
    first_open = float(df.Open.iloc[0]) * (1 + slippage)
    bh_shares = int(capital / (first_open * (1 + commission)))
    bh_left = capital - bh_shares * first_open * (1 + commission)
    benchmark = bh_shares * df.Close * (1 - slippage) + bh_left
    bh_final = bh_shares * float(df.Close.iloc[-1]) * (1 - slippage) * (1 - commission - sell_tax) + bh_left
    bh_return = (bh_final / capital - 1) * 100
    capped_shares = int(capital * position_pct / (first_open * (1 + commission)))
    capped_cash = capital - capped_shares * first_open * (1 + commission)
    capped_final = capped_shares * float(df.Close.iloc[-1]) * (1 - slippage) * (1 - commission - sell_tax) + capped_cash
    capped_benchmark = capped_shares * df.Close * (1 - slippage) + capped_cash
    capped_bh_return = (capped_final / capital - 1) * 100
    capped_bh_mdd = float(((capped_benchmark / capped_benchmark.cummax()) - 1).min() * 100)
    proxy_prices = df.MarketProxyClose.ffill().bfill()
    proxy_shares = int(capital / float(proxy_prices.iloc[0]))
    proxy_cash = capital - proxy_shares * float(proxy_prices.iloc[0])
    market_benchmark = proxy_shares * proxy_prices + proxy_cash
    market_return = (market_benchmark.iloc[-1] / capital - 1) * 100
    market_mdd = float(((market_benchmark / market_benchmark.cummax()) - 1).min() * 100)
    monthly_equity = equity.resample("M").last().pct_change().dropna()
    monthly_hold = benchmark.resample("M").last().pct_change().reindex(monthly_equity.index).dropna()
    monthly_market = market_benchmark.resample("M").last().pct_change().reindex(monthly_equity.index).dropna()
    common_hold = monthly_equity.index.intersection(monthly_hold.index)
    common_market = monthly_equity.index.intersection(monthly_market.index)
    mdd_bh = float(((benchmark / benchmark.cummax()) - 1).min() * 100)
    winners = trades_df[trades_df["損益(扣費)"] > 0] if not trades_df.empty else pd.DataFrame()
    wins = trades_df.loc[trades_df["損益(扣費)"] > 0, "損益(扣費)"].sum() if not trades_df.empty else 0
    losses = abs(trades_df.loc[trades_df["損益(扣費)"] < 0, "損益(扣費)"].sum()) if not trades_df.empty else 0
    return {"equity":equity, "benchmark":benchmark, "market_benchmark":market_benchmark,
        "market_proxy_name":str(df.get("MarketProxyName", pd.Series(["市場ETF"])).iloc[0]),
        "trades":trades_df, "return":(equity.iloc[-1]/capital-1)*100,
        "cagr":cagr, "mdd":float(drawdown.min()), "buy_hold":bh_return, "benchmark_mdd":mdd_bh,
        "capped_buy_hold":capped_bh_return, "capped_benchmark_mdd":capped_bh_mdd,
        "market_return":market_return, "market_mdd":market_mdd,
        "monthly_win_hold":float((monthly_equity.loc[common_hold] > monthly_hold.loc[common_hold]).mean()*100) if len(common_hold) else 0,
        "monthly_win_market":float((monthly_equity.loc[common_market] > monthly_market.loc[common_market]).mean()*100) if len(common_market) else 0,
        "monthly_count":min(len(common_hold),len(common_market)),
        "win_rate":len(winners)/len(trades_df)*100 if not trades_df.empty else 0,
        "profit_factor":wins/losses if losses else (float("inf") if wins else 0), "count":len(trades_df),
        "exposure":float(np.mean(exposure)*100), "start":equity.index[0], "end":equity.index[-1]}


def market_allocation_pct(score, position_rank=0.5, bias=0.0, reversal=False,
                          bear_trend=False, uptrend=False, sma60_change=0.0,
                          recovery_extension=0.0):
    """Continuous regime targets calibrated on training data and checked out of sample."""
    _, rank_arr, bias_arr, reversal_arr, bear_arr, uptrend_arr, slope_arr, recovery_arr = np.broadcast_arrays(
        np.asarray(score, dtype=float), np.asarray(position_rank, dtype=float),
        np.asarray(bias, dtype=float), np.asarray(reversal, dtype=bool),
        np.asarray(bear_trend, dtype=bool), np.asarray(uptrend, dtype=bool),
        np.asarray(sma60_change, dtype=float), np.asarray(recovery_extension, dtype=float),
    )
    allocation = np.full(rank_arr.shape, 45.0, dtype=float)  # neutral/waiting
    up_allocation = 65.0 + np.clip(slope_arr / 4.0, 0.0, 1.0) * 5.0
    recovery_allocation = 65.0 + np.clip(recovery_arr / 3.0, 0.0, 1.0) * 5.0
    bear_allocation = 45.0 - np.clip(-slope_arr / 8.0, 0.0, 1.0) * 10.0
    allocation = np.where(uptrend_arr, up_allocation, allocation)
    allocation = np.where(reversal_arr, recovery_allocation, allocation)
    allocation = np.where(bear_arr, bear_allocation, allocation)
    overheated_high = uptrend_arr & (rank_arr >= 0.90) & (bias_arr >= 7.0)
    heat_trim = (8.0 * np.clip((rank_arr - 0.90) / 0.10, 0.0, 1.0)
                 + 7.0 * np.clip((bias_arr - 7.0) / 8.0, 0.0, 1.0))
    allocation = np.where(overheated_high, np.maximum(55.0, allocation - heat_trim), allocation)
    result = np.round(allocation, 1)
    return float(result) if result.ndim == 0 else result


@st.cache_data(ttl=1800, show_spinner=False)
def get_cnn_fear_greed() -> dict | None:
    """Best-effort read of CNN's public page data; the endpoint is not a supported public API."""
    url = "https://production.dataviz.cnn.io/index/fearandgreed/graphdata"
    headers = {
        "User-Agent": "Mozilla/5.0 StockAssistant/1.0",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.cnn.com",
        "Referer": "https://www.cnn.com/markets/fear-and-greed",
    }
    try:
        response = requests.get(url, headers=headers, timeout=8)
        response.raise_for_status()
        payload = response.json()
        current = payload.get("fear_and_greed", {})
        score = current.get("score")
        updated = current.get("timestamp")
        if score is None:
            history = payload.get("fear_and_greed_historical", {}).get("data", [])
            if history:
                score = history[-1].get("y")
                updated = history[-1].get("x")
        score = float(score)
        if not np.isfinite(score) or not 0 <= score <= 100:
            return None
        rating = str(current.get("rating") or (
            "極度恐懼" if score < 25 else "恐懼" if score < 45 else
            "中性" if score <= 55 else "貪婪" if score <= 75 else "極度貪婪"
        ))
        return {"score": score, "rating": rating, "updated": str(updated or "")}
    except Exception:
        return None


@st.cache_data(ttl=1800, show_spinner=False)
def get_ptt_stock_sentiment() -> dict | None:
    """Small, title-only sample from PTT Stock Atom feed; not full-board sentiment."""
    url = "https://www.ptt.cc/atom/Stock.xml"
    try:
        response = requests.get(
            url, headers={"User-Agent": "Mozilla/5.0 StockAssistant/1.0"}, timeout=8
        )
        response.raise_for_status()
        root = ET.fromstring(response.content)
        atom_ns = "{http://www.w3.org/2005/Atom}"
        entries = root.findall(f".//{atom_ns}entry")
        titles = []
        for entry in entries:
            title = entry.findtext(f"{atom_ns}title", default="").strip()
            if title:
                titles.append(html.unescape(title))
        if not titles:
            titles = [item.findtext("title", default="").strip() for item in root.findall(".//item")]
            titles = [html.unescape(title) for title in titles if title]
        if not titles:
            return None
        bullish_terms = ("看多", "偏多", "多頭", "做多", "買進", "買點", "噴出", "起飛", "創高", "利多", "上看", "轉強")
        bearish_terms = ("看空", "偏空", "空頭", "做空", "賣出", "賣點", "崩跌", "重挫", "跌停", "利空", "下看", "轉弱")
        bullish = bearish = neutral = 0
        for title in titles:
            has_bull = any(term in title for term in bullish_terms)
            has_bear = any(term in title for term in bearish_terms)
            if has_bull and not has_bear:
                bullish += 1
            elif has_bear and not has_bull:
                bearish += 1
            else:
                neutral += 1
        directional = bullish + bearish
        ratio = bullish / directional if directional else None
        return {"sample_count": len(titles), "bullish": bullish, "bearish": bearish,
                "neutral": neutral, "bullish_ratio": ratio}
    except Exception:
        return None


def market_score(symbol: str = "2330.TW"):
    is_taiwan = str(symbol).upper().endswith((".TW", ".TWO"))
    benchmark_symbol = "^TWII" if is_taiwan else "^GSPC"
    benchmark_name = "台灣加權指數" if is_taiwan else "S&P 500"
    index_data = get_history(benchmark_symbol, "1y")
    vix = get_history("^VIX", "3mo")
    if index_data.empty or len(index_data) < 60:
        return None, f"無法取得足夠的{benchmark_name}資料"
    index_data["SMA20"] = index_data.Close.rolling(20).mean()
    index_data["SMA60"] = index_data.Close.rolling(60).mean()
    last = index_data.iloc[-1]
    trend60_points = 40 if last.Close > last.SMA60 else 10
    trend20_points = 20 if last.Close > last.SMA20 else 5
    bias = (last.Close / last.SMA20 - 1) * 100
    bias_points = 20 if -3 <= bias <= 5 else (10 if -7 <= bias <= 8 else 0)
    vix_value = float(vix.Close.iloc[-1]) if not vix.empty else np.nan
    if np.isfinite(vix_value):
        vix_points = 20 if vix_value < 20 else (10 if vix_value < 30 else 0)
    else:
        vix_points = 10
    score = trend60_points + trend20_points + bias_points + vix_points
    rolling_high = index_data.Close.rolling(252, min_periods=60).max().iloc[-1]
    rolling_low = index_data.Close.rolling(252, min_periods=60).min().iloc[-1]
    position_rank = float((last.Close - rolling_low) / (rolling_high - rolling_low)) if rolling_high > rolling_low else 0.5
    sma20_rising = bool(index_data["SMA20"].iloc[-1] > index_data["SMA20"].iloc[-6])
    sma60_rising = bool(index_data["SMA60"].iloc[-1] > index_data["SMA60"].iloc[-21])
    reversal_confirmed = bool(last.Close > last.SMA20 and sma20_rising and last.Close < last.SMA60)
    bear_trend = bool(last.Close < last.SMA20 and last.Close < last.SMA60
                      and last.SMA20 < last.SMA60 and not sma60_rising)
    confirmed_uptrend = bool(last.Close > last.SMA20 and last.Close > last.SMA60 and sma60_rising)
    overheated_high = bool(confirmed_uptrend and position_rank >= 0.90 and bias >= 7.0)
    sma60_change = float((last.SMA60 / index_data["SMA60"].iloc[-21] - 1) * 100)
    recovery_extension = float((last.Close / last.SMA20 - 1) * 100)
    base_alloc = float(market_allocation_pct(score, position_rank, bias, reversal_confirmed,
                                             bear_trend, confirmed_uptrend, sma60_change,
                                             recovery_extension))
    fear_greed = get_cnn_fear_greed()
    # PTT Stock board is a Taiwan-specific auxiliary signal.
    ptt_sentiment = get_ptt_stock_sentiment() if is_taiwan else None
    trend_confirmed = bool(last.Close > last.SMA20)
    cnn_adjustment = 0.0
    if fear_greed:
        fg_score = fear_greed["score"]
        if fg_score >= 75:
            cnn_adjustment = -4.0 * (fg_score - 75.0) / 25.0
        elif fg_score <= 25 and trend_confirmed:
            cnn_adjustment = 3.0 * (25.0 - fg_score) / 25.0
    ptt_adjustment = 0.0
    ptt_directional_count = (ptt_sentiment["bullish"] + ptt_sentiment["bearish"]) if ptt_sentiment else 0
    if ptt_sentiment and ptt_sentiment["sample_count"] >= 20 and ptt_directional_count >= 10 and ptt_sentiment["bullish_ratio"] is not None:
        bullish_ratio = ptt_sentiment["bullish_ratio"]
        if bullish_ratio >= 0.75:
            ptt_adjustment = -3.0 * (bullish_ratio - 0.75) / 0.25
        elif bullish_ratio <= 0.25 and trend_confirmed:
            ptt_adjustment = 2.0 * (0.25 - bullish_ratio) / 0.25
    alloc_adjustment = cnn_adjustment + ptt_adjustment
    alloc = float(np.clip(base_alloc + alloc_adjustment, 10.0, 85.0).round(1))
    market_state = (
        f"上升趨勢高位過熱，技術水位 {base_alloc:.1f}%" if overheated_high else
        f"弱勢轉強確認，技術水位 {base_alloc:.1f}%" if reversal_confirmed else
        f"確認空頭趨勢，技術水位 {base_alloc:.1f}%" if bear_trend else
        f"趨勢確認向上，技術水位 {base_alloc:.1f}%" if confirmed_uptrend else
        f"中性／等待確認，技術水位 {base_alloc:.1f}%"
    )
    return {"score": score, "alloc": alloc, "bias": bias, "vix": vix_value,
            "base_alloc": base_alloc, "alloc_adjustment": alloc_adjustment,
            "cnn_adjustment": cnn_adjustment, "ptt_adjustment": ptt_adjustment,
            "position_rank": position_rank, "reversal_confirmed": reversal_confirmed,
            "bear_trend": bear_trend,
            "confirmed_uptrend": confirmed_uptrend,
            "market_state": market_state,
            "allocation_parts": {"技術狀態水位": base_alloc},
            "fear_greed": fear_greed, "ptt_sentiment": ptt_sentiment,
            "score_parts": {"60日趨勢": (trend60_points, 40), "20日趨勢": (trend20_points, 20),
                            "20日乖離": (bias_points, 20), "VIX風險情緒": (vix_points, 20)},
            "date": index_data.index[-1], "benchmark_name": benchmark_name}, ""


def market_state_label(snapshot: dict) -> str:
    raw_state = str(snapshot.get("market_state", ""))
    if raw_state.startswith("上升趨勢高位過熱"):
        return "趨勢偏多、位階偏熱"
    if snapshot.get("reversal_confirmed"):
        return "弱勢中出現轉強訊號"
    if snapshot.get("bear_trend"):
        return "技術趨勢偏弱"
    if snapshot.get("confirmed_uptrend"):
        return "技術趨勢偏多"
    return "趨勢尚未明確"


def market_state_color_class(snapshot: dict) -> str:
    """Return a restrained direction/strength color for the market status badge."""
    raw_state = str(snapshot.get("market_state", ""))
    if raw_state.startswith("上升趨勢高位過熱"):
        return "market-state-bull-hot"
    if snapshot.get("reversal_confirmed"):
        return "market-state-transition"
    if snapshot.get("bear_trend"):
        return "market-state-bear-strong"
    if snapshot.get("confirmed_uptrend"):
        return "market-state-bull-strong"
    return "market-state-neutral"


def combine_market_snapshots(taiwan: dict | None, us: dict | None):
    """Build a 50/50 Taiwan-US equity exposure reference from regional snapshots."""
    if not taiwan or not us:
        return None
    tw_weight = us_weight = 0.5
    base_alloc = tw_weight * taiwan["base_alloc"] + us_weight * us["base_alloc"]
    # CNN is global sentiment; average regional trend-gated adjustments so it is included once.
    cnn_adjustment = tw_weight * taiwan["cnn_adjustment"] + us_weight * us["cnn_adjustment"]
    # PTT only represents Taiwan: scale its local adjustment by the Taiwan market weight.
    ptt_adjustment = tw_weight * taiwan["ptt_adjustment"]
    alloc = float(np.clip(base_alloc + cnn_adjustment + ptt_adjustment, 10.0, 85.0).round(1))
    score = int(round(tw_weight * taiwan["score"] + us_weight * us["score"]))
    score_parts = {
        name: (round(tw_weight * taiwan["score_parts"][name][0] + us_weight * us["score_parts"][name][0], 1), weight)
        for name, (_, weight) in taiwan["score_parts"].items()
    }
    tw_state = market_state_label(taiwan)
    us_state = market_state_label(us)
    if taiwan["confirmed_uptrend"] and us["confirmed_uptrend"]:
        combined_state = "台美技術趨勢皆偏多"
    elif taiwan["bear_trend"] and us["bear_trend"]:
        combined_state = "台美技術趨勢皆偏弱"
    else:
        combined_state = "台美市場走勢分歧或尚待確認"
    return {
        "score": score, "alloc": alloc, "base_alloc": base_alloc,
        "cnn_adjustment": cnn_adjustment, "ptt_adjustment": ptt_adjustment,
        "bias": tw_weight * taiwan["bias"] + us_weight * us["bias"],
        "position_rank": tw_weight * taiwan["position_rank"] + us_weight * us["position_rank"],
        "vix": us["vix"], "market_state": combined_state,
        "date": max(pd.Timestamp(taiwan["date"]), pd.Timestamp(us["date"])),
        "fear_greed": us.get("fear_greed") or taiwan.get("fear_greed"),
        "ptt_sentiment": taiwan.get("ptt_sentiment"), "score_parts": score_parts,
        "taiwan": taiwan, "us": us, "taiwan_state": tw_state, "us_state": us_state,
        "taiwan_weight": tw_weight, "us_weight": us_weight,
    }


def add_market_columns(stock: pd.DataFrame, symbol: str, start_date: str, end_date: str) -> pd.DataFrame:
    """Historical market regime; same-day close is used only for next-session sizing."""
    benchmark_symbol = "^TWII" if symbol.endswith((".TW", ".TWO")) else "^GSPC"
    market = get_history_range(benchmark_symbol, start_date, end_date)
    out = stock.copy()
    proxy_symbol = "0050.TW" if symbol.endswith((".TW", ".TWO")) else "SPY"
    proxy = get_history_range(proxy_symbol, start_date, end_date)
    if not proxy.empty:
        out["MarketProxyClose"] = proxy.Close.reindex(out.index, method="ffill").bfill()
    else:
        out["MarketProxyClose"] = out.Close
    out["MarketProxyName"] = proxy_symbol
    if market.empty:
        out["MarketScore"], out["MarketAlloc"] = 60.0, 45.0
        return out
    close = market.Close
    sma20, sma60 = close.rolling(20).mean(), close.rolling(60).mean()
    bias = (close / sma20 - 1) * 100
    vix = get_history_range("^VIX", start_date, end_date)
    if not vix.empty:
        vix_close = vix.Close.reindex(close.index, method="ffill")
        vix_points = pd.Series(np.where(vix_close < 20, 20, np.where(vix_close < 30, 10, 0)), index=close.index)
        vix_points = vix_points.where(vix_close.notna(), 10)
    else:
        vix_points = pd.Series(10, index=close.index)
    score = (close.gt(sma60) * 40 + close.le(sma60) * 10
             + close.gt(sma20) * 20 + close.le(sma20) * 5
             + bias.between(-3, 5) * 20 + ((bias.between(-7, 8)) & ~bias.between(-3, 5)) * 10
             + vix_points)
    rolling_high = close.rolling(252, min_periods=60).max()
    rolling_low = close.rolling(252, min_periods=60).min()
    position_rank = ((close - rolling_low) / (rolling_high - rolling_low)).replace([np.inf, -np.inf], np.nan).fillna(0.5)
    reversal = close.gt(sma20) & sma20.gt(sma20.shift(5)) & close.lt(sma60)
    sma60_rising = sma60.gt(sma60.shift(20))
    bear_trend = close.lt(sma20) & close.lt(sma60) & sma20.lt(sma60) & ~sma60_rising
    uptrend = close.gt(sma20) & close.gt(sma60) & sma60_rising
    sma60_change = (sma60 / sma60.shift(20) - 1) * 100
    recovery_extension = (close / sma20 - 1) * 100
    alloc = pd.Series(market_allocation_pct(score.to_numpy(), position_rank.to_numpy(), bias.to_numpy(),
                                            reversal.to_numpy(), bear_trend.to_numpy(),
                                            uptrend.to_numpy(), sma60_change.to_numpy(),
                                            recovery_extension.to_numpy()), index=score.index)
    aligned_score = score.reindex(out.index, method="ffill").fillna(50)
    aligned_alloc = alloc.reindex(out.index, method="ffill").fillna(45)
    out["MarketScore"], out["MarketAlloc"] = aligned_score, aligned_alloc
    return out


def load_personal_watchlist(owner_id: str) -> dict:
    """Read one user's workspace using the server-only Supabase key."""
    config = st.secrets["supabase"]
    base_url = str(config["url"]).rstrip("/")
    secret_key = str(config["secret_key"])
    response = requests.get(
        f"{base_url}/rest/v1/user_watchlists",
        params={"owner_id": f"eq.{owner_id}", "select": "symbols,signal_state,signal_events", "limit": "1"},
        headers={"apikey": secret_key},
        timeout=15,
    )
    response.raise_for_status()
    rows = response.json()
    if not rows:
        return {"symbols": ["2330.TW", "2454.TW", "0050.TW"], "signal_state": {}, "signal_events": []}
    row = rows[0]
    return {
        "symbols": row.get("symbols") if isinstance(row.get("symbols"), list) else [],
        "signal_state": row.get("signal_state") if isinstance(row.get("signal_state"), dict) else {},
        "signal_events": row.get("signal_events") if isinstance(row.get("signal_events"), list) else [],
    }


def save_personal_watchlist(owner_id: str) -> None:
    """Save only the authenticated user's data; keep the Supabase key server-side."""
    config = st.secrets["supabase"]
    base_url = str(config["url"]).rstrip("/")
    secret_key = str(config["secret_key"])
    payload = [{
        "owner_id": owner_id,
        "symbols": st.session_state.get("watchlist_symbols", []),
        "signal_state": st.session_state.get("watchlist_signal_state", {}),
        "signal_events": st.session_state.get("watchlist_signal_events", []),
        "updated_at": pd.Timestamp.now(tz="UTC").isoformat(),
    }]
    response = requests.post(
        f"{base_url}/rest/v1/user_watchlists",
        params={"on_conflict": "owner_id"},
        headers={
            "apikey": secret_key,
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates,return=minimal",
        },
        json=payload,
        timeout=15,
    )
    response.raise_for_status()

with st.sidebar:
    st.header("🔎 分析設定")
    stock_directory = load_taiwan_stock_directory()
    ticker_query = st.text_input(
        "股票代號或名稱",
        value="2330.TW",
        placeholder="例：2330、2330.TW、台積電",
        help="台股可輸入代號或中文簡稱；美股等其他市場請輸入代號，例如 AAPL。",
    )
    matches = search_taiwan_stocks(ticker_query, stock_directory)
    query_is_code = bool(ticker_query.strip()) and all(
        char.isascii() and (char.isalnum() or char in ".^=-") for char in ticker_query.strip()
    )
    if not matches.empty:
        symbols = matches["symbol"].tolist()
        name_by_symbol = dict(zip(matches["symbol"], matches["簡稱"]))
        market_by_symbol = dict(zip(matches["symbol"], matches["市場"]))
        ticker = st.selectbox(
            "搜尋結果",
            symbols,
            format_func=lambda symbol: f"{symbol}｜{name_by_symbol[symbol]}｜{market_by_symbol[symbol]}",
            key="stock_search_result",
        )
        company_name = name_by_symbol.get(ticker, ticker)
    elif query_is_code:
        ticker = ticker_query.strip().upper()
        company_name = None
        if not stock_directory.empty and ticker.replace(".TWO", "").replace(".TW", "").isdigit():
            st.warning("目錄中找不到這個台股代號；請確認代號，或輸入其他市場的股票代號。")
    else:
        st.error("沒有找到相符的台股公司名稱。可試輸入股票代號，或輸入較短的公司簡稱。")
        st.stop()
    # Load five years so the chart's own range buttons can zoom out beyond its one-year default view.
    history_period = "5y"
    st.markdown("#### 短線判斷設定")
    with st.expander("進階設定：買賣訊號門檻", expanded=False):
        st.caption("一般使用者可先保留預設值 60。調低買進門檻會增加買進候選；調高賣出風險門檻會減少賣出提示。這些門檻會影響目前判讀、多股掃描與回測結果。")
        buy_threshold = st.slider(
            "買進候選最低分數", 40, 90, 60, 5,
            help="尚未持有時，個股仍須先符合價值區與大盤條件；偏多分數再達此門檻，才列為買進候選。回測也使用此門檻決定進場。",
        )
        sell_threshold = st.slider(
            "減碼／賣出風險分數", 40, 90, 60, 5,
            help="目前持有時，偏弱風險分數達此門檻會提示減碼／賣出候選；回測則模擬下一交易日開盤出場。設得越低，越容易出現風險提示。",
        )
    if st.button("🔄 更新行情", type="primary", width="stretch"):
        st.cache_data.clear()

try:
    data = get_history(ticker.strip().upper(), history_period)
    if data.empty or len(data) < 120:
        st.error("取得的歷史資料不足。請檢查代碼或稍後再試。")
        st.stop()
    df = add_indicators(data)
    df = df.dropna(subset=["SMA120", "RSI14", "K", "D", "VolMA20", "ADX14"])
    total, sell_score, label, breakdown = score_stock(df)
    last = df.iloc[-1]
    stock_name = company_name or get_symbol_name(ticker, stock_directory)

    try:
        market_snapshot, market_error = market_score(ticker.strip().upper())
    except Exception as market_exc:
        market_snapshot, market_error = None, str(market_exc)

    try:
        if str(ticker).upper().endswith((".TW", ".TWO")):
            taiwan_market_snapshot = market_snapshot
            taiwan_market_error = market_error
            us_market_snapshot, us_market_error = market_score("AAPL")
        else:
            us_market_snapshot = market_snapshot
            us_market_error = market_error
            taiwan_market_snapshot, taiwan_market_error = market_score("2330.TW")
    except Exception as market_exc:
        taiwan_market_snapshot, us_market_snapshot = None, None
        taiwan_market_error = us_market_error = str(market_exc)
    combined_market_snapshot = combine_market_snapshots(taiwan_market_snapshot, us_market_snapshot)

    if combined_market_snapshot:
        market_state_tw = html.escape(combined_market_snapshot["taiwan_state"])
        market_state_us = html.escape(combined_market_snapshot["us_state"])
        market_state_class_tw = market_state_color_class(taiwan_market_snapshot)
        market_state_class_us = market_state_color_class(us_market_snapshot)
        market_alloc_display = f"{combined_market_snapshot['alloc']:.1f}%"
        market_alloc_tw = f"{taiwan_market_snapshot['alloc']:.1f}%"
        market_alloc_us = f"{us_market_snapshot['alloc']:.1f}%"
        market_date_tw = pd.Timestamp(taiwan_market_snapshot["date"]).strftime("%Y-%m-%d")
        market_date_us = pd.Timestamp(us_market_snapshot["date"]).strftime("%Y-%m-%d")
        st.markdown(
            f"""
<style>
  .market-allocation-banner {{
    display: flex; align-items: center; gap: 22px; box-sizing: border-box;
    width: 100%; margin: 4px 0 14px; padding: 12px 18px;
    border: 1px solid #dbe5f1; border-left: 5px solid #2563eb; border-radius: 14px;
    background: linear-gradient(105deg, #f5f9ff 0%, #ffffff 72%);
    box-shadow: 0 3px 12px rgba(30, 64, 175, .06);
  }}
  .market-allocation-title {{ color: #334155; font-size: 14px; font-weight: 700; white-space: nowrap; }}
  .market-allocation-note {{ margin-top: 3px; color: #64748b; font-size: 11px; }}
  .market-allocation-value {{ color: #1d4ed8; font-size: 30px; font-weight: 750; line-height: 1; white-space: nowrap; }}
  .market-allocation-divider {{ width: 1px; height: 38px; background: #dbe5f1; }}
  .market-allocation-state {{ display: flex; flex: 1; flex-direction: column; gap: 6px; min-width: 0; color: #1e293b; font-size: 13px; font-weight: 650; }}
  .market-allocation-row {{ display: flex; align-items: center; gap: 8px; min-width: 0; }}
  .market-allocation-region {{ min-width: 38px; color: #475569; font-weight: 750; }}
  .market-allocation-regional-value {{ padding: 2px 7px; border-radius: 999px; color: #1d4ed8; background: #eaf1ff; font-weight: 750; white-space: nowrap; }}
  .market-state-badge {{ display: inline-flex; align-items: center; padding: 3px 9px; border: 1px solid transparent; border-radius: 999px; font-size: 12px; font-weight: 700; line-height: 1.35; }}
  .market-state-bull-strong {{ color: #991b1b; background: #fee2e2; border-color: #fecaca; }}
  .market-state-bear-strong {{ color: #14532d; background: #dcfce7; border-color: #86efac; }}
  .market-state-transition {{ color: #92400e; background: #fef3c7; border-color: #fcd34d; }}
  .market-state-neutral {{ color: #a16207; background: #fefce8; border-color: #fde68a; }}
  .market-state-bull-hot {{ color: #b91c1c; background: #fff7ed; border-color: #fdba74; }}
  .market-allocation-date {{ margin-left: auto; color: #64748b; font-size: 11px; font-weight: 400; white-space: nowrap; }}
  @media (max-width: 700px) {{
    .market-allocation-banner {{ gap: 10px; padding: 11px 12px; flex-wrap: wrap; }}
    .market-allocation-value {{ font-size: 25px; }}
    .market-allocation-divider {{ display: none; }}
    .market-allocation-state {{ flex-basis: 55%; font-size: 12px; }}
    .market-allocation-row {{ gap: 5px; flex-wrap: wrap; }}
    .market-allocation-date {{ font-size: 10px; }}
    .market-allocation-note {{ font-size: 10px; }}
  }}
</style>
<div class="market-allocation-banner">
  <div>
    <div class="market-allocation-title">🌏 台美合併配置參考</div>
    <div class="market-allocation-note">台股／美股各 50% 權重</div>
  </div>
  <div class="market-allocation-value">{market_alloc_display}</div>
  <div class="market-allocation-divider"></div>
  <div class="market-allocation-state">
    <div class="market-allocation-row"><span class="market-allocation-region">台股：</span><span class="market-allocation-regional-value">{market_alloc_tw}</span><span class="market-state-badge {market_state_class_tw}">{market_state_tw}</span><span class="market-allocation-date">資料日 {market_date_tw}</span></div>
    <div class="market-allocation-row"><span class="market-allocation-region">美股：</span><span class="market-allocation-regional-value">{market_alloc_us}</span><span class="market-state-badge {market_state_class_us}">{market_state_us}</span><span class="market-allocation-date">資料日 {market_date_us}</span></div>
  </div>
</div>
<div style="margin:-9px 0 13px 5px;color:#64748b;font-size:11px;">
  台股與美股水位為各自市場參考值；合併值按 50% 權重計算，非相加。顏色表示技術方向與狀態強弱；橘框代表偏多但位階偏熱。個股判讀仍依所屬市場；整體配置非單筆下單比例。
</div>
""",
            unsafe_allow_html=True,
        )
    else:
        st.info("台股與美股資料需同時可用，才能計算合併配置水位。" + (f"台股：{taiwan_market_error}；美股：{us_market_error}" if taiwan_market_error or us_market_error else ""))

    st.subheader(f"{stock_name}（{ticker}）")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("最新收盤價", f"{last.Close:,.2f}", f"{last.Close / df.Close.iloc[-2] - 1:+.2%}")
    c2.metric("偏多評分", f"{total} / 100", label)
    c3.metric("偏弱風險", f"{sell_score} / 100", "分數越高，偏弱訊號越多")
    c4.metric("資料日期", pd.Timestamp(df.index[-1]).strftime("%Y-%m-%d"))

    overview_tab, indicators_tab, scanner_tab, backtest_tab, market_tab, guide_tab = st.tabs(
        ["📊 總覽", "🧭 指標明細", "🔍 多股掃描", "🧪 策略回測", "🌏 市場水位", "💡 使用說明"])
    zone, zone_stage, zone_reasons, zone_action = classify_zone(df, total, sell_score)
    with overview_tab:
        st.subheader("個股狀態與判斷理由")
        st.caption("操作狀態依最近完整交易日收盤資料判斷，非盤中即時建議；目前只用技術指標與市場水位，不含基本面及籌碼，也不會自動下單。")
        z1, z2, z3 = st.columns([1, 2, 2])
        zone_style = {"價值區": "zone-value", "升溫區": "zone-hot", "空檔": "zone-neutral"}.get(zone, "zone-neutral")
        z1.markdown(
            f"""<div class="zone-status-card {zone_style}"><div class="zone-status-label">短線區域</div><div class="zone-status-value">{html.escape(str(zone))}</div><div class="zone-status-stage">{html.escape(str(zone_stage))}</div></div>""",
            unsafe_allow_html=True,
        )
        z2.markdown("**為什麼判成這個區域**\n\n" + "\n\n".join(f"- {reason}" for reason in zone_reasons))
        z3.markdown(f"**區域解讀**\n\n{zone_action}")
        st.info("資料涵蓋狀態：目前區域依技術指標判斷；基本面、財報與籌碼面尚未接入。")

        st.subheader("現在偏向買進、等待，還是賣出？")
        holding_choice = st.radio(
            "請先告訴系統你目前是否持有這檔股票",
            ["尚未持有（判斷是否買進）", "目前持有（判斷是否減碼／賣出）"],
            horizontal=True, key="live_holding_state",
        )
        holding = holding_choice.startswith("目前持有")
        timing_label, timing_detail, timing_reasons, timing_style = current_timing_action(
            zone, zone_stage, total, sell_score, holding, market_snapshot,
            buy_threshold, sell_threshold,
        )
        timing_message = f"### {timing_label}\n\n{timing_detail}"
        if timing_style == "success":
            st.success(timing_message)
        elif timing_style == "error":
            st.error(timing_message)
        elif timing_style == "warning":
            st.warning(timing_message)
        else:
            st.info(timing_message)
        timing_summary = concise_timing_reason(
            timing_label, zone, zone_stage, total, sell_score, holding,
            market_snapshot, buy_threshold, sell_threshold,
        )
        stock_data_date = pd.Timestamp(df.index[-1]).strftime("%Y-%m-%d")
        st.caption(f"主要依據：{timing_summary}　·　資料日 {stock_data_date}")
        with st.expander("查看完整判斷依據", expanded=False):
            st.markdown("**市場與操作條件**")
            st.markdown("\n".join(f"- {reason}" for reason in timing_reasons))
            zone_detail_reasons = [
                reason for reason in zone_reasons
                if not (("偏多" in reason or "偏弱" in reason) and "/100" in reason)
            ]
            if zone_detail_reasons:
                st.markdown("**區域判斷原因**")
                st.markdown("\n".join(f"- {reason}" for reason in zone_detail_reasons))
        with st.expander("短線交易計畫參考", expanded=False):
            if timing_label != "偏向買進候選":
                st.info(f"目前判讀為「{timing_label}」，不提供新的進場情境。等訊號符合買進候選後，再查看進場觀察帶與風險試算。")
            else:
                raw_plan_data = data.copy()
                for column in ("Open", "High", "Low", "Close"):
                    raw_column = f"Raw{column}"
                    if raw_column in raw_plan_data:
                        raw_plan_data[column] = raw_plan_data[raw_column]
                raw_plan_df = add_indicators(raw_plan_data).dropna(subset=["SMA20", "ATR14"])
                plan_row = raw_plan_df.iloc[-1]
                plan_center = float(plan_row.SMA20)
                plan_atr = float(plan_row.ATR14)
                entry_low = max(plan_center - 0.5 * plan_atr, 0.0)
                entry_high = plan_center + 0.5 * plan_atr
                plan_entry = (entry_low + entry_high) / 2
                plan_stop = max(plan_entry - LIVE_STOP_REFERENCE_ATR_MULTIPLE * plan_atr, 0.0)
                risk_per_share = max(plan_entry - plan_stop, 0.0)
                target_1r = plan_entry + risk_per_share
                target_2r = plan_entry + 2 * risk_per_share
                current_raw_close = float(plan_row.Close)
                if current_raw_close > entry_high:
                    entry_note = "現價高於觀察帶，避免只因訊號成立就追價。"
                elif current_raw_close < entry_low:
                    entry_note = "現價低於觀察帶，先重新確認訊號是否仍成立。"
                else:
                    entry_note = "現價位於觀察帶內；仍需依下一交易日實際價格決定是否進場。"
                st.caption("這是透明的試算規則，尚未做獨立策略驗證：進場觀察帶＝MA20 ± 0.5×ATR(14)；停損參考＝觀察帶中點 − 2×ATR；目標情境以 1R／2R 計算。採未還原報價，避免把除權息調整價當作可成交價。")
                p1, p2, p3 = st.columns(3)
                p1.metric("MA20 周邊觀察帶", f"{entry_low:,.2f} – {entry_high:,.2f}")
                p2.metric("停損參考價", f"{plan_stop:,.2f}", f"每股風險約 {risk_per_share:,.2f}")
                p3.metric("目標情境（非預測）", f"1R {target_1r:,.2f} / 2R {target_2r:,.2f}")
                st.caption(entry_note + " 停損只作風險規劃參考；跳空或流動性不足可能造成更大損失。")
                planned_amount = st.number_input(
                    "預計投入金額（依該股報價幣別；填 0 可略過）",
                    min_value=0, value=0, step=10000, key=f"plan_amount_{ticker}",
                )
                if planned_amount > 0 and plan_entry > 0:
                    planned_shares = int(float(planned_amount) // plan_entry)
                    estimated_risk = planned_shares * risk_per_share
                    r1, r2, r3 = st.columns(3)
                    r1.metric("參考股數", f"{planned_shares:,} 股")
                    r2.metric("參考買入金額", f"{planned_shares * plan_entry:,.2f}")
                    r3.metric("觸及停損時估計風險", f"{estimated_risk:,.2f}")
                    st.caption("股數為整股數學試算，未計手續費、稅費、台股交易單位或跳空風險；實際可買股數可能不同。")
        with st.expander("歷史訊號後續表現（5–20 個交易日）", expanded=False):
            st.caption("以過去相似訊號追蹤後續報酬與回撤；同一區域每 20 個交易日最多取一個訊號，報酬扣除估計交易成本。歷史樣本僅供參考，不代表未來。")
            with st.spinner("正在檢查歷史區域訊號…"):
                outcome_table = evaluate_zone_outcomes(
                    df, ZONE_STATS_COMMISSION_RATE, ZONE_STATS_SELL_TAX_RATE, ZONE_STATS_SLIPPAGE_RATE
                )
            if outcome_table.empty:
                st.warning("目前載入的近 5 年資料樣本不足，暫時無法評估區域訊號；股票上市時間較短時可能出現此情況。")
            else:
                st.dataframe(outcome_table, hide_index=True, width="stretch")
                sample_counts = outcome_table.groupby("訊號區域／狀態")["去重後訊號數"].first().reindex(
                    ["價值區", "升溫區｜偏熱觀察", "升溫區｜轉弱確認"], fill_value=0
                )
                min_n = int(sample_counts.min())
                if min_n < 20:
                    st.caption(f"目前至少一類訊號只有 {min_n} 個去重樣本，統計不穩定；建議累積更多股票與市場期間再調整門檻。")

        section_heading_with_help(
            "K 線圖",
            "滑鼠資料固定顯示在圖表上方；滾輪縮放 Y 軸，上方期間按鈕與下方滑桿調整 X 軸。K 棒、均線與布林通道採原始交易價格，評分和回測則使用還原價格，除權息附近可能略有差異。紅 K 代表收盤高於開盤，綠 K 代表收盤低於開盤。底色紅色為價值區、綠色為升溫區；空檔不著色。區域依當日與之前資料判斷。",
        )
        chart = go.Figure()
        with st.spinner("正在計算 K 棒的歷史區域背景…"):
            zone_history = classify_zone_history(df)
        zone_values = zone_history.tolist()
        for zone_name, fill in [
            ("價值區", "rgba(239, 68, 68, 0.13)"),
            ("升溫區", "rgba(34, 197, 94, 0.13)"),
        ]:
            segment_start = None
            for i in range(len(zone_values) + 1):
                active = i < len(zone_values) and zone_values[i] == zone_name
                if active and segment_start is None:
                    segment_start = i
                elif not active and segment_start is not None:
                    segment_end = i - 1
                    x1 = df.index[i] if i < len(df) else df.index[segment_end] + pd.Timedelta(days=1)
                    chart.add_vrect(
                        x0=df.index[segment_start], x1=x1,
                        fillcolor=fill, line_width=0, layer="below",
                    )
                    segment_start = None
        chart_df = df.copy()
        for column in ("Open", "High", "Low", "Close"):
            raw_column = f"Raw{column}"
            if raw_column in chart_df:
                chart_df[column] = chart_df[raw_column]
        chart_df = add_indicators(chart_df)
        ohlc_custom = chart_df[["Open", "High", "Low", "Close", "SMA5", "SMA20", "SMA60", "BB_Upper", "BB_Lower"]].to_numpy()
        chart.add_trace(go.Candlestick(
            x=chart_df.index, open=chart_df.Open, high=chart_df.High, low=chart_df.Low, close=chart_df.Close,
            customdata=ohlc_custom,
            name="K線",
            increasing={"line": {"color": "#e53935"}, "fillcolor": "#e53935"},
            decreasing={"line": {"color": "#159447"}, "fillcolor": "#159447"},
        ))
        chart.add_trace(go.Scatter(x=chart_df.index, y=chart_df.SMA5, customdata=ohlc_custom, name="MA5", line={"width": 1.8, "color": "#f59e0b"}))
        chart.add_trace(go.Scatter(x=chart_df.index, y=chart_df.SMA20, customdata=ohlc_custom, name="MA20", line={"width": 1.8, "color": "#2563eb"}))
        chart.add_trace(go.Scatter(x=chart_df.index, y=chart_df.SMA60, customdata=ohlc_custom, name="MA60", line={"width": 1.8, "color": "#8b5cf6"}))
        chart.add_trace(go.Scatter(x=chart_df.index, y=chart_df.BB_Upper, customdata=ohlc_custom, name="布林上軌", line={"width": 1.4, "dash": "dot", "color": "#475569"}))
        chart.add_trace(go.Scatter(x=chart_df.index, y=chart_df.BB_Lower, customdata=ohlc_custom, name="布林下軌", line={"width": 1.4, "dash": "dot", "color": "#06b6d4"}))
        chart.add_trace(go.Scatter(x=[None], y=[None], mode="markers", name="價值區範圍（紅底）",
                                   marker={"symbol": "square", "size": 11, "color": "rgba(239, 68, 68, 0.55)"}, hoverinfo="skip"))
        chart.add_trace(go.Scatter(x=[None], y=[None], mode="markers", name="升溫區範圍（綠底）",
                                   marker={"symbol": "square", "size": 11, "color": "rgba(34, 197, 94, 0.55)"}, hoverinfo="skip"))
        trading_days = pd.DatetimeIndex(chart_df.index).normalize().unique()
        calendar_days = pd.date_range(trading_days.min(), trading_days.max(), freq="D")
        missing_weekdays = calendar_days[(calendar_days.dayofweek < 5) & (~calendar_days.isin(trading_days))]
        default_start = max(chart_df.index[0], chart_df.index[-1] - pd.DateOffset(years=1))
        chart.update_layout(
            height=720,
            showlegend=False,
            dragmode="zoom",
            hovermode="closest",
            margin={"l": 72, "r": 16, "t": 28, "b": 18},
        )
        chart.update_xaxes(
            type="date",
            range=[default_start.strftime("%Y-%m-%d"), (chart_df.index[-1] + pd.Timedelta(days=1)).strftime("%Y-%m-%d")],
            rangebreaks=[
                {"bounds": ["sat", "mon"]},
                {"values": missing_weekdays.strftime("%Y-%m-%d").tolist()},
            ],
            rangeslider={"visible": True, "thickness": 0.08},
            rangeselector={
                "buttons": [
                    {"count": 14, "label": "近2週", "step": "day", "stepmode": "backward"},
                    {"count": 30, "label": "近1月", "step": "day", "stepmode": "backward"},
                    {"count": 90, "label": "近3月", "step": "day", "stepmode": "backward"},
                    {"count": 6, "label": "6個月", "step": "month", "stepmode": "backward"},
                    {"count": 1, "label": "1年", "step": "year", "stepmode": "backward"},
                    {"count": 2, "label": "2年", "step": "year", "stepmode": "backward"},
                    {"step": "all", "label": "全部"},
                ],
                "x": 0,
                "xanchor": "left",
                "y": 1.12,
                "yanchor": "top",
            },
        )
        chart.update_yaxes(
            autorange=True,
            fixedrange=False,
            automargin=True,
            tickformat=",.0f" if chart_df.Close.iloc[-1] >= 500 else ",.1f" if chart_df.Close.iloc[-1] >= 50 else ",.2f",
            showspikes=True,
            spikemode="across",
            spikesnap="cursor",
            spikethickness=1,
            spikecolor="#64748b",
        )
        chart.update_xaxes(
            showspikes=True,
            spikemode="across",
            spikesnap="data",
            spikethickness=1,
            spikecolor="#64748b",
        )
        y_wheel_script = """
const gd = document.getElementById('{plot_id}');
const readoutLine1 = document.getElementById('stock-chart-readout-line1');
const readoutLine2 = document.getElementById('stock-chart-readout-line2');
const priceTag = document.getElementById('stock-y-cursor-price');
const fmt = value => (value === undefined || value === null || !Number.isFinite(Number(value)))
  ? '--' : Number(value).toLocaleString('zh-TW', {minimumFractionDigits: 2, maximumFractionDigits: 2});
const fmtPrice = value => {
  if (value === undefined || value === null || !Number.isFinite(Number(value))) return '--';
  const n = Number(value);
  const tick = n >= 1000 ? 5 : n >= 500 ? 1 : n >= 100 ? 0.5 : n >= 50 ? 0.1 : n >= 10 ? 0.05 : 0.01;
  const digits = tick >= 1 ? 0 : tick >= 0.1 ? 1 : 2;
  const tradablePrice = Math.round(n / tick) * tick;
  return tradablePrice.toLocaleString('zh-TW', {minimumFractionDigits: digits, maximumFractionDigits: digits});
};
const rangeControls = document.getElementById('stock-chart-mobile-controls');
const dateText = date => date.toISOString().slice(0, 10);
if (rangeControls) {
  rangeControls.addEventListener('click', function(event) {
    const button = event.target.closest('button[data-range]');
    if (!button) return;
    const candle = gd.data.find(trace => trace.name === 'K線');
    if (!candle || !candle.x || !candle.x.length) return;
    if (button.dataset.range === 'all') {
      Plotly.relayout(gd, {'xaxis.autorange': true});
      return;
    }
    const lastDate = new Date(candle.x[candle.x.length - 1]);
    const endDate = new Date(lastDate);
    endDate.setDate(endDate.getDate() + 1);
    const startDate = new Date(lastDate);
    const periods = {'2w': ['day', 14], '1m': ['day', 30], '3m': ['day', 90], '6m': ['month', 6], '1y': ['year', 1], '2y': ['year', 2]};
    const period = periods[button.dataset.range];
    if (!period) return;
    if (period[0] === 'day') startDate.setDate(startDate.getDate() - period[1]);
    else if (period[0] === 'month') startDate.setMonth(startDate.getMonth() - period[1]);
    else startDate.setFullYear(startDate.getFullYear() - period[1]);
    Plotly.relayout(gd, {'xaxis.autorange': false, 'xaxis.range': [dateText(startDate), dateText(endDate)]});
  });
}
let resizeTimer;
const fitChartToViewport = function() {
  if (!gd || !gd._fullLayout) return;
  const mobile = window.matchMedia('(max-width: 720px)').matches;
  Plotly.relayout(gd, {
    height: mobile ? 560 : 720,
    'margin.l': mobile ? 54 : 72,
    'margin.r': mobile ? 10 : 16,
    'margin.t': mobile ? 12 : 28,
    'margin.b': mobile ? 30 : 18,
    'xaxis.rangeslider.visible': !mobile
  });
};
window.addEventListener('resize', function() {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(fitChartToViewport, 180);
});
setTimeout(fitChartToViewport, 0);
gd.on('plotly_hover', function(eventData) {
  if (!readoutLine1 || !readoutLine2 || !eventData || !eventData.points || !eventData.points.length) return;
  const points = eventData.points;
  const hovered = points[0];
  const i = hovered.pointNumber;
  const candle = gd.data.find(trace => trace.name === 'K線');
  if (!candle || i === undefined) return;
  const values = hovered.customdata || candle.customdata[i];
  const val = column => values && values[column];
  const date = String(hovered.x || candle.x[i] || '').slice(0, 10);
  readoutLine1.textContent = `${date}　開 ${fmtPrice(val(0))}　收 ${fmtPrice(val(3))}　高 ${fmtPrice(val(1))}　低 ${fmtPrice(val(2))}`;
  readoutLine2.innerHTML = `<span style="color:#334155">均線：</span><span style="color:#f59e0b">MA5 ${fmt(val(4))}</span>　<span style="color:#2563eb">MA20 ${fmt(val(5))}</span>　<span style="color:#8b5cf6">MA60 ${fmt(val(6))}</span>　<span style="color:#334155">布林通道：</span><span style="color:#475569">上軌 ${fmt(val(7))}</span>　<span style="color:#06b6d4">下軌 ${fmt(val(8))}</span>`;
});
gd.addEventListener('mousemove', function(event) {
  const size = gd._fullLayout && gd._fullLayout._size;
  const wrap = document.getElementById('stock-chart-wrap');
  if (!size || !wrap || !priceTag) return;
  const rect = gd.getBoundingClientRect();
  const y = event.clientY - rect.top;
  if (y < size.t || y > size.t + size.h) { priceTag.style.display = 'none'; return; }
  const range = gd._fullLayout.yaxis.range.map(Number);
  const value = range[0] + ((size.t + size.h - y) / size.h) * (range[1] - range[0]);
  const wrapRect = wrap.getBoundingClientRect();
  priceTag.textContent = fmtPrice(value);
  priceTag.style.left = `${rect.left - wrapRect.left + size.l - priceTag.offsetWidth - 4}px`;
  priceTag.style.top = `${rect.top - wrapRect.top + y - priceTag.offsetHeight / 2}px`;
  priceTag.style.display = 'block';
});
gd.addEventListener('mouseleave', function() { if (priceTag) priceTag.style.display = 'none'; });
gd.addEventListener('wheel', function(event) {
  const size = gd._fullLayout && gd._fullLayout._size;
  if (!size) return;
  const rect = gd.getBoundingClientRect();
  const x = event.clientX - rect.left;
  const y = event.clientY - rect.top;
  if (x < size.l || x > size.l + size.w || y < size.t || y > size.t + size.h) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  const range = gd._fullLayout.yaxis.range;
  const center = (Number(range[0]) + Number(range[1])) / 2;
  const oldSpan = Math.abs(Number(range[1]) - Number(range[0]));
  const factor = event.deltaY < 0 ? 0.85 : 1 / 0.85;
  const minSpan = Math.max(Math.abs(center) * 0.0001, 1e-8);
  const newSpan = Math.max(oldSpan * factor, minSpan);
  Plotly.relayout(gd, {'yaxis.range': [center - newSpan / 2, center + newSpan / 2]});
}, {passive: false, capture: true});
"""
        chart_html = chart.to_html(
            full_html=False,
            include_plotlyjs="cdn",
            config={"scrollZoom": False, "displayModeBar": True, "responsive": True},
            post_script=y_wheel_script,
        )
        chart_panel = """
<style>
  html, body { margin: 0; padding: 0; overflow: hidden; font-family: Arial, sans-serif; }
  #stock-chart-mobile-controls { display: none; }
  #stock-chart-mobile-controls button {
    min-height: 34px; padding: 5px 8px; border: 1px solid #d9e1ec; border-radius: 7px;
    color: #334155; background: #f8fafc; font: 600 12px Arial, sans-serif;
  }
  #stock-chart-legend {
    display: flex; flex-wrap: wrap; align-items: center; gap: 6px 14px;
    padding: 5px 12px; color: #475569; font: 12px Arial, sans-serif;
  }
  #stock-chart-legend .legend-item { display: inline-flex; align-items: center; gap: 5px; white-space: nowrap; }
  #stock-chart-legend .legend-line { display: inline-block; width: 18px; border-top: 2px solid; }
  #stock-chart-legend .legend-dash { border-top-style: dashed; }
  #stock-chart-legend .legend-candle { display: inline-block; width: 14px; height: 10px; background: linear-gradient(135deg,#e53935 0 50%,#159447 50%); }
  #stock-chart-legend .legend-zone { display: inline-block; width: 11px; height: 11px; }
  #stock-chart-readout {
    box-sizing: border-box; min-height: 60px; padding: 7px 14px; margin: 0 8px;
    overflow: hidden; white-space: normal;
    color: #1f2937; background: #f1f5f9; border-left: 4px solid #2563eb;
    border-radius: 4px; font-size: 13px; font-weight: 600; line-height: 22px;
  }
  #stock-chart-wrap { position: relative; width: 100%; max-width: 1150px; margin: 0 auto; }
  #stock-chart-wrap .plotly-graph-div { width: 100% !important; }
  #stock-y-cursor-price {
    display: none; position: absolute; z-index: 30; min-width: 72px; box-sizing: border-box;
    padding: 3px 5px; color: white; background: #334155; border: 1px solid white;
    border-radius: 2px; text-align: right; font: 600 12px Arial, sans-serif;
    pointer-events: none; white-space: nowrap;
  }
  .hoverlayer .hovertext { display: none !important; }
  @media (max-width: 720px) {
    #stock-chart-mobile-controls { display: flex; flex-wrap: wrap; gap: 5px; padding: 2px 8px 8px; }
    #stock-chart-mobile-controls button { flex: 1 0 auto; }
    #stock-chart-readout { min-height: 0; padding: 7px 10px; margin: 0 4px; font-size: 11px; line-height: 1.55; }
    #stock-chart-readout > div { overflow-wrap: anywhere; }
    #stock-chart-legend { gap: 5px 10px; padding: 5px 8px; font-size: 10px; }
    #stock-chart-wrap { max-width: none; }
    #stock-chart-wrap .modebar-container, #stock-chart-wrap .rangeselector-container,
    #stock-chart-wrap g.rangeselector { display: none !important; }
  }
</style>
<div id="stock-chart-mobile-controls" aria-label="手機圖表期間選擇">
  <button data-range="2w">近2週</button><button data-range="1m">近1月</button>
  <button data-range="3m">近3月</button><button data-range="6m">6個月</button>
  <button data-range="1y">1年</button><button data-range="2y">2年</button>
  <button data-range="all">全部</button>
</div>
<div id="stock-chart-readout">
  <div id="stock-chart-readout-line1">日期　開 --　收 --　高 --　低 --</div>
  <div id="stock-chart-readout-line2">均線：MA5 --　MA20 --　MA60 --　布林通道：下軌 --　上軌 --</div>
</div>
<div id="stock-chart-legend">
  <span class="legend-item"><i class="legend-candle"></i>K線</span>
  <span class="legend-item"><i class="legend-line" style="border-color:#f59e0b"></i>MA5</span>
  <span class="legend-item"><i class="legend-line" style="border-color:#2563eb"></i>MA20</span>
  <span class="legend-item"><i class="legend-line" style="border-color:#8b5cf6"></i>MA60</span>
  <span class="legend-item"><i class="legend-line legend-dash" style="border-color:#475569"></i>布林上軌</span>
  <span class="legend-item"><i class="legend-line legend-dash" style="border-color:#06b6d4"></i>布林下軌</span>
  <span class="legend-item"><i class="legend-zone" style="background:rgba(239,68,68,.55)"></i>價值區</span>
  <span class="legend-item"><i class="legend-zone" style="background:rgba(34,197,94,.55)"></i>升溫區</span>
</div>
<div id="stock-chart-wrap"><div id="stock-y-cursor-price"></div>""" + chart_html + """</div>
"""
        components.html(chart_panel, height=920, scrolling=False)
        st.markdown(f"**快速摘要：** RSI {last.RSI14:.1f}；MACD {'在訊號線上方' if last.MACD > last.MACDSignal else '在訊號線下方'}；K {'高於' if last.K > last.D else '低於'} D；成交量為20日均量 {last.Volume / last.VolMA20:.2f} 倍。")

    with indicators_tab:
        section_heading_with_help(
            "各項指標分數與解讀",
            "單項分數以「實得分／該項滿分」呈現；上方偏多與偏弱總分會換算成 100 分制，與回測門檻一致。偏多分數低不等於應賣出；偏弱分數高表示需要留意風險。",
        )
        st.dataframe(breakdown, hide_index=True, width="stretch")
        quick = pd.DataFrame([
            {"指標": "RSI(14)", "目前值": f"{last.RSI14:.1f}", "白話說明": "衡量近期價格強弱；高檔可能偏熱，低檔不保證反彈。"},
            {"指標": "KD", "目前值": f"K {last.K:.1f}／D {last.D:.1f}", "白話說明": "K>D 偏強、K<D 偏弱，常用來觀察短線動能。"},
            {"指標": "MACD", "目前值": f"{last.MACD:.2f}／訊號 {last.MACDSignal:.2f}", "白話說明": "比較中短期動能方向，MACD 在訊號線上方通常偏強。"},
            {"指標": "布林通道", "目前值": f"上 {last.BB_Upper:.2f}／中 {last.SMA20:.2f}／下 {last.BB_Lower:.2f}", "白話說明": "觀察價格在近期波動區間的位置。"},
            {"指標": "OBV", "目前值": "高於20日均線" if last.OBV > last.OBV_MA20 else "低於20日均線", "白話說明": "用成交量配合漲跌方向，觀察量能累積變化。"},
            {"指標": "ADX / +DI / -DI", "目前值": f"{last.ADX14:.1f}／{last.PlusDI:.1f}／{last.MinusDI:.1f}", "白話說明": "ADX看趨勢強度，+DI/-DI協助判斷趨勢方向。"},
            {"指標": "ATR(14)", "目前值": f"{last.ATR14:.2f}", "白話說明": "估計近期波動幅度，不判斷漲跌方向。"},
            {"指標": "20日乖離", "目前值": f"{last.Bias20:+.1f}%", "白話說明": "顯示股價距離20日均線的百分比。"},
        ])
        st.dataframe(quick, hide_index=True, width="stretch")

    with scanner_tab:
        st.subheader("多股指標掃描")
        watchlist_owner = str(st.user.get("sub", "")).strip()
        if not watchlist_owner:
            st.error("Google 登入沒有提供唯一使用者識別碼，無法安全載入個人清單。")
            st.stop()
        if st.session_state.get("watchlist_loaded_for") != watchlist_owner:
            try:
                personal_data = load_personal_watchlist(watchlist_owner)
                st.session_state["watchlist_symbols"] = personal_data["symbols"]
                st.session_state["watchlist_signal_state"] = personal_data["signal_state"]
                st.session_state["watchlist_signal_events"] = personal_data["signal_events"]
                st.session_state["watchlist_loaded_for"] = watchlist_owner
            except Exception as exc:
                st.error(f"無法載入你的自選股資料庫資料：{exc}")
                st.stop()

        with st.expander("⭐ 我的自選股與訊號變化", expanded=True):
            st.caption("已使用 Google 帳戶同步自選清單與訊號變化紀錄；每位使用者只會讀寫自己的資料。")
            add_col, add_button_col = st.columns([4, 1])
            with add_col:
                watchlist_input = st.text_input("新增代號", placeholder="例如：AAPL 或 2330.TW", key="watchlist_add_input")
            with add_button_col:
                st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
                if st.button("加入", key="watchlist_add_button", width="stretch"):
                    symbol_to_add = watchlist_input.strip().upper()
                    if symbol_to_add and symbol_to_add not in st.session_state["watchlist_symbols"]:
                        st.session_state["watchlist_symbols"].append(symbol_to_add)
                        try:
                            save_personal_watchlist(watchlist_owner)
                            st.rerun()
                        except Exception as exc:
                            st.error(f"自選股已暫存在此工作階段，但同步資料庫失敗：{exc}")
                    elif symbol_to_add:
                        st.info("這檔股票已在自選清單中。")

            current_watchlist = st.session_state["watchlist_symbols"]
            remove_symbols = st.multiselect("目前自選股", current_watchlist, key="watchlist_remove_selection")
            if st.button("移除選取", key="watchlist_remove_button", disabled=not remove_symbols):
                st.session_state["watchlist_symbols"] = [s for s in current_watchlist if s not in remove_symbols]
                try:
                    save_personal_watchlist(watchlist_owner)
                    st.rerun()
                except Exception as exc:
                    st.error(f"移除已套用在此工作階段，但同步資料庫失敗：{exc}")

            if st.session_state["watchlist_signal_events"]:
                st.markdown("**訊號變化紀錄**")
                event_frame = pd.DataFrame(st.session_state["watchlist_signal_events"])
                st.dataframe(event_frame, hide_index=True, width="stretch")
                st.download_button(
                    "下載訊號變化紀錄 CSV",
                    event_frame.to_csv(index=False).encode("utf-8-sig"),
                    file_name="stock_signal_changes.csv", mime="text/csv", key="download_signal_events",
                )
            else:
                st.caption("尚無訊號紀錄。每次掃描自選股時，系統會記錄狀態相較上次掃描的變化。")

        if st.button("將自選股帶入掃描清單", key="use_watchlist_for_scan", disabled=not st.session_state["watchlist_symbols"]):
            st.session_state["scan_symbols_text"] = ", ".join(st.session_state["watchlist_symbols"])
        scan_text = st.text_area("輸入股票代碼（逗號分隔）", "2330.TW, 2454.TW, 0050.TW", height=80, key="scan_symbols_text")
        sort_mode = st.radio(
            "掃描用途",
            ["找買點", "看賣出風險"],
            horizontal=True,
            help="買點模式將每檔當作尚未持有來判斷；賣出風險模式假設你已持有每檔，檢查是否該減碼或賣出。",
        )
        if sort_mode == "找買點":
            st.caption("新增「操作時點建議」欄：假設目前未持有，依個股區域與大盤環境逐檔判斷是否列為買進候選。結果依價值區及歷史統計排序。")
        else:
            st.caption("新增「操作時點建議」欄：假設目前持有每檔，依轉弱階段與偏弱分數判斷續抱、減碼或賣出；程式不知道你的實際持倉。")
        st.caption("歷史統計使用近 5 年資料、訊號每20個交易日去重，並扣除估計交易成本。成交金額為近20日平均估值。目前只含技術資料。")
        if st.button("開始掃描", key="run_scan", type="primary"):
            scan_rows = []
            tickers = list(dict.fromkeys(x.strip().upper() for x in scan_text.split(",") if x.strip()))
            bar = st.progress(0)
            for i, symbol in enumerate(tickers):
                try:
                    hist = get_history(symbol, history_period)
                    if not hist.empty and len(hist) >= 120:
                        scored = add_indicators(hist).dropna(subset=["SMA120", "RSI14", "K", "D", "VolMA20", "ADX14"])
                        buy, sell, label, _ = score_stock(scored)
                        zone, stage, reasons, _ = classify_zone(scored, buy, sell)
                        scan_market, _ = market_score(symbol)
                        assume_holding = sort_mode == "看賣出風險"
                        timing_label, _, _, _ = current_timing_action(
                            zone, stage, buy, sell, assume_holding,
                            scan_market, buy_threshold, sell_threshold,
                        )
                        operation_reason = concise_timing_reason(
                            timing_label, zone, stage, buy, sell, assume_holding,
                            scan_market, buy_threshold, sell_threshold,
                        )
                        latest = scored.iloc[-1]
                        history_stats = evaluate_zone_outcomes(
                            scored, ZONE_STATS_COMMISSION_RATE, ZONE_STATS_SELL_TAX_RATE, ZONE_STATS_SLIPPAGE_RATE
                        )
                        value20 = history_stats.loc[
                            (history_stats["訊號區域／狀態"] == "價值區")
                            & (history_stats["觀察期間(交易日)"] == 20)
                        ]
                        weak20 = history_stats.loc[
                            (history_stats["訊號區域／狀態"] == "升溫區｜轉弱確認")
                            & (history_stats["觀察期間(交易日)"] == 20)
                        ]
                        value_row = value20.iloc[0] if not value20.empty else None
                        weak_row = weak20.iloc[0] if not weak20.empty else None
                        scan_name = get_symbol_name(symbol, stock_directory)
                        avg_turnover_m = float((scored.Close * scored.Volume).tail(20).mean() / 1_000_000)
                        scan_rows.append({
                            "代號": symbol, "股票名稱": scan_name,
                            "收盤": round(float(latest.Close), 2),
                            "短線區域": zone, "區域階段": stage,
                            "偏多分數": buy, "偏弱風險": sell,
                            "價值區20日樣本": int(value_row["去重後訊號數"]) if value_row is not None else 0,
                            "價值區20日正報酬率(%)": float(value_row["扣估計成本後正報酬率"]) if value_row is not None else np.nan,
                            "價值區20日平均淨報酬(%)": float(value_row["平均淨報酬"]) if value_row is not None else np.nan,
                            "轉弱區20日樣本": int(weak_row["去重後訊號數"]) if weak_row is not None else 0,
                            "轉弱後平均最大跌幅(%)": float(weak_row["20日內平均最大盤中跌幅"]) if weak_row is not None else np.nan,
                            "20日均成交額(百萬元)": round(avg_turnover_m, 1),
                            "判讀": label, "操作時點建議": timing_label,
                            "操作建議理由": operation_reason,
                            "大盤環境分數": scan_market["score"] if scan_market else np.nan,
                            "資料日期": pd.Timestamp(scored.index[-1]).strftime("%Y-%m-%d"),
                        })
                    else:
                        scan_rows.append({"代號": symbol, "錯誤": "行情資料不足，無法計算短線操作建議"})
                except Exception as exc:
                    scan_rows.append({"代號": symbol, "錯誤": str(exc)})
                bar.progress((i + 1) / max(len(tickers), 1))
            if scan_rows:
                results = pd.DataFrame(scan_rows)
                previous_signal_state = st.session_state["watchlist_signal_state"]
                new_signal_events = []
                scanned_at = pd.Timestamp.now(tz="Asia/Taipei").strftime("%Y-%m-%d %H:%M")
                watchlist_set = set(st.session_state["watchlist_symbols"])
                for _, signal_row in results.iterrows():
                    symbol = str(signal_row.get("代號", "")).upper()
                    if symbol not in watchlist_set or pd.isna(signal_row.get("短線區域")):
                        continue
                    current_state = {
                        "區域": str(signal_row.get("短線區域", "")),
                        "階段": str(signal_row.get("區域階段", "")),
                        "操作判讀": str(signal_row.get("操作時點建議", "")),
                        "掃描用途": sort_mode,
                        "偏多分數": int(signal_row.get("偏多分數", 0)),
                        "偏弱風險": int(signal_row.get("偏弱風險", 0)),
                        "資料日期": str(signal_row.get("資料日期", "")),
                    }
                    old_state = previous_signal_state.get(symbol)
                    if old_state:
                        compared_fields = ["區域", "階段"]
                        if old_state.get("掃描用途") == sort_mode:
                            compared_fields.append("操作判讀")
                        changed = [key for key in compared_fields if old_state.get(key) != current_state[key]]
                        if changed:
                            new_signal_events.append({
                                "掃描時間": scanned_at,
                                "代號": symbol,
                                "股票名稱": str(signal_row.get("股票名稱", symbol)),
                                "資料日期": current_state["資料日期"],
                                "變化項目": "、".join(changed),
                                "前次狀態": f"{old_state.get('區域', '')}｜{old_state.get('階段', '')}｜{old_state.get('操作判讀', '')}",
                                "目前狀態": f"{current_state['區域']}｜{current_state['階段']}｜{current_state['操作判讀']}",
                                "偏多／偏弱": f"{current_state['偏多分數']}／{current_state['偏弱風險']}",
                            })
                    previous_signal_state[symbol] = current_state
                st.session_state["watchlist_signal_state"] = previous_signal_state
                if new_signal_events:
                    st.session_state["watchlist_signal_events"] = (
                        new_signal_events + st.session_state["watchlist_signal_events"]
                    )[:300]
                    st.markdown("**本次掃描發現的訊號變化**")
                    st.dataframe(pd.DataFrame(new_signal_events), hide_index=True, width="stretch")
                else:
                    st.caption("本次沒有發現區域或階段變化；首次掃描會建立比較基準。若切換掃描用途，操作判讀會從該用途重新建立基準。")
                try:
                    save_personal_watchlist(watchlist_owner)
                except Exception as exc:
                    st.warning(f"本次訊號已在畫面更新，但同步資料庫失敗：{exc}")
                if "短線區域" in results.columns:
                    if sort_mode == "找買點":
                        zone_order = {"價值區": 0, "空檔": 1, "升溫區": 2}
                        results["_區域排序"] = results["短線區域"].map(zone_order).fillna(9)
                        results["_樣本充足"] = results["價值區20日樣本"] >= 20
                        results["_可信正報酬率"] = results["價值區20日正報酬率(%)"].where(results["_樣本充足"], -1)
                        results["_可信平均報酬"] = results["價值區20日平均淨報酬(%)"].where(results["_樣本充足"], -9999)
                        results = results.sort_values(
                            ["_區域排序", "_樣本充足", "_可信正報酬率",
                             "_可信平均報酬", "20日均成交額(百萬元)", "偏多分數"],
                            ascending=[True, False, False, False, False, False], na_position="last",
                        )
                    else:
                        confirmed = (results["短線區域"] == "升溫區") & (results["區域階段"] == "轉弱確認")
                        observation = (results["短線區域"] == "升溫區") & (results["區域階段"] == "偏熱觀察")
                        results["_區域排序"] = np.select([confirmed, observation], [0, 1], default=2)
                        results["_樣本充足"] = results["轉弱區20日樣本"] >= 20
                        results["_可信風險分數"] = results["偏弱風險"].where(results["_樣本充足"], -1)
                        results["_可信平均跌幅"] = results["轉弱後平均最大跌幅(%)"].where(results["_樣本充足"], 0)
                        results = results.sort_values(
                            ["_區域排序", "_樣本充足", "_可信風險分數",
                             "_可信平均跌幅", "20日均成交額(百萬元)"],
                            ascending=[True, False, False, True, False], na_position="last",
                        )
                    results = results.drop(columns=[c for c in results.columns if c.startswith("_")], errors="ignore")
                    preferred_columns = [
                        "代號", "股票名稱", "收盤", "短線區域", "區域階段", "判讀", "操作時點建議",
                        "操作建議理由", "大盤環境分數",
                        "偏多分數", "偏弱風險", "價值區20日樣本", "價值區20日正報酬率(%)",
                        "價值區20日平均淨報酬(%)", "轉弱區20日樣本", "轉弱後平均最大跌幅(%)",
                        "20日均成交額(百萬元)", "資料日期",
                    ]
                    ordered_columns = [c for c in preferred_columns if c in results.columns]
                    ordered_columns.extend(c for c in results.columns if c not in ordered_columns)
                    results = results[ordered_columns]
                st.dataframe(results, hide_index=True, width="stretch")
                st.caption("歷史正報酬率與平均跌幅是過去條件的結果，不是未來預測；樣本不足20筆時，排名不採信該歷史比例。")
            else:
                st.warning("沒有取得可用資料。請確認代碼格式或稍後再試。")

    with backtest_tab:
        st.subheader("買賣分數策略模擬")
        st.markdown("訊號在每日收盤後計算，成交模擬在**下一交易日開盤**，並計入手續費、賣出交易稅與滑價。")
        with st.expander("回測設定", expanded=True):
            st.caption("以下參數只用於本頁歷史模擬，不會改變看盤圖表或多股掃描。買進與賣出分數門檻可在左側「進階設定：買賣訊號門檻」調整。")
            years_available = list(range(2010, pd.Timestamp.now().year + 1))
            default_start = max(2010, pd.Timestamp.now().year - 5)
            year_left, year_right = st.columns(2)
            start_year = year_left.selectbox("回測起始年", years_available, index=years_available.index(default_start))
            end_year = year_right.selectbox("回測結束年", years_available, index=len(years_available)-1)
            strategy_mode = st.selectbox(
                "回測訊號策略",
                ["多指標分數策略", "均線趨勢實驗：站上50日線與200日線進場／跌破200日線出場"],
                help="只切換歷史回測規則；首頁目前時點判讀固定使用多指標區域與大盤濾網。均線趨勢候選先要求股價站上50日與200日線。",
            )
            setting_left, setting_right = st.columns(2)
            position_pct = setting_left.slider("單筆最高投入資金比例", 10, 80, 30, 5, help="回測單筆配置上限；實際投入還會依市場水位與個股分數下修。") / 100
            risk_pct = setting_right.slider("每筆預估最大虧損（本金比例）", 0.1, 2.0, 0.5, 0.1,
                help="預設 0.5% 是可調整的保守起始值。系統以停損距離估算股數；跳空或流動性不足時，實際虧損可能超出估算。")
            setting_left, setting_right = st.columns(2)
            atr_stop_multiple = setting_left.slider("初始停損距離（ATR 倍數）", 0.5, 4.0, 2.0, 0.25,
                help="停損價以隔日模擬買進價減去訊號日 ATR(14) × 此倍數估算；這是回測參數，不保證能按停損價成交。")
            max_holding_days = setting_right.slider("最長持有期間（交易日）", 5, 20, 20, 1,
                help="達到設定天數後，回測於下一個交易日開盤模擬出場。實際訊號或停損可能更早出場。")
            initial_capital = st.number_input("回測本金", min_value=10000, value=1000000, step=100000)
            fee_left, fee_right, fee_third = st.columns(3)
            commission_pct = fee_left.number_input("單邊手續費率 (%)", min_value=0.0, max_value=1.0, value=0.1425, step=0.0025)
            tax_pct = fee_right.number_input("賣出交易稅率 (%)", min_value=0.0, max_value=1.0, value=0.3, step=0.05,
                help="台灣一般股票預設 0.3%；ETF 等商品可能不同，請依商品與券商設定調整。")
            slippage_pct = fee_third.number_input("每次成交滑價 (%)", min_value=0.0, max_value=1.0, value=0.05, step=0.05)
            if start_year > end_year:
                st.warning("回測起始年需早於或等於結束年。")
        if strategy_mode.startswith("均線趨勢"):
            st.caption(f"策略邏輯：收盤同時站上50日與200日均線才形成買進訊號；持倉後收盤跌破200日均線形成賣出訊號。均線模式使用固定60分觸發門檻；單筆投入上限 {position_pct:.0%}，另依大盤水位下修。")
        else:
            st.caption(f"策略邏輯：偏多分數向上穿越 {buy_threshold} 進場；持倉時偏弱分數達 {sell_threshold} 出場。單筆投入上限為本金 {position_pct:.0%}，實際比例依當日市場分數與個股分數調整。")
        st.caption(f"部位與退出：每筆預估風險上限 {risk_pct:.1f}%（本金 {float(initial_capital) * risk_pct / 100:,.0f} 元）；停損距離為 ATR(14) × {atr_stop_multiple:.2f}；最多持有 {max_holding_days} 個交易日。實際投入金額同時受風險上限、單筆投入上限及市場配置限制。")
        if st.button("執行歷史模擬", key="run_backtest", type="primary"):
            if start_year > end_year:
                st.error("回測起始年不可晚於結束年。")
                st.stop()
            bt_start = f"{start_year}-01-01"
            bt_end = min(f"{end_year}-12-31", pd.Timestamp.now().strftime("%Y-%m-%d"))
            with st.spinner("下載並計算所選年份的股票與大盤資料…"):
                bt_raw = get_history_range(ticker.strip().upper(), bt_start, bt_end)
                if bt_raw.empty or len(bt_raw) < 130:
                    st.error("這段期間可用資料不足，請確認代碼或選擇較長期間。")
                    st.stop()
                prepared = add_score_columns(add_indicators(bt_raw))
                if strategy_mode.startswith("均線趨勢"):
                    prepared["BuyScore"] = ((prepared.Close > prepared.SMA50) & (prepared.Close > prepared.SMA200)).astype(float) * 100
                    prepared["SellScore"] = (prepared.Close < prepared.SMA200).astype(float) * 100
                    prepared["BuyReason"] = prepared.apply(lambda r: f"收盤 {r.Close:.2f} 同時站上50日均線 {r.SMA50:.2f} 與200日均線 {r.SMA200:.2f}，長期趨勢轉強", axis=1)
                    prepared["SellReason"] = prepared.apply(lambda r: f"收盤 {r.Close:.2f} 跌破200日均線 {r.SMA200:.2f}，長期趨勢轉弱", axis=1)
                scored_all = prepared.copy()
                scored_all = add_market_columns(scored_all, ticker.strip().upper(), bt_start, bt_end)
                scored_all = scored_all.loc[(scored_all.index >= pd.Timestamp(bt_start)) & (scored_all.index <= pd.Timestamp(bt_end))]
                scored_all = scored_all.dropna(subset=["SMA120", "BuyScore", "SellScore"])
            if len(scored_all) < 30:
                st.error("所選年份有效交易資料不足，無法可靠地計算回測。")
                st.stop()
            warmup_rows = prepared.loc[prepared.index < scored_all.index[0], "BuyScore"]
            warmup_score = float(warmup_rows.iloc[-1]) if not warmup_rows.empty else 0.0
            result = run_backtest(scored_all, buy_threshold, sell_threshold, position_pct, float(initial_capital),
                                  commission_pct / 100, tax_pct / 100, slippage_pct / 100, prior_buy_score=warmup_score,
                                  risk_fraction=risk_pct / 100, atr_stop_multiple=atr_stop_multiple,
                                  max_holding_days=max_holding_days)
            split = int(len(scored_all) * 0.7)
            oos = None
            if split < len(scored_all) - 20:
                oos = run_backtest(scored_all.iloc[split:], buy_threshold, sell_threshold, position_pct,
                                   float(initial_capital), commission_pct / 100, tax_pct / 100,
                                   slippage_pct / 100, prior_buy_score=float(scored_all.BuyScore.iloc[split - 1]),
                                   risk_fraction=risk_pct / 100, atr_stop_multiple=atr_stop_multiple,
                                   max_holding_days=max_holding_days)
            if result:
                st.markdown(f"**所選區間：{pd.Timestamp(result['start']).strftime('%Y-%m-%d')} 至 {pd.Timestamp(result['end']).strftime('%Y-%m-%d')}**")
                a, b, c, d = st.columns(4)
                a.metric("策略總報酬（扣估計成本）", f"{result['return']:+.2f}%", f"年化 {result['cagr']:+.2f}%")
                b.metric("買進持有", f"{result['buy_hold']:+.2f}%", f"策略差 {result['return'] - result['buy_hold']:+.2f}%")
                c.metric("策略最大回撤", f"{result['mdd']:.2f}%", f"買進持有 {result['benchmark_mdd']:.2f}%")
                d.metric("交易數／勝率", f"{result['count']} 筆", f"勝率 {result['win_rate']:.1f}%")
                st.caption(f"另以同一單筆投入上限 {position_pct:.0%} 持有不賣：{result['capped_buy_hold']:+.2f}%（策略差 {result['return'] - result['capped_buy_hold']:+.2f}%）。這可協助分辨策略是否只是因現金部位不同而落後全額持有。")
                st.caption(f"市場基準 {result['market_proxy_name']}（含調整後價格、未扣交易費用）：{result['market_return']:+.2f}%；逐月勝過個股持有 {result['monthly_win_hold']:.1f}%、勝過市場基準 {result['monthly_win_market']:.1f}%（共 {result['monthly_count']} 個可比較月份）。逐月勝出比例不同於總報酬，也不等於交易勝率。")
                if oos:
                    st.markdown("#### 後段 30% 區間（簡單樣本外檢視）")
                    st.caption(f"區間 {pd.Timestamp(oos['start']).strftime('%Y-%m-%d')} 至 {pd.Timestamp(oos['end']).strftime('%Y-%m-%d')}；此區間獨立以現金開始，不沿用前段持倉。")
                    x, y, z, w = st.columns(4)
                    x.metric("策略報酬", f"{oos['return']:+.2f}%")
                    y.metric("買進持有", f"{oos['buy_hold']:+.2f}%")
                    z.metric("策略最大回撤", f"{oos['mdd']:.2f}%", f"買進持有 {oos['benchmark_mdd']:.2f}%")
                    w.metric("交易數／勝率", f"{oos['count']} 筆", f"勝率 {oos['win_rate']:.1f}%")
                    st.caption(f"後段以相同單筆投入上限 {position_pct:.0%} 持有不賣：{oos['capped_buy_hold']:+.2f}%（策略差 {oos['return'] - oos['capped_buy_hold']:+.2f}%）。")
                    st.caption(f"後段市場基準 {oos['market_proxy_name']}：{oos['market_return']:+.2f}%；逐月勝過個股持有 {oos['monthly_win_hold']:.1f}%、勝過市場基準 {oos['monthly_win_market']:.1f}%（{oos['monthly_count']} 個月）。")
                    if result["return"] > result["buy_hold"] and oos["return"] <= oos["buy_hold"]:
                        st.warning("初步判讀：完整期勝過個股持有，但後段沒有勝過；目前不能說策略有效，還要檢查大盤 ETF 基準和其他市場期間。")
                    elif result["return"] > result["buy_hold"] and result["return"] > result["market_return"] and oos["return"] > oos["buy_hold"] and oos["return"] > oos["market_return"]:
                        st.info("初步判讀：完整期與後段都勝過個股持有及市場 ETF，初步通過這兩個歷史檢查；仍需跨標的驗證，不能推定未來會重現。")
                    else:
                        st.warning("初步判讀：這組規則尚未同時勝過個股持有與市場 ETF；請勿把較高勝率或較小回撤當作報酬已勝出。")
                curve_fig = go.Figure()
                curve_fig.add_trace(go.Scatter(x=result["equity"].index, y=result["equity"], name="策略資產", line={"width": 2.5, "color": "#1769aa"}))
                curve_fig.add_trace(go.Scatter(x=result["benchmark"].index, y=result["benchmark"], name="買進持有基準", line={"width": 1.5, "dash": "dot", "color": "#888"}))
                curve_fig.add_trace(go.Scatter(x=result["market_benchmark"].index, y=result["market_benchmark"], name=f"市場基準 {result['market_proxy_name']}", line={"width":1.5,"dash":"dash","color":"#d08b24"}))
                curve_fig.update_layout(height=380, title="資產曲線（同本金比較）", hovermode="x unified", margin={"l": 8, "r": 8, "t": 45, "b": 8})
                st.plotly_chart(curve_fig, width="stretch")
                if not result["trades"].empty:
                    price_fig = go.Figure()
                    price_fig.add_trace(go.Scatter(x=scored_all.index, y=scored_all.Close, name="收盤價", line={"color":"#6b7280", "width":1.5}))
                    buys = result["trades"]
                    buy_dates = pd.to_datetime(buys["買進成交日"])
                    sell_dates = pd.to_datetime(buys["賣出成交日"])
                    price_fig.add_trace(go.Scatter(x=buy_dates, y=buys["買進價"], mode="markers", name="模擬買進",
                        marker={"symbol":"triangle-up", "size":11, "color":"#159947"},
                        text=buys["進場理由"], hovertemplate="買進成交 %{x|%Y-%m-%d}<br>價格 %{y:.2f}<br>%{text}<extra></extra>"))
                    price_fig.add_trace(go.Scatter(x=sell_dates, y=buys["賣出價"], mode="markers", name="模擬賣出／停損／期末結算",
                        marker={"symbol":"triangle-down", "size":11, "color":"#d94841"},
                        text=buys["出場理由"], hovertemplate="賣出成交 %{x|%Y-%m-%d}<br>價格 %{y:.2f}<br>%{text}<extra></extra>"))
                    price_fig.update_layout(height=350, title="價格走勢與模擬成交點（滑鼠停留可看理由）", hovermode="closest", xaxis_rangeslider_visible=False,
                        margin={"l":8,"r":8,"t":45,"b":8})
                    st.plotly_chart(price_fig, width="stretch")
                st.markdown(f"平均投入水位約 {result['exposure']:.1f}%；獲利因子 {result['profit_factor']:.2f}。")
                st.caption("比較基準為本金全額買進持有；策略會保留現金，因此兩者的市場曝險不同。策略回撤較小不代表風險調整後報酬已勝出。歷史配置採大盤指數的均線與乖離計算，訊號日收盤資料只用於下一交易日的下單比例。")
                if result["trades"].empty:
                    st.warning("此門檻下沒有完成交易，請檢查分數門檻或延長歷史區間。")
                else:
                    st.markdown("#### 每筆交易時間與原因")
                    st.caption("手機可展開單筆交易查看；表格可下載 CSV，亦可左右捲動查看更多欄位。")
                    for _, trade in result["trades"].iterrows():
                        with st.expander(f"{trade['買進成交日']} 買進 → {trade['賣出成交日']} 賣出／結算｜{trade['報酬率(扣費)']:+.2f}%"):
                            st.markdown(f"**進場訊號日：** {trade['進場訊號日']}　**進場成交日：** {trade['買進成交日']}　**買進價：** {trade['買進價']}　**估算停損價：** {trade['估算停損價']}")
                            st.markdown(f"**每筆風險上限：** {trade['單筆風險上限']:,.0f} 元　**依股數估計停損虧損：** {trade['估計停損虧損']:,.0f} 元（未含跳空超額風險）")
                            st.markdown(f"**進場理由：** {trade['進場理由']}")
                            st.markdown(f"**出場訊號日：** {trade['出場訊號日']}　**出場成交日：** {trade['賣出成交日']}　**賣出價：** {trade['賣出價']}")
                            st.markdown(f"**出場理由：** {trade['出場理由']}")
                            st.markdown(f"**市場分數：** {trade['當時市場分數']}　**偏多分數：** {trade['進場偏多分數']}　**建議投入比例：** {trade['建議投入比例']}　**實際投入：** {trade['實際投入金額']:,.0f} 元　**損益：** {trade['損益(扣費)']:+,.0f} 元")
                    st.dataframe(result["trades"], hide_index=True, width="stretch")
                st.info("投入金額按選定本金計算：市場配置上限 × 個股訊號強度，再受單筆最高比例限制。勝率高或逐月常勝出不代表累積報酬較高；目前規則需同時比較完整期和後段的個股持有與市場 ETF，並在其他標的與不同市場環境重複驗證。")

    with market_tab:
        st.subheader("台美合併股市資金配置參考")
        st.markdown("這個試行版把全部可投資資金視為一個股票部位，台股與美股各占模型權重 50%。總水位是整體股票部位的參考比例；不是台股、美股各自都投入這個比例，也不是單筆下單比例。")
        if combined_market_snapshot:
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("台美綜合環境分數", f"{combined_market_snapshot['score']} / 100")
            m2.metric("總股票配置參考", f"{combined_market_snapshot['alloc']:.1f}%")
            m3.metric("綜合一年區間位置", f"{combined_market_snapshot['position_rank'] * 100:.1f}%")
            m4.metric("綜合20日乖離", f"{combined_market_snapshot['bias']:.1f}%")
            m5.metric("VIX", f"{combined_market_snapshot['vix']:.1f}" if np.isfinite(combined_market_snapshot["vix"]) else "無資料")
            st.caption(f"合併狀態：{combined_market_snapshot['market_state']}。這是依技術資料產生的狀態描述，不代表預測後續漲跌。")

            regional_rows = []
            for region_name, snapshot, weight in [
                ("台股（加權指數）", taiwan_market_snapshot, combined_market_snapshot["taiwan_weight"]),
                ("美股（S&P 500）", us_market_snapshot, combined_market_snapshot["us_weight"]),
            ]:
                regional_rows.append({
                    "市場": region_name,
                    "合併權重": f"{weight:.0%}",
                    "技術狀態": market_state_label(snapshot),
                    "市場分數": f"{snapshot['score']} / 100",
                    "該市場模型水位": f"{snapshot['alloc']:.1f}%",
                    "一年區間位置": f"{snapshot['position_rank'] * 100:.1f}%",
                    "20日乖離": f"{snapshot['bias']:.1f}%",
                    "資料日期": pd.Timestamp(snapshot["date"]).strftime("%Y-%m-%d"),
                })
            st.dataframe(pd.DataFrame(regional_rows), hide_index=True, width="stretch")
            st.caption("表中的「該市場模型水位」是各地區各自的參考值；合併水位依下方公式計算，不是把兩個水位相加。")

            st.markdown("**合併水位計算明細**")
            st.write(
                f"台美技術水位等權平均 **{combined_market_snapshot['base_alloc']:.1f}%**　"
                f"+ CNN 情緒（兩地趨勢條件平均） **{combined_market_snapshot['cnn_adjustment']:+.1f} 點**　"
                f"+ PTT 台股情緒 × 50% **{combined_market_snapshot['ptt_adjustment']:+.1f} 點**　"
                f"= 總股票配置參考 **{combined_market_snapshot['alloc']:.1f}%**"
            )

            fg_col, ptt_col = st.columns(2)
            fear_greed = combined_market_snapshot.get("fear_greed")
            with fg_col:
                st.markdown("**CNN 美股貪婪／恐懼指數**")
                if fear_greed:
                    st.metric("情緒讀數", f"{fear_greed['score']:.1f} / 100 · {fear_greed['rating']}")
                    st.caption(f"資料時間：{fear_greed.get('updated') or '來源未提供'}。CNN 情緒納入合併水位一次，不代表全球散戶調查。")
                else:
                    st.info("目前無法取得資料；本次投入水位未因此調整。")
            with ptt_col:
                st.markdown("**PTT Stock 標題情緒（實驗性）**")
                ptt = combined_market_snapshot.get("ptt_sentiment")
                if ptt:
                    n = ptt["sample_count"]
                    if ptt["bullish_ratio"] is None:
                        st.metric("標題樣本", f"{n} 篇 · 無明確多空詞")
                    else:
                        st.metric("多方占明確方向標題", f"{ptt['bullish_ratio'] * 100:.1f}% · {n} 篇標題")
                    directional_n = ptt["bullish"] + ptt["bearish"]
                    ptt_usable = n >= 20 and directional_n >= 10
                    st.caption(f"多方詞 {ptt['bullish']}、空方詞 {ptt['bearish']}、中性／混合 {ptt['neutral']}；需至少 20 篇總標題且 10 篇明確多空標題才會影響台股端，目前 {'已按 50% 權重納入' if ptt_usable else '有效樣本不足，未調整'}。")
                else:
                    st.info("目前無法取得 PTT 標題；本次投入水位未因此調整。")

            score_rows = [
                {"合併市場參考項目": name, "加權得分／權重": f"{points}／{weight}"}
                for name, (points, weight) in combined_market_snapshot["score_parts"].items()
            ]
            st.dataframe(pd.DataFrame(score_rows), hide_index=True, width="stretch")
            st.caption("此畫面先以台美等權作示意；模型水位仍屬實驗規則，不保證能辨識市場高低點。台美價格以各自指數計算，未將美元匯率變動納入水位。CNN 含 VIX 等市場因子，與技術環境可能重疊，因此只作有限度修正；PTT 僅分析近期標題，無法辨別反諷、文章品質或重複討論。Dcard 未接入；X 需官方 API。情緒值不回填歷史回測。")
        else:
            st.warning("目前無法同時取得台股與美股大盤資料，暫時無法計算合併水位。台股：" + str(taiwan_market_error or "無資料") + "；美股：" + str(us_market_error or "無資料"))

    with guide_tab:
        st.markdown("""
### 首頁的操作時點建議
- **偏向買進候選**：個股符合價值區、買進分數達設定門檻，且大盤環境分數至少40；以最近收盤資料判讀，隔日仍要確認價格和風險。
- **等待／不追買**：續漲條件不足、大盤資料不足或短線過熱；目前不列為新買進候選。
- **偏向減碼／賣出候選**：你選擇「目前持有」，且出現升溫轉弱確認或偏弱分數達賣出門檻。
- **偏向續抱觀察**：你選擇「目前持有」，續漲條件仍成立且未達賣出門檻；這不代表應該加碼。

### 怎麼使用選股區域與分數
- **價值區**：20／50日趨勢向上，近10日動能轉強，或出現放量突破／上升趨勢回踩等續漲條件；並需偏多分數至少60、偏弱風險低於40。這是短線候選，不是基本面低估或獲利保證。
- **升溫區｜偏熱觀察**：RSI、乖離、布林上軌或近20日漲幅顯示能量偏高，但尚未確認轉弱。避免追價，可考慮分批鎖利；不直接判定全部賣出。
- **升溫區｜轉弱確認**：能量偏高，且收盤跌破20日線或同時出現至少兩項轉弱訊號。持股者可檢查風險與分批減碼規則。
- **空檔**：續漲條件不足，也未形成明確升溫警示，先觀察。
- **歷史訊號檢查**：顯示各區域訊號後5、10、20個交易日的淨報酬統計；訊號每20日去重，並扣估計交易成本。樣本少時命中率不穩定。
- **資料涵蓋**：目前個股評分只涵蓋技術面；基本面、財報、法人與融資融券尚未納入。

### 怎麼使用這些分數
- **偏多分數**：彙整趨勢、動能、量價與乖離等偏多條件。
- **偏弱分數**：彙整趨勢轉弱、動能轉弱或過熱等風險訊號；高分時可進一步檢查持倉與自己的出場規則。
- **市場資金配置水位**：參考大盤均線趨勢、近一年區間位置與乖離；上升趨勢保留較高配置，確認空頭時降低，弱勢轉強時提高。它表示總資金配置比例，不是單筆下單比例，也不保證能抓到高低點。

目前權重是透明的示範規則，尚未經過完整歷史驗證。行情也可能延遲；這些訊號不會自動下單。
""")
except Exception as exc:
    st.error(f"分析時發生問題：{exc}")
    st.info("請確認網路連線與股票代碼格式。")

