import re
import time
import requests
import pandas as pd
import streamlit as st
from io import BytesIO
from thefuzz import fuzz

# ── Page config ───────────────────────────────────────────────────
st.set_page_config(
    page_title="Ticker Lookup Bot",
    page_icon="📈",
    layout="centered",
)

st.title("📈 Ticker Symbol Lookup Bot")
st.markdown(
    "Upload your Excel file and this bot will automatically find the "
    "**NYSE / NASDAQ ticker symbol** for each company name — "
    "even if names have extra words like *NPS*, *Inc.*, *Report*, etc."
)
st.divider()

# ── Constants ─────────────────────────────────────────────────────
NOISE_PATTERNS = [
    r"\bNPS\b", r"\bPXY\b", r"\bCSR\b", r"\bRevcon\b",
    r"\bCombo\b", r"\bOptimized\b", r"\bPDF\b",
    r"Asset\s+Handbook", r"Corporate\s+Report",
    r"Service\s+Support", r"CSR\s+Report",
    r"2\d{3}",
]
CORP_SUFFIXES = [
    r"\bIncorporated\b", r"\bInc\.?", r"\bCorporation\b", r"\bCorp\.?",
    r"\bLimited\b", r"\bLtd\.?", r"\bPLC\b", r"\bLLC\b",
]
NOT_LISTED = {
    "ypo", "bdo private bank", "bdo unibank", "bdo network bank",
    "bdo insure bank", "lus", "dmc", "international",
    "automotive", "northern",
}
MAJOR_EXCHANGES = {"NMS", "NYQ", "NYS", "NGM", "NCM", "ASE"}
MIN_SCORE = 50


# ── Helper functions ──────────────────────────────────────────────
def clean_name(raw: str) -> str:
    name = raw.strip()
    name = re.split(r"\s+[-]\s+", name)[0].strip()
    for pat in NOISE_PATTERNS:
        name = re.sub(pat, " ", name, flags=re.IGNORECASE).strip()
    for pat in CORP_SUFFIXES:
        name = re.sub(pat, " ", name, flags=re.IGNORECASE).strip()
    name = re.sub(r"[.\-,;]+\s*$", "", name).strip()
    name = re.sub(r"\s{2,}", " ", name).strip()
    return name


def yahoo_search(query: str) -> list:
    url = "https://query2.finance.yahoo.com/v1/finance/search"
    params = {"q": query, "quotesCount": 6, "newsCount": 0, "enableFuzzyQuery": True}
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json",
        "Referer": "https://finance.yahoo.com/",
    }
    try:
        r = requests.get(url, params=params, headers=headers, timeout=10)
        r.raise_for_status()
        return [q for q in r.json().get("quotes", []) if q.get("quoteType") == "EQUITY"]
    except Exception:
        return []


def pick_best_quote(cleaned: str, quotes: list) -> tuple:
    best_score, best = 0, None
    for q in quotes:
        candidate = q.get("longname") or q.get("shortname") or ""
        score = fuzz.token_set_ratio(cleaned.lower(), candidate.lower())
        extra_words = len(candidate.split()) - len(cleaned.split())
        if extra_words > 1:
            score -= extra_words * 4
        if q.get("exchange", "") in MAJOR_EXCHANGES:
            score += 5
        score = max(0, min(100, score))
        if score > best_score:
            best_score, best = score, q
    if best:
        display = best.get("longname") or best.get("shortname") or ""
        return best.get("symbol", ""), display, best.get("exchange", ""), best_score
    return "", "", "", 0


def to_excel(found_df: pd.DataFrame, review_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        found_df.to_excel(writer, index=False, sheet_name="Tickers")
        review_df.to_excel(writer, index=False, sheet_name="Manual Review")
    return output.getvalue()


# ── Step 1 — Upload ───────────────────────────────────────────────
st.subheader("Step 1 — Upload your Excel file")
uploaded = st.file_uploader("Choose an .xlsx file", type=["xlsx"])

if uploaded:
    try:
        xl = pd.ExcelFile(uploaded)
        sheet = st.selectbox("Select sheet", xl.sheet_names)
        df_raw = xl.parse(sheet, header=None)

        st.markdown("**Preview (first 5 rows):**")
        st.dataframe(df_raw.head(5), use_container_width=True)

        header_row = st.number_input(
            "Which row is the header? (0 = first row)",
            min_value=0, max_value=20, value=0, step=1,
        )
        df = xl.parse(sheet, header=int(header_row))
        col = st.selectbox("Which column has company names?", df.columns.tolist())

    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

    st.divider()

    # ── Step 2 — Run ──────────────────────────────────────────────
    st.subheader("Step 2 — Find tickers")
    delay = st.slider(
        "Delay between requests (seconds) — increase if you get errors",
        min_value=0.2, max_value=2.0, value=0.5, step=0.1,
    )

    if st.button("🔍  Start Lookup", type="primary", use_container_width=True):

        names = df[col].dropna().astype(str).str.strip().tolist()
        names = [n for n in names if n and n.lower() != str(col).lower()]

        found_rows, review_rows = [], []
        seen = set()

        progress_bar = st.progress(0, text="Starting...")
        log_area = st.empty()
        log_lines = []

        def add_log(msg, icon=""):
            entry = f"{icon}  {msg}" if icon else msg
            log_lines.append(entry)
            log_area.code("\n".join(log_lines[-30:]), language=None)

        total = len(names)

        for i, raw in enumerate(names):
            progress_bar.progress((i + 1) / total, text=f"Processing {i + 1} of {total}...")

            if raw.lower() in NOT_LISTED:
                review_rows.append({"Original Name": raw, "Cleaned Name": raw,
                                    "Reason": "Known non-listed entity", "Ticker (fill manually)": ""})
                add_log(f"Skipped (non-listed): {raw}", "⊘")
                continue

            cleaned = clean_name(raw)
            if not cleaned:
                review_rows.append({"Original Name": raw, "Cleaned Name": "",
                                    "Reason": "Name empty after cleaning", "Ticker (fill manually)": ""})
                continue

            key = cleaned.lower()
            if key in seen:
                add_log(f"Duplicate skipped: {raw}", "⟳")
                continue
            seen.add(key)

            add_log(f'Searching: "{cleaned}"', "→")
            quotes = yahoo_search(cleaned)
            time.sleep(delay)

            if not quotes:
                review_rows.append({"Original Name": raw, "Cleaned Name": cleaned,
                                    "Reason": "No results from Yahoo Finance", "Ticker (fill manually)": ""})
                add_log(f"No results: {cleaned}", "⚠")
                continue

            ticker, long_name, exchange, score = pick_best_quote(cleaned, quotes)

            if not ticker or score < MIN_SCORE:
                review_rows.append({"Original Name": raw, "Cleaned Name": cleaned,
                                    "Reason": f"Low confidence match (score: {score}%)",
                                    "Ticker (fill manually)": ""})
                add_log(f"Low confidence ({score}%): {cleaned}", "⚠")
            else:
                confidence = "High" if score >= 75 else "Medium"
                found_rows.append({"Original Name": raw, "Cleaned Name": cleaned,
                                   "Ticker": ticker, "Company (Yahoo)": long_name,
                                   "Exchange": exchange, "Confidence": confidence, "Score": score})
                icon = "✅" if confidence == "High" else "🟡"
                add_log(f"{ticker}  |  {long_name}  [{confidence}, {score}%]", icon)

        progress_bar.progress(1.0, text="Done!")

        found_df = pd.DataFrame(found_rows)
        review_df = pd.DataFrame(review_rows)

        # ── Step 3 — Results ──────────────────────────────────────
        st.divider()
        st.subheader("Step 3 — Results")

        m1, m2, m3 = st.columns(3)
        m1.metric("✅ Tickers found",    len(found_rows))
        m2.metric("⚠️  Manual review",  len(review_rows))
        m3.metric("📋 Total processed", len(found_rows) + len(review_rows))

        tab1, tab2 = st.tabs(["✅ Tickers Found", "⚠️ Manual Review"])
        with tab1:
            if not found_df.empty:
                st.dataframe(found_df, use_container_width=True)
            else:
                st.info("No tickers found automatically. Check the Manual Review tab.")
        with tab2:
            if not review_df.empty:
                st.dataframe(review_df, use_container_width=True)
                st.caption("Fill in the 'Ticker (fill manually)' column for these entries after downloading.")
            else:
                st.success("All companies were matched automatically!")

        # ── Step 4 — Download ─────────────────────────────────────
        st.divider()
        st.subheader("Step 4 — Download results")
        st.download_button(
            label="⬇️  Download results as Excel (.xlsx)",
            data=to_excel(found_df, review_df),
            file_name="companies_with_tickers.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(
            "The Excel file has two sheets: **Tickers** (matched results) "
            "and **Manual Review** (entries that need manual lookup)."
        )

# ── Footer ────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Powered by Yahoo Finance search API. "
    "Always verify ticker symbols before use in financial decisions."
)
