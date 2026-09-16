# -*- coding: utf-8 -*-
"""
Agentic 1D Move Predictor — Dashboard
────────────────────────────────────────────────────────────────────────────
Standalone Streamlit app. Pulls 1D bars through your existing analyzer,
then runs the historical-analogue agent in agentic_predictor.py.

Run:  streamlit run agentic_predictor_dashboard.py
Needs: agentic_predictor.py + one of your analyzer modules in the same folder.
"""

import importlib.util
import os
from datetime import datetime
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st

import agentic_predictor as ap

st.set_page_config(page_title="Agentic 1D Predictor", layout="wide", page_icon="🔮")

st.markdown("""
<style>
 .p-up   { background:#DCFCE7;color:#166534;padding:3px 10px;border-radius:20px;font-weight:700;font-size:12px }
 .p-down { background:#FEE2E2;color:#991B1B;padding:3px 10px;border-radius:20px;font-weight:700;font-size:12px }
 .p-abs  { background:#F1F5F9;color:#475569;padding:3px 10px;border-radius:20px;font-weight:700;font-size:12px }
 .t-title{ font-size:26px;font-weight:800 }
 .t-sub  { font-size:13px;color:#666;margin-top:-6px }
</style>""", unsafe_allow_html=True)

st.markdown('<div class="t-title">🔮 Agentic 1D Move Predictor</div>', unsafe_allow_html=True)
st.markdown('<div class="t-sub">Finds historical bars whose LSMA-WMA diff pattern, Stoch RSI, '
            'volume conviction, Bollinger position and Gann placement matched today — then reports '
            'what actually happened next.</div>', unsafe_allow_html=True)
st.write("")


# ── Load whichever analyzer module is present ────────────────────────────────
@st.cache_resource
def load_analyzer():
    for fname in ("nifty50_analyzer.py", "FNOlist_Yfinancedata.py"):
        if os.path.exists(fname):
            spec = importlib.util.spec_from_file_location(fname[:-3], fname)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod, fname
    return None, None


mod, mod_name = load_analyzer()
if mod is None:
    st.error("Couldn't find `nifty50_analyzer.py` or `FNOlist_Yfinancedata.py` next to this file. "
             "Put one of them in the same folder.")
    st.stop()

REGISTRY = getattr(mod, "_TICKER_LIST_REGISTRY", {"Nifty": mod.NIFTY50_TICKERS})

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.caption(f"Data source: `{mod_name}`")
    list_name = st.selectbox("Ticker list", list(REGISTRY.keys()))
    n_days = st.slider("History (calendar days)", 180, 1460, 730, 30,
                        help="More history = more analogues = more reliable base rates. "
                             "Two years or more is strongly recommended.")
    min_samples = st.slider("Min analogues before the agent will call", 10, 100, 25, 5)
    edge_thr = st.slider("Required edge over 50/50", 0.02, 0.25, 0.08, 0.01)
    run = st.button("▶ Run Agent", type="primary", width="stretch")

ap.MIN_SAMPLES = min_samples
ap.EDGE_THRESHOLD = edge_thr

if "ran" not in st.session_state:
    st.session_state.ran = False
if run:
    st.session_state.ran = True
if not st.session_state.ran:
    st.info("Choose a ticker list on the left and hit **▶ Run Agent**.")
    st.stop()


@st.cache_data(ttl=900, show_spinner=False)
def fetch(list_name, n_days):
    return mod.fetch_nifty50_data(n_days, "1d", 2, -2, ticker_list=REGISTRY[list_name])


with st.spinner("📡 Fetching 1D history..."):
    try:
        df = fetch(list_name, n_days)
    except Exception as e:
        st.error(f"Data fetch failed: {e}")
        st.stop()

if df is None or df.empty:
    st.error("No data returned.")
    st.stop()

with st.spinner("🧠 Agent is mining historical analogues..."):
    preds, feat, ev = ap.run_predictions(df)

if preds.empty:
    st.warning("Not enough history per stock to build analogues. Increase the history window.")
    st.stop()

# ═════════════════════════════════════════════════════════════════════
# Honesty panel — shown FIRST, deliberately
# ═════════════════════════════════════════════════════════════════════
st.markdown("## 📊 Does this actually work? (out-of-sample test)")
if ev.get("n_calls"):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Out-of-sample calls", f"{ev['n_calls']:,}")
    c2.metric("Hit rate", f"{ev['hit_rate']*100:.1f}%")
    c3.metric("Majority-guess baseline", f"{ev['baseline']*100:.1f}%")
    c4.metric("Coverage", f"{ev['coverage']*100:.1f}%",
              help="Share of bars the agent was willing to call at all. Low is fine — "
                   "abstaining beats guessing.")
    beat = ev.get("wilson_lo", 0) > ev.get("baseline", 1)
    (st.success if beat else st.warning)(ev["verdict"])
    if not beat:
        st.caption("Because the backtest does not establish an edge, the predictions below are best "
                   "read as *descriptions of historical base rates*, not as forecasts. "
                   "Sizing real money off them is not supported by this evidence.")
else:
    st.warning(ev.get("verdict", "No evaluation available."))

st.divider()

# ═════════════════════════════════════════════════════════════════════
# Predictions
# ═════════════════════════════════════════════════════════════════════
n_up = (preds["Prediction"] == "UP").sum()
n_dn = (preds["Prediction"] == "DOWN").sum()
n_ab = (preds["Prediction"] == "ABSTAIN").sum()

st.markdown("## 🎯 Next-Bar Calls")
k1, k2, k3 = st.columns(3)
k1.metric("🟢 UP", n_up)
k2.metric("🔴 DOWN", n_dn)
k3.metric("⚪ Abstained", n_ab)

called = preds[preds["Prediction"] != "ABSTAIN"]
if called.empty:
    st.info("The agent abstained on every stock — no indicator state today has a strong enough "
            "historical precedent. That's a legitimate output, not a failure.")
else:
    for _, r in called.iterrows():
        cls = "p-up" if r["Prediction"] == "UP" else "p-down"
        with st.container(border=True):
            a, b, c = st.columns([2, 1.2, 5])
            a.markdown(f"**{r['Stock Name']}**  \n₹{r['Close']}")
            b.markdown(f'<span class="{cls}">{r["Prediction"]}</span><br>'
                       f'Confidence: <b>{r["Confidence"]}%</b>', unsafe_allow_html=True)
            c.markdown(f"{r['Why']}")

with st.expander(f"⚪ Abstained ({n_ab}) — why the agent declined to call these"):
    for _, r in preds[preds["Prediction"] == "ABSTAIN"].iterrows():
        st.markdown(f"**{r['Stock Name']}** — {r['Why']}")

st.divider()
st.markdown("## 📋 Full Table")
st.dataframe(preds[["Stock Name", "Close", "Prediction", "Confidence", "Analogues",
                    "Hist Up Rate", "Match Layer", "State"]],
             width="stretch", hide_index=True)

csv = preds.to_csv(index=False).encode()
st.download_button("⬇️ Download predictions (CSV)", csv,
                   file_name=f"agentic_1d_predictions_"
                             f"{datetime.now(ZoneInfo('Asia/Kolkata')).strftime('%Y%m%d_%H%M')}.csv",
                   mime="text/csv", width="stretch")

st.caption("Predictions are historical base rates for matching indicator states, not forecasts. "
           "Next-day direction is close to a coin flip in liquid markets; any edge here is small, "
           "unstable, and ignores costs and slippage. Not investment advice.")
