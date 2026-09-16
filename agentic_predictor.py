# -*- coding: utf-8 -*-
"""
Agentic 1D Move Predictor — historical pattern analogue engine
────────────────────────────────────────────────────────────────────────────
Consumes the dataframe produced by nifty50_analyzer.fetch_nifty50_data()
(or FNOlist_Yfinancedata) and predicts the NEXT 1D bar's direction by
finding historical bars whose indicator state matched today's, then
measuring what actually happened next in those cases.

Feature families used (all already computed upstream):
  1. LSMA-WMA_Diff pattern  — sign + consecutive streak (momentum of the gap)
  2. Stoch RSI              — bucketed 0-100 zone
  3. Volume_Signal          — Strong_Buy_Vol / Strong_Sell_Vol / none
  4. BB_Position            — Upper/Lower breakout, inside bands
  5. Gann levels            — where Close sits between Gann_Support/Resistance

DESIGN PRINCIPLE — NO LOOKAHEAD:
Every prediction for bar i is built ONLY from bars < i. The analogue
search, the base rates, everything. This is what makes the reported
accuracy meaningful rather than a curve-fit illusion. The evaluation
below is a genuine walk-forward test, and it will happily report
"no edge" when there isn't one — that's the point.

Nothing here is investment advice. Technical pattern base rates on a
few hundred bars are noisy, and the honest expectation for next-day
direction prediction is an edge close to zero.
"""

from __future__ import annotations

import numpy as np
import pandas as pd

# ═════════════════════════════════════════════════════════════════════════
# 1. FEATURE DISCRETISATION
#    Each bar is reduced to a small set of categorical "state" tokens.
#    Coarse buckets are deliberate: fine-grained states produce unique
#    signatures that never repeat, so there'd be nothing to learn from.
# ═════════════════════════════════════════════════════════════════════════

def _streak(sign_series: pd.Series) -> pd.Series:
    """Consecutive-bar run length of the same sign (capped at 4+)."""
    s = sign_series.fillna(0).astype(int)
    grp = (s != s.shift(1)).cumsum()
    run = s.groupby(grp).cumcount() + 1
    return run.clip(upper=4) * s


def _bucket_stoch_rsi(v):
    if pd.isna(v):
        return "na"
    if v <= 20:
        return "oversold"
    if v <= 40:
        return "low"
    if v <= 60:
        return "mid"
    if v <= 80:
        return "high"
    return "overbought"


def _bucket_gann(row):
    """Where Close sits relative to that bar's own Gann band."""
    c, r, s = row.get("CLOSE"), row.get("Gann_Resistance"), row.get("Gann_Support")
    if pd.isna(c) or pd.isna(r) or pd.isna(s) or r <= s:
        return "na"
    pos = (c - s) / (r - s)
    if pos >= 0.85:
        return "near_resist"
    if pos <= 0.15:
        return "near_support"
    return "mid_band"


def build_features(df: pd.DataFrame) -> pd.DataFrame:
    """Add discretised state columns + the forward target. Operates on ONE
    stock's bars, already sorted oldest→newest."""
    d = df.copy().reset_index(drop=True)

    # 1. LSMA-WMA diff pattern: sign and how long it's persisted
    if "LSMA-WMA_Diff" in d.columns:
        diff = pd.to_numeric(d["LSMA-WMA_Diff"], errors="coerce")
    else:
        lw = pd.to_numeric(d.get("LSMA-WMA"), errors="coerce")
        diff = lw - lw.shift(1)
    d["_diff_sign"] = np.sign(diff).fillna(0)
    d["_diff_streak"] = _streak(d["_diff_sign"])
    d["f_diff"] = d["_diff_streak"].map(
        lambda x: f"up{int(abs(x))}" if x > 0 else (f"dn{int(abs(x))}" if x < 0 else "flat")
    )

    # 2. Stoch RSI zone
    d["f_srsi"] = pd.to_numeric(d.get("RSI"), errors="coerce").map(_bucket_stoch_rsi)

    # 3. Volume conviction
    vs = d.get("Volume_Signal", pd.Series([""] * len(d))).fillna("")
    d["f_vol"] = vs.map(lambda v: "buyvol" if v == "Strong_Buy_Vol"
                        else ("sellvol" if v == "Strong_Sell_Vol" else "none"))

    # 4. Bollinger position
    bb = d.get("BB_Position", pd.Series([""] * len(d))).fillna("")
    d["f_bb"] = bb.map(lambda v: {"Upper_Breakout": "bb_up",
                                   "Lower_Breakout": "bb_dn"}.get(v, "bb_in"))

    # 5. Gann band position
    d["f_gann"] = d.apply(_bucket_gann, axis=1)

    # Forward target: did the NEXT bar close higher than this bar's close?
    close = pd.to_numeric(d["CLOSE"], errors="coerce")
    d["_fwd_ret"] = close.shift(-1) / close - 1.0
    d["_fwd_up"] = (d["_fwd_ret"] > 0).astype(float)
    d.loc[d["_fwd_ret"].isna(), "_fwd_up"] = np.nan

    return d


# Signature layers, from most specific to most general. The agent backs off
# through these until it finds enough historical analogues to say anything.
_LAYERS = [
    ("full",     ["f_diff", "f_srsi", "f_vol", "f_bb", "f_gann"]),
    ("core",     ["f_diff", "f_srsi", "f_bb"]),
    ("momentum", ["f_diff", "f_srsi"]),
    ("diffonly", ["f_diff"]),
]


def _sig(row, cols):
    return "|".join(str(row[c]) for c in cols)


# ═════════════════════════════════════════════════════════════════════════
# 2. THE AGENT
# ═════════════════════════════════════════════════════════════════════════

MIN_SAMPLES = 12          # below this, a base rate is meaningless noise
EDGE_THRESHOLD = 0.08     # must beat 50/50 by this much to call a direction


def _wilson_lower(k, n, z=1.96):
    """Lower bound of the Wilson confidence interval — a sample-size-aware
    way of asking 'is this hit rate real or just a small-sample fluke?'
    12 wins out of 15 is far weaker evidence than 120 out of 150."""
    if n == 0:
        return 0.0
    p = k / n
    denom = 1 + z**2 / n
    centre = p + z**2 / (2 * n)
    margin = z * np.sqrt(p * (1 - p) / n + z**2 / (4 * n**2))
    return (centre - margin) / denom


def agent_predict(hist: pd.DataFrame, current_row: pd.Series) -> dict:
    """Given a stock's PAST bars (hist, strictly before the bar we're
    predicting from) and the current bar's state, find analogues and
    decide. Returns a verdict dict including an explicit ABSTAIN when
    the evidence doesn't support a call."""
    usable = hist.dropna(subset=["_fwd_up"])
    if len(usable) < MIN_SAMPLES:
        return {"direction": "ABSTAIN", "confidence": 0.0, "n": len(usable),
                "layer": "none", "up_rate": np.nan, "avg_ret": np.nan,
                "reason": "Not enough history yet to establish any base rate."}

    base_rate = usable["_fwd_up"].mean()

    for layer_name, cols in _LAYERS:
        target = _sig(current_row, cols)
        sigs = usable.apply(lambda r: _sig(r, cols), axis=1)
        match = usable[sigs == target]
        n = len(match)
        if n < MIN_SAMPLES:
            continue

        up_rate = match["_fwd_up"].mean()
        avg_ret = match["_fwd_ret"].mean()
        k_up = int(match["_fwd_up"].sum())

        # Evidence strength, adjusted for sample size, in both directions
        lo_up = _wilson_lower(k_up, n)
        lo_dn = _wilson_lower(n - k_up, n)

        edge = up_rate - base_rate
        if up_rate >= 0.5 + EDGE_THRESHOLD and lo_up > 0.5:
            direction, conf = "UP", lo_up
        elif up_rate <= 0.5 - EDGE_THRESHOLD and lo_dn > 0.5:
            direction, conf = "DOWN", lo_dn
        else:
            # Pattern found, but it isn't separated from a coin flip.
            return {"direction": "ABSTAIN", "confidence": 0.0, "n": n,
                    "layer": layer_name, "up_rate": up_rate, "avg_ret": avg_ret,
                    "base_rate": base_rate, "signature": target,
                    "reason": (f"Found {n} historical analogues of this exact state, but they "
                               f"resolved up only {up_rate*100:.0f}% of the time — not "
                               f"distinguishable from chance, so no directional call.")}

        return {"direction": direction, "confidence": conf, "n": n,
                "layer": layer_name, "up_rate": up_rate, "avg_ret": avg_ret,
                "base_rate": base_rate, "signature": target,
                "reason": (f"In {n} past bars where this stock showed the same state "
                           f"({target.replace('|', ', ')}), the next bar closed "
                           f"{'higher' if direction=='UP' else 'lower'} "
                           f"{(up_rate if direction=='UP' else 1-up_rate)*100:.0f}% of the time "
                           f"(vs a {base_rate*100:.0f}% baseline), averaging "
                           f"{avg_ret*100:+.2f}% per bar. Matched at the '{layer_name}' "
                           f"specificity layer.")}

    return {"direction": "ABSTAIN", "confidence": 0.0, "n": 0, "layer": "none",
            "up_rate": np.nan, "avg_ret": np.nan, "base_rate": base_rate,
            "reason": "Today's indicator combination has no sufficiently-repeated precedent "
                      "in this stock's history."}


# ═════════════════════════════════════════════════════════════════════════
# 3. WALK-FORWARD BACKTEST — the honesty check
# ═════════════════════════════════════════════════════════════════════════

def walkforward_eval(feat: pd.DataFrame, warmup: int = 60) -> dict:
    """Replay the stock bar by bar. At each step the agent may only see
    bars strictly before it. This is the number that tells you whether
    the engine has any real predictive edge, as opposed to having
    memorised its own training data."""
    rows = []
    for i in range(warmup, len(feat) - 1):
        hist = feat.iloc[:i]
        cur = feat.iloc[i]
        if pd.isna(cur["_fwd_up"]):
            continue
        p = agent_predict(hist, cur)
        if p["direction"] == "ABSTAIN":
            continue
        correct = ((p["direction"] == "UP" and cur["_fwd_up"] == 1) or
                   (p["direction"] == "DOWN" and cur["_fwd_up"] == 0))
        rows.append({"i": i, "direction": p["direction"], "correct": bool(correct),
                     "confidence": p["confidence"], "fwd_ret": cur["_fwd_ret"]})

    if not rows:
        return {"n_calls": 0, "hit_rate": np.nan, "baseline": np.nan,
                "coverage": 0.0, "verdict": "Agent abstained on every bar — no testable edge."}

    r = pd.DataFrame(rows)
    hit = r["correct"].mean()
    tested = len(feat.iloc[warmup:-1].dropna(subset=["_fwd_up"]))
    baseline = feat["_fwd_up"].iloc[warmup:].mean()   # always-guess-up accuracy
    # naive always-up baseline is whichever of up/down is more common
    baseline = max(baseline, 1 - baseline)

    n = len(r)
    lo = _wilson_lower(int(r["correct"].sum()), n)
    if lo > baseline:
        verdict = (f"Hit rate {hit*100:.1f}% over {n} calls, and even the pessimistic "
                   f"end of its confidence interval ({lo*100:.1f}%) clears the "
                   f"{baseline*100:.1f}% always-guess-the-majority baseline. Weak but real edge.")
    elif hit > baseline:
        verdict = (f"Hit rate {hit*100:.1f}% over {n} calls edges past the {baseline*100:.1f}% "
                   f"baseline, but the sample is too small to rule out luck. Treat as unproven.")
    else:
        verdict = (f"Hit rate {hit*100:.1f}% over {n} calls does NOT beat simply always "
                   f"guessing the majority direction ({baseline*100:.1f}%). No demonstrated edge "
                   f"on this stock.")

    return {"n_calls": n, "hit_rate": hit, "baseline": baseline,
            "coverage": n / tested if tested else 0.0,
            "wilson_lo": lo, "avg_ret_when_called": r["fwd_ret"].mean(),
            "verdict": verdict}


# ═════════════════════════════════════════════════════════════════════════
# 4. POOLED (CROSS-SECTIONAL) ANALOGUE LIBRARY
#
# A single stock only has a few hundred bars, so a specific indicator
# state rarely repeats often enough to say anything statistically. Pooling
# the same state across the whole basket gives sample sizes in the
# hundreds or thousands, which is what makes the base rates meaningful.
#
# The time-ordering discipline is preserved: a prediction made on date t
# only ever counts analogues whose outcome was already observable strictly
# BEFORE t. Same-date bars from other stocks are excluded too, since their
# next-bar outcome isn't known yet at decision time.
# ═════════════════════════════════════════════════════════════════════════

def _pooled_tables(all_feat: pd.DataFrame):
    """For each signature layer, build a date-indexed cumulative table of
    (occurrences, up-count) that can be queried 'as of strictly before
    date t' in O(1)."""
    tables = {}
    for layer_name, cols in _LAYERS:
        d = all_feat.dropna(subset=["_fwd_up"]).copy()
        d["_sig"] = d.apply(lambda r: _sig(r, cols), axis=1)
        # aggregate per (signature, date), then cumulate over dates and
        # shift by one date so the current date contributes nothing.
        per = (d.groupby(["_sig", "Date"])["_fwd_up"]
                 .agg(["count", "sum"]).reset_index()
                 .sort_values("Date"))
        per["cum_n"] = per.groupby("_sig")["count"].cumsum() - per["count"]
        per["cum_up"] = per.groupby("_sig")["sum"].cumsum() - per["sum"]
        tables[layer_name] = per.set_index(["_sig", "Date"])[["cum_n", "cum_up"]]
    # overall base rate as of each date
    b = (all_feat.dropna(subset=["_fwd_up"]).groupby("Date")["_fwd_up"]
         .agg(["count", "sum"]).sort_index())
    b["cum_n"] = b["count"].cumsum() - b["count"]
    b["cum_up"] = b["sum"].cumsum() - b["sum"]
    return tables, b[["cum_n", "cum_up"]]


def _lookup(tables, layer, sig, date):
    try:
        row = tables[layer].loc[(sig, date)]
        return int(row["cum_n"]), int(row["cum_up"])
    except KeyError:
        # signature never seen on this exact date — find the latest earlier date
        try:
            sub = tables[layer].xs(sig, level="_sig")
            sub = sub[sub.index < date]
            if sub.empty:
                return 0, 0
            last = sub.iloc[-1]
            return int(last["cum_n"] + 0), int(last["cum_up"] + 0)
        except KeyError:
            return 0, 0


def agent_predict_pooled(tables, base_tbl, current_row) -> dict:
    """Same decision logic as agent_predict, but the analogue counts come
    from the whole basket's history rather than one stock's."""
    date = current_row["Date"]
    prior = base_tbl[base_tbl.index < date]
    if prior.empty or prior.iloc[-1]["cum_n"] + prior.iloc[-1].get("count", 0) < MIN_SAMPLES:
        base_n = prior.iloc[-1]["cum_n"] if not prior.empty else 0
        if base_n < MIN_SAMPLES:
            return {"direction": "ABSTAIN", "confidence": 0.0, "n": 0, "layer": "none",
                    "up_rate": np.nan, "avg_ret": np.nan,
                    "reason": "Not enough pooled history yet to establish a base rate."}
    base_row = prior.iloc[-1]
    base_rate = base_row["cum_up"] / base_row["cum_n"] if base_row["cum_n"] else 0.5

    for layer_name, cols in _LAYERS:
        sig = _sig(current_row, cols)
        n, k_up = _lookup(tables, layer_name, sig, date)
        if n < MIN_SAMPLES:
            continue

        up_rate = k_up / n
        lo_up = _wilson_lower(k_up, n)
        lo_dn = _wilson_lower(n - k_up, n)

        if up_rate >= 0.5 + EDGE_THRESHOLD and lo_up > 0.5:
            direction, conf = "UP", lo_up
        elif up_rate <= 0.5 - EDGE_THRESHOLD and lo_dn > 0.5:
            direction, conf = "DOWN", lo_dn
        else:
            return {"direction": "ABSTAIN", "confidence": 0.0, "n": n,
                    "layer": layer_name, "up_rate": up_rate, "avg_ret": np.nan,
                    "base_rate": base_rate, "signature": sig,
                    "reason": (f"{n} pooled historical analogues of this state resolved up "
                               f"{up_rate*100:.0f}% of the time — not separable from chance, "
                               f"so no directional call.")}

        return {"direction": direction, "confidence": conf, "n": n,
                "layer": layer_name, "up_rate": up_rate, "avg_ret": np.nan,
                "base_rate": base_rate, "signature": sig,
                "reason": (f"Across {n} past bars in this basket with the same state "
                           f"({sig.replace('|', ', ')}), the next bar closed "
                           f"{'higher' if direction=='UP' else 'lower'} "
                           f"{(up_rate if direction=='UP' else 1-up_rate)*100:.0f}% of the time, "
                           f"against a {base_rate*100:.0f}% baseline. Matched at the "
                           f"'{layer_name}' specificity layer.")}

    return {"direction": "ABSTAIN", "confidence": 0.0, "n": 0, "layer": "none",
            "up_rate": np.nan, "avg_ret": np.nan, "base_rate": base_rate,
            "reason": "This indicator combination has no sufficiently-repeated precedent "
                      "anywhere in the basket's history."}


# ═════════════════════════════════════════════════════════════════════════
# 5. TOP-LEVEL DRIVER
# ═════════════════════════════════════════════════════════════════════════

def run_predictions(df: pd.DataFrame, warmup: int = 60,
                    evaluate: bool = True) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """df = multi-stock analyzer output (1D bars).
    Returns (predictions, per-stock detail, pooled evaluation)."""
    frames = []
    for name, g in df.groupby("Stock Name"):
        g = g.sort_values("Date") if "Date" in g.columns else g
        g = g.reset_index(drop=True)
        if len(g) < 40:
            continue
        frames.append(build_features(g))
    if not frames:
        return pd.DataFrame(), pd.DataFrame(), {"verdict": "Not enough data."}

    all_feat = pd.concat(frames, ignore_index=True)
    if "Date" not in all_feat.columns:
        all_feat["Date"] = all_feat.groupby("Stock Name").cumcount()
    all_feat = all_feat.sort_values("Date").reset_index(drop=True)

    tables, base_tbl = _pooled_tables(all_feat)

    # ── Live prediction: latest bar of each stock ──
    preds = []
    for name, g in all_feat.groupby("Stock Name"):
        cur = g.sort_values("Date").iloc[-1]
        p = agent_predict_pooled(tables, base_tbl, cur)
        preds.append({
            "Stock Name": name,
            "Close": cur.get("CLOSE"),
            "Prediction": p["direction"],
            "Confidence": round(p["confidence"] * 100, 1) if p["confidence"] else 0.0,
            "Analogues": p["n"],
            "Hist Up Rate": round(p["up_rate"] * 100, 1) if pd.notna(p.get("up_rate")) else None,
            "Match Layer": p["layer"],
            "State": p.get("signature", ""),
            "Why": p["reason"],
        })
    pred_df = pd.DataFrame(preds)
    if not pred_df.empty:
        order = {"UP": 0, "DOWN": 1, "ABSTAIN": 2}
        pred_df = (pred_df.assign(_o=pred_df["Prediction"].map(order))
                          .sort_values(["_o", "Confidence"], ascending=[True, False])
                          .drop(columns="_o").reset_index(drop=True))

    # ── Pooled walk-forward evaluation ──
    ev = {}
    if evaluate:
        test = all_feat.dropna(subset=["_fwd_up"])
        dates = sorted(test["Date"].unique())
        cutoff = dates[int(len(dates) * 0.35)] if len(dates) > 10 else dates[0]
        test = test[test["Date"] > cutoff]
        rows = []
        for _, r in test.iterrows():
            p = agent_predict_pooled(tables, base_tbl, r)
            if p["direction"] == "ABSTAIN":
                continue
            correct = ((p["direction"] == "UP" and r["_fwd_up"] == 1) or
                       (p["direction"] == "DOWN" and r["_fwd_up"] == 0))
            rows.append({"correct": bool(correct), "fwd_ret": r["_fwd_ret"],
                         "direction": p["direction"]})
        if rows:
            rr = pd.DataFrame(rows)
            n, hit = len(rr), rr["correct"].mean()
            up_share = test["_fwd_up"].mean()
            baseline = max(up_share, 1 - up_share)
            lo = _wilson_lower(int(rr["correct"].sum()), n)
            if lo > baseline:
                verdict = (f"Hit rate {hit*100:.1f}% across {n} out-of-sample calls; even the "
                           f"pessimistic end of the interval ({lo*100:.1f}%) clears the "
                           f"{baseline*100:.1f}% majority-guess baseline. Weak but measurable edge.")
            elif hit > baseline:
                verdict = (f"Hit rate {hit*100:.1f}% across {n} calls edges past the "
                           f"{baseline*100:.1f}% baseline, but not by enough to rule out luck. "
                           f"Treat as unproven.")
            else:
                verdict = (f"Hit rate {hit*100:.1f}% across {n} calls does not beat the "
                           f"{baseline*100:.1f}% majority-guess baseline. No demonstrated edge — "
                           f"treat the predictions as descriptive, not predictive.")
            ev = {"n_calls": n, "hit_rate": hit, "baseline": baseline, "wilson_lo": lo,
                  "coverage": n / len(test) if len(test) else 0,
                  "avg_ret_when_called": rr["fwd_ret"].mean(), "verdict": verdict}
        else:
            ev = {"n_calls": 0, "hit_rate": np.nan, "baseline": np.nan, "coverage": 0.0,
                  "verdict": "The agent abstained on every out-of-sample bar — no testable edge."}

    return pred_df, all_feat, ev
