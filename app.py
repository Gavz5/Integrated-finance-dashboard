# app.py
# Integrated Financial Performance & Governance Dashboard
# + Management "Screenshot-style" One-Page Dashboard (for presentation)
#
# Run:
#   pip install -r requirements.txt
#   streamlit run app.py

from __future__ import annotations

import os
from pathlib import Path
from calendar import monthrange

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="Integrated Financial Performance & Governance Dashboard",
    layout="wide",
)

BASE_DIR = Path(__file__).parent

# Put your model-map image inside repo, e.g. ./assets/model_map.png
ASSETS_DIR = BASE_DIR / "assets"
MODEL_MAP_CANDIDATES = [
    ASSETS_DIR / "model_map.png",
    ASSETS_DIR / "overview.png",
    ASSETS_DIR / "model.png",
]

INR_SYMBOL = "₹"


# -----------------------------
# GLOBAL CSS (Management look)
# -----------------------------
st.markdown(
    """
<style>
/* Reduce top padding */
.block-container { padding-top: 1rem !important; }

/* Hide Streamlit footer */
footer { visibility: hidden; }

/* Management title bar */
.mgmt-title {
    background: linear-gradient(180deg, #0b2d56 0%, #0a2a52 55%, #082545 100%);
    border-radius: 16px;
    padding: 22px 24px;
    color: white;
    font-weight: 900;
    font-size: 34px;
    letter-spacing: 0.2px;
    text-align: center;
    box-shadow: 0 10px 24px rgba(15, 23, 42, 0.18);
    margin-top: 6px;
    margin-bottom: 16px;
}

/* KPI tile row */
.kpi-tile {
    border-radius: 14px;
    padding: 16px 14px;
    color: #fff;
    font-weight: 800;
    text-align: center;
    box-shadow: 0 10px 22px rgba(15, 23, 42, 0.10);
}
.kpi-label { font-size: 15px; opacity: 0.95; }
.kpi-value { font-size: 20px; margin-top: 6px; }

/* Section header bar like screenshot */
.section-header {
    background: #0b2d56;
    color: white;
    padding: 12px 14px;
    border-radius: 12px 12px 0 0;
    font-weight: 900;
    font-size: 18px;
}

/* Panel container */
.panel {
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    background: #fff;
    box-shadow: 0 10px 22px rgba(15,23,42,0.06);
    overflow: hidden;
    margin-bottom: 16px;
}
.panel-body {
    padding: 12px 14px 14px 14px;
}

/* Thin formula strip */
.formula-strip {
    background: #154277;
    color: white;
    padding: 12px 14px;
    border-radius: 14px;
    font-weight: 900;
    text-align: center;
    margin-top: 14px;
    margin-bottom: 16px;
    box-shadow: 0 10px 22px rgba(15,23,42,0.10);
}

/* Control strip container */
.control-strip {
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    background: #fbfbfd;
    padding: 12px 14px;
    margin-bottom: 10px;
}

/* Small helper text */
.small-note { color: #6b7280; font-size: 12px; }
</style>
""",
    unsafe_allow_html=True,
)


# -----------------------------
# HELPERS
# -----------------------------
def inr_full(x) -> str:
    """Full INR (no L/Cr truncation)."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        x = float(x)
    except Exception:
        return str(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    return f"{sign}{INR_SYMBOL}{x:,.0f}"


def inr_compact(x) -> str:
    """Compact INR for tiles (L / Cr)."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        v = float(x)
    except Exception:
        return str(x)

    sign = "-" if v < 0 else ""
    v = abs(v)

    # Lakhs & Crores
    if v >= 1e7:
        return f"{sign}{INR_SYMBOL}{v/1e7:.1f} Cr"
    if v >= 1e5:
        return f"{sign}{INR_SYMBOL}{v/1e5:.1f} L"
    return f"{sign}{INR_SYMBOL}{v:,.0f}"


def pct(x, digits=1) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        return f"{100 * float(x):.{digits}f}%"
    except Exception:
        return "—"


def safe_div(a, b):
    try:
        if b in (0, 0.0) or (isinstance(b, float) and np.isnan(b)):
            return np.nan
        return a / b
    except Exception:
        return np.nan


def month_start_end(month_dt: pd.Timestamp):
    ms = pd.Timestamp(month_dt.year, month_dt.month, 1)
    me = pd.Timestamp(month_dt.year, month_dt.month, monthrange(month_dt.year, month_dt.month)[1])
    return ms, me


def to_month_ts(x):
    if isinstance(x, pd.Timestamp):
        return pd.Timestamp(x.year, x.month, 1)
    try:
        t = pd.to_datetime(x)
        return pd.Timestamp(t.year, t.month, 1)
    except Exception:
        return pd.NaT


def compute_wc_days(ar, ap, inv, revenue, cogs, days_in_month):
    dso = safe_div(ar, revenue) * days_in_month if revenue else np.nan
    dpo = safe_div(ap, abs(cogs)) * days_in_month if cogs else np.nan
    dio = safe_div(inv, abs(cogs)) * days_in_month if cogs else np.nan

    if np.isnan(dso) and np.isnan(dpo) and np.isnan(dio):
        return np.nan, np.nan, np.nan, np.nan

    ccc = (0 if np.isnan(dso) else dso) + (0 if np.isnan(dio) else dio) - (0 if np.isnan(dpo) else dpo)
    return dso, dpo, dio, ccc


def rag_from_threshold(value, green_max=None, green_min=None, amber_band=0.1, higher_is_bad=True):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "AMBER", "Insufficient data (missing/zero denominator)."

    if higher_is_bad:
        if green_max is None:
            return "AMBER", "No threshold set."
        if value > green_max:
            return "RED", "Breached limit."
        if value > (1 - amber_band) * green_max:
            return "AMBER", "Near limit."
        return "GREEN", "Within comfort zone."
    else:
        if green_min is None:
            return "AMBER", "No threshold set."
        if value < green_min:
            return "RED", "Below minimum."
        if value < (1 + amber_band) * green_min:
            return "AMBER", "Barely above minimum."
        return "GREEN", "Comfortably above minimum."


def rag_text_to_color(rag):
    # match screenshot vibes
    if rag == "GREEN":
        return "#1f7a3a"
    if rag == "AMBER":
        return "#b45309"
    if rag == "RED":
        return "#b42318"
    return "#334155"


def tile_html(title, value, bg):
    return f"""
    <div class="kpi-tile" style="background:{bg};">
        <div class="kpi-label">{title}</div>
        <div class="kpi-value">{value}</div>
    </div>
    """


# -----------------------------
# FILE LOADERS
# -----------------------------
@st.cache_data(show_spinner=False)
def load_kpi_template(excel_bytes_or_path):
    tx = pd.read_excel(excel_bytes_or_path, sheet_name="Transactions")
    tx.columns = [str(c).strip() for c in tx.columns]
    required = ["Date", "Entity", "AccountType", "Category", "Counterparty", "Amount", "CashFlag", "IntercompanyFlag"]
    missing = [c for c in required if c not in tx.columns]
    if missing:
        raise ValueError(f"KPI template Transactions sheet missing columns: {missing}")

    tx["Date"] = pd.to_datetime(tx["Date"], errors="coerce")
    tx["Entity"] = tx["Entity"].astype(str).str.strip()
    tx["AccountType"] = tx["AccountType"].astype(str).str.strip()
    tx["Category"] = tx["Category"].astype(str).str.strip()
    tx["Counterparty"] = tx["Counterparty"].astype(str).str.strip()
    tx["CashFlag"] = tx["CashFlag"].astype(str).str.strip()
    tx["IntercompanyFlag"] = tx["IntercompanyFlag"].astype(str).str.strip()

    tx["Amount"] = pd.to_numeric(tx["Amount"], errors="coerce").fillna(0.0)

    # Enforce sign convention:
    # Revenue (+), all other types (-)
    tx["Amount_norm"] = tx["Amount"].copy()
    is_rev = tx["AccountType"].str.upper().eq("REVENUE")
    tx.loc[is_rev, "Amount_norm"] = tx.loc[is_rev, "Amount"].abs()
    tx.loc[~is_rev, "Amount_norm"] = -tx.loc[~is_rev, "Amount"].abs()

    tx["Month"] = tx["Date"].apply(to_month_ts)
    return tx


@st.cache_data(show_spinner=False)
def load_sde_soe_tables(excel_bytes_or_path):
    def extract_table(sheet_name, header_row_idx):
        raw = pd.read_excel(excel_bytes_or_path, sheet_name=sheet_name, header=None)
        header = raw.iloc[header_row_idx].astype(str).tolist()
        df = raw.iloc[header_row_idx + 1 :].copy()
        df.columns = header
        df = df.dropna(how="all")
        df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", na=False)]
        df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]
        return df

    def find_header_row(sheet_name, needle="Sr. No."):
        raw = pd.read_excel(excel_bytes_or_path, sheet_name=sheet_name, header=None)
        for i in range(len(raw)):
            row = raw.iloc[i].astype(str).tolist()
            if any(needle.lower() in str(x).lower() for x in row):
                return i
        return None

    cd_sheet = "CDOE(26.02.26)"
    comb_sheet = "26.02.26"

    cd_hr = find_header_row(cd_sheet, "Sr. No.")
    comb_hr = find_header_row(comb_sheet, "Sr. No.")

    if cd_hr is None or comb_hr is None:
        raise ValueError("Could not locate header rows in SDE/SOE file.")

    cd = extract_table(cd_sheet, cd_hr)
    comb = extract_table(comb_sheet, comb_hr)

    def clean_year(df):
        if "Year" not in df.columns:
            ycol = next((c for c in df.columns if "year" in c.lower()), None)
            if ycol:
                df = df.rename(columns={ycol: "Year"})
        df["Year"] = df["Year"].astype(str).str.strip()
        return df

    cd = clean_year(cd)
    comb = clean_year(comb)

    for df in [cd, comb]:
        for c in df.columns:
            if c.lower() in {"sr. no.", "sr no", "sr.no."}:
                continue
            if c.lower() == "year":
                continue
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return {"CDOE": cd, "COMBINED": comb}


# -----------------------------
# KPI COMPUTATION
# -----------------------------
def compute_kpi_for_entity_month(tx, entity, month_ts, eliminate_ic=False):
    ms, me = month_start_end(month_ts)
    days = (me - ms).days + 1

    df = tx[(tx["Date"] >= ms) & (tx["Date"] <= me)].copy()
    if entity != "__CONSOLIDATED__":
        df = df[df["Entity"] == entity].copy()

    if eliminate_ic:
        df = df[~df["IntercompanyFlag"].str.upper().eq("YES")].copy()

    def sum_type(t):
        return df.loc[df["AccountType"].str.upper().eq(t.upper()), "Amount_norm"].sum()

    revenue = sum_type("Revenue")
    cogs = sum_type("COGS")
    opex = sum_type("Opex")
    interest = sum_type("Interest")
    tax = sum_type("Tax")

    gross_profit = revenue + cogs
    ebitda = gross_profit + opex
    net_profit = ebitda + interest + tax

    ocf = df.loc[df["CashFlag"].str.upper().eq("CASH"), "Amount_norm"].sum()

    ic_sum = df.loc[df["IntercompanyFlag"].str.upper().eq("YES"), "Amount_norm"].sum()
    ic_pct_rev = safe_div(abs(ic_sum), revenue) if revenue != 0 else np.nan

    exp_ratio = safe_div(abs(cogs) + abs(opex), revenue) if revenue != 0 else np.nan

    gross_margin = safe_div(gross_profit, revenue) if revenue != 0 else np.nan
    ebitda_margin = safe_div(ebitda, revenue) if revenue != 0 else np.nan
    net_margin = safe_div(net_profit, revenue) if revenue != 0 else np.nan

    out = {
        "days_in_month": days,
        "Revenue": revenue,
        "COGS": cogs,
        "Opex": opex,
        "Interest": interest,
        "Tax": tax,
        "Gross Profit": gross_profit,
        "EBITDA": ebitda,
        "Net Profit": net_profit,
        "Operating Cash Flow": ocf,
        "Intercompany Sum": ic_sum,
        "Intercompany % of Revenue": ic_pct_rev,
        "Expense Ratio": exp_ratio,
        "Gross Margin %": gross_margin,
        "EBITDA Margin %": ebitda_margin,
        "Net Margin %": net_margin,
        "df_month": df,
    }
    return out


def monthly_pnl_table(tx, entity, eliminate_ic=False):
    df = tx.copy()
    if entity != "__CONSOLIDATED__":
        df = df[df["Entity"] == entity].copy()
    if eliminate_ic:
        df = df[~df["IntercompanyFlag"].str.upper().eq("YES")].copy()

    months = sorted([m for m in df["Month"].dropna().unique()])
    if not months:
        return pd.DataFrame()

    rows = []
    for m in months:
        k = compute_kpi_for_entity_month(tx, entity, pd.Timestamp(m), eliminate_ic=eliminate_ic)
        rows.append(
            {
                "Month": pd.Timestamp(m),
                "Revenue": k["Revenue"],
                "COGS": k["COGS"],
                "Gross Profit": k["Gross Profit"],
                "Opex": k["Opex"],
                "EBITDA": k["EBITDA"],
                "Interest": k["Interest"],
                "Tax": k["Tax"],
                "Net Profit": k["Net Profit"],
            }
        )
    out = pd.DataFrame(rows).sort_values("Month")
    out["MonthLabel"] = out["Month"].dt.strftime("%b-%Y")
    return out


# -----------------------------
# TOP CONTROL STRIP
# -----------------------------
st.markdown('<div class="control-strip">', unsafe_allow_html=True)
c0, c1, c2, c3, c4, c5 = st.columns([1.4, 1.2, 1.2, 1.2, 1.2, 1.0], vertical_alignment="bottom")

with c0:
    kpi_upload = st.file_uploader("Upload KPI Template (xlsx)", type=["xlsx"], key="kpi_upl_top")
with c1:
    sde_upload = st.file_uploader("Upload SDE/SOE file (xlsx)", type=["xlsx"], key="sde_upl_top")
with c2:
    view_mode = st.selectbox("View Mode", ["Management View (Presentation)", "Detailed View (Operations)"], index=0)
with c3:
    eliminate_ic = st.selectbox("Consolidation Basis", ["Include Intercompany", "Eliminate Intercompany"], index=1)
    eliminate_ic = (eliminate_ic == "Eliminate Intercompany")
with c4:
    revenue_shock = st.slider("Revenue shock %", min_value=-30, max_value=30, value=-5, step=1)
with c5:
    cost_increase = st.slider("Cost increase %", min_value=0, max_value=30, value=10, step=1)

st.markdown('</div>', unsafe_allow_html=True)


# -----------------------------
# LOAD FILES
# -----------------------------
kpi_loaded = False
sde_loaded = False
tx = None
sde_tables = None

try:
    if kpi_upload is None:
        st.info("Upload KPI Template to activate KPIs (Transactions sheet).")
    else:
        tx = load_kpi_template(kpi_upload)
        kpi_loaded = True
except Exception as e:
    st.error(f"KPI template load failed: {e}")

try:
    if sde_upload is not None:
        sde_tables = load_sde_soe_tables(sde_upload)
        sde_loaded = True
except Exception as e:
    st.warning(f"SDE/SOE file not loaded: {e}")


# -----------------------------
# GLOBAL CONTROLS (Entity / Month / WC)
# -----------------------------
if kpi_loaded:
    entities = sorted(tx["Entity"].dropna().unique().tolist())
    months = sorted([m for m in tx["Month"].dropna().unique()])
    month_labels = [pd.Timestamp(m).strftime("%b-%Y") for m in months]
else:
    entities = []
    months = []
    month_labels = []

ctrl_row1 = st.columns([1.2, 1.2, 1, 1, 1], vertical_alignment="bottom")
with ctrl_row1[0]:
    entity_opt = ["Consolidated (All Entities)"] + entities if kpi_loaded else ["Consolidated (All Entities)"]
    sel_entity_label = st.selectbox("Entity", entity_opt, index=0)
    sel_entity = "__CONSOLIDATED__" if sel_entity_label.startswith("Consolidated") else sel_entity_label

with ctrl_row1[1]:
    if kpi_loaded and month_labels:
        sel_month_label = st.selectbox("Month", month_labels, index=len(month_labels) - 1)
        sel_month = pd.Timestamp(months[month_labels.index(sel_month_label)])
    else:
        sel_month = pd.Timestamp.today().replace(day=1)
        st.selectbox("Month", ["(upload KPI file)"], index=0, disabled=True)

with ctrl_row1[2]:
    ar = st.number_input("AR (₹)", min_value=0.0, value=450000.0, step=10000.0, format="%.0f")
with ctrl_row1[3]:
    ap = st.number_input("AP (₹)", min_value=0.0, value=380000.0, step=10000.0, format="%.0f")
with ctrl_row1[4]:
    inv = st.number_input("Inventory (₹)", min_value=0.0, value=250000.0, step=10000.0, format="%.0f")

ctrl_row2 = st.columns([1.0, 1.0, 1.0, 2.0], vertical_alignment="bottom")
with ctrl_row2[0]:
    cash_bal = st.number_input("Cash (₹)", min_value=0.0, value=200000.0, step=10000.0, format="%.0f")
with ctrl_row2[1]:
    min_cash_alert = st.number_input("Min Cash Alert (₹)", min_value=0.0, value=100000.0, step=10000.0, format="%.0f")
with ctrl_row2[2]:
    max_ccc_days = st.number_input("Max CCC Days", min_value=0.0, value=60.0, step=5.0, format="%.0f")
with ctrl_row2[3]:
    with st.expander("Governance thresholds (click to expand)", expanded=False):
        max_ic_pct = st.number_input("Max Inter-company % of Revenue", min_value=0.0, max_value=1.0, value=0.20, step=0.01, format="%.2f")
        max_exp_ratio = st.number_input("Max Expense Ratio", min_value=0.0, max_value=2.0, value=0.75, step=0.01, format="%.2f")
        min_net_margin = st.number_input("Min Net Margin", min_value=-1.0, max_value=1.0, value=0.10, step=0.01, format="%.2f")

st.markdown('<div class="small-note">Tip: choose <b>Consolidated (All Entities)</b> to show group numbers to management.</div>', unsafe_allow_html=True)


# -----------------------------
# COMPUTE SELECTED KPI
# -----------------------------
k = None
dso = dpo = dio = ccc = np.nan
if kpi_loaded:
    k = compute_kpi_for_entity_month(tx, sel_entity, sel_month, eliminate_ic=eliminate_ic)
    dso, dpo, dio, ccc = compute_wc_days(ar, ap, inv, k["Revenue"], k["COGS"], k["days_in_month"])


# =============================================================================
# MANAGEMENT VIEW (Screenshot-style)
# =============================================================================
def render_management_view():
    st.markdown('<div class="mgmt-title">Integrated Financial Performance &amp; Governance Dashboard</div>', unsafe_allow_html=True)

    if not kpi_loaded or k is None:
        st.warning("Upload KPI Template to render the Management View.")
        return

    # KPI Tiles row (like screenshot)
    total_rev = k["Revenue"]
    total_exp = abs(k["COGS"]) + abs(k["Opex"]) + abs(k["Interest"]) + abs(k["Tax"])
    net_profit = k["Net Profit"]
    cash_pos = cash_bal

    # Working capital status (simple rules)
    wc_rag, _ = rag_from_threshold(ccc, green_max=max_ccc_days, higher_is_bad=True)
    wc_text = "Working Capital Healthy" if wc_rag == "GREEN" else ("Working Capital Needs Attention" if wc_rag == "RED" else "Working Capital Watchlist")
    wc_bg = "#2f6f52" if wc_rag == "GREEN" else ("#8a3a10" if wc_rag == "RED" else "#806000")

    tile_cols = st.columns(5)
    tile_cols[0].markdown(tile_html("Total Revenue", inr_compact(total_rev), "#163b73"), unsafe_allow_html=True)
    tile_cols[1].markdown(tile_html("Total Expenses", inr_compact(total_exp), "#c05613"), unsafe_allow_html=True)
    tile_cols[2].markdown(tile_html("Net Profit", inr_compact(net_profit), "#1f7a3a"), unsafe_allow_html=True)
    tile_cols[3].markdown(tile_html("Cash Position", inr_compact(cash_pos), "#0f766e"), unsafe_allow_html=True)
    tile_cols[4].markdown(tile_html(wc_text, "", wc_bg), unsafe_allow_html=True)

    st.markdown(
        '<div class="formula-strip">Revenue — COGS — Operating Expenses — Interest &amp; Taxes = Net Profit</div>',
        unsafe_allow_html=True,
    )

    # 3 panels top row
    left, mid, right = st.columns([1.25, 1.25, 1.5])

    # -------- Revenue Analysis
    with left:
        st.markdown('<div class="panel"><div class="section-header">Revenue Analysis</div><div class="panel-body">', unsafe_allow_html=True)
        pnl = monthly_pnl_table(tx, sel_entity, eliminate_ic=eliminate_ic)
        if pnl.empty:
            st.info("No revenue history found.")
        else:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=pnl["MonthLabel"], y=pnl["Revenue"], name="Revenue"))
            fig.add_trace(go.Scatter(x=pnl["MonthLabel"], y=pnl["Revenue"].rolling(3, min_periods=1).mean(), name="Trend"))
            fig.update_layout(height=280, margin=dict(l=10, r=10, t=10, b=30), showlegend=True)
            st.plotly_chart(fig, use_container_width=True)

            # Growth + target achievement (simple)
            if len(pnl) >= 2:
                cur = pnl.iloc[-1]["Revenue"]
                prev = pnl.iloc[-2]["Revenue"]
                gr = safe_div(cur - prev, abs(prev)) if prev != 0 else np.nan
                st.markdown(f"**Growth Rate:** {pct(gr)}")
            else:
                st.markdown("**Growth Rate:** —")

            # target achievement: assume target = 1.15 * previous month (demo style)
            if len(pnl) >= 2:
                target = pnl.iloc[-2]["Revenue"] * 1.15
                ach = safe_div(pnl.iloc[-1]["Revenue"], target) if target else np.nan
                st.markdown(f"**Target Achievement:** {pct(ach)}")
            else:
                st.markdown("**Target Achievement:** —")

        st.markdown("</div></div>", unsafe_allow_html=True)

    # -------- Cost & Expenditure
    with mid:
        st.markdown('<div class="panel"><div class="section-header">Cost & Expenditure</div><div class="panel-body">', unsafe_allow_html=True)
        dfm = k["df_month"].copy()

        # Expense breakdown donut: use Opex categories (fallback: all non-revenue)
        opex = dfm[dfm["AccountType"].str.upper().eq("OPEX")].copy()
        if opex.empty:
            st.info("No Opex breakdown for selected period.")
        else:
            g = opex.groupby("Category")["Amount_norm"].sum().abs().reset_index()
            g = g.rename(columns={"Amount_norm": "Value"})
            fig_d = px.pie(g, names="Category", values="Value", hole=0.65)
            fig_d.update_layout(height=260, margin=dict(l=10, r=10, t=10, b=10), showlegend=True)
            st.plotly_chart(fig_d, use_container_width=True)

            exp_ratio = k["Expense Ratio"]
            st.markdown(f"✅ **Expense Ratio:** {pct(exp_ratio)}")
            top2 = g.sort_values("Value", ascending=False).head(2)
            st.markdown("✅ **Top Cost Drivers:** " + ", ".join([f"{r.Category} ({inr_compact(r.Value)})" for _, r in top2.iterrows()]))

        # Budget vs Actual (benchmark): use scenario sliders
        budget = total_exp * (1 + max(cost_increase, 0) / 100.0)  # simulated
        actual = total_exp
        bdf = pd.DataFrame({"Type": ["Budget", "Actual"], "Value": [budget, actual]})
        fig_ba = px.bar(bdf, x="Type", y="Value", text="Value", title="Budget vs Actual (Benchmark)")
        fig_ba.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        fig_ba.update_layout(height=280, margin=dict(l=10, r=10, t=40, b=10), showlegend=False)
        st.plotly_chart(fig_ba, use_container_width=True)

        st.markdown("</div></div>", unsafe_allow_html=True)

    # -------- Profitability Metrics
    with right:
        st.markdown('<div class="panel"><div class="section-header">Profitability Metrics</div><div class="panel-body">', unsafe_allow_html=True)

        m1, m2, m3 = st.columns(3)
        m1.metric("Gross Margin", pct(k["Gross Margin %"]), help="Gross Profit / Revenue")
        m2.metric("EBITDA Margin", pct(k["EBITDA Margin %"]), help="EBITDA / Revenue")
        m3.metric("Net Profit Margin", pct(k["Net Margin %"]), help="Net Profit / Revenue")

        pnl = monthly_pnl_table(tx, sel_entity, eliminate_ic=eliminate_ic)
        if pnl.empty:
            st.info("No profit trend data.")
        else:
            fig_np = go.Figure()
            fig_np.add_trace(go.Scatter(x=pnl["MonthLabel"], y=pnl["Net Profit"], mode="lines+markers", name="Net Profit"))
            fig_np.update_layout(height=240, margin=dict(l=10, r=10, t=10, b=30), showlegend=False)
            st.plotly_chart(fig_np, use_container_width=True)

        # Break-even quick status (FIXED: no inline-if syntax error)
        if net_profit >= 0:
            st.success("Operating above break-even for selected period.")
        else:
            st.warning("Below break-even for selected period. Focus on cost + pricing actions.")

        st.markdown("</div></div>", unsafe_allow_html=True)

    # Bottom row: Cash Flow & Liquidity + Sister Concerns + Risk/Scenario
    b1, b2, b3 = st.columns([1.25, 1.25, 1.5])

    # Cash Flow & Liquidity
    with b1:
        st.markdown('<div class="panel"><div class="section-header">Cash Flow & Liquidity</div><div class="panel-body">', unsafe_allow_html=True)
        st.metric("Operating Cash Flow", inr_full(k["Operating Cash Flow"]))
        st.metric("AR Days (DSO)", "—" if np.isnan(dso) else f"{dso:.0f}")
        st.metric("AP Days (DPO)", "—" if np.isnan(dpo) else f"{dpo:.0f}")
        st.metric("Cash Balance", inr_full(cash_bal))

        # mini bars for AR/AP days
        mini = pd.DataFrame({"Metric": ["AR Days", "AP Days"], "Days": [0 if np.isnan(dso) else dso, 0 if np.isnan(dpo) else dpo]})
        fig_m = px.bar(mini, x="Metric", y="Days", text="Days")
        fig_m.update_traces(textposition="outside")
        fig_m.update_layout(height=220, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(fig_m, use_container_width=True)

        st.markdown("</div></div>", unsafe_allow_html=True)

    # Sister Concerns Overview
    with b2:
        st.markdown('<div class="panel"><div class="section-header">Sister Concerns Overview</div><div class="panel-body">', unsafe_allow_html=True)
        dfm = k["df_month"].copy()
        ic = dfm[dfm["IntercompanyFlag"].str.upper().eq("YES")].copy()
        if ic.empty:
            st.info("No inter-company transactions for this period.")
        else:
            g = ic.groupby("Entity")["Amount_norm"].sum().reset_index()
            g["Inter-Transfers"] = g["Amount_norm"].apply(inr_full)
            g = g[["Entity", "Inter-Transfers"]]
            st.dataframe(g, use_container_width=True, hide_index=True)
            st.caption("Tip: switch to Consolidated + Eliminate IC to show true group performance.")
        st.markdown("</div></div>", unsafe_allow_html=True)

    # Risk & Scenario Analysis
    with b3:
        st.markdown('<div class="panel"><div class="section-header">Risk & Scenario Analysis</div><div class="panel-body">', unsafe_allow_html=True)

        # Scenario impact on Net Profit (simple what-if)
        rev2 = total_rev * (1 + revenue_shock / 100.0)
        exp2 = total_exp * (1 + cost_increase / 100.0)
        np2 = rev2 - exp2  # simplified for scenario narration

        s_df = pd.DataFrame(
            {
                "Scenario": ["Base", "What-if"],
                "Revenue": [total_rev, rev2],
                "Expenses": [total_exp, exp2],
                "Net Profit": [net_profit, np2],
            }
        )
        fig_s = px.bar(
            s_df.melt(id_vars=["Scenario"], value_vars=["Revenue", "Expenses", "Net Profit"]),
            x="Scenario",
            y="value",
            color="variable",
            barmode="group",
            text="value",
        )
        fig_s.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
        fig_s.update_layout(height=320, margin=dict(l=10, r=10, t=10, b=10), legend_title_text="")
        st.plotly_chart(fig_s, use_container_width=True)

        runway_months = np.nan
        if cash_bal and (exp2 / 12) > 0:
            runway_months = cash_bal / (exp2 / 12)

        st.markdown(f"**Break-even shift:** {inr_compact(np2 - net_profit)} impact vs base")
        st.markdown(f"**Cash runway:** {'—' if np.isnan(runway_months) else f'{runway_months:.1f} months'}")

        st.markdown("</div></div>", unsafe_allow_html=True)


# =============================================================================
# DETAILED VIEW (Operations)
# =============================================================================
def render_detailed_view():
    tabs = st.tabs(
        [
            "🏠 Home / Model Map",
            "📊 KPI Template Dashboard",
            "🏛 Governance (RAG + Improvements)",
            "🧾 Consolidation (Group View)",
            "📈 KPI Charts + Narration",
            "🎓 SDE/SOE Dashboard",
        ]
    )

    # Home / model map
    with tabs[0]:
        st.subheader("Overview / Model Map")
        shown = False
        for p in MODEL_MAP_CANDIDATES:
            if p.exists():
                st.image(str(p), use_container_width=True)
                shown = True
                break
        if not shown:
            st.info("Add model map image to ./assets/model_map.png (or overview.png) and redeploy.")

        c1, c2 = st.columns(2)
        c1.metric("KPI Template Loaded", "YES" if kpi_loaded else "NO")
        c2.metric("SDE/SOE Loaded", "YES" if sde_loaded else "NO")

    # KPI Template dashboard (table-first)
    with tabs[1]:
        st.subheader("KPI Template Dashboard (Ledger-driven)")
        if not kpi_loaded or k is None:
            st.warning("Upload KPI template file to enable this section.")
        else:
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total Revenue", inr_full(k["Revenue"]))
            c2.metric("COGS", inr_full(k["COGS"]))
            c3.metric("Opex", inr_full(k["Opex"]))
            c4.metric("Net Profit", inr_full(k["Net Profit"]))
            c5.metric("Expense Ratio", pct(k["Expense Ratio"]))
            c6.metric("Inter-company %", pct(k["Intercompany % of Revenue"]))

            st.caption("Sign convention enforced: Revenue (+); Costs/Interest/Tax negative; Profit is direct sum.")

            sub = st.tabs(["Transactions", "P&L (Monthly)", "Working Capital", "Raw KPI Table"])
            with sub[0]:
                dfm = k["df_month"].copy().sort_values("Date")
                st.dataframe(dfm, use_container_width=True, hide_index=True)

            with sub[1]:
                pnl = monthly_pnl_table(tx, sel_entity, eliminate_ic=eliminate_ic)
                if pnl.empty:
                    st.info("No monthly history available.")
                else:
                    show = pnl[["MonthLabel", "Revenue", "COGS", "Opex", "Net Profit"]].copy()
                    for col in ["Revenue", "COGS", "Opex", "Net Profit"]:
                        show[col] = show[col].apply(inr_full)
                    st.dataframe(show, use_container_width=True, hide_index=True)

            with sub[2]:
                st.metric("DSO (AR Days)", "—" if np.isnan(dso) else f"{dso:.1f}")
                st.metric("DPO (AP Days)", "—" if np.isnan(dpo) else f"{dpo:.1f}")
                st.metric("DIO (Inv Days)", "—" if np.isnan(dio) else f"{dio:.1f}")
                st.metric("CCC (Days)", "—" if np.isnan(ccc) else f"{ccc:.1f}")

            with sub[3]:
                t = pd.DataFrame(
                    [
                        ("Gross Margin %", pct(k["Gross Margin %"])),
                        ("EBITDA Margin %", pct(k["EBITDA Margin %"])),
                        ("Net Margin %", pct(k["Net Margin %"])),
                        ("Expense Ratio", pct(k["Expense Ratio"])),
                        ("IC % of Revenue", pct(k["Intercompany % of Revenue"])),
                    ],
                    columns=["KPI", "Value"],
                )
                st.dataframe(t, use_container_width=True, hide_index=True)

    # Governance
    with tabs[2]:
        st.subheader("Governance (RAG + Actions)")
        if not kpi_loaded or k is None:
            st.warning("Upload KPI template to enable governance checks.")
        else:
            exp_ratio = k["Expense Ratio"]
            net_margin = k["Net Margin %"]
            ic_pct = k["Intercompany % of Revenue"]

            rag_exp, _ = rag_from_threshold(exp_ratio, green_max=max_exp_ratio, higher_is_bad=True)
            rag_ic, _ = rag_from_threshold(ic_pct, green_max=max_ic_pct, higher_is_bad=True)
            rag_cash, _ = rag_from_threshold(cash_bal, green_min=min_cash_alert, higher_is_bad=False)
            rag_nm, _ = rag_from_threshold(net_margin, green_min=min_net_margin, higher_is_bad=False)
            rag_ccc, _ = rag_from_threshold(ccc, green_max=max_ccc_days, higher_is_bad=True)

            checks = [
                ("Expense Ratio", pct(exp_ratio), f"Max {pct(max_exp_ratio)}", rag_exp),
                ("Inter-company %", pct(ic_pct), f"Max {pct(max_ic_pct)}", rag_ic),
                ("Cash Balance", inr_full(cash_bal), f"Min {inr_full(min_cash_alert)}", rag_cash),
                ("Net Margin", pct(net_margin), f"Min {pct(min_net_margin)}", rag_nm),
                ("CCC Days", "—" if np.isnan(ccc) else f"{ccc:.1f}", f"Max {max_ccc_days:.0f}", rag_ccc),
            ]
            df = pd.DataFrame(checks, columns=["Check", "Current", "Threshold", "RAG"])
            st.dataframe(df, use_container_width=True, hide_index=True)

            st.markdown("### Recommended actions")
            if rag_exp in ("RED", "AMBER"):
                st.markdown("- Tighten category-wise budgets, freeze discretionary spend, renegotiate vendors.")
            if rag_nm in ("RED", "AMBER"):
                st.markdown("- Improve pricing / fee realization; reduce low-margin activities; cut overheads.")
            if rag_cash in ("RED", "AMBER"):
                st.markdown("- Accelerate collections; defer non-critical payments; weekly cash forecast cadence.")
            if rag_ic in ("RED", "AMBER"):
                st.markdown("- Reconcile intercompany weekly; approvals for IC transfers; enforce IC caps.")
            if rag_ccc in ("RED", "AMBER"):
                st.markdown("- Reduce AR days, optimize inventory, negotiate longer AP terms.")

    # Consolidation
    with tabs[3]:
        st.subheader("Consolidation (Group View)")
        if not kpi_loaded:
            st.warning("Upload KPI template to enable consolidation.")
        else:
            kc = compute_kpi_for_entity_month(tx, "__CONSOLIDATED__", sel_month, eliminate_ic=eliminate_ic)
            st.markdown(f"**Month:** {sel_month.strftime('%b-%Y')} | **IC:** {'Eliminated' if eliminate_ic else 'Included'}")

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Revenue", inr_full(kc["Revenue"]))
            c2.metric("Total Expenses", inr_full(abs(kc["COGS"]) + abs(kc["Opex"]) + abs(kc["Interest"]) + abs(kc["Tax"])))
            c3.metric("Net Profit", inr_full(kc["Net Profit"]))
            c4.metric("Expense Ratio", pct(kc["Expense Ratio"]))

            pnl = monthly_pnl_table(tx, "__CONSOLIDATED__", eliminate_ic=eliminate_ic)
            if not pnl.empty:
                fig = px.line(pnl, x="MonthLabel", y=["Revenue", "Net Profit"])
                fig.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=30))
                st.plotly_chart(fig, use_container_width=True)

    # KPI charts
    with tabs[4]:
        st.subheader("KPI Charts + Narration")
        if not kpi_loaded:
            st.warning("Upload KPI template to enable charts.")
        else:
            pnl = monthly_pnl_table(tx, sel_entity, eliminate_ic=eliminate_ic)
            if pnl.empty:
                st.info("No data for charts.")
            else:
                melt = pnl.melt(id_vars=["MonthLabel"], value_vars=["Revenue", "Opex", "Net Profit"], var_name="Metric", value_name="Value")
                melt.loc[melt["Metric"].str.upper().eq("OPEX"), "Value"] = melt.loc[melt["Metric"].str.upper().eq("OPEX"), "Value"].abs()
                fig = px.bar(melt, x="MonthLabel", y="Value", color="Metric", barmode="group", text="Value")
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
                fig.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=40), legend_title_text="")
                st.plotly_chart(fig, use_container_width=True)

    # SDE/SOE
    with tabs[5]:
        st.subheader("SDE/SOE Dashboard")
        if not sde_loaded:
            st.info("Upload SDE/SOE file to enable this section.")
        else:
            cd = sde_tables["CDOE"].copy()
            rename_map = {}
            for c in cd.columns:
                cl = str(c).lower()
                if "enrollment" in cl:
                    rename_map[c] = "Enrollment"
                elif cl == "income":
                    rename_map[c] = "Income"
                elif "expenditure" in cl:
                    rename_map[c] = "Expenditure"
                elif "surplus" in cl:
                    rename_map[c] = "Surplus"
                elif "transferred" in cl or "transfer" in cl:
                    rename_map[c] = "Transfer_to_BVDU"
            cd = cd.rename(columns=rename_map)

            keep = ["Year", "Enrollment", "Income", "Expenditure", "Surplus", "Transfer_to_BVDU"]
            for c in keep:
                if c not in cd.columns:
                    cd[c] = np.nan
            cd2 = cd[keep].dropna(subset=["Year"]).copy()

            st.dataframe(cd2, use_container_width=True, hide_index=True)

            df_ie = cd2.melt(id_vars=["Year"], value_vars=["Income", "Expenditure"], var_name="Metric", value_name="Value")
            fig_ie = px.bar(df_ie, x="Year", y="Value", color="Metric", barmode="group", text="Value")
            fig_ie.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
            fig_ie.update_layout(height=420, margin=dict(l=10, r=10, t=10, b=30), legend_title_text="")
            st.plotly_chart(fig_ie, use_container_width=True)


# -----------------------------
# ROUTING (two dashboards in one app)
# -----------------------------
if view_mode.startswith("Management"):
    render_management_view()
else:
    render_detailed_view()