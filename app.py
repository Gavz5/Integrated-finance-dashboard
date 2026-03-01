# app.py
# =========================================================
# AMFGIE | Adaptive Multi-Entity Financial Governance Intelligence Engine
# WHITE / POWER BI-LIKE THEME + CLEAN DONUT + GAUGES + ANOMALY CARDS
# + FORECAST & PREDICTIONS (Revenue Forecast + Expense Projection)
#
# Run:
#   pip install streamlit pandas openpyxl plotly numpy
#   streamlit run app.py
# =========================================================

import os
from calendar import monthrange
from dataclasses import dataclass
from typing import Dict, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# -----------------------------
# PAGE CONFIG
# -----------------------------
st.set_page_config(
    page_title="AMFGIE | Multi-Entity Financial Governance Dashboard",
    layout="wide",
)

INR_SYMBOL = "₹"
DEFAULT_KPI_PATH = "Business_Financial_KPI_Dashboard_Template.xlsx"
DEFAULT_SDE_PATH = "SDE & SOE Graph (2).xlsx"

# -----------------------------
# POWER BI-LIKE PALETTE (LIGHT)
# -----------------------------
PBI = {
    "blue": "#118DFF",
    "dark_blue": "#0F3D91",
    "orange": "#E66C37",
    "purple": "#744EC2",
    "magenta": "#E044A7",
    "teal": "#1BA39C",
    "green": "#2CA02C",
    "red": "#D64550",
    "yellow": "#F2C80F",
    "gray": "#6B7280",
    "bg": "#F5F7FB",
    "card": "#FFFFFF",
    "stroke": "rgba(15, 23, 42, 0.10)",
    "stroke2": "rgba(17, 141, 255, 0.18)",
    "text": "#0F172A",
    "muted": "#64748B",
}
PBI_SEQ = [PBI["blue"], PBI["orange"], PBI["purple"], PBI["magenta"], PBI["teal"], PBI["yellow"], PBI["red"], PBI["green"]]


# -----------------------------
# WHITE THEME CSS (POWER BI LOOK)
# -----------------------------
def inject_white_theme():
    st.markdown(
        f"""
        <style>
        html, body, [data-testid="stAppViewContainer"] {{
            background: {PBI["bg"]};
            color: {PBI["text"]};
        }}
        header[data-testid="stHeader"] {{ background: transparent; }}
        .block-container {{
            padding-top: 1.0rem;
            padding-bottom: 4.4rem; /* space for bottom bar */
        }}

        .hero {{
            background: linear-gradient(90deg, rgba(17,141,255,0.14), rgba(116,78,194,0.12));
            border: 1px solid {PBI["stroke2"]};
            border-radius: 18px;
            padding: 16px 18px;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.10);
        }}
        .hero h1 {{
            margin: 0;
            font-size: 28px;
            font-weight: 900;
            letter-spacing: 0.3px;
            color: {PBI["text"]};
        }}
        .hero .sub {{
            margin-top: 6px;
            font-size: 13px;
            color: {PBI["muted"]};
            font-weight: 600;
        }}

        .control-strip {{
            margin-top: 12px;
            background: {PBI["card"]};
            border: 1px solid {PBI["stroke"]};
            border-radius: 18px;
            padding: 14px 14px 6px 14px;
            box-shadow: 0 10px 25px rgba(15, 23, 42, 0.08);
        }}

        [data-testid="stTabs"] button {{
            color: {PBI["muted"]} !important;
            font-weight: 800 !important;
        }}
        [data-testid="stTabs"] button[aria-selected="true"] {{
            color: {PBI["text"]} !important;
            border-bottom: 3px solid {PBI["blue"]} !important;
        }}

        .section-title {{
            margin-top: 14px;
            padding: 10px 12px;
            border-radius: 14px;
            background: linear-gradient(90deg, rgba(17,141,255,0.16), rgba(255,255,255,1));
            border: 1px solid {PBI["stroke2"]};
            font-weight: 900;
            letter-spacing: 0.2px;
            color: {PBI["text"]};
        }}

        .panel {{
            background: {PBI["card"]};
            border: 1px solid {PBI["stroke"]};
            border-radius: 18px;
            padding: 14px 14px 10px 14px;
            box-shadow: 0 10px 26px rgba(15, 23, 42, 0.08);
        }}
        .panel h3 {{
            margin: 0 0 8px 0;
            font-size: 16px;
            color: {PBI["text"]};
            font-weight: 900;
        }}
        .muted {{ color: {PBI["muted"]}; }}

        .pill-row {{ display:flex; gap:10px; flex-wrap:wrap; margin-top:12px; }}
        .pill {{
            flex: 1 1 180px;
            min-width: 180px;
            background: {PBI["card"]};
            border: 1px solid {PBI["stroke"]};
            border-radius: 16px;
            padding: 12px 14px;
            box-shadow: 0 10px 22px rgba(15, 23, 42, 0.07);
        }}
        .pill .k {{ color: {PBI["muted"]}; font-size: 12px; font-weight: 900; }}
        .pill .v {{ font-size: 22px; font-weight: 950; margin-top: 4px; color: {PBI["text"]}; }}
        .pill .s {{ color: {PBI["muted"]}; font-size: 12px; margin-top: 6px; }}

        .rag {{
            display:inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            font-weight: 950;
            font-size: 12px;
            margin-right: 8px;
            border: 1px solid rgba(15,23,42,0.12);
        }}
        .rag.green{{ background: rgba(44,160,44,0.14); color: #0B3D0B; }}
        .rag.amber{{ background: rgba(242,200,15,0.16); color: #5A4A00; }}
        .rag.red{{ background: rgba(214,69,80,0.14); color: #5A1016; }}

        [data-testid="stDataFrame"] {{
            border: 1px solid {PBI["stroke"]};
            border-radius: 14px;
            overflow: hidden;
        }}

        .an-item {{
            display:flex;
            justify-content: space-between;
            gap: 12px;
            padding: 10px 12px;
            border-radius: 14px;
            border: 1px solid rgba(15,23,42,0.08);
            background: rgba(255,255,255,0.80);
            margin-bottom: 10px;
        }}
        .an-left {{
            display:flex;
            gap: 10px;
            align-items: center;
            font-weight: 900;
            color: {PBI["text"]};
        }}
        .an-dot {{
            width: 10px; height: 10px; border-radius: 99px;
            background: {PBI["orange"]};
            box-shadow: 0 0 0 3px rgba(230,108,55,0.16);
        }}
        .an-right {{
            font-weight: 950;
            color: {PBI["red"]};
        }}

        .bottom-nav {{
            position: fixed;
            left: 0; right: 0; bottom: 0;
            padding: 10px 14px;
            background: rgba(255,255,255,0.92);
            border-top: 1px solid rgba(15,23,42,0.10);
            backdrop-filter: blur(8px);
            z-index: 9999;
        }}
        .bottom-nav-inner {{
            max-width: 1200px;
            margin: 0 auto;
            display:flex;
            gap: 10px;
            justify-content: center;
            align-items: center;
            flex-wrap: wrap;
        }}
        .nav-pill {{
            border: 1px solid rgba(15,23,42,0.10);
            background: white;
            border-radius: 999px;
            padding: 8px 12px;
            font-weight: 900;
            color: {PBI["text"]};
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.07);
            font-size: 12px;
        }}
        .nav-pill.active {{
            border-color: rgba(17,141,255,0.40);
            box-shadow: 0 10px 22px rgba(17,141,255,0.12);
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

inject_white_theme()


# -----------------------------
# HELPERS
# -----------------------------
def inr(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        x = float(x)
    except Exception:
        return str(x)
    sign = "-" if x < 0 else ""
    x = abs(x)
    return f"{sign}{INR_SYMBOL}{x:,.0f}"

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

def to_month_ts(x):
    t = pd.to_datetime(x, errors="coerce")
    if pd.isna(t):
        return pd.NaT
    return pd.Timestamp(t.year, t.month, 1)

def month_start_end(month_ts: pd.Timestamp):
    ms = pd.Timestamp(month_ts.year, month_ts.month, 1)
    me = pd.Timestamp(month_ts.year, month_ts.month, monthrange(month_ts.year, month_ts.month)[1])
    return ms, me

def rag_label(value, green_max=None, green_min=None, amber_band=0.10, higher_is_bad=True):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "AMBER", "Insufficient data."
    if higher_is_bad:
        if green_max is None:
            return "AMBER", "Threshold not set."
        if value > green_max:
            return "RED", "Breached limit."
        if value > (1 - amber_band) * green_max:
            return "AMBER", "Near limit."
        return "GREEN", "Within range."
    else:
        if green_min is None:
            return "AMBER", "Threshold not set."
        if value < green_min:
            return "RED", "Below minimum."
        if value < (1 + amber_band) * green_min:
            return "AMBER", "Barely above minimum."
        return "GREEN", "Above minimum."

def rag_badge_html(rag: str):
    r = rag.upper()
    cls = "amber"
    if r == "GREEN":
        cls = "green"
    elif r == "RED":
        cls = "red"
    return f'<span class="rag {cls}">{r}</span>'


# -----------------------------
# PLOTLY LIGHT LAYOUT (NO "undefined")
# -----------------------------
def plotly_light_layout(fig, height=320, title_text=""):
    fig.update_layout(
        height=height,
        paper_bgcolor="white",
        plot_bgcolor="white",
        colorway=PBI_SEQ,
        font=dict(color=PBI["text"]),
        margin=dict(l=14, r=14, t=56, b=14),

        # hard-fix title
        title=dict(text=title_text or "", x=0.0, xanchor="left", font=dict(color=PBI["text"])),

        legend=dict(
            bgcolor="rgba(255,255,255,0.85)",
            bordercolor="rgba(15,23,42,0.10)",
            borderwidth=1,
            orientation="h",
            yanchor="bottom",
            y=1.10,
            xanchor="left",
            x=0.0,
            font=dict(color=PBI["text"]),
        ),
        xaxis=dict(
            title="",
            gridcolor="rgba(15,23,42,0.08)",
            zerolinecolor="rgba(15,23,42,0.10)",
            tickfont=dict(color=PBI["text"]),
        ),
        yaxis=dict(
            title="",
            gridcolor="rgba(15,23,42,0.08)",
            zerolinecolor="rgba(15,23,42,0.10)",
            tickfont=dict(color=PBI["text"]),
        ),
    )
    return fig


# -----------------------------
# GAUGE
# -----------------------------
def gauge_percent(title: str, value: float, color: str, height: int = 220):
    value = float(np.clip(value, 0, 100))
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=value,
        number={"suffix": "%", "font": {"size": 44, "color": PBI["text"]}},
        title={"text": title, "font": {"size": 14, "color": PBI["muted"]}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1, "tickcolor": "rgba(15,23,42,0.35)"},
            "bar": {"color": color, "thickness": 0.35},
            "bgcolor": "white",
            "borderwidth": 1,
            "bordercolor": "rgba(15,23,42,0.10)",
            "steps": [
                {"range": [0, 50], "color": "rgba(44,160,44,0.12)"},
                {"range": [50, 75], "color": "rgba(242,200,15,0.14)"},
                {"range": [75, 100], "color": "rgba(214,69,80,0.14)"},
            ],
            "threshold": {"line": {"color": "rgba(15,23,42,0.35)", "width": 3}, "thickness": 0.75, "value": value}
        }
    ))
    fig.update_layout(height=height, margin=dict(l=12, r=12, t=40, b=10), paper_bgcolor="white")
    return fig


# -----------------------------
# ANOMALY COUNTS (simple)
# -----------------------------
def anomaly_counts(df_month: pd.DataFrame) -> Dict[str, int]:
    out = {
        "Fee Discrepancy": 0,
        "Fake Invigilator Entry": 0,
        "Unclaimed Certificates": 0,
        "Inter-Entity Transfer": 0,
    }
    if df_month is None or df_month.empty:
        return out

    rev = df_month[df_month["AccountType"].str.upper().eq("REVENUE")].copy()
    if not rev.empty:
        out["Fee Discrepancy"] = int((pd.to_numeric(rev["Amount"], errors="coerce").fillna(0) < 0).sum())

    opex = df_month[df_month["AccountType"].str.upper().eq("OPEX")].copy()
    if not opex.empty:
        cp = opex["Counterparty"].astype(str).str.strip().str.lower()
        suspicious = cp.str.contains("invigilator|faculty|staff|contractor", regex=True, na=False)
        missing = cp.eq("") | cp.eq("nan")
        out["Fake Invigilator Entry"] = int((suspicious & missing).sum())

    cat = df_month["Category"].astype(str).str.lower()
    cp2 = df_month["Counterparty"].astype(str).str.lower()
    out["Unclaimed Certificates"] = int((cat.str.contains("certificate|unclaimed|refund", na=False) | cp2.str.contains("certificate|unclaimed|refund", na=False)).sum())

    out["Inter-Entity Transfer"] = int(df_month["IntercompanyFlag"].astype(str).str.upper().eq("YES").sum())
    return out


# -----------------------------
# FUND FLOW (CLEAN DONUT)
# -----------------------------
def fund_flow_donut(k_revenue: float, k_opex: float, k_cogs: float, k_interest: float, k_tax: float):
    revenue = float(k_revenue) if not np.isnan(k_revenue) else 0.0
    out_cogs = abs(float(k_cogs)) if not np.isnan(k_cogs) else 0.0
    out_opex = abs(float(k_opex)) if not np.isnan(k_opex) else 0.0
    out_int = abs(float(k_interest)) if not np.isnan(k_interest) else 0.0
    out_tax = abs(float(k_tax)) if not np.isnan(k_tax) else 0.0

    out_total = out_cogs + out_opex + out_int + out_tax
    remaining = max(revenue - out_total, 0)

    df = pd.DataFrame({
        "Bucket": ["COGS", "Opex", "Interest", "Tax", "Remaining"],
        "Value": [out_cogs, out_opex, out_int, out_tax, remaining]
    })

    df = df[df["Value"] > 0].copy()
    if df.empty:
        df = pd.DataFrame({"Bucket": ["No Data"], "Value": [1]})

    fig = px.pie(
        df,
        names="Bucket",
        values="Value",
        hole=0.68,
        title="",
        color="Bucket",
        color_discrete_map={
            "COGS": PBI["orange"],
            "Opex": PBI["magenta"],
            "Interest": PBI["purple"],
            "Tax": PBI["yellow"],
            "Remaining": PBI["green"],
            "No Data": PBI["gray"],
        }
    )
    fig.update_traces(
        textinfo="label+percent",
        textposition="inside",
        insidetextorientation="radial",
        hovertemplate="%{label}<br>%{value:,.0f}<br>%{percent}<extra></extra>"
    )

    fig.update_layout(
        margin=dict(l=10, r=10, t=18, b=40),
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.10,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0.90)",
            bordercolor="rgba(15,23,42,0.10)",
            borderwidth=1,
        ),
        paper_bgcolor="white",
    )
    fig = plotly_light_layout(fig, height=330, title_text="")
    fig.update_layout(margin=dict(l=10, r=10, t=18, b=48))
    return fig


# -----------------------------
# FORECAST HELPERS
# -----------------------------
def next_months(start_month: pd.Timestamp, n: int) -> list:
    months = []
    cur = pd.Timestamp(start_month.year, start_month.month, 1)
    for _ in range(n):
        cur = (cur + pd.offsets.MonthBegin(1))
        months.append(pd.Timestamp(cur.year, cur.month, 1))
    return months

def simple_forecast(series: pd.Series, n: int = 6):
    """
    Very stable forecast:
    - Use last 6 points
    - Calculate avg growth rate (bounded)
    - Apply forward
    """
    y = pd.to_numeric(series, errors="coerce").fillna(0).values
    y = y[-6:] if len(y) >= 6 else y
    if len(y) < 2:
        growth = 0.05
        last = float(y[-1]) if len(y) else 0.0
    else:
        diffs = []
        for i in range(1, len(y)):
            prev = y[i-1]
            cur = y[i]
            if prev == 0:
                continue
            diffs.append((cur - prev) / abs(prev))
        growth = float(np.nanmean(diffs)) if diffs else 0.05
        growth = float(np.clip(growth, -0.10, 0.25))
        last = float(y[-1])

    out = []
    cur = last
    for _ in range(n):
        cur = max(cur * (1 + growth), 0)
        out.append(cur)
    return out

def forecast_tables(monthly_df: pd.DataFrame, n: int = 6):
    if monthly_df is None or monthly_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    monthly_df = monthly_df.sort_values("Month")
    last_month = pd.Timestamp(monthly_df["Month"].iloc[-1])
    f_months = next_months(last_month, n)
    labels = [m.strftime("%b-%Y") for m in f_months]

    rev_fc = simple_forecast(monthly_df["Revenue"], n=n)
    exp_fc = simple_forecast(monthly_df["Expenses"], n=n)

    df_rev = pd.DataFrame({"Month": labels, "Revenue Forecast": rev_fc})
    df_exp = pd.DataFrame({"Month": labels, "Expense Projection": exp_fc})
    return df_rev, df_exp


# -----------------------------
# DATA LOADERS
# -----------------------------
@st.cache_data(show_spinner=False)
def load_kpi_transactions(excel_src) -> pd.DataFrame:
    tx = pd.read_excel(excel_src, sheet_name="Transactions")
    tx.columns = [str(c).strip() for c in tx.columns]

    required = ["Date", "Entity", "AccountType", "Category", "Counterparty", "Amount", "CashFlag", "IntercompanyFlag"]
    missing = [c for c in required if c not in tx.columns]
    if missing:
        raise ValueError(f"Transactions sheet missing columns: {missing}")

    tx["Date"] = pd.to_datetime(tx["Date"], errors="coerce")
    tx["Entity"] = tx["Entity"].astype(str).str.strip()
    tx["AccountType"] = tx["AccountType"].astype(str).str.strip()
    tx["Category"] = tx["Category"].astype(str).str.strip()
    tx["Counterparty"] = tx["Counterparty"].astype(str).str.strip()
    tx["CashFlag"] = tx["CashFlag"].astype(str).str.strip()
    tx["IntercompanyFlag"] = tx["IntercompanyFlag"].astype(str).str.strip()
    tx["Amount"] = pd.to_numeric(tx["Amount"], errors="coerce").fillna(0.0)

    is_rev = tx["AccountType"].str.upper().eq("REVENUE")
    tx["Amount_norm"] = tx["Amount"].copy()
    tx.loc[is_rev, "Amount_norm"] = tx.loc[is_rev, "Amount"].abs()
    tx.loc[~is_rev, "Amount_norm"] = -tx.loc[~is_rev, "Amount"].abs()

    tx["Month"] = tx["Date"].apply(to_month_ts)
    return tx

@st.cache_data(show_spinner=False)
def load_sde_tables(excel_src) -> Dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(excel_src)
    out = {}
    for sh in xl.sheet_names:
        try:
            df = pd.read_excel(excel_src, sheet_name=sh)
            if df.shape[0] > 0:
                out[sh] = df
        except Exception:
            pass
    return out


# -----------------------------
# KPI COMPUTATION
# -----------------------------
@dataclass
class KPIResult:
    Revenue: float
    COGS: float
    Opex: float
    Interest: float
    Tax: float
    GrossProfit: float
    EBITDA: float
    NetProfit: float
    OCF: float
    ExpenseRatio: float
    GrossMargin: float
    EBITDAMargin: float
    NetMargin: float
    IntercompanySum: float
    IntercompanyPct: float
    df_month: pd.DataFrame
    days_in_month: int

def compute_kpis(tx: pd.DataFrame, month_ts: pd.Timestamp, entity: Optional[str], consolidated: bool, eliminate_ic: bool) -> KPIResult:
    ms, me = month_start_end(month_ts)
    days = (me - ms).days + 1

    df = tx[(tx["Date"] >= ms) & (tx["Date"] <= me)].copy()
    if (not consolidated) and entity:
        df = df[df["Entity"] == entity].copy()

    if consolidated and eliminate_ic:
        df = df[~df["IntercompanyFlag"].str.upper().eq("YES")].copy()

    def sum_type(t):
        return df.loc[df["AccountType"].str.upper().eq(t.upper()), "Amount_norm"].sum()

    revenue = sum_type("Revenue")
    cogs = sum_type("COGS")
    opex = sum_type("Opex")
    interest = sum_type("Interest")
    tax = sum_type("Tax")

    gp = revenue + cogs
    ebitda = gp + opex
    npf = ebitda + interest + tax

    ocf = df.loc[df["CashFlag"].str.upper().eq("CASH"), "Amount_norm"].sum()

    ic_sum = df.loc[df["IntercompanyFlag"].str.upper().eq("YES"), "Amount_norm"].sum()
    ic_pct = safe_div(abs(ic_sum), revenue) if revenue != 0 else np.nan

    exp_ratio = safe_div(abs(cogs) + abs(opex), revenue) if revenue != 0 else np.nan
    gm = safe_div(gp, revenue) if revenue != 0 else np.nan
    em = safe_div(ebitda, revenue) if revenue != 0 else np.nan
    nm = safe_div(npf, revenue) if revenue != 0 else np.nan

    return KPIResult(
        Revenue=revenue, COGS=cogs, Opex=opex, Interest=interest, Tax=tax,
        GrossProfit=gp, EBITDA=ebitda, NetProfit=npf, OCF=ocf,
        ExpenseRatio=exp_ratio, GrossMargin=gm, EBITDAMargin=em, NetMargin=nm,
        IntercompanySum=ic_sum, IntercompanyPct=ic_pct,
        df_month=df, days_in_month=days
    )

def compute_wc_days(ar, ap, inv, revenue, cogs, days_in_month):
    dso = safe_div(ar, revenue) * days_in_month if revenue else np.nan
    dpo = safe_div(ap, abs(cogs)) * days_in_month if cogs else np.nan
    dio = safe_div(inv, abs(cogs)) * days_in_month if cogs else np.nan
    if np.isnan(dso) and np.isnan(dpo) and np.isnan(dio):
        return np.nan, np.nan, np.nan, np.nan
    ccc = (0 if np.isnan(dso) else dso) + (0 if np.isnan(dio) else dio) - (0 if np.isnan(dpo) else dpo)
    return dso, dpo, dio, ccc

def monthly_series(tx, consolidated: bool, entity: Optional[str], eliminate_ic: bool):
    df = tx.copy()
    if (not consolidated) and entity:
        df = df[df["Entity"] == entity].copy()

    months = sorted([m for m in df["Month"].dropna().unique()])
    rows = []
    for m in months:
        k = compute_kpis(df, pd.Timestamp(m), entity, consolidated, eliminate_ic)
        rows.append({
            "Month": pd.Timestamp(m),
            "Revenue": k.Revenue,
            "Expenses": abs(k.COGS) + abs(k.Opex),
            "Net Profit": k.NetProfit,
            "OCF": k.OCF,
        })
    out = pd.DataFrame(rows).sort_values("Month")
    if out.empty:
        return out
    out["MonthLabel"] = out["Month"].dt.strftime("%b-%Y")
    return out


# =========================================================
# HERO
# =========================================================
st.markdown(
    """
    <div class="hero">
      <h1>ADAPTIVE MULTI-ENTITY FINANCIAL GOVERNANCE INTELLIGENCE ENGINE (AMFGIE)</h1>
      <div class="sub">Automated • Predictive • Compliant • Multi-Entity Integration</div>
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# CONTROL STRIP
# =========================================================
st.markdown('<div class="control-strip">', unsafe_allow_html=True)

c0, c1, c2, c3, c4, c5 = st.columns([1.4, 1.4, 1.0, 1.0, 1.0, 1.2])

with c0:
    kpi_upload = st.file_uploader("KPI Template (xlsx)", type=["xlsx"], key="kpi_upl_top")
with c1:
    sde_upload = st.file_uploader("SDE/SOE (xlsx)", type=["xlsx"], key="sde_upl_top")

tx = None
kpi_loaded = False
kpi_src = kpi_upload if kpi_upload is not None else (DEFAULT_KPI_PATH if os.path.exists(DEFAULT_KPI_PATH) else None)
if kpi_src is not None:
    try:
        tx = load_kpi_transactions(kpi_src)
        kpi_loaded = True
    except Exception as e:
        st.error(f"KPI load failed: {e}")

sde_tables = None
sde_loaded = False
sde_src = sde_upload if sde_upload is not None else (DEFAULT_SDE_PATH if os.path.exists(DEFAULT_SDE_PATH) else None)
if sde_src is not None:
    try:
        sde_tables = load_sde_tables(sde_src)
        sde_loaded = True
    except Exception as e:
        st.error(f"SDE/SOE load failed: {e}")

view_mode = "Entity View"
entity = None
month_ts = None
eliminate_ic = True

with c2:
    view_mode = st.selectbox("View Mode", ["Entity View", "Consolidated View"], index=0)

entities = []
months = []
if kpi_loaded:
    entities = sorted(tx["Entity"].dropna().unique().tolist())
    months = sorted([m for m in tx["Month"].dropna().unique()])

with c3:
    if view_mode == "Entity View":
        entity = st.selectbox("Entity", entities, index=0 if entities else 0, disabled=(not kpi_loaded or not entities))
    else:
        st.selectbox("Entity", ["All Entities"], index=0, disabled=True)

with c4:
    if months:
        month_labels = [pd.Timestamp(m).strftime("%b-%Y") for m in months]
        sel = st.selectbox("Month", month_labels, index=len(month_labels) - 1)
        month_ts = pd.Timestamp(months[month_labels.index(sel)])
    else:
        st.selectbox("Month", ["—"], index=0, disabled=True)

with c5:
    eliminate_ic = st.toggle("Eliminate Inter-company", value=True, help="Recommended for consolidated view")

st.markdown("</div>", unsafe_allow_html=True)

# Controls row (numbers + thresholds)
cA, cB, cC, cD, cE, cF = st.columns([1, 1, 1, 1, 1, 1.2])
with cA:
    ar = st.number_input("AR (₹)", min_value=0.0, value=450000.0, step=10000.0, format="%.0f")
with cB:
    ap = st.number_input("AP (₹)", min_value=0.0, value=380000.0, step=10000.0, format="%.0f")
with cC:
    inv = st.number_input("Inventory (₹)", min_value=0.0, value=250000.0, step=10000.0, format="%.0f")
with cD:
    cash_bal = st.number_input("Cash (₹)", min_value=0.0, value=200000.0, step=10000.0, format="%.0f")
with cE:
    min_cash_alert = st.number_input("Min Cash Alert (₹)", min_value=0.0, value=100000.0, step=10000.0, format="%.0f")
with cF:
    with st.expander("Governance thresholds (expand)"):
        max_exp_ratio = st.number_input("Max Expense Ratio", min_value=0.0, max_value=2.0, value=0.75, step=0.01, format="%.2f")
        min_net_margin = st.number_input("Min Net Margin", min_value=-1.0, max_value=1.0, value=0.10, step=0.01, format="%.2f")
        max_ic_pct = st.number_input("Max Inter-company % of Revenue", min_value=0.0, max_value=1.0, value=0.20, step=0.01, format="%.2f")
        max_ccc_days = st.number_input("Max CCC Days", min_value=0.0, value=60.0, step=5.0, format="%.0f")

st.divider()

# =========================================================
# MAIN NAV
# =========================================================
tabs = st.tabs(
    ["Executive Dashboard (Management)", "Integrated Dashboard (Detailed)", "Data Tables", "SDE/SOE (If uploaded)"]
)

# =========================================================
# TAB 1: EXECUTIVE DASHBOARD
# =========================================================
with tabs[0]:
    if not kpi_loaded or month_ts is None:
        st.warning("Upload KPI Template and ensure it has a 'Transactions' sheet.")
    else:
        consolidated = (view_mode == "Consolidated View")
        k = compute_kpis(tx, month_ts, entity, consolidated, eliminate_ic)

        dso, dpo, dio, ccc = compute_wc_days(ar, ap, inv, k.Revenue, k.COGS, k.days_in_month)

        wc_rag, wc_why = rag_label(ccc, green_max=max_ccc_days, higher_is_bad=True)
        cash_rag, cash_why = rag_label(cash_bal, green_min=min_cash_alert, higher_is_bad=False)
        exp_rag, exp_why = rag_label(k.ExpenseRatio, green_max=max_exp_ratio, higher_is_bad=True)
        nm_rag, nm_why = rag_label(k.NetMargin, green_min=min_net_margin, higher_is_bad=False)
        ic_rag, ic_why = rag_label(k.IntercompanyPct, green_max=max_ic_pct, higher_is_bad=True)

        # Pills
        st.markdown('<div class="pill-row">', unsafe_allow_html=True)

        def pill(title, value, subtitle=""):
            st.markdown(
                f"""
                <div class="pill">
                  <div class="k">{title}</div>
                  <div class="v">{value}</div>
                  <div class="s">{subtitle}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        context = f"{month_ts.strftime('%b-%Y')} | {'All Entities' if consolidated else entity}"
        pill("Context", context, "Selected period / entity")
        pill("Total Revenue", inr(k.Revenue), "Ledger-derived")
        pill("Total Expenses", inr(abs(k.COGS) + abs(k.Opex)), "COGS + Opex")
        pill("Net Profit", inr(k.NetProfit), f"Net Margin: {pct(k.NetMargin)}")
        pill("Operating Cash Flow", inr(k.OCF), "CashFlag = CASH")
        pill("Inter-company %", pct(k.IntercompanyPct), f"IC Sum: {inr(k.IntercompanySum)}")
        st.markdown("</div>", unsafe_allow_html=True)

        # OVERVIEW
        st.markdown("<div class='section-title'>OVERVIEW</div>", unsafe_allow_html=True)
        o1, o2, o3 = st.columns([1.2, 1.0, 1.0])

        with o1:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Revenue vs Expenses</h3>", unsafe_allow_html=True)

            s = monthly_series(tx, consolidated, entity, eliminate_ic)
            if s.empty:
                st.markdown('<div class="muted">No monthly series available.</div>', unsafe_allow_html=True)
            else:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=s["MonthLabel"], y=s["Revenue"],
                    name="Revenue",
                    marker_color=PBI["green"]
                ))
                fig.add_trace(go.Scatter(
                    x=s["MonthLabel"], y=s["Expenses"],
                    name="Expenses",
                    mode="lines+markers",
                    line=dict(color=PBI["red"], width=3),
                    marker=dict(color=PBI["red"], size=7),
                ))
                fig = plotly_light_layout(fig, height=340, title_text="")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("</div>", unsafe_allow_html=True)

        with o2:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Fund Flow Analysis</h3>", unsafe_allow_html=True)
            fig_ff = fund_flow_donut(k.Revenue, k.Opex, k.COGS, k.Interest, k.Tax)
            st.plotly_chart(fig_ff, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with o3:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Alerts & Flags</h3>", unsafe_allow_html=True)

            issues = 0
            issues += 1 if exp_rag in ("RED", "AMBER") else 0
            issues += 1 if nm_rag in ("RED", "AMBER") else 0
            issues += 1 if cash_rag in ("RED", "AMBER") else 0
            issues += 1 if ic_rag in ("RED", "AMBER") else 0
            issues += 1 if wc_rag in ("RED", "AMBER") else 0

            st.metric("Issues Triggered", f"{issues}", help="Count of governance checks in AMBER/RED")
            st.markdown("<div class='muted'>Top Governance Status</div>", unsafe_allow_html=True)

            st.markdown(
                f"{rag_badge_html(exp_rag)} <b>Expense Ratio</b> <span class='muted'>({pct(k.ExpenseRatio)} vs {pct(max_exp_ratio)})</span>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"{rag_badge_html(nm_rag)} <b>Net Margin</b> <span class='muted'>({pct(k.NetMargin)} vs min {pct(min_net_margin)})</span>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"{rag_badge_html(cash_rag)} <b>Cash Position</b> <span class='muted'>({inr(cash_bal)} vs min {inr(min_cash_alert)})</span>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"{rag_badge_html(ic_rag)} <b>Inter-company %</b> <span class='muted'>({pct(k.IntercompanyPct)} vs {pct(max_ic_pct)})</span>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"{rag_badge_html(wc_rag)} <b>CCC</b> <span class='muted'>({'—' if np.isnan(ccc) else f'{ccc:.1f}'} vs {max_ccc_days:.0f})</span>",
                unsafe_allow_html=True
            )
            st.markdown("</div>", unsafe_allow_html=True)

        # RISK INSIGHTS
        st.markdown("<div class='section-title'>RISK INSIGHTS</div>", unsafe_allow_html=True)
        r1, r2 = st.columns([1.2, 1.0])

        with r1:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Anomaly Alerts</h3>", unsafe_allow_html=True)

            anomalies = anomaly_counts(k.df_month)
            order = ["Fee Discrepancy", "Fake Invigilator Entry", "Unclaimed Certificates", "Inter-Entity Transfer"]

            for name in order:
                count = anomalies.get(name, 0)
                dot_color = PBI["orange"] if count > 0 else "rgba(15,23,42,0.16)"
                right_color = PBI["red"] if count > 0 else PBI["muted"]
                st.markdown(
                    f"""
                    <div class="an-item">
                      <div class="an-left">
                        <span class="an-dot" style="background:{dot_color}; box-shadow: 0 0 0 3px rgba(230,108,55,0.14);"></span>
                        {name}
                      </div>
                      <div class="an-right" style="color:{right_color};">{count} Cases</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

            st.markdown("<div class='muted'>Counts are auto-derived from ledger heuristics.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with r2:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Risk Gauges</h3>", unsafe_allow_html=True)

            fraud_proxy = 0.0
            fraud_proxy += 45 * (0 if np.isnan(k.IntercompanyPct) else float(np.clip(k.IntercompanyPct / max_ic_pct, 0, 2)))
            fraud_proxy += 35 * (0 if np.isnan(k.ExpenseRatio) else float(np.clip(k.ExpenseRatio / max_exp_ratio, 0, 2)))
            fraud_proxy += 20 * (1.0 if cash_rag == "RED" else 0.5 if cash_rag == "AMBER" else 0.0)
            fraud_proxy = float(np.clip(fraud_proxy, 0, 100))

            s = monthly_series(tx, consolidated, entity, eliminate_ic)
            pr_proxy = 50.0
            if not s.empty and len(s) >= 4:
                vol = float(np.nanstd(s["Net Profit"].tail(6)))
                base = float(np.nanmean(np.abs(s["Net Profit"].tail(6))) + 1.0)
                pr_proxy += 35 * float(np.clip(vol / base, 0, 1.5))
            an_intensity = float(sum(anomalies.values()))
            pr_proxy += 2.0 * min(an_intensity, 20)
            pr_proxy = float(np.clip(pr_proxy, 0, 100))

            st.plotly_chart(gauge_percent("Fraud Risk", fraud_proxy, color=PBI["red"], height=220), use_container_width=True)
            st.plotly_chart(gauge_percent("Pass Rate Spike", pr_proxy, color=PBI["orange"], height=220), use_container_width=True)

            st.markdown("</div>", unsafe_allow_html=True)

        # FORECAST & PREDICTIONS (ADDED BACK)
        st.markdown("<div class='section-title'>FORECAST & PREDICTIONS</div>", unsafe_allow_html=True)

        fL, fR = st.columns([1, 1])

        s = monthly_series(tx, consolidated, entity, eliminate_ic)
        df_rev, df_exp = forecast_tables(s, n=6)

        with fL:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Revenue Forecast (Next 6 Months)</h3>", unsafe_allow_html=True)

            if df_rev.empty:
                st.markdown("<div class='muted'>Not enough history to forecast.</div>", unsafe_allow_html=True)
            else:
                fig_rf = px.bar(
                    df_rev,
                    x="Month",
                    y="Revenue Forecast",
                    text="Revenue Forecast",
                    color_discrete_sequence=[PBI["green"]],
                )
                fig_rf.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
                fig_rf = plotly_light_layout(fig_rf, height=330, title_text="")
                st.plotly_chart(fig_rf, use_container_width=True)

            st.markdown("</div>", unsafe_allow_html=True)

        with fR:
            st.markdown('<div class="panel">', unsafe_allow_html=True)
            st.markdown("<h3>Expense Projection (Next 6 Months)</h3>", unsafe_allow_html=True)

            if df_exp.empty:
                st.markdown("<div class='muted'>Not enough history to forecast.</div>", unsafe_allow_html=True)
            else:
                fig_ep = px.bar(
                    df_exp,
                    x="Month",
                    y="Expense Projection",
                    text="Expense Projection",
                    color_discrete_sequence=[PBI["blue"]],
                )
                fig_ep.update_traces(texttemplate="%{text:,.0f}", textposition="outside", cliponaxis=False)
                fig_ep = plotly_light_layout(fig_ep, height=330, title_text="")
                st.plotly_chart(fig_ep, use_container_width=True)

            st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# TAB 2: INTEGRATED DASHBOARD
# =========================================================
with tabs[1]:
    st.markdown("<div class='section-title'>Integrated Financial Performance & Governance Dashboard (Detailed)</div>", unsafe_allow_html=True)

    if not kpi_loaded or month_ts is None:
        st.warning("Upload KPI Template and ensure it has Transactions data.")
    else:
        consolidated = (view_mode == "Consolidated View")
        k = compute_kpis(tx, month_ts, entity, consolidated, eliminate_ic)

        dso, dpo, dio, ccc = compute_wc_days(ar, ap, inv, k.Revenue, k.COGS, k.days_in_month)

        t1, t2, t3, t4 = st.tabs(["KPI Summary", "P&L Table", "Inter-company", "Narration (Mgmt)"])

        with t1:
            col1, col2, col3, col4, col5, col6 = st.columns(6)
            col1.metric("Revenue", inr(k.Revenue))
            col2.metric("COGS", inr(k.COGS))
            col3.metric("Opex", inr(k.Opex))
            col4.metric("EBITDA", inr(k.EBITDA))
            col5.metric("Net Profit", inr(k.NetProfit))
            col6.metric("OCF", inr(k.OCF))

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Expense Ratio", pct(k.ExpenseRatio))
            m2.metric("Gross Margin", pct(k.GrossMargin))
            m3.metric("EBITDA Margin", pct(k.EBITDAMargin))
            m4.metric("Net Margin", pct(k.NetMargin))

            w1, w2, w3, w4 = st.columns(4)
            w1.metric("DSO", "—" if np.isnan(dso) else f"{dso:.1f}")
            w2.metric("DPO", "—" if np.isnan(dpo) else f"{dpo:.1f}")
            w3.metric("DIO", "—" if np.isnan(dio) else f"{dio:.1f}")
            w4.metric("CCC", "—" if np.isnan(ccc) else f"{ccc:.1f}")

        with t2:
            dfm = k.df_month.copy()
            if dfm.empty:
                st.info("No ledger rows for selected context.")
            else:
                def sum_type(df, t):
                    return df.loc[df["AccountType"].str.upper().eq(t.upper()), "Amount_norm"].sum()

                pnl = pd.DataFrame(
                    [
                        ["Revenue", sum_type(dfm, "Revenue")],
                        ["COGS", sum_type(dfm, "COGS")],
                        ["Gross Profit", sum_type(dfm, "Revenue") + sum_type(dfm, "COGS")],
                        ["Opex", sum_type(dfm, "Opex")],
                        ["EBITDA", (sum_type(dfm, "Revenue") + sum_type(dfm, "COGS") + sum_type(dfm, "Opex"))],
                        ["Interest", sum_type(dfm, "Interest")],
                        ["Tax", sum_type(dfm, "Tax")],
                        ["Net Profit", k.NetProfit],
                    ],
                    columns=["Metric", "Value"],
                )
                pnl["Value"] = pnl["Value"].apply(inr)
                st.dataframe(pnl, use_container_width=True, hide_index=True)

                st.markdown("<div class='muted'>Raw ledger (audit trail)</div>", unsafe_allow_html=True)
                show_cols = ["Date","Entity","AccountType","Category","Counterparty","Amount","Amount_norm","CashFlag","IntercompanyFlag"]
                st.dataframe(dfm[show_cols].sort_values("Date"), use_container_width=True, hide_index=True)

        with t3:
            dfm = k.df_month.copy()
            ic = dfm[dfm["IntercompanyFlag"].str.upper().eq("YES")].copy()
            if ic.empty:
                st.info("No inter-company rows for selected context.")
            else:
                by_cp = ic.groupby("Counterparty")["Amount_norm"].sum().abs().sort_values(ascending=False).head(12).reset_index()
                by_cp = by_cp.rename(columns={"Amount_norm": "Absolute IC"})
                fig_ic = px.bar(
                    by_cp, x="Counterparty", y="Absolute IC",
                    title="",
                    color_discrete_sequence=[PBI["orange"]]
                )
                fig_ic = plotly_light_layout(fig_ic, height=360, title_text="")
                st.plotly_chart(fig_ic, use_container_width=True)

        with t4:
            context = f"{month_ts.strftime('%b-%Y')} | {'All Entities' if consolidated else entity}"
            bullets = [
                f"Context: {context}. Inter-company elimination: {'ON' if (consolidated and eliminate_ic) else 'OFF'}.",
                f"Revenue {inr(k.Revenue)}; Expenses {inr(abs(k.COGS)+abs(k.Opex))}; Net Profit {inr(k.NetProfit)} (Margin {pct(k.NetMargin)}).",
                f"Operating Cash Flow {inr(k.OCF)}. Cash buffer: {inr(cash_bal)} vs alert: {inr(min_cash_alert)}.",
                f"Expense Ratio {pct(k.ExpenseRatio)} vs limit {pct(max_exp_ratio)}; IC% {pct(k.IntercompanyPct)} vs limit {pct(max_ic_pct)}.",
                f"Working capital: DSO {'—' if np.isnan(dso) else f'{dso:.1f}'} | DPO {'—' if np.isnan(dpo) else f'{dpo:.1f}'} | DIO {'—' if np.isnan(dio) else f'{dio:.1f}'} | CCC {'—' if np.isnan(ccc) else f'{ccc:.1f}'} (limit {max_ccc_days:.0f}).",
            ]
            st.markdown("<div class='panel'>", unsafe_allow_html=True)
            st.markdown("<h3>Management Narration</h3>", unsafe_allow_html=True)
            for b in bullets:
                st.markdown(f"- {b}")
            st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# TAB 3: DATA TABLES
# =========================================================
with tabs[2]:
    st.markdown("<div class='section-title'>Data Tables</div>", unsafe_allow_html=True)
    if not kpi_loaded:
        st.info("Upload KPI template to see tables.")
    else:
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<h3>KPI Transactions (sample)</h3>", unsafe_allow_html=True)
        st.dataframe(tx.head(300), use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# TAB 4: SDE/SOE
# =========================================================
with tabs[3]:
    st.markdown("<div class='section-title'>SDE/SOE Dashboard (Optional)</div>", unsafe_allow_html=True)
    if not sde_loaded or not sde_tables:
        st.info("Upload SDE/SOE xlsx to enable this section.")
    else:
        st.markdown("<div class='panel'>", unsafe_allow_html=True)
        st.markdown("<h3>Available Sheets</h3>", unsafe_allow_html=True)
        sheet = st.selectbox("Select sheet", list(sde_tables.keys()))
        st.dataframe(sde_tables[sheet].head(250), use_container_width=True, hide_index=True)
        st.markdown("<div class='muted'>Tell me which SDE/SOE columns you want as charts and I’ll wire them.</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# BOTTOM NAV MENU (visual)
# =========================================================
active = st.session_state.get("bottom_active", "Dashboard")
items = ["Dashboard", "Finance", "Compliance", "Audit Trail", "Reports", "AI Analytics", "Forecast", "Settings"]

st.markdown('<div class="bottom-nav"><div class="bottom-nav-inner">', unsafe_allow_html=True)
for it in items:
    cls = "nav-pill active" if it == active else "nav-pill"
    st.markdown(f'<span class="{cls}">{it}</span>', unsafe_allow_html=True)
st.markdown("</div></div>", unsafe_allow_html=True)

st.markdown(
    f"<div class='muted' style='margin-top:14px;'>Status: KPI Loaded = {kpi_loaded} | SDE/SOE Loaded = {sde_loaded}</div>",
    unsafe_allow_html=True,
)