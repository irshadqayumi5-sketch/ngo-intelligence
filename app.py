from pathlib import Path
from datetime import datetime, date
from io import BytesIO
import hashlib
import json
import urllib.request
import urllib.error

import pandas as pd
import plotly.express as px
import streamlit as st
import numpy as np
from docx import Document
from docx.shared import Inches
from sklearn.compose import ColumnTransformer
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (accuracy_score, balanced_accuracy_score, precision_score, recall_score, f1_score, roc_auc_score, average_precision_score, confusion_matrix)
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

APP_DIR = Path(__file__).resolve().parent
DATA_FILE = APP_DIR / "data" / "beneficiary_data.xlsx"
REPORTS_DIR = APP_DIR / "reports"
REQUIRED_COLUMNS = ["Beneficiary_ID", "Province", "Activity", "Satisfaction_Score", "Date"]
LOW_SCORE_VALUES = [1, 2]

st.set_page_config(page_title="AI NGO Intelligence | V22.0", page_icon="✦", layout="wide", initial_sidebar_state="expanded")
st.markdown("""
<style>
:root { --navy:#10243A; --navy2:#173B59; --teal:#16B8A6; --ink:#172B3A; --muted:#657789; --line:#E3EAF0; --surface:#FFFFFF; --canvas:#F4F7FA; }
html, body, [class*="css"] { font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
.stApp { background:linear-gradient(180deg,#F3F7FA 0%,#F7F9FC 100%); color:var(--ink); }
.block-container { max-width:1600px; padding-top:1.15rem; padding-bottom:3rem; }
[data-testid="stHeader"] { background:rgba(244,247,250,.82); }
[data-testid="stSidebar"] { background:linear-gradient(180deg,#10243A 0%,#163750 100%); border-right:1px solid rgba(255,255,255,.08); }
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"], [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span { color:#EAF2F8; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 { color:#FFFFFF; }
[data-testid="stSidebar"] hr { border-color:rgba(255,255,255,.16); }
[data-testid="stSidebar"] [data-testid="stRadio"] label { padding:.42rem .55rem; border-radius:9px; }
[data-testid="stSidebar"] [data-testid="stRadio"] label:hover { background:rgba(255,255,255,.08); }
[data-testid="stSidebar"] [data-testid="stFileUploader"] section { background:rgba(255,255,255,.06); border:1px dashed rgba(255,255,255,.35); }
[data-testid="stMetric"] { background:var(--surface); border:1px solid var(--line); border-radius:16px; padding:17px 18px; box-shadow:0 5px 18px rgba(16,36,58,.045); min-height:118px; }
[data-testid="stMetricLabel"] { color:#657789!important; font-size:.82rem!important; font-weight:650!important; }
[data-testid="stMetricValue"] { color:var(--navy)!important; font-size:1.85rem!important; font-weight:780!important; }
[data-testid="stMetricDelta"] { font-size:.8rem!important; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius:16px; }
.stButton > button, .stDownloadButton > button { border-radius:10px; border:1px solid #D4E0E8; font-weight:650; transition:all .16s ease; }
.stButton > button[kind="primary"] { background:linear-gradient(135deg,#117D83,#16B8A6); border:0; color:#fff; }
.stButton > button:hover, .stDownloadButton > button:hover { border-color:#16B8A6; color:#117D83; box-shadow:0 4px 12px rgba(22,184,166,.12); }
[data-testid="stDataFrame"], [data-testid="stTable"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
[data-testid="stExpander"] { background:rgba(255,255,255,.8); border:1px solid var(--line); border-radius:12px; }
[data-testid="stAlert"] { border-radius:12px; }
hr { border-color:var(--line); }
h1,h2,h3 { color:var(--navy); letter-spacing:-.025em; }
h2 { font-size:1.45rem!important; font-weight:760!important; }
h3 { font-size:1.08rem!important; font-weight:730!important; }
p, li { line-height:1.58; }
[data-testid="stPlotlyChart"] { background:#fff; border:1px solid var(--line); border-radius:16px; padding:8px 8px 0 8px; box-shadow:0 5px 18px rgba(16,36,58,.035); }
.premium-hero { background:linear-gradient(115deg,#10243A 0%,#17465B 63%,#117D83 100%); padding:27px 30px; border-radius:22px; color:white; position:relative; overflow:hidden; box-shadow:0 14px 34px rgba(16,36,58,.15); margin-bottom:17px; }
.premium-hero:after { content:""; position:absolute; width:270px; height:270px; border:1px solid rgba(255,255,255,.12); border-radius:50%; right:-60px; top:-135px; box-shadow:0 0 0 28px rgba(255,255,255,.035),0 0 0 58px rgba(255,255,255,.025); }
.premium-kicker { text-transform:uppercase; letter-spacing:.17em; font-size:.72rem; font-weight:750; color:#8DE5DA; margin-bottom:8px; }
.premium-title { color:#fff; font-size:2rem; font-weight:820; letter-spacing:-.045em; line-height:1.13; margin:0; }
.premium-subtitle { color:#D9E9F0; font-size:.94rem; margin-top:9px; max-width:760px; }
.premium-chip { display:inline-block; margin-top:17px; border:1px solid rgba(255,255,255,.22); border-radius:99px; padding:5px 11px; color:#F2FBFC; font-size:.75rem; background:rgba(255,255,255,.08); }
.section-intro { color:var(--muted); font-size:.91rem; margin-top:-.35rem; margin-bottom:1rem; }
.small-note { color:var(--muted); font-size:.8rem; }
@media (max-width: 700px) { .premium-hero { padding:22px 20px; border-radius:16px; } .premium-title {font-size:1.55rem;} .block-container {padding-left:1rem;padding-right:1rem;} }
</style>
""", unsafe_allow_html=True)
st.markdown("""<div class="premium-hero">
  <div class="premium-kicker">Programme Monitoring · MEAL · Decision Support</div>
  <div class="premium-title">AI NGO Intelligence</div>
  <div class="premium-subtitle">A unified workspace for beneficiary insights, programme performance, data governance and evidence-based management reporting.</div>
  <span class="premium-chip">V22.0 · PREMIUM INTERFACE</span>
</div>""", unsafe_allow_html=True)
st.caption("Evidence-led analytics • Human-reviewed recommendations • Privacy-conscious reporting. Descriptive patterns do not establish causes; experimental model estimates must not determine assistance eligibility.")



INDICATOR_FILE = APP_DIR / "data" / "indicator_registry_v22_0.csv"
INDICATOR_COLUMNS = [
    "Indicator_ID", "Indicator_Name", "Definition", "Unit", "Baseline", "Target", "Actual",
    "Direction", "Reporting_Period", "Responsible_Team", "Due_Date", "Status", "Notes"
]

def empty_indicator_registry():
    return pd.DataFrame(columns=INDICATOR_COLUMNS)

def load_indicator_registry():
    if not INDICATOR_FILE.exists():
        return empty_indicator_registry()
    try:
        df = pd.read_csv(INDICATOR_FILE, dtype={"Indicator_ID": "string"})
        for col in INDICATOR_COLUMNS:
            if col not in df.columns:
                df[col] = pd.NA
        return df[INDICATOR_COLUMNS]
    except Exception as exc:
        st.warning(f"Could not load the saved indicator registry: {exc}")
        return empty_indicator_registry()

def clean_indicator_registry(df):
    cleaned = df.copy()
    for col in INDICATOR_COLUMNS:
        if col not in cleaned.columns:
            cleaned[col] = pd.NA
    cleaned = cleaned[INDICATOR_COLUMNS]
    text_cols = ["Indicator_ID", "Indicator_Name", "Definition", "Unit", "Direction", "Reporting_Period", "Responsible_Team", "Due_Date", "Status", "Notes"]
    for col in text_cols:
        cleaned[col] = cleaned[col].fillna("").astype(str).str.strip()
    for col in ["Baseline", "Target", "Actual"]:
        cleaned[col] = pd.to_numeric(cleaned[col], errors="coerce")
    # Drop untouched blank rows; retain rows with at least a name or ID.
    cleaned = cleaned[(cleaned["Indicator_Name"] != "") | (cleaned["Indicator_ID"] != "")].copy()
    if not cleaned.empty:
        used = set(cleaned.loc[cleaned["Indicator_ID"] != "", "Indicator_ID"].tolist())
        next_num = 1
        for idx in cleaned.index:
            if not cleaned.at[idx, "Indicator_ID"]:
                while f"IND-{next_num:03d}" in used:
                    next_num += 1
                new_id = f"IND-{next_num:03d}"
                cleaned.at[idx, "Indicator_ID"] = new_id
                used.add(new_id)
                next_num += 1
        cleaned["Direction"] = cleaned["Direction"].where(cleaned["Direction"].isin(["Higher is better", "Lower is better"]), "Higher is better")
        cleaned["Status"] = cleaned["Status"].where(cleaned["Status"].isin(["Not started", "On track", "At risk", "Delayed", "Completed", "Needs review"]), "Needs review")
    return cleaned.reset_index(drop=True)

def calculate_indicator_progress(df):
    rows = []
    for _, row in df.iterrows():
        baseline, target, actual = row.get("Baseline"), row.get("Target"), row.get("Actual")
        direction = row.get("Direction", "Higher is better")
        progress = None
        gap = None
        if pd.notna(target) and pd.notna(actual):
            gap = float(target - actual) if direction == "Higher is better" else float(actual - target)
        if pd.notna(baseline) and pd.notna(target) and pd.notna(actual):
            denominator = (target - baseline) if direction == "Higher is better" else (baseline - target)
            numerator = (actual - baseline) if direction == "Higher is better" else (baseline - actual)
            if denominator != 0:
                progress = float(numerator / denominator * 100)
        elif pd.notna(target) and pd.notna(actual) and float(target) != 0 and direction == "Higher is better":
            # Without a baseline, report target attainment, not progress from baseline.
            progress = float(actual / target * 100)
        rows.append({"Progress toward target (%)": round(progress, 1) if progress is not None else None,
                     "Gap to target": round(gap, 3) if gap is not None else None})
    return pd.DataFrame(rows, index=df.index)

def load_workbook(uploaded):
    if uploaded is not None:
        return pd.read_excel(uploaded), "Uploaded Excel workbook"
    if DATA_FILE.exists():
        return pd.read_excel(DATA_FILE), str(DATA_FILE)
    return None, None


def prepare_data(raw):
    df = raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        return df, missing
    df["Satisfaction_Score"] = pd.to_numeric(df["Satisfaction_Score"], errors="coerce")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for col in ["Beneficiary_ID", "Province", "Activity"]:
        df[col] = df[col].astype("string").str.strip()
        df[col] = df[col].replace("", pd.NA)
    return df, []


def valid_scores(df):
    """Return satisfaction scores within the expected 1–5 scale only."""
    return df.loc[df["Satisfaction_Score"].between(1, 5), "Satisfaction_Score"]


def quality_checks(df):
    checks = [
        ("Total rows", len(df)),
        ("Total columns", len(df.columns)),
        ("Rows with any missing cell", int(df.isna().any(axis=1).sum())),
        ("Missing cells (all columns)", int(df.isna().sum().sum())),
        ("Duplicate full rows", int(df.duplicated().sum())),
        ("Missing Beneficiary_ID", int(df["Beneficiary_ID"].isna().sum())),
        ("Duplicate nonblank Beneficiary_ID values", int(df.loc[df["Beneficiary_ID"].notna(), "Beneficiary_ID"].duplicated().sum())),
        ("Missing Province", int(df["Province"].isna().sum())),
        ("Missing Activity", int(df["Activity"].isna().sum())),
        ("Missing/invalid satisfaction scores", int((df["Satisfaction_Score"].isna() | ~df["Satisfaction_Score"].between(1, 5)).sum())),
        ("Scores outside 1–5", int((df["Satisfaction_Score"].notna() & ~df["Satisfaction_Score"].between(1, 5)).sum())),
        ("Missing/invalid dates", int(df["Date"].isna().sum())),
        ("Future dates", int((df["Date"].notna() & (df["Date"].dt.date > date.today())).sum())),
    ]
    # Additional domain checks are conditional on the source workbook's available columns.
    if "Age" in df.columns:
        age = pd.to_numeric(df["Age"], errors="coerce")
        checks.extend([("Missing/invalid Age", int(age.isna().sum())), ("Age outside 0–120", int((age.notna() & ~age.between(0, 120)).sum()))])
    if "Household_Size" in df.columns:
        hh = pd.to_numeric(df["Household_Size"], errors="coerce")
        checks.extend([("Missing/invalid Household_Size", int(hh.isna().sum())), ("Household_Size below 1", int((hh.notna() & (hh < 1)).sum()))])
    for col in ["Target", "Actual", "Attendance"]:
        if col in df.columns:
            num = pd.to_numeric(df[col], errors="coerce")
            checks.append((f"Missing/non-numeric {col}", int(num.isna().sum())))
            checks.append((f"Negative {col}", int((num.notna() & (num < 0)).sum())))
    if "Actual" in df.columns and "Target" in df.columns:
        target = pd.to_numeric(df["Target"], errors="coerce")
        actual = pd.to_numeric(df["Actual"], errors="coerce")
        checks.append(("Actual exceeds Target", int((target.notna() & actual.notna() & (actual > target)).sum())))
    if "Sex" in df.columns:
        allowed = {"female", "male", "other", "unknown", "prefer not to say", "prefer_not_to_say", "non-binary", "nonbinary"}
        values = df["Sex"].dropna().astype(str).str.strip().str.lower()
        checks.append(("Unrecognized Sex category (review only)", int((~values.isin(allowed)).sum())))
    if "Assistance_Received" in df.columns:
        values = df["Assistance_Received"].dropna().astype(str).str.strip().str.lower()
        allowed = {"yes", "no", "true", "false", "1", "0", "pending", "unknown", "not applicable", "n/a"}
        checks.append(("Unrecognized Assistance_Received category (review only)", int((~values.isin(allowed)).sum())))
    return pd.DataFrame(checks, columns=["Check", "Result"])


def grouped_analysis(df, group_col):
    valid = df[df["Satisfaction_Score"].between(1, 5)].dropna(subset=[group_col]).copy()
    if valid.empty:
        return pd.DataFrame(columns=[group_col, "Scored records", "Average satisfaction", "Low-score records", "Low-satisfaction rate (%)"])
    result = valid.groupby(group_col, dropna=False).agg(
        **{"Scored records": ("Satisfaction_Score", "size"),
           "Average satisfaction": ("Satisfaction_Score", "mean"),
           "Low-score records": ("Satisfaction_Score", lambda s: int(s.isin(LOW_SCORE_VALUES).sum()))}
    ).reset_index()
    result["Average satisfaction"] = result["Average satisfaction"].round(2)
    result["Low-satisfaction rate (%)"] = (result["Low-score records"] / result["Scored records"] * 100).round(2)
    return result.sort_values("Low-satisfaction rate (%)", ascending=False)


def make_insights(df, province_table, activity_table):
    scores = valid_scores(df)
    insights = []
    if scores.empty:
        return [{"Priority": "High", "Area": "Satisfaction data", "Finding": "No valid satisfaction scores are available in the selected records.", "Recommended action": "Check source forms, score coding, and import mapping before interpreting satisfaction.", "Evidence / denominator": "0 valid scores"}]
    low_n = int(scores.isin(LOW_SCORE_VALUES).sum())
    low_pct = low_n / len(scores) * 100
    insights.append({
        "Priority": "High" if low_pct >= 30 else "Medium" if low_pct >= 15 else "Monitor",
        "Area": "Overall beneficiary satisfaction",
        "Finding": f"{low_n:,} of {len(scores):,} valid satisfaction scores ({low_pct:.1f}%) are low (1–2). Mean score: {scores.mean():.2f}/5.",
        "Recommended action": "Review a sample of low-score records, document feedback themes, and assign follow-up actions only to issues confirmed by evidence.",
        "Evidence / denominator": f"{len(scores):,} valid satisfaction scores in the current selection",
    })
    for table, col, label in [(province_table, "Province", "Province"), (activity_table, "Activity", "Activity")]:
        if not table.empty:
            eligible = table[table["Scored records"] >= 20]
            candidate = (eligible if not eligible.empty else table).iloc[0]
            insights.append({
                "Priority": "High" if candidate["Low-satisfaction rate (%)"] >= 35 else "Medium",
                "Area": f"{label}: {candidate[col]}",
                "Finding": f"{candidate[col]} has a low-satisfaction rate of {candidate['Low-satisfaction rate (%)']:.1f}% ({int(candidate['Low-score records']):,}/{int(candidate['Scored records']):,} scored records).",
                "Recommended action": f"Review disaggregated feedback and delivery records for {candidate[col]}; consider sample size and context before changing implementation.",
                "Evidence / denominator": f"{int(candidate['Scored records']):,} valid scored records; descriptive comparison only",
            })
    return insights


def build_action_plan(insights):
    return pd.DataFrame([{
        "Action_ID": f"V18.0-ACT-{i:03d}", "Priority": item["Priority"], "Area": item["Area"],
        "Finding": item["Finding"], "Recommended_Action": item["Recommended action"],
        "Responsible_Role": "Programme / MEAL team (confirm locally)", "Due_Date": "",
        "Status": "Not Started", "Evidence_or_Completion_Note": ""
    } for i, item in enumerate(insights, 1)])


ACTION_PLAN_FILE = APP_DIR / "data" / "management_action_plan_v18_0.csv"

def load_persistent_action_plan(insights):
    """Load the local action register and append newly generated insight areas without overwriting edits."""
    generated = build_action_plan(insights)
    if ACTION_PLAN_FILE.exists():
        try:
            saved = pd.read_csv(ACTION_PLAN_FILE).fillna("")
            required = list(generated.columns)
            if all(c in saved.columns for c in required):
                saved = saved[required].copy()
                existing_areas = set(saved["Area"].astype(str))
                additions = generated[~generated["Area"].astype(str).isin(existing_areas)]
                if not additions.empty:
                    saved = pd.concat([saved, additions], ignore_index=True)
                return saved
        except Exception:
            pass
    return generated

def save_action_plan(action_plan):
    ACTION_PLAN_FILE.parent.mkdir(parents=True, exist_ok=True)
    action_plan.to_csv(ACTION_PLAN_FILE, index=False, encoding="utf-8-sig")

def make_management_narrative(df, province_table, activity_table):
    """Create a transparent, deterministic narrative from observed values; no external AI calls."""
    scores = valid_scores(df)
    lines = []
    lines.append(f"The selected dataset contains {len(df):,} records.")
    if scores.empty:
        lines.append("No valid satisfaction scores are available; satisfaction conclusions should be withheld until score quality is reviewed.")
    else:
        low = int(scores.isin(LOW_SCORE_VALUES).sum())
        rate = low / len(scores) * 100
        lines.append(f"Among {len(scores):,} valid satisfaction scores, {low:,} ({rate:.1f}%) are low (score 1–2); the mean score is {scores.mean():.2f}/5.")
        if rate >= 30:
            lines.append("Management attention: the observed low-satisfaction rate is elevated under the system's descriptive screening threshold. Review sampled feedback and corroborating delivery records.")
        elif rate >= 15:
            lines.append("Management attention: review low-satisfaction feedback themes and monitor changes across reporting periods.")
        else:
            lines.append("The overall low-satisfaction rate is below the system's 15% descriptive attention threshold; continue routine monitoring and check subgroup variation.")
    for table, col, label in [(province_table, "Province", "province"), (activity_table, "Activity", "activity")]:
        if not table.empty and col in table.columns:
            eligible = table[table["Scored records"] >= 20]
            if not eligible.empty:
                row = eligible.iloc[0]
                lines.append(f"Among {label} groups with at least 20 scored records, {row[col]} has the highest observed low-satisfaction rate ({row['Low-satisfaction rate (%)']:.1f}%; n={int(row['Scored records'])}). This is a descriptive flag, not proof of cause.")
    if {"Target", "Actual"}.issubset(df.columns):
        valid_pairs = df[["Target", "Actual"]].apply(pd.to_numeric, errors="coerce").dropna()
        valid_pairs = valid_pairs[(valid_pairs["Target"] > 0) & (valid_pairs["Actual"] >= 0)]
        if not valid_pairs.empty:
            ratio = valid_pairs["Actual"] / valid_pairs["Target"] * 100
            lines.append(f"Programme performance has {len(valid_pairs):,} valid Target/Actual observations; the median per-record achievement rate is {ratio.median():.1f}%. Do not sum these values unless each row is an independent planning/achievement observation.")
    lines.append("Recommended follow-up: validate the flagged patterns with programme teams, document beneficiary feedback themes, assign owners and deadlines, and recheck progress in the next reporting cycle.")
    return lines


def word_report(df, source, quality_df, insights, action_plan, filter_summary):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(.65)
    section.bottom_margin = Inches(.65)
    doc.add_heading("Beneficiary Analysis and Management Insights", 0)
    doc.add_paragraph(f"AI_NGO Reporting System V18.0 | Generated: {datetime.now():%d %B %Y, %H:%M}")
    doc.add_paragraph(f"Data source: {source}")
    doc.add_paragraph(f"Filters applied: {filter_summary}")
    scores = valid_scores(df)
    doc.add_heading("1. Executive Summary", level=1)
    doc.add_paragraph(f"Selected records: {len(df):,}. Valid satisfaction scores: {len(scores):,}.")
    if len(scores):
        low_n = int(scores.isin(LOW_SCORE_VALUES).sum())
        doc.add_paragraph(f"Mean satisfaction: {scores.mean():.2f}/5. Low satisfaction (score 1–2): {low_n:,} ({low_n/len(scores)*100:.1f}% of valid scores).")
    doc.add_paragraph("These descriptive statistics do not establish the causes of reported satisfaction scores.")
    doc.add_heading("2. Evidence-based Insights", level=1)
    for item in insights:
        doc.add_heading(f"{item['Priority']} priority — {item['Area']}", level=2)
        doc.add_paragraph(item["Finding"])
        doc.add_paragraph("Recommended action: " + item["Recommended action"])
        doc.add_paragraph("Evidence / denominator: " + item["Evidence / denominator"])
    doc.add_heading("3. Management Action Plan", level=1)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Shading Accent 1"
    for cell, value in zip(table.rows[0].cells, ["Action ID", "Priority", "Area", "Recommended action", "Status"]):
        cell.text = value
    for _, row in action_plan.iterrows():
        cells = table.add_row().cells
        for cell, value in zip(cells, [row["Action_ID"], row["Priority"], row["Area"], row["Recommended_Action"], row["Status"]]):
            cell.text = str(value)
    doc.add_heading("4. Data Quality Checks", level=1)
    for _, row in quality_df.iterrows():
        doc.add_paragraph(f"{row['Check']}: {row['Result']}", style="List Bullet")
    doc.add_paragraph("Validate results against source records and with programme teams before external or donor-facing publication.")
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def monthly_trend(df):
    """Monthly descriptive trend; invalid dates/scores are excluded from score metrics."""
    data = df.copy()
    data = data.dropna(subset=["Date"])
    data = data[data["Satisfaction_Score"].between(1, 5)].copy()
    if data.empty:
        return pd.DataFrame(columns=["Month", "Scored records", "Average satisfaction", "Low-satisfaction rate (%)"])
    data["Month"] = data["Date"].dt.to_period("M").astype(str)
    result = data.groupby("Month").agg(
        **{"Scored records": ("Satisfaction_Score", "size"),
           "Average satisfaction": ("Satisfaction_Score", "mean"),
           "Low-score records": ("Satisfaction_Score", lambda x: int(x.isin(LOW_SCORE_VALUES).sum()))}
    ).reset_index()
    result["Average satisfaction"] = result["Average satisfaction"].round(2)
    result["Low-satisfaction rate (%)"] = (result["Low-score records"] / result["Scored records"] * 100).round(2)
    return result


def risk_flags(group_table, overall_rate, group_col):
    """Flag groups with adequate sample size and a materially higher descriptive rate."""
    if group_table.empty:
        return pd.DataFrame(columns=[group_col, "Scored records", "Low-satisfaction rate (%)", "Difference vs overall (pp)", "Risk flag"])
    result = group_table.copy()
    result["Difference vs overall (pp)"] = (result["Low-satisfaction rate (%)"] - overall_rate).round(2)
    result["Risk flag"] = np.where(
        result["Scored records"] < 20, "Insufficient sample (<20)",
        np.where(result["Difference vs overall (pp)"] >= 5, "Review priority", "No elevated-rate flag")
    )
    return result.sort_values(["Risk flag", "Difference vs overall (pp)"], ascending=[True, False])


def prepare_ml_features(df):
    """Build leakage-resistant predictors. Exclude outcome, direct identifiers and free-text fields."""
    x = df.copy()
    y_score = pd.to_numeric(x["Satisfaction_Score"], errors="coerce")
    y = y_score.isin(LOW_SCORE_VALUES).astype(int)
    usable_label = y_score.between(1, 5)
    x = x.drop(columns=["Satisfaction_Score"], errors="ignore")
    drop_cols = []
    sensitive_tokens = ("beneficiary_id", "person_id", "respondent_id", "phone", "mobile", "name", "address", "comment", "feedback", "note", "token", "barcode", "serial")
    for col in x.columns:
        low_name = str(col).strip().lower()
        if any(token in low_name for token in sensitive_tokens):
            drop_cols.append(col)
    x = x.drop(columns=drop_cols, errors="ignore")
    # Expand date into calendar features rather than feeding raw timestamps.
    for col in list(x.columns):
        if pd.api.types.is_datetime64_any_dtype(x[col]) or str(col).strip().lower() in {"date", "interview_date", "survey_date", "registration_date"}:
            parsed = pd.to_datetime(x[col], errors="coerce")
            x[f"{col}_year"] = parsed.dt.year
            x[f"{col}_month"] = parsed.dt.month
            x = x.drop(columns=[col])
    # Remove near-unique string columns likely to behave as identifiers.
    for col in list(x.columns):
        if not pd.api.types.is_numeric_dtype(x[col]):
            nunique = x[col].nunique(dropna=True)
            if len(x) >= 20 and nunique > max(30, int(0.8 * len(x))):
                x = x.drop(columns=[col])
    # Convert numeric-looking columns where most nonblank values parse numerically.
    for col in x.columns:
        if not pd.api.types.is_numeric_dtype(x[col]):
            converted = pd.to_numeric(x[col], errors="coerce")
            nonblank = x[col].notna().sum()
            if nonblank and converted.notna().sum() / nonblank >= 0.9:
                x[col] = converted
            else:
                x[col] = x[col].astype("string").astype(object).where(x[col].notna(), np.nan)
    x = x.replace([np.inf, -np.inf], np.nan)
    return x, y, usable_label


def programme_performance(df, group_col=None):
    """Summarize target/actual fields without silently double-counting beneficiary-level rows.

    Totals are intentionally not computed here: a beneficiary-level file may repeat a
    project target on multiple records. Show means/medians and row counts unless the
    user explicitly confirms each row is an independent planning/achievement observation.
    """
    if not {"Target", "Actual"}.issubset(df.columns):
        return pd.DataFrame(), pd.DataFrame()
    data = df.copy()
    data["Target"] = pd.to_numeric(data["Target"], errors="coerce")
    data["Actual"] = pd.to_numeric(data["Actual"], errors="coerce")
    data = data.dropna(subset=["Target", "Actual"])
    data = data[(data["Target"] > 0) & (data["Actual"] >= 0)]
    if data.empty:
        return data, pd.DataFrame()
    data["Achievement_Gap"] = data["Target"] - data["Actual"]
    data["Achievement_Rate_Pct"] = data["Actual"] / data["Target"] * 100
    if group_col and group_col in data.columns:
        summary = data.groupby(group_col, dropna=False).agg(
            **{"Observations": ("Target", "size"),
               "Mean Target per record": ("Target", "mean"),
               "Mean Actual per record": ("Actual", "mean"),
               "Median Achievement Rate (%)": ("Achievement_Rate_Pct", "median"),
               "Mean Achievement Gap per record": ("Achievement_Gap", "mean")}
        ).reset_index()
        summary["Mean Target per record"] = summary["Mean Target per record"].round(2)
        summary["Mean Actual per record"] = summary["Mean Actual per record"].round(2)
        summary["Median Achievement Rate (%)"] = summary["Median Achievement Rate (%)"].round(2)
        summary["Mean Achievement Gap per record"] = summary["Mean Achievement Gap per record"].round(2)
        summary = summary.sort_values("Median Achievement Rate (%)", ascending=True)
    else:
        summary = pd.DataFrame([{
            "Valid observations": len(data),
            "Mean Target per record": round(data["Target"].mean(), 2),
            "Mean Actual per record": round(data["Actual"].mean(), 2),
            "Median Achievement Rate (%)": round(data["Achievement_Rate_Pct"].median(), 2),
            "Mean Achievement Gap per record": round(data["Achievement_Gap"].mean(), 2),
        }])
    return data, summary


def build_priority_ranking(province_table, activity_table, overall_rate):
    """Combine group-level descriptive flags into a transparent review queue."""
    rows = []
    for table, col, group_type in [(province_table, "Province", "Province"), (activity_table, "Activity", "Activity")]:
        if table.empty or col not in table.columns:
            continue
        for _, row in table.iterrows():
            n = int(row["Scored records"])
            rate = float(row["Low-satisfaction rate (%)"])
            diff = round(rate - overall_rate, 2)
            if n < 20:
                level = "Monitor — small sample"
                rationale = "Fewer than 20 valid scored records; do not rank as a confirmed priority."
            elif diff >= 5:
                level = "High review priority" if rate >= 35 else "Review priority"
                rationale = f"At least 20 scored records and {diff:.1f} percentage points above the selected overall rate."
            elif rate >= 30:
                level = "Monitor"
                rationale = f"Rate is {rate:.1f}%; however, it is not at least 5 points above the selected overall rate."
            else:
                level = "Routine monitoring"
                rationale = "No elevated-rate flag under the current descriptive screening rule."
            rows.append({"Group type": group_type, "Group": str(row[col]), "Scored records": n,
                         "Low-score records": int(row["Low-score records"]),
                         "Low-satisfaction rate (%)": rate, "Difference vs overall (pp)": diff,
                         "Review level": level, "Why it is ranked here": rationale,
                         "Recommended next step": "Review a sample of feedback and verify delivery records; document confirmed issues before assigning corrective actions."})
    columns = ["Group type", "Group", "Scored records", "Low-score records", "Low-satisfaction rate (%)", "Difference vs overall (pp)", "Review level", "Why it is ranked here", "Recommended next step"]
    if not rows:
        return pd.DataFrame(columns=columns)
    result = pd.DataFrame(rows, columns=columns)
    level_order = {"High review priority": 0, "Review priority": 1, "Monitor": 2, "Monitor — small sample": 3, "Routine monitoring": 4}
    result["_order"] = result["Review level"].map(level_order).fillna(9)
    return result.sort_values(["_order", "Difference vs overall (pp)", "Low-satisfaction rate (%)", "Scored records"], ascending=[True, False, False, False]).drop(columns="_order").reset_index(drop=True)


def pdf_report(df, source, quality_df, insights, action_plan, filter_summary, priority_table):
    """Create a compact PDF management brief using ReportLab Platypus."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
    from xml.sax.saxutils import escape
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=14*mm, leftMargin=14*mm, topMargin=12*mm, bottomMargin=12*mm)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="SmallBody", parent=styles["BodyText"], fontSize=8, leading=10, alignment=TA_LEFT, spaceAfter=3))
    styles.add(ParagraphStyle(name="Tiny", parent=styles["BodyText"], fontSize=6.5, leading=8))
    story = [Paragraph("AI_NGO Reporting System V18.0 — Management Brief", styles["Title"]),
             Paragraph(f"Generated: {datetime.now():%d %B %Y, %H:%M} | Source: {escape(str(source))}", styles["SmallBody"]),
             Paragraph("Filters: " + escape(filter_summary), styles["SmallBody"]), Spacer(1, 4*mm)]
    scores = valid_scores(df)
    low_n = int(scores.isin(LOW_SCORE_VALUES).sum()) if len(scores) else 0
    low_pct = low_n / len(scores) * 100 if len(scores) else 0
    summary = [["Selected records", "Valid scores", "Mean satisfaction", "Low scores", "Low-satisfaction rate"],
               [f"{len(df):,}", f"{len(scores):,}", f"{scores.mean():.2f}/5" if len(scores) else "N/A", f"{low_n:,}", f"{low_pct:.1f}%" if len(scores) else "N/A"]]
    t = Table(summary, colWidths=[36*mm]*5)
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("GRID",(0,0),(-1,-1),0.4,colors.grey),("ALIGN",(0,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),5)]))
    story += [t, Spacer(1, 5*mm), Paragraph("Executive interpretation", styles["Heading1"])]
    if len(scores):
        story.append(Paragraph(f"{low_n:,} of {len(scores):,} valid satisfaction scores ({low_pct:.1f}%) are low (1–2). The mean score is {scores.mean():.2f}/5. These descriptive figures do not establish causes.", styles["SmallBody"]))
    else:
        story.append(Paragraph("No valid satisfaction scores are available; do not draw satisfaction conclusions until the source data is reviewed.", styles["SmallBody"]))
    story.append(Paragraph("Priority ranking (descriptive screening only)", styles["Heading1"]))
    if priority_table.empty:
        story.append(Paragraph("No group-level ranking could be calculated.", styles["SmallBody"]))
    else:
        headers = ["Type", "Group", "n", "Low n", "Low %", "Diff. pp", "Review level"]
        data = [[Paragraph(escape(h), styles["Tiny"]) for h in headers]]
        for _, r in priority_table.head(12).iterrows():
            vals = [r["Group type"], r["Group"], str(r["Scored records"]), str(r["Low-score records"]), f'{r["Low-satisfaction rate (%)"]:.1f}', f'{r["Difference vs overall (pp)"]:.1f}', r["Review level"]]
            data.append([Paragraph(escape(str(v)), styles["Tiny"]) for v in vals])
        t = Table(data, colWidths=[20*mm,32*mm,13*mm,13*mm,17*mm,17*mm,43*mm], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.35,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),4)]))
        story += [t, Spacer(1, 4*mm)]
    story.append(Paragraph("Recommended actions", styles["Heading1"]))
    for item in insights[:8]:
        story.append(Paragraph(f"<b>{escape(str(item['Priority']))} — {escape(str(item['Area']))}</b>: {escape(str(item['Finding']))} Recommended action: {escape(str(item['Recommended action']))} Evidence: {escape(str(item['Evidence / denominator']))}", styles["SmallBody"]))
    story.append(Paragraph("Action register snapshot", styles["Heading1"]))
    if action_plan.empty:
        story.append(Paragraph("No action rows available.", styles["SmallBody"]))
    else:
        data = [[Paragraph(escape(x), styles["Tiny"]) for x in ["Action ID", "Priority", "Area", "Due date", "Status", "Owner"]]]
        for _, r in action_plan.head(15).iterrows():
            data.append([Paragraph(escape(str(r.get(k, "")))[:180], styles["Tiny"]) for k in ["Action_ID", "Priority", "Area", "Due_Date", "Status", "Responsible_Role"]])
        t = Table(data, colWidths=[28*mm,20*mm,48*mm,22*mm,24*mm,42*mm], repeatRows=1)
        t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.35,colors.grey),("VALIGN",(0,0),(-1,-1),"TOP"),("PADDING",(0,0),(-1,-1),4)]))
        story += [t, Spacer(1, 4*mm)]
    story.append(Paragraph("Data-quality check results", styles["Heading1"]))
    qdata = [[Paragraph("Check", styles["Tiny"]), Paragraph("Result", styles["Tiny"])]]
    for _, r in quality_df.iterrows():
        qdata.append([Paragraph(escape(str(r["Check"])), styles["Tiny"]), Paragraph(escape(str(r["Result"])), styles["Tiny"])])
    qt = Table(qdata, colWidths=[110*mm, 25*mm], repeatRows=1)
    qt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#17365D")),("TEXTCOLOR",(0,0),(-1,0),colors.white),("GRID",(0,0),(-1,-1),0.3,colors.grey),("PADDING",(0,0),(-1,-1),3)]))
    story += [qt, Spacer(1, 3*mm), Paragraph("Limitations: Group flags are descriptive and depend on sample size and selected filters. They do not prove causes, should not be used to deny assistance, and require validation with beneficiary feedback and programme records. Target/Actual values are not summed because repeated project totals may be double-counted.", styles["SmallBody"])]
    doc.build(story)
    return buffer.getvalue()


def build_data_dictionary(df):
    """Return privacy-conscious column metadata without exposing beneficiary values."""
    rows = []
    for col in df.columns:
        series = df[col]
        rows.append({
            "Column": str(col),
            "Data type": str(series.dtype),
            "Non-missing records": int(series.notna().sum()),
            "Missing records": int(series.isna().sum()),
            "Missing (%)": round(float(series.isna().mean() * 100), 2) if len(series) else 0.0,
            "Distinct non-missing values": int(series.nunique(dropna=True)),
            "Required for dashboard": "Yes" if col in REQUIRED_COLUMNS else "No",
        })
    return pd.DataFrame(rows)


def dataframe_fingerprint(df):
    """Stable SHA-256 fingerprint of the loaded table's values and column labels."""
    try:
        hashed = pd.util.hash_pandas_object(df, index=True, categorize=True).values.tobytes()
        schema = "\x1f".join(map(str, df.columns)).encode("utf-8")
        return hashlib.sha256(schema + hashed).hexdigest()
    except Exception:
        fallback = df.astype("string").fillna("<NA>").to_csv(index=True).encode("utf-8")
        return hashlib.sha256(fallback).hexdigest()


def build_audit_manifest(source, full_df, filtered_df, filter_summary):
    scores = valid_scores(filtered_df)
    manifest = {
        "Application": "AI_NGO Reporting System",
        "Version": "17.0",
        "Report generated (local time)": datetime.now().astimezone().isoformat(timespec="seconds"),
        "Source description": str(source),
        "Source dataset SHA-256 fingerprint": dataframe_fingerprint(full_df),
        "Source row count": int(len(full_df)),
        "Source column count": int(len(full_df.columns)),
        "Selected row count": int(len(filtered_df)),
        "Selected column count": int(len(filtered_df.columns)),
        "Valid satisfaction scores in selection": int(len(scores)),
        "Low-satisfaction records (score 1-2)": int(scores.isin(LOW_SCORE_VALUES).sum()),
        "Filter summary": filter_summary,
        "Interpretation note": "Fingerprint identifies the loaded table content; it is not proof that source records are accurate. Findings are descriptive and require human review.",
    }
    return pd.DataFrame([{"Field": k, "Value": v} for k, v in manifest.items()])


def build_ml_pipeline(x):
    numeric_cols = [c for c in x.columns if pd.api.types.is_numeric_dtype(x[c])]
    categorical_cols = [c for c in x.columns if c not in numeric_cols]
    transformers = []
    if numeric_cols:
        numeric_pipe = Pipeline([("imputer", SimpleImputer(strategy="median", add_indicator=True)), ("scale", StandardScaler())])
        transformers.append(("numeric", numeric_pipe, numeric_cols))
    if categorical_cols:
        categorical_pipe = Pipeline([("imputer", SimpleImputer(strategy="most_frequent")), ("onehot", OneHotEncoder(handle_unknown="ignore", min_frequency=2))])
        transformers.append(("categorical", categorical_pipe, categorical_cols))
    if not transformers:
        return None
    preprocessor = ColumnTransformer(transformers=transformers, remainder="drop")
    return Pipeline([("preprocess", preprocessor), ("model", LogisticRegression(max_iter=1500, class_weight="balanced", random_state=42))])



def aggregate_report_payload(df, province_table, activity_table, quality_table, filter_text):
    """Build a compact aggregate-only payload; never include beneficiary-level rows or IDs."""
    scores = valid_scores(df)
    low_n = int(scores.isin(LOW_SCORE_VALUES).sum())
    def clean_records(table, limit=15):
        if table is None or table.empty:
            return []
        safe = table.head(limit).copy()
        return safe.replace({np.nan: None}).to_dict(orient="records")
    quality_summary = []
    if quality_table is not None and not quality_table.empty:
        for _, row in quality_table.iterrows():
            label = str(row.get("Check", ""))
            value = row.get("Result")
            # Include only counts relevant to data-quality context; no row values.
            if any(term in label.lower() for term in ["missing", "invalid", "duplicate", "future", "negative", "exceeds", "unrecognized"]):
                try:
                    quality_summary.append({"check": label, "count": int(value)})
                except (TypeError, ValueError):
                    continue
    return {
        "reporting_scope": {
            "selected_record_count": int(len(df)),
            "filters": str(filter_text),
            "valid_satisfaction_count": int(len(scores)),
            "mean_satisfaction_out_of_5": round(float(scores.mean()), 2) if len(scores) else None,
            "low_satisfaction_definition": "score 1 or 2 on a 1–5 scale",
            "low_satisfaction_count": low_n,
            "low_satisfaction_rate_percent": round(low_n / len(scores) * 100, 2) if len(scores) else None,
        },
        "province_aggregates": clean_records(province_table),
        "activity_aggregates": clean_records(activity_table),
        "quality_issue_counts": quality_summary[:30],
        "safeguards": [
            "No beneficiary-level rows, names, IDs, contact details, or individual predictions are included.",
            "Group-level differences are descriptive and do not establish causes.",
            "Recommendations must request verification rather than assert unsupported explanations.",
            "Do not recommend denying assistance or making individual eligibility decisions from these metrics."
        ],
    }


def local_evidence_draft(payload):
    scope = payload["reporting_scope"]
    lines = [
        "# NGO Programme Management Brief (Evidence-Based Draft)",
        f"Generated: {datetime.now():%Y-%m-%d %H:%M}",
        "",
        "## Scope and key indicators",
        f"- Selected records: {scope['selected_record_count']:,}",
        f"- Valid satisfaction scores: {scope['valid_satisfaction_count']:,}",
        f"- Mean satisfaction: {scope['mean_satisfaction_out_of_5'] if scope['mean_satisfaction_out_of_5'] is not None else 'Not available'} / 5",
        f"- Low satisfaction (score 1–2): {scope['low_satisfaction_count']:,} records ({scope['low_satisfaction_rate_percent'] if scope['low_satisfaction_rate_percent'] is not None else 'Not available'}%)",
        f"- Filters: {scope['filters']}",
        "",
        "## Findings requiring attention",
    ]
    candidates=[]
    for group_type, key in [("Province", "province_aggregates"), ("Activity", "activity_aggregates")]:
        for row in payload.get(key, []):
            try:
                n=int(row.get("Scored records", 0) or 0)
                rate=float(row.get("Low-satisfaction rate (%)", 0) or 0)
                name=row.get(group_type, "Unknown")
                if n >= 20 and scope["low_satisfaction_rate_percent"] is not None and rate >= scope["low_satisfaction_rate_percent"] + 5:
                    candidates.append((rate, group_type, name, n))
            except (TypeError, ValueError):
                continue
    if candidates:
        for rate, group_type, name, n in sorted(candidates, reverse=True)[:8]:
            lines.append(f"- {group_type}: {name} — low-satisfaction rate {rate:.1f}% among {n} scored records. Validate sample feedback and delivery records before concluding why this pattern occurred.")
    else:
        lines.append("- No province/activity group met the current descriptive screening rule in the displayed aggregate tables. This does not prove that no programme issue exists.")
    quality = [x for x in payload.get("quality_issue_counts", []) if x.get("count", 0) > 0]
    lines += ["", "## Data-quality follow-up"]
    if quality:
        for item in sorted(quality, key=lambda x: x["count"], reverse=True)[:8]:
            lines.append(f"- Review {item['check']}: {item['count']:,} records/cases flagged.")
    else:
        lines.append("- No positive counts were identified in the selected data-quality checks; review the full data-quality panel for context.")
    lines += [
        "", "## Recommended management actions",
        "1. Sample and review low-satisfaction feedback with appropriate consent and confidentiality safeguards.",
        "2. Triangulate feedback with monitoring records, implementation timelines, and partner explanations.",
        "3. Assign a responsible owner and due date to each verified issue; document closure evidence.",
        "4. Compare future reporting periods while documenting changes in sample size, coverage, and data collection.",
        "", "## Limitations",
        "These are descriptive, filtered results. They do not establish causation or statistical significance. Missing or invalid data may affect results. Do not use these indicators or experimental ML outputs to deny assistance or determine individual eligibility.",
    ]
    return "\n".join(lines)


def call_openai_responses_api(api_key, model, payload):
    """Send aggregate-only JSON to the OpenAI Responses API; API key is not persisted."""
    request_body = {
        "model": model.strip() or "gpt-4.1-mini",
        "instructions": (
            "You are a careful humanitarian MEAL and programme reporting analyst. Draft a concise, professional management brief. "
            "Use only supplied aggregate figures. Do not invent facts, causes, quotes, sample details, or commitments. "
            "Distinguish observed findings from hypotheses and recommendations. Mention sample sizes and data limitations. "
            "Recommend verification, triangulation, and accountable follow-up. Never recommend denying assistance or making individual eligibility decisions from these metrics. "
            "Do not claim statistical significance unless it is explicitly provided. Use Markdown headings and concise language."
        ),
        "input": "Prepare an evidence-based NGO management report from this aggregate-only analysis payload:\n" + json.dumps(payload, ensure_ascii=False, default=str),
        "max_output_tokens": 1600,
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(request_body).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=75) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:1200]
        raise RuntimeError(f"API request failed (HTTP {exc.code}). Check API key, account access, model name, and API billing. Details: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the API. Check internet connectivity. Details: {exc.reason}") from exc
    chunks=[]
    for item in result.get("output", []):
        if item.get("type") == "message":
            for content in item.get("content", []):
                if content.get("type") == "output_text" and content.get("text"):
                    chunks.append(content["text"])
    text = "\n".join(chunks).strip()
    if not text:
        raise RuntimeError("The API returned no text output. Try another model or review the API response status.")
    return text

def local_data_answer(question, df, province_table, activity_table, priority_table, action_plan, filter_text):
    """Answer common questions using calculated aggregates only; no generative claims."""
    q = str(question or "").strip().lower()
    scores = valid_scores(df)
    low_n = int(scores.isin(LOW_SCORE_VALUES).sum())
    low_pct = (low_n / len(scores) * 100) if len(scores) else None
    mean_score = float(scores.mean()) if len(scores) else None
    lines = [f"**Analysis scope:** {len(df):,} selected records. {filter_text}."]

    if any(k in q for k in ["average satisfaction", "mean satisfaction", "average score", "mean score", "رضایت متوسط", "میانگین رضایت"]):
        lines.append(f"Mean satisfaction is **{mean_score:.2f}/5** based on **{len(scores):,} valid scores**." if mean_score is not None else "No valid satisfaction scores are available in the current selection.")
    elif any(k in q for k in ["low satisfaction", "dissatisfied", "score 1", "score 2", "نارضایتی", "رضایت پایین"]):
        lines.append(f"Low satisfaction (score 1–2) accounts for **{low_n:,} of {len(scores):,} valid scores** ({low_pct:.2f}%)." if low_pct is not None else "No valid satisfaction scores are available in the current selection.")
    elif any(k in q for k in ["which province", "province highest", "highest province", "worst province", "ولایت", "کدام ولایت"]):
        if province_table.empty:
            lines.append("No province-level satisfaction results are available for the current selection.")
        else:
            r=province_table.sort_values(["Low-satisfaction rate (%)", "Scored records"], ascending=[False, False]).iloc[0]
            lines.append(f"**{r['Province']}** has the highest observed low-satisfaction rate at **{r['Low-satisfaction rate (%)']:.2f}%** ({int(r['Low-score records']):,}/{int(r['Scored records']):,} valid scores). This is descriptive and does not explain the cause.")
    elif any(k in q for k in ["which activity", "activity highest", "highest activity", "worst activity", "فعالیت"]):
        if activity_table.empty:
            lines.append("No activity-level satisfaction results are available for the current selection.")
        else:
            r=activity_table.sort_values(["Low-satisfaction rate (%)", "Scored records"], ascending=[False, False]).iloc[0]
            lines.append(f"**{r['Activity']}** has the highest observed low-satisfaction rate at **{r['Low-satisfaction rate (%)']:.2f}%** ({int(r['Low-score records']):,}/{int(r['Scored records']):,} valid scores). This does not establish why the rate is higher.")
    elif any(k in q for k in ["how many records", "record count", "number of beneficiaries", "how many beneficiaries", "تعداد مستفید", "چند نفر"]):
        lines.append(f"There are **{len(df):,} selected records** in the current filtered dataset. Records should not automatically be interpreted as unique beneficiaries unless IDs are complete and unique.")
    elif any(k in q for k in ["priority", "priorities", "what should we do", "recommend", "اقدام", "اولویت"]):
        if priority_table.empty:
            lines.append("No group-level priority rows are available. Review the current filters and valid satisfaction records.")
        else:
            top=priority_table.head(5)
            lines.append("The current descriptive review queue is:")
            for _, r in top.iterrows():
                lines.append(f"- **{r['Group type']}: {r['Group']}** — {r['Review level']}; low-satisfaction rate {r['Low-satisfaction rate (%)']:.2f}% across {int(r['Scored records']):,} scored records.")
            lines.append("Recommended next step: validate the pattern, review feedback with appropriate safeguards, and assign actions only to confirmed issues.")
    elif any(k in q for k in ["missing data", "data quality", "invalid", "missing values", "کیفیت معلومات", "معلومات ناقص"]):
        lines.append("Review the **Data Quality** section for missing/invalid fields, duplicate IDs, date issues, and domain checks. These checks identify records for review; they do not prove that the remaining data are accurate.")
    elif any(k in q for k in ["target", "actual", "achievement", "هدف", "دستاورد"]):
        if not {"Target", "Actual"}.issubset(df.columns):
            lines.append("The selected workbook does not contain both Target and Actual columns.")
        else:
            target=pd.to_numeric(df["Target"], errors="coerce")
            actual=pd.to_numeric(df["Actual"], errors="coerce")
            valid=target.gt(0) & actual.ge(0)
            if valid.any():
                rates=(actual[valid]/target[valid])*100
                lines.append(f"There are **{int(valid.sum()):,} valid Target/Actual row pairs**. Median row-level achievement rate is **{rates.median():.2f}%**. Do not sum these fields unless each row is an independent planning/achievement observation; repeated totals can be double-counted.")
            else:
                lines.append("No valid Target/Actual pairs were found (Target must be >0 and Actual must be ≥0).")
    else:
        lines.append("The local assistant supports questions about average satisfaction, low satisfaction, highest-rate province/activity, record counts, review priorities, data quality, and Target vs Actual. For a free-form question, use the optional AI answer after reviewing the aggregate-only payload and authorizing external sharing.")
    lines.append("**Limitation:** results describe the selected data only. Group differences do not establish causation and must not be used to determine individual assistance eligibility.")
    return "\n\n".join(lines)


def call_openai_data_question(api_key, model, question, payload):
    """Send a user question and aggregate-only payload to the Responses API."""
    request_body = {
        "model": model.strip() or "gpt-4.1-mini",
        "instructions": (
            "You are a careful humanitarian MEAL and programme data analyst. Answer the user's question using only the aggregate JSON provided. "
            "Do not invent figures, causes, records, quotes, or commitments. If the payload cannot answer the question, say what is missing. "
            "Show relevant denominators and filters where available. Distinguish observed facts from hypotheses. Never infer individual characteristics, "
            "recommend denial of assistance, or use these metrics for individual eligibility. Keep the answer concise and actionable."
        ),
        "input": "User question:\n" + str(question) + "\n\nAggregate-only dataset summary:\n" + json.dumps(payload, ensure_ascii=False, default=str),
        "max_output_tokens": 1100,
    }
    req = urllib.request.Request(
        "https://api.openai.com/v1/responses",
        data=json.dumps(request_body).encode("utf-8"),
        headers={"Authorization": f"Bearer {api_key.strip()}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=75) as response:
            result = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")[:1000]
        raise RuntimeError(f"API request failed (HTTP {exc.code}). Check API key, billing, account access, and model name. Details: {body}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not reach the API. Check internet connectivity. Details: {exc.reason}") from exc
    chunks=[]
    for item in result.get("output", []):
        if item.get("type") == "message":
            for content in item.get("content", []):
                if content.get("type") == "output_text" and content.get("text"):
                    chunks.append(content["text"])
    answer="\n".join(chunks).strip()
    if not answer:
        raise RuntimeError("The API returned no text answer. Review the API response or try another supported model.")
    return answer

with st.sidebar:
    st.markdown("""<div style="padding:8px 4px 16px 4px;">
<div style="font-size:.68rem;letter-spacing:.18em;font-weight:800;color:#8DE5DA;text-transform:uppercase;">Workspace</div>
<div style="font-size:1.24rem;font-weight:820;color:white;letter-spacing:-.03em;margin-top:5px;">AI NGO <span style="color:#8DE5DA;">INTELLIGENCE</span></div>
<div style="font-size:.77rem;color:#BBD0DE;margin-top:5px;">Monitoring &amp; management suite</div>
</div>""", unsafe_allow_html=True)
    st.markdown("<div style='font-size:.7rem;font-weight:750;letter-spacing:.12em;text-transform:uppercase;color:#8EA9BB;margin:4px 0 6px;'>Navigation</div>", unsafe_allow_html=True)
    section = st.radio("Open section", ["Executive Dashboard", "Priority Ranking", "Programme Performance", "Early Warning & Narrative", "Insights & Action Plan", "Province & Activity", "Data Quality", "Data Governance & Audit", "Generate Reports", "Ask Your Data", "Report Studio", "Indicator Tracker", "AI-Assisted Report Drafting", "Trends & ML Risk"])
    uploaded = st.file_uploader("Optional: upload an Excel workbook", type=["xlsx"])
    st.caption("Uploaded workbooks are used for this session only.")

try:
    raw, source = load_workbook(uploaded)
except Exception as exc:
    st.error(f"Unable to read the Excel workbook: {exc}")
    st.stop()
if raw is None:
    st.error(f"Dataset not found. Expected file: {DATA_FILE}")
    st.info("Place beneficiary_data.xlsx in the data folder next to this app, or upload it in the sidebar. The workbook is not bundled with this ZIP.")
    st.stop()

df, missing = prepare_data(raw)
if missing:
    st.error("Required columns missing: " + ", ".join(missing))
    st.write("Columns found:", list(df.columns))
    st.stop()
if df.empty:
    st.error("The workbook contains no records.")
    st.stop()

# Interactive filters; date filtering excludes invalid dates from the filtered date range only when a range is applied.
st.sidebar.divider()
st.sidebar.subheader("Analysis filters")
filtered = df.copy()
provinces = sorted(filtered["Province"].dropna().astype(str).unique().tolist())
activities = sorted(filtered["Activity"].dropna().astype(str).unique().tolist())
selected_provinces = st.sidebar.multiselect("Province", provinces, default=provinces)
selected_activities = st.sidebar.multiselect("Activity", activities, default=activities)
if provinces:
    filtered = filtered[filtered["Province"].astype("string").isin(selected_provinces)]
if activities:
    filtered = filtered[filtered["Activity"].astype("string").isin(selected_activities)]
valid_dates = filtered["Date"].dropna()
if not valid_dates.empty:
    min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
    date_range = st.sidebar.date_input("Date range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
    if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered = filtered[filtered["Date"].isna() | filtered["Date"].dt.date.between(start_date, end_date)]
else:
    st.sidebar.caption("Date filter unavailable: no valid dates in the selected province/activity records.")

filter_summary = f"Provinces: {', '.join(selected_provinces) if selected_provinces else 'none'}; Activities: {', '.join(selected_activities) if selected_activities else 'none'}"
quality = quality_checks(filtered)
province_table = grouped_analysis(filtered, "Province")
activity_table = grouped_analysis(filtered, "Activity")
insights = make_insights(filtered, province_table, activity_table)
action_plan = load_persistent_action_plan(insights)
scores = valid_scores(filtered)
low_count = int(scores.isin(LOW_SCORE_VALUES).sum())
low_rate = low_count / len(scores) * 100 if len(scores) else 0
priority_table = build_priority_ranking(province_table, activity_table, low_rate)
data_dictionary = build_data_dictionary(df)
audit_manifest = build_audit_manifest(source, df, filtered, filter_summary)

if filtered.empty:
    st.warning("No records match the selected filters. Change the filters in the sidebar.")
else:
    if section == "Executive Dashboard":
        st.subheader("Executive Overview")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Selected records", f"{len(filtered):,}")
        c2.metric("Valid satisfaction scores", f"{len(scores):,}")
        c3.metric("Average satisfaction", f"{scores.mean():.2f}/5" if len(scores) else "N/A")
        c4.metric("Low-satisfaction rate", f"{low_rate:.1f}%" if len(scores) else "N/A")
        a, b = st.columns(2)
        with a:
            distribution = scores.value_counts().reindex([1, 2, 3, 4, 5], fill_value=0).rename_axis("Score").reset_index(name="Records")
            st.plotly_chart(px.bar(distribution, x="Score", y="Records", text="Records", title="Satisfaction score distribution"), use_container_width=True, config={"displayModeBar": False})
        with b:
            pcounts = filtered["Province"].fillna("Missing").value_counts().rename_axis("Province").reset_index(name="Records")
            st.plotly_chart(px.bar(pcounts, x="Province", y="Records", title="Records by province"), use_container_width=True, config={"displayModeBar": False})
        trend = monthly_trend(filtered)
        if not trend.empty:
            st.markdown("#### Satisfaction trend over time")
            st.plotly_chart(px.line(trend, x="Month", y="Average satisfaction", markers=True, title="Monthly average satisfaction (valid scores only)"), use_container_width=True, config={"displayModeBar": False})
        st.markdown("#### Executive management summary")
        for paragraph in make_management_narrative(filtered, province_table, activity_table):
            st.write(paragraph)
        st.markdown("#### Priority insights")
        for item in insights:
            st.markdown(f"**{item['Priority']} — {item['Area']}**")
            st.write(item["Finding"])
        st.download_button("Export selected records (CSV)", filtered.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_selected_records_{datetime.now():%Y%m%d}.csv", mime="text/csv")

    elif section == "Priority Ranking":
        st.subheader("Evidence-based Management Priority Ranking")
        st.caption("Ranking is transparent and descriptive. It uses sample size, low-satisfaction rate, and percentage-point difference from the selected overall rate; it does not infer causes.")
        st.info("Review priority rule: at least 20 valid satisfaction records and a low-satisfaction rate at least 5 percentage points above the selected overall rate. Small groups are explicitly marked for monitoring rather than promoted as confirmed priorities.")
        if priority_table.empty:
            st.warning("No group-level priority ranking is available for the current filters.")
        else:
            high_n = int(priority_table["Review level"].isin(["High review priority", "Review priority"]).sum())
            small_n = int((priority_table["Review level"] == "Monitor — small sample").sum())
            m1, m2, m3 = st.columns(3)
            m1.metric("Groups needing review", f"{high_n:,}")
            m2.metric("Groups assessed", f"{len(priority_table):,}")
            m3.metric("Small-sample groups", f"{small_n:,}")
            st.dataframe(priority_table, use_container_width=True, hide_index=True)
            st.plotly_chart(px.bar(priority_table, x="Group", y="Low-satisfaction rate (%)", color="Review level", facet_col="Group type", title="Low-satisfaction rates with review levels"), use_container_width=True, config={"displayModeBar": False})
            ranking_xlsx = BytesIO()
            with pd.ExcelWriter(ranking_xlsx, engine="openpyxl") as writer:
                priority_table.to_excel(writer, sheet_name="Priority Ranking", index=False)
            st.download_button("Download priority ranking (Excel)", ranking_xlsx.getvalue(), file_name=f"v20_0_priority_ranking_{datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.download_button("Download priority ranking (CSV)", priority_table.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_priority_ranking_{datetime.now():%Y%m%d}.csv", mime="text/csv")
        st.markdown("#### How to use this ranking")
        st.write("1. Check denominators and review the underlying records. 2. Compare patterns with programme delivery data and documented feedback. 3. Assign an owner and due date only after the issue is verified. 4. Reassess in the next reporting cycle.")

    elif section == "Programme Performance":
        st.subheader("Programme Performance: Target vs Actual")
        if not {"Target", "Actual"}.issubset(filtered.columns):
            st.warning("This workbook does not contain both Target and Actual columns. Upload a workbook with these fields to use programme performance analysis.")
        else:
            st.write("Target = planned number of beneficiaries to reach; Actual = number reached so far, as clarified by the project owner.")
            st.warning("Data-granularity check: this workbook has beneficiary-level IDs. If Target/Actual values are repeated project or activity totals on each beneficiary row, summing them would double-count. By default, this section reports per-record averages and median achievement rates rather than summed totals.")
            performance_rows, overall_perf = programme_performance(filtered)
            if performance_rows.empty:
                st.info("No valid Target/Actual pairs found. Target must be greater than zero and Actual must be non-negative numeric data.")
            else:
                p1, p2, p3, p4 = st.columns(4)
                p1.metric("Valid Target/Actual observations", f"{len(performance_rows):,}")
                p2.metric("Mean target per record", f"{performance_rows['Target'].mean():,.2f}")
                p3.metric("Mean actual per record", f"{performance_rows['Actual'].mean():,.2f}")
                p4.metric("Median achievement rate", f"{performance_rows['Achievement_Rate_Pct'].median():.1f}%")
                st.caption("Achievement rate per record = Actual ÷ Target × 100. Negative gaps mean Actual exceeds Target; rates above 100% are retained and should be checked against the source definitions.")
                show_cols = [c for c in ["Beneficiary_ID", "Province", "District", "Activity", "Target", "Actual", "Achievement_Gap", "Achievement_Rate_Pct"] if c in performance_rows.columns]
                st.markdown("#### Record-level calculations")
                st.dataframe(performance_rows[show_cols].round(2), use_container_width=True, hide_index=True)
                group_options = [c for c in ["Province", "District", "Activity"] if c in performance_rows.columns]
                if group_options:
                    group_by = st.selectbox("Summarize by", group_options)
                    _, grouped_perf = programme_performance(filtered, group_by)
                    st.markdown(f"#### Performance summary by {group_by}")
                    st.dataframe(grouped_perf, use_container_width=True, hide_index=True)
                    if not grouped_perf.empty:
                        st.plotly_chart(px.bar(grouped_perf, x=group_by, y="Median Achievement Rate (%)", title=f"Median target achievement rate by {group_by}"), use_container_width=True, config={"displayModeBar": False})
                st.markdown("#### Optional aggregate totals")
                totals_confirm = st.checkbox("I confirm each row is an independent planning/achievement observation, not a repeated beneficiary-level copy of a shared target.")
                if totals_confirm:
                    total_target = performance_rows["Target"].sum()
                    total_actual = performance_rows["Actual"].sum()
                    t1, t2, t3 = st.columns(3)
                    t1.metric("Sum of targets", f"{total_target:,.0f}")
                    t2.metric("Sum of actual reached", f"{total_actual:,.0f}")
                    t3.metric("Aggregate achievement rate", f"{(total_actual / total_target * 100):.1f}%" if total_target > 0 else "N/A")
                    st.caption("These totals are shown because you confirmed that each row is an independent observation. Validate this assumption against the data dictionary before reporting externally.")
                else:
                    st.info("Aggregate totals are hidden to prevent possible double-counting. Use the row-level and grouped summaries until the data granularity is confirmed.")
                out = BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    performance_rows.to_excel(writer, sheet_name="Record Performance", index=False)
                    if group_options:
                        _, export_group = programme_performance(filtered, group_options[0])
                        export_group.to_excel(writer, sheet_name=f"By {group_options[0]}"[:31], index=False)
                st.download_button("Download programme performance (Excel)", out.getvalue(), file_name=f"v20_0_programme_performance_{datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif section == "Early Warning & Narrative":
        st.subheader("Early Warning & Management Narrative")
        st.caption("Transparent rules summarize observed patterns. Flags support review; they are not causal findings, automated decisions, or beneficiary eligibility criteria.")
        narrative = make_management_narrative(filtered, province_table, activity_table)
        st.markdown("#### Draft management narrative")
        st.info("This narrative is generated from explicit rules and observed data; it does not call an external generative-AI service.")
        for paragraph in narrative:
            st.write(paragraph)
        narrative_text = "\n\n".join(narrative)
        st.download_button("Download narrative (TXT)", narrative_text.encode("utf-8"), file_name=f"v20_0_management_narrative_{datetime.now():%Y%m%d}.txt", mime="text/plain")
        st.markdown("#### Early-warning queue")
        valid_score = valid_scores(filtered)
        overall_rate = (valid_score.isin(LOW_SCORE_VALUES).mean() * 100) if len(valid_score) else 0.0
        province_flags = risk_flags(province_table, overall_rate, "Province")
        activity_flags = risk_flags(activity_table, overall_rate, "Activity")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Province review queue**")
            st.dataframe(province_flags, use_container_width=True, hide_index=True)
        with col2:
            st.markdown("**Activity review queue**")
            st.dataframe(activity_flags, use_container_width=True, hide_index=True)
        combined = []
        for table, group_col, group_type in [(province_flags, "Province", "Province"), (activity_flags, "Activity", "Activity")]:
            if not table.empty:
                subset = table[table["Risk flag"] == "Review priority"]
                for _, row in subset.iterrows():
                    combined.append({"Group type": group_type, "Group": row[group_col], "Scored records": row["Scored records"], "Low-satisfaction rate (%)": row["Low-satisfaction rate (%)"], "Difference vs overall (pp)": row["Difference vs overall (pp)"], "Suggested follow-up": "Review feedback sample, verify delivery records, and assign an action owner."})
        queue = pd.DataFrame(combined, columns=["Group type", "Group", "Scored records", "Low-satisfaction rate (%)", "Difference vs overall (pp)", "Suggested follow-up"])
        if queue.empty:
            st.success("No groups currently meet the descriptive early-warning rule, or group sample sizes are insufficient.")
        else:
            st.dataframe(queue, use_container_width=True, hide_index=True)
            st.download_button("Download early-warning queue (CSV)", queue.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_early_warning_queue_{datetime.now():%Y%m%d}.csv", mime="text/csv")
        st.markdown("#### Management action register")
        st.caption("Edit cells directly, then save. To remove a row, select its Action_ID below and confirm deletion. On Streamlit Community Cloud, CSV storage may be temporary and should not be treated as a durable database.")
        edited_actions = st.data_editor(
            action_plan,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key="v22_0_early_warning_action_register",
        )
        save_col, export_col = st.columns([1, 1])
        with save_col:
            if st.button("Save management action register", type="primary", key="v22_0_save_action_register"):
                try:
                    save_action_plan(edited_actions)
                    st.success(f"Saved {len(edited_actions):,} action rows to {ACTION_PLAN_FILE}.")
                except Exception as exc:
                    st.error(f"Could not save the action register: {exc}")
        with export_col:
            st.download_button(
                "Export action register (CSV)",
                edited_actions.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"v22_0_action_register_{datetime.now():%Y%m%d}.csv",
                mime="text/csv",
                key="v22_0_export_action_register",
            )

        st.markdown("##### Delete an action")
        if edited_actions.empty:
            st.info("There are no action rows to delete.")
        elif "Action_ID" not in edited_actions.columns:
            st.error("Deletion is unavailable because the Action_ID column is missing.")
        else:
            delete_options = edited_actions["Action_ID"].astype(str).tolist()
            selected_delete_id = st.selectbox(
                "Select the Action_ID to delete",
                options=delete_options,
                key="v22_0_delete_action_id",
            )
            typed_delete_id = st.text_input(
                f"Type {selected_delete_id} to confirm deletion",
                key="v22_0_confirm_delete_action_id",
            )
            if st.button("Delete selected action", type="secondary", disabled=(typed_delete_id.strip() != str(selected_delete_id)), key="v22_0_delete_action_button"):
                try:
                    remaining_actions = edited_actions[edited_actions["Action_ID"].astype(str) != str(selected_delete_id)].copy()
                    save_action_plan(remaining_actions)
                    st.success(f"Deleted {selected_delete_id}. {len(remaining_actions):,} action rows remain and have been saved.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"Could not delete the selected action: {exc}")

    elif section == "Insights & Action Plan":
        st.subheader("Evidence-based Insights & Management Action Plan")
        st.info("Insights use explicit statistical rules. They are not an LLM diagnosis and do not infer unobserved causes.")
        for item in insights:
            with st.container(border=True):
                st.markdown(f"### {item['Priority']} priority: {item['Area']}")
                st.write("**Finding:** " + item["Finding"])
                st.write("**Recommended action:** " + item["Recommended action"])
                st.caption("Evidence / denominator: " + item["Evidence / denominator"])
        st.markdown("#### Action plan monitoring")
        monitor = action_plan.copy()
        if not monitor.empty:
            today = pd.Timestamp(date.today())
            due = pd.to_datetime(monitor["Due_Date"], errors="coerce")
            statuses = monitor["Status"].astype(str).str.strip().str.lower()
            completed = statuses.isin(["completed", "closed", "done"])
            overdue = due.notna() & (due < today) & ~completed
            open_actions = ~completed
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Total actions", f"{len(monitor):,}")
            a2.metric("Open actions", f"{int(open_actions.sum()):,}")
            a3.metric("Overdue actions", f"{int(overdue.sum()):,}")
            a4.metric("Completed actions", f"{int(completed.sum()):,}")
            st.caption("Overdue status is calculated only when Due_Date is a valid date earlier than today and Status is not Completed/Closed/Done. Blank deadlines are not classified as overdue.")
        st.markdown("#### Editable action plan")
        edited = st.data_editor(action_plan, use_container_width=True, hide_index=True, num_rows="dynamic", key="v20_0_action_plan")
        if st.button("Save action plan locally", key="save_v17_action_plan"):
            try:
                save_action_plan(edited)
                st.success(f"Saved {len(edited):,} action rows locally.")
            except Exception as exc:
                st.error(f"Could not save the action plan: {exc}")
        st.download_button("Download action plan (CSV)", edited.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_management_action_plan_{datetime.now():%Y%m%d}.csv", mime="text/csv")

    elif section == "Province & Activity":
        tab1, tab2 = st.tabs(["By Province", "By Activity"])
        with tab1:
            st.dataframe(province_table, use_container_width=True, hide_index=True)
            if not province_table.empty:
                st.plotly_chart(px.bar(province_table, x="Province", y="Low-satisfaction rate (%)", hover_data=["Scored records", "Low-score records"], title="Low-satisfaction rate by province"), use_container_width=True, config={"displayModeBar": False})
        with tab2:
            st.dataframe(activity_table, use_container_width=True, hide_index=True)
            if not activity_table.empty:
                st.plotly_chart(px.bar(activity_table, x="Activity", y="Low-satisfaction rate (%)", hover_data=["Scored records", "Low-score records"], title="Low-satisfaction rate by activity"), use_container_width=True, config={"displayModeBar": False})
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            province_table.to_excel(writer, sheet_name="By Province", index=False)
            activity_table.to_excel(writer, sheet_name="By Activity", index=False)
        st.download_button("Download province/activity analysis (Excel)", output.getvalue(), file_name=f"v20_0_group_analysis_{datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif section == "Data Quality":
        st.subheader("Data Quality Diagnostics")
        st.caption("Checks identify potential issues; they do not prove that records are substantively accurate.")
        st.dataframe(quality, use_container_width=True, hide_index=True)
        st.markdown("#### Selected records preview")
        st.dataframe(filtered.head(100), use_container_width=True, hide_index=True)
        st.download_button("Download quality checks (CSV)", quality.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_data_quality_{datetime.now():%Y%m%d}.csv", mime="text/csv")

    elif section == "Data Governance & Audit":
        st.subheader("Data Governance & Reproducibility Audit")
        st.caption("This section documents the dataset schema, selected analysis scope, and a content fingerprint for traceability. The fingerprint is not a validation of data truth or source authenticity.")
        st.markdown("#### Analysis run manifest")
        st.dataframe(audit_manifest, use_container_width=True, hide_index=True)
        manifest_dict = dict(zip(audit_manifest["Field"], audit_manifest["Value"]))
        st.download_button("Download audit manifest (JSON)", json.dumps(manifest_dict, ensure_ascii=False, indent=2, default=str).encode("utf-8"), file_name=f"v20_0_audit_manifest_{datetime.now():%Y%m%d_%H%M%S}.json", mime="application/json")
        st.download_button("Download audit manifest (CSV)", audit_manifest.to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_audit_manifest_{datetime.now():%Y%m%d_%H%M%S}.csv", mime="text/csv")
        st.markdown("#### Data dictionary")
        st.write("The dictionary shows column names, data types, missingness, distinct-value counts, and whether each field is required. It intentionally excludes sample beneficiary values.")
        st.dataframe(data_dictionary, use_container_width=True, hide_index=True)
        st.download_button("Download data dictionary (CSV)", data_dictionary.to_csv(index=False).encode("utf-8-sig"), file_name="v20_0_data_dictionary.csv", mime="text/csv")
        audit_xlsx = BytesIO()
        with pd.ExcelWriter(audit_xlsx, engine="openpyxl") as writer:
            data_dictionary.to_excel(writer, sheet_name="Data Dictionary", index=False)
            audit_manifest.to_excel(writer, sheet_name="Audit Manifest", index=False)
            quality.to_excel(writer, sheet_name="Quality Checks", index=False)
        st.download_button("Download governance workbook (Excel)", audit_xlsx.getvalue(), file_name=f"v20_0_governance_audit_{datetime.now():%Y%m%d_%H%M%S}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("#### Review checklist")
        st.markdown("- Confirm the source workbook and reporting period before sharing results.\n- Check the filter summary and selected record count.\n- Review missingness and duplicate identifiers with the data owner.\n- Keep beneficiary-level exports in access-controlled storage.\n- Treat the fingerprint as a reproducibility aid, not proof of accuracy.")

    elif section == "Generate Reports":
        st.subheader("Automated Management Reports")
        st.write("Reports use the currently filtered records and show the filter summary.")
        if st.button("Prepare Word report", type="primary"):
            try:
                data = word_report(filtered, source, quality, insights, action_plan, filter_summary)
                st.download_button("Download Word report", data=data, file_name=f"v20_0_management_report_{datetime.now():%Y%m%d_%H%M}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except Exception as exc:
                st.error(f"Word report generation failed: {exc}")
        if st.button("Prepare Excel workbook"):
            try:
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    filtered.to_excel(writer, sheet_name="Selected Data", index=False)
                    quality.to_excel(writer, sheet_name="Data Quality", index=False)
                    province_table.to_excel(writer, sheet_name="By Province", index=False)
                    activity_table.to_excel(writer, sheet_name="By Activity", index=False)
                    pd.DataFrame(insights).to_excel(writer, sheet_name="Insights", index=False)
                    priority_table.to_excel(writer, sheet_name="Priority Ranking", index=False)
                    action_plan.to_excel(writer, sheet_name="Action Plan", index=False)
                    data_dictionary.to_excel(writer, sheet_name="Data Dictionary", index=False)
                    audit_manifest.to_excel(writer, sheet_name="Audit Manifest", index=False)
                st.download_button("Download Excel workbook", data=output.getvalue(), file_name=f"v20_0_management_analysis_{datetime.now():%Y%m%d_%H%M}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as exc:
                st.error(f"Excel report generation failed: {exc}")
        if st.button("Prepare PDF management brief"):
            try:
                pdf_data = pdf_report(filtered, source, quality, insights, action_plan, filter_summary, priority_table)
                st.download_button("Download PDF management brief", data=pdf_data, file_name=f"v20_0_management_brief_{datetime.now():%Y%m%d_%H%M}.pdf", mime="application/pdf")
            except Exception as exc:
                st.error(f"PDF report generation failed: {exc}")


    elif section == "Ask Your Data":
        st.subheader("Ask Your Data")
        st.markdown("Ask questions about the current filtered dataset. Local answers are calculated from the dashboard tables. Optional AI answers use aggregate summaries only.")
        st.info("The local assistant is free and works without an API key. AI answers require separately billed API access and explicit authorization.")
        examples = [
            "What is the average satisfaction?",
            "How many records have low satisfaction?",
            "Which province has the highest low-satisfaction rate?",
            "Which activity has the highest low-satisfaction rate?",
            "What are the current review priorities?",
            "What data-quality issues should I review?",
            "What does Target vs Actual show?",
            "How many records are in the current selection?",
        ]
        chosen_question = st.selectbox("Choose a question", examples, key="v20_local_question")
        custom_question = st.text_input("Or write your own question", placeholder="e.g., What should management review first?", key="v20_custom_question")
        question = custom_question.strip() or chosen_question
        payload = aggregate_report_payload(filtered, province_table, activity_table, quality, filter_summary)
        with st.expander("Review the aggregate-only summary", expanded=False):
            st.json(payload)
        left, right = st.columns(2)
        with left:
            if st.button("Answer locally", type="primary", key="v20_answer_local"):
                st.session_state["v20_data_answer"] = local_data_answer(question, filtered, province_table, activity_table, priority_table, action_plan, filter_summary)
                st.session_state["v20_data_answer_source"] = "Local calculated answer — no external AI service"
        with right:
            st.caption("For open-ended interpretation, the optional API answer sends the question and displayed aggregate summary to OpenAI.")
        with st.expander("Optional AI answer via OpenAI API", expanded=False):
            api_key = st.text_input("OpenAI API key (used for this request only)", type="password", key="v20_qa_api_key")
            model_name = st.text_input("API model", value="gpt-4.1-mini", key="v20_qa_model")
            consent = st.checkbox("I reviewed the aggregate summary and authorize sending it with my question to the OpenAI API.", key="v20_qa_consent")
            if st.button("Ask AI", key="v20_ask_ai"):
                if not question.strip():
                    st.error("Enter or select a question first.")
                elif not api_key.strip():
                    st.error("Enter an API key or use the local answer option.")
                elif not consent:
                    st.error("Review the summary and authorize external sharing before making an API request.")
                else:
                    with st.spinner("Preparing an answer from aggregate statistics…"):
                        try:
                            st.session_state["v20_data_answer"] = call_openai_data_question(api_key, model_name, question, payload)
                            st.session_state["v20_data_answer_source"] = f"OpenAI API answer — model {model_name.strip()}"
                            st.success("Answer generated. Verify all figures against the dashboard before use.")
                        except Exception as exc:
                            st.error(f"AI question failed: {exc}")
        if st.session_state.get("v20_data_answer"):
            st.markdown("#### Answer")
            st.caption(st.session_state.get("v20_data_answer_source", "Answer"))
            st.markdown(st.session_state["v20_data_answer"])
            st.download_button("Download answer (Markdown)", st.session_state["v20_data_answer"].encode("utf-8"), file_name=f"v20_0_data_answer_{datetime.now():%Y%m%d_%H%M}.md", mime="text/markdown")
            st.download_button("Download answer (TXT)", st.session_state["v20_data_answer"].encode("utf-8"), file_name=f"v20_0_data_answer_{datetime.now():%Y%m%d_%H%M}.txt", mime="text/plain")
        st.markdown("#### Safe-use reminders")
        st.markdown("- Local answers are deterministic and only cover supported question types.\n- AI answers can be wrong or overgeneralize; verify each figure and interpretation.\n- Do not send beneficiary-level rows, names, IDs, phone numbers, or case details to an external model.\n- Do not use group-level metrics or AI output to deny assistance or determine individual eligibility.")

    elif section == "Report Studio":
        st.subheader("Report Studio")
        st.markdown("Build, customize, save, and revisit locally generated report drafts from the current filtered analysis.")
        st.info("This studio creates a rule-based draft locally. It does not call an external AI service and does not save beneficiary-level records.")
        payload = aggregate_report_payload(filtered, province_table, activity_table, quality, filter_summary)
        audience = st.selectbox("Report audience", ["Programme Management", "MEAL Team", "Senior Leadership", "Donor / External Stakeholder"], key="v21_report_audience")
        template = st.selectbox("Report template", ["Monthly Programme Review", "Satisfaction & Feedback Brief", "Data Quality Review", "Executive Summary"], key="v21_report_template")
        language = st.selectbox("Report language", ["English", "English with Dari key-points placeholder"], key="v21_report_language")
        reporting_period = st.text_input("Reporting period / report label", value=datetime.now().strftime("%B %Y"), key="v21_report_period")
        extra_context = st.text_area("Optional context for the report (do not include personal or beneficiary-identifying information)", placeholder="e.g., Quarterly programme review; note any known data-collection limitations.", key="v21_report_context")
        if st.button("Generate report draft", type="primary", key="v21_generate_report"):
            core = local_evidence_draft(payload)
            audience_notes = {
                "Programme Management": "Focus: operational decisions, verified issues, owners, and deadlines.",
                "MEAL Team": "Focus: indicator definitions, denominators, data quality, triangulation, and learning questions.",
                "Senior Leadership": "Focus: concise strategic overview, material risks, and decisions requiring leadership attention.",
                "Donor / External Stakeholder": "Focus: factual, cautious reporting. Validate all figures and contractual reporting requirements before external submission."
            }
            template_notes = {
                "Monthly Programme Review": "Use this draft to structure the monthly programme review and assign follow-up actions.",
                "Satisfaction & Feedback Brief": "Prioritize satisfaction patterns, feedback follow-up, and the need to verify underlying reasons.",
                "Data Quality Review": "Prioritize missing, invalid, duplicate, or otherwise flagged records and document corrective actions.",
                "Executive Summary": "Keep the summary concise; retain key figures, priorities, limitations, and decisions needed."
            }
            title = f"{template} — {reporting_period.strip() or datetime.now().strftime('%B %Y')}"
            additions = [
                f"# {title}",
                f"**Audience:** {audience}",
                f"**Generated:** {datetime.now():%Y-%m-%d %H:%M}",
                f"**Selected records:** {payload['reporting_scope']['selected_record_count']:,}",
                f"**Applied filters:** {payload['reporting_scope']['filters']}",
                "", "## Purpose and use", template_notes[template], audience_notes[audience],
            ]
            if extra_context.strip():
                additions += ["", "## Context supplied by report author", extra_context.strip()]
            if language == "English with Dari key-points placeholder":
                additions += ["", "## Dari key points", "Add a human-reviewed Dari translation of the verified key findings before sharing."]
            # Remove the generic heading and duplicate generated/scope lines from the reusable core draft.
            core_lines = core.splitlines()
            if core_lines and core_lines[0].startswith("# "):
                core_lines = core_lines[1:]
            draft = "\n".join(additions + ["", "---", ""] + core_lines)
            st.session_state["v21_report_draft"] = draft
            st.session_state["v21_report_filename"] = re.sub(r"[^A-Za-z0-9_-]+", "_", f"{reporting_period}_{template}").strip("_")[:80] or "management_report"
        if st.session_state.get("v21_report_draft"):
            st.markdown("#### Review and edit draft")
            edited = st.text_area("Report draft (editable)", value=st.session_state["v21_report_draft"], height=430, key="v21_editable_report_draft")
            st.session_state["v21_report_draft"] = edited
            st.download_button("Download Markdown", edited.encode("utf-8"), file_name=f"{st.session_state.get('v21_report_filename','management_report')}.md", mime="text/markdown")
            st.download_button("Download TXT", edited.encode("utf-8"), file_name=f"{st.session_state.get('v21_report_filename','management_report')}.txt", mime="text/plain")
            if st.button("Save draft to local report history", key="v21_save_report"):
                try:
                    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    base = st.session_state.get("v21_report_filename", "management_report")
                    target = REPORTS_DIR / f"{timestamp}_{base}.md"
                    target.write_text(edited, encoding="utf-8")
                    st.success(f"Draft saved locally: reports/{target.name}")
                except Exception as exc:
                    st.error(f"Could not save report draft: {exc}")
        st.markdown("#### Local report history")
        try:
            REPORTS_DIR.mkdir(parents=True, exist_ok=True)
            saved_reports = sorted(REPORTS_DIR.glob("*.md"), key=lambda p: p.stat().st_mtime, reverse=True)
            if saved_reports:
                choices = {f"{p.name} — {datetime.fromtimestamp(p.stat().st_mtime):%Y-%m-%d %H:%M}": p for p in saved_reports[:30]}
                selected_saved = st.selectbox("Saved reports (latest 30)", list(choices.keys()), key="v21_saved_report_choice")
                saved_path = choices[selected_saved]
                st.download_button("Download selected saved report", saved_path.read_bytes(), file_name=saved_path.name, mime="text/markdown")
                with st.expander("Preview selected saved report"):
                    st.markdown(saved_path.read_text(encoding="utf-8"))
            else:
                st.caption("No saved report drafts yet. Generate and save a draft to create local history.")
        except Exception as exc:
            st.warning(f"Could not load report history: {exc}")
        st.markdown("#### Review checklist")
        st.markdown("- Verify every statistic and reporting denominator against the dashboard.\n- Confirm that the reporting period and filters are correct.\n- Treat descriptive differences as signals for investigation, not proof of causes.\n- Review recommendations and remove unsupported claims before sharing.\n- Do not include beneficiary names, IDs, contact details, or case-level information.")

    elif section == "Indicator Tracker":
        st.subheader("Indicator Tracker & Results Framework")
        st.markdown("Maintain a separate indicator registry for programme logframes and monitoring plans. This registry is not inferred from beneficiary-level Target/Actual columns.")
        st.info("Enter the approved indicator definition, unit, baseline, target, latest actual, reporting period, owner, and status. Progress is a planning aid; confirm the indicator's calculation method and denominator before reporting.")
        if "v22_indicator_registry" not in st.session_state:
            st.session_state["v22_indicator_registry"] = load_indicator_registry()
        current_registry = st.session_state["v22_indicator_registry"].copy()
        if current_registry.empty:
            blank_row = {c: "" for c in INDICATOR_COLUMNS}
            for numeric_col in ["Baseline", "Target", "Actual"]:
                blank_row[numeric_col] = None
            blank_row["Unit"] = "Number"
            blank_row["Direction"] = "Higher is better"
            blank_row["Status"] = "Needs review"
            current_registry = pd.DataFrame([blank_row], columns=INDICATOR_COLUMNS)
        edited_registry = st.data_editor(
            current_registry,
            key="v22_indicator_editor",
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Indicator_ID": st.column_config.TextColumn("Indicator ID", help="Leave blank to auto-assign an ID when saving."),
                "Indicator_Name": st.column_config.TextColumn("Indicator name", required=True, width="large"),
                "Definition": st.column_config.TextColumn("Definition / calculation", width="large"),
                "Unit": st.column_config.SelectboxColumn("Unit", options=["Number", "%", "People", "Households", "Sessions", "Score", "Other"], required=False),
                "Baseline": st.column_config.NumberColumn("Baseline", format="%.3f"),
                "Target": st.column_config.NumberColumn("Target", format="%.3f"),
                "Actual": st.column_config.NumberColumn("Latest actual", format="%.3f"),
                "Direction": st.column_config.SelectboxColumn("Direction", options=["Higher is better", "Lower is better"]),
                "Reporting_Period": st.column_config.TextColumn("Reporting period"),
                "Responsible_Team": st.column_config.TextColumn("Responsible team"),
                "Due_Date": st.column_config.TextColumn("Due date (YYYY-MM-DD)"),
                "Status": st.column_config.SelectboxColumn("Status", options=["Not started", "On track", "At risk", "Delayed", "Completed", "Needs review"]),
                "Notes": st.column_config.TextColumn("Notes", width="large"),
            },
        )
        clean_preview = clean_indicator_registry(edited_registry)
        calculated = calculate_indicator_progress(clean_preview) if not clean_preview.empty else pd.DataFrame(columns=["Progress toward target (%)", "Gap to target"])
        if not clean_preview.empty:
            result_view = pd.concat([clean_preview, calculated], axis=1)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Indicators", len(result_view))
            c2.metric("Completed", int((result_view["Status"] == "Completed").sum()))
            c3.metric("At risk / delayed", int(result_view["Status"].isin(["At risk", "Delayed"]).sum()))
            valid_progress = result_view["Progress toward target (%)"].dropna()
            c4.metric("Mean progress", f"{valid_progress.mean():.1f}%" if not valid_progress.empty else "N/A")
            st.markdown("#### Calculated monitoring view")
            st.dataframe(result_view, use_container_width=True, hide_index=True)
            st.download_button("Download indicator registry (CSV)", result_view.to_csv(index=False).encode("utf-8-sig"), file_name=f"indicator_registry_v22_{datetime.now():%Y%m%d}.csv", mime="text/csv")
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_view.to_excel(writer, sheet_name="Indicator Tracker", index=False)
                pd.DataFrame([{"Generated at": datetime.now().isoformat(timespec="seconds"), "Indicator count": len(result_view), "Source": "Locally maintained indicator registry"}]).to_excel(writer, sheet_name="Export Notes", index=False)
            st.download_button("Download indicator tracker (Excel)", output.getvalue(), file_name=f"indicator_tracker_v22_{datetime.now():%Y%m%d}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.caption("No indicators entered yet. Add a row in the editor to begin.")
        save_col, reset_col = st.columns([1, 1])
        if save_col.button("Save indicator registry locally", type="primary", key="v22_save_indicator_registry"):
            try:
                to_save = clean_indicator_registry(edited_registry)
                INDICATOR_FILE.parent.mkdir(parents=True, exist_ok=True)
                to_save.to_csv(INDICATOR_FILE, index=False, encoding="utf-8-sig")
                st.session_state["v22_indicator_registry"] = to_save
                st.success(f"Saved {len(to_save)} indicators to data/indicator_registry_v22_0.csv")
            except Exception as exc:
                st.error(f"Could not save the indicator registry: {exc}")
        if reset_col.button("Reload saved registry", key="v22_reload_indicator_registry"):
            st.session_state["v22_indicator_registry"] = load_indicator_registry()
            st.rerun()
        with st.expander("How progress is calculated"):
            st.markdown("- When baseline, target, and actual are present, progress is calculated relative to the baseline and target, respecting whether higher or lower values are better. It may be below 0% or above 100%.")
            st.markdown("- When no baseline is supplied, higher-is-better indicators use actual ÷ target as target attainment. This is not the same as progress from baseline.")
            st.markdown("- Lower-is-better indicators require baseline, target, and actual for a meaningful progress calculation. Gap is shown as a positive value when the target has not yet been reached in the expected direction.")
            st.markdown("- Confirm indicator definitions, disaggregation, denominator, reporting frequency, and source documents before using figures in donor reports.")
        st.markdown("#### Blank indicator template")
        template = pd.DataFrame(columns=INDICATOR_COLUMNS)
        st.download_button("Download blank indicator template (CSV)", template.to_csv(index=False).encode("utf-8-sig"), file_name="indicator_registry_template_v22.csv", mime="text/csv")

    elif section == "AI-Assisted Report Drafting":
        st.subheader("AI-Assisted Management Report Drafting")
        st.write("Generate a management brief from the current filters. The AI option sends aggregate statistics only—not beneficiary-level rows or IDs.")
        st.warning("External AI use is optional. Before using it, confirm your organisation permits this service and that sharing aggregate province/activity statistics is appropriate. Do not paste API keys into code or share them with others.")
        payload = aggregate_report_payload(filtered, province_table, activity_table, quality, filter_summary)
        with st.expander("Review the aggregate-only data that may be sent", expanded=False):
            st.json(payload)
        st.markdown("#### Option A — Local draft (no external service)")
        if st.button("Generate local evidence-based draft", key="v18_local_draft"):
            st.session_state["v18_report_draft"] = local_evidence_draft(payload)
            st.session_state["v18_report_source"] = "Local rule-based draft"
        st.markdown("#### Option B — AI-generated draft via OpenAI API")
        api_key = st.text_input("OpenAI API key (used for this request only; not saved by this app)", type="password", key="v18_openai_key")
        model_name = st.text_input("API model", value="gpt-4.1-mini", help="Use a model available to your API project.", key="v18_openai_model")
        consent = st.checkbox("I reviewed the aggregate payload and authorize sending it to the OpenAI API for report drafting.", key="v18_api_consent")
        if st.button("Generate AI management report", type="primary", key="v18_ai_draft"):
            if not api_key.strip():
                st.error("Enter an API key, or use the local draft option.")
            elif not consent:
                st.error("Review the payload and check the authorization box before making an external API request.")
            else:
                with st.spinner("Generating a draft from aggregate statistics…"):
                    try:
                        st.session_state["v18_report_draft"] = call_openai_responses_api(api_key, model_name, payload)
                        st.session_state["v18_report_source"] = f"OpenAI API draft — model {model_name.strip()}"
                        st.success("Draft generated. Review all statements and calculations before external use.")
                    except Exception as exc:
                        st.error(f"AI report generation failed: {exc}")
        if st.session_state.get("v18_report_draft"):
            st.markdown("#### Generated draft")
            st.caption(st.session_state.get("v18_report_source", "Generated draft"))
            st.markdown(st.session_state["v18_report_draft"])
            report_text = st.session_state["v18_report_draft"]
            st.download_button("Download report draft (Markdown)", report_text.encode("utf-8"), file_name=f"v20_0_ai_management_draft_{datetime.now():%Y%m%d_%H%M}.md", mime="text/markdown")
            st.download_button("Download report draft (TXT)", report_text.encode("utf-8"), file_name=f"v20_0_ai_management_draft_{datetime.now():%Y%m%d_%H%M}.txt", mime="text/plain")
        st.markdown("#### Human review checklist")
        st.markdown("- Verify every figure against the dashboard and source workbook.\n- Remove or correct any unsupported statement.\n- Confirm the reporting period, denominator, and filters.\n- Obtain programme/MEAL review before sharing with donors or external stakeholders.\n- Keep beneficiary-level data out of prompts and report drafts unless separately authorised and protected.")

    elif section == "Trends & ML Risk":
        st.subheader("Trends, Risk Flags & Machine-Learning Evaluation")
        st.warning("Risk flags and model estimates are decision-support signals only. They do not prove causation and must not be used to deny assistance or target individuals without human review.")
        trend = monthly_trend(filtered)
        if trend.empty:
            st.info("No valid dated satisfaction records are available for a trend chart in this selection.")
        else:
            st.markdown("#### Monthly trend")
            st.dataframe(trend, use_container_width=True, hide_index=True)
            metric = st.selectbox("Trend metric", ["Average satisfaction", "Low-satisfaction rate (%)", "Scored records"])
            st.plotly_chart(px.line(trend, x="Month", y=metric, markers=True, title=f"Monthly {metric.lower()}"), use_container_width=True, config={"displayModeBar": False})
            st.caption("Monthly comparisons can be affected by changes in sample size, coverage, activity mix, or data collection methods.")

        valid_score = valid_scores(filtered)
        overall_rate = (valid_score.isin(LOW_SCORE_VALUES).mean() * 100) if len(valid_score) else 0.0
        st.markdown("#### Descriptive risk flags")
        st.caption("Groups with at least 20 scored records and a low-satisfaction rate ≥5 percentage points above the selected overall rate are flagged for review. This is a screening rule, not a statistical proof.")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("**Province flags**")
            st.dataframe(risk_flags(province_table, overall_rate, "Province"), use_container_width=True, hide_index=True)
        with col_b:
            st.markdown("**Activity flags**")
            st.dataframe(risk_flags(activity_table, overall_rate, "Activity"), use_container_width=True, hide_index=True)

        st.markdown("#### Optional supervised ML evaluation")
        st.write("The model predicts whether a record's satisfaction score is low (1–2) using available non-outcome fields. Satisfaction_Score itself and direct identifiers are excluded from predictors. Holdout metrics are reported; results may not generalize to other provinces, periods, or programmes.")
        if len(filtered) < 50:
            st.info("At least 50 selected records are required before attempting a holdout evaluation. Select a broader dataset or remove restrictive filters.")
        else:
            x_all, y_all, labeled_mask = prepare_ml_features(filtered)
            labeled_x = x_all.loc[labeled_mask].copy()
            labeled_y = y_all.loc[labeled_mask].astype(int)
            unlabeled_x = x_all.loc[~labeled_mask].copy()
            counts = labeled_y.value_counts()
            if labeled_x.shape[1] == 0:
                st.info("No eligible predictor columns remain after excluding IDs and outcome leakage. Add suitable non-identifying fields such as province, activity, age band, or household size.")
            elif len(labeled_x) < 50 or len(counts) < 2 or counts.min() < 10:
                st.info("Insufficient labeled records or class balance for a responsible holdout evaluation. Aim for ≥50 labeled records and at least 10 examples in each class.")
            else:
                x_train, x_test, y_train, y_test = train_test_split(
                    labeled_x, labeled_y, test_size=0.25, random_state=42, stratify=labeled_y
                )
                model = build_ml_pipeline(x_train)
                if model is None:
                    st.info("No eligible predictors available for modelling.")
                else:
                    try:
                        model.fit(x_train, y_train)
                        pred = model.predict(x_test)
                        proba = model.predict_proba(x_test)[:, 1]
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("Holdout records", f"{len(y_test):,}")
                        m2.metric("Balanced accuracy", f"{balanced_accuracy_score(y_test, pred):.3f}")
                        m3.metric("Precision", f"{precision_score(y_test, pred, zero_division=0):.3f}")
                        m4.metric("Recall", f"{recall_score(y_test, pred, zero_division=0):.3f}")
                        m5, m6, m7 = st.columns(3)
                        m5.metric("F1 score", f"{f1_score(y_test, pred, zero_division=0):.3f}")
                        try:
                            m6.metric("ROC AUC", f"{roc_auc_score(y_test, proba):.3f}")
                        except Exception:
                            m6.metric("ROC AUC", "N/A")
                        try:
                            m7.metric("Average precision", f"{average_precision_score(y_test, proba):.3f}")
                        except Exception:
                            m7.metric("Average precision", "N/A")
                        # Compare with a simple prevalence-based baseline so model metrics have context.
                        baseline_pred = np.zeros(len(y_test), dtype=int)
                        st.markdown("#### Baseline comparison")
                        baseline_col1, baseline_col2, baseline_col3 = st.columns(3)
                        baseline_col1.metric("Majority baseline balanced accuracy", f"{balanced_accuracy_score(y_test, baseline_pred):.3f}")
                        baseline_col2.metric("Majority baseline F1", f"{f1_score(y_test, baseline_pred, zero_division=0):.3f}")
                        baseline_col3.metric("Holdout low-satisfaction prevalence", f"{y_test.mean()*100:.1f}%")
                        st.caption("The majority baseline predicts every holdout record as not-low. Compare balanced accuracy/F1 and other metrics together; a model is not useful merely because one metric exceeds this baseline.")
                        cm = confusion_matrix(y_test, pred, labels=[0, 1])
                        cm_df = pd.DataFrame(cm, index=["Actual not-low", "Actual low"], columns=["Predicted not-low", "Predicted low"])
                        st.markdown("**Holdout confusion matrix**")
                        st.dataframe(cm_df, use_container_width=True)
                        st.caption("Metrics depend on this random split. Review class balance, subgroup performance, and data quality; do not treat this evaluation as proof of operational readiness.")
                        if not unlabeled_x.empty:
                            st.markdown("#### Screening of records without a valid satisfaction score")
                            st.caption("These are model-estimated probabilities for records with missing/invalid scores only; they are not observed satisfaction outcomes.")
                            all_labeled_x = labeled_x
                            final_model = build_ml_pipeline(all_labeled_x)
                            final_model.fit(all_labeled_x, labeled_y)
                            missing_probs = final_model.predict_proba(unlabeled_x)[:, 1]
                            scored = filtered.loc[~labeled_mask].copy()
                            scored["Estimated low-satisfaction probability (%)"] = (missing_probs * 100).round(1)
                            scored["Screening band"] = np.where(missing_probs >= 0.7, "Higher estimate — review", np.where(missing_probs >= 0.4, "Intermediate estimate — review", "Lower estimate — do not interpret as no need"))
                            show_cols = [c for c in ["Province", "Activity", "Date", "Estimated low-satisfaction probability (%)", "Screening band"] if c in scored.columns]
                            st.dataframe(scored[show_cols], use_container_width=True, hide_index=True)
                            st.download_button("Download screening estimates (CSV)", scored[show_cols].to_csv(index=False).encode("utf-8-sig"), file_name=f"v20_0_screening_estimates_{datetime.now():%Y%m%d}.csv", mime="text/csv")
                    except Exception as exc:
                        st.error(f"Model evaluation could not be completed: {exc}")


st.divider()
st.caption("AI NGO Intelligence V22.0 • Descriptive findings and model estimates require human review before operational decisions or external publication.")
