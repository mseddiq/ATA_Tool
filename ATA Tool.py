import io
import json
import os
import sys
import tempfile
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote
import matplotlib.pyplot as plt
import pandas as pd
import gspread
import streamlit as st
from google.oauth2.service_account import Credentials
import streamlit.components.v1 as components
import extra_streamlit_components as stx
from fpdf import FPDF
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pptx import Presentation
from pptx.util import Inches
# -------------------- APP CONFIG --------------------
st.set_page_config(page_title="DAMAC | ATA Tool", layout="wide")
def get_data_dir() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    else:
        base = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    data_dir = base / "ATA_Tool"
    data_dir.mkdir(parents=True, exist_ok=True)
    return data_dir
DATA_DIR = get_data_dir()
PARAMETERS_JSON = str(DATA_DIR / "parameters.json")
EXPORT_XLSX = str(DATA_DIR / "ATA_Audit_Log.xlsx")
SUMMARY_COLUMNS = [
    "Evaluation ID",
    "Evaluation Date",
    "Audit Date",
    "Reaudit",
    "QA Name",
    "Auditor",
    "Call ID",
    "Call Duration",
    "Call Disposition",
    "Overall Score %",
    "Passed Points",
    "Failed Points",
    "Total Points",
    "Last Updated",
]
DETAILS_COLUMNS = [
    "Evaluation ID",
    "Evaluation Date",
    "Audit Date",
    "Reaudit",
    "QA Name",
    "Auditor",
    "Call ID",
    "Overall Score %",
    "Group",
    "Parameter",
    "Points",
    "Description",
    "Result",
    "Comment",
]


def _norm_col_name(name: str) -> str:
    return str(name).strip().lower().replace("_", " ")


def _standardize_columns(df: pd.DataFrame, expected_columns: list[str]) -> pd.DataFrame:
    if df is None:
        df = pd.DataFrame()
    source = df.copy()
    rename_map = {_norm_col_name(col): col for col in source.columns}
    output = pd.DataFrame()
    for expected in expected_columns:
        actual = rename_map.get(_norm_col_name(expected))
        output[expected] = source[actual] if actual is not None else ""
    return output


def _rewrite_google_worksheet(ws, df: pd.DataFrame, expected_columns: list[str]) -> None:
    out = _standardize_columns(df, expected_columns)
    ws.clear()
    ws.append_row(expected_columns)
    if not out.empty:
        ws.append_rows(out.values.tolist())


def connect_google_sheet():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = dict(st.secrets["gcp_service_account"])
    # Convert escaped newlines into real newlines
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key("1ojy6fWLX9Cil6Wnmdb5P2Xh6RinGfnTI2ZaylINU2ko")
    return sheet


@st.cache_data
def read_google_summary():
    sheet = connect_google_sheet()
    ws = sheet.worksheet("Summary")
    data = ws.get_all_records()
    return _standardize_columns(pd.DataFrame(data), SUMMARY_COLUMNS)


@st.cache_data
def read_google_details():
    sheet = connect_google_sheet()
    ws = sheet.worksheet("Details")
    data = ws.get_all_records()
    return _standardize_columns(pd.DataFrame(data), DETAILS_COLUMNS)
DAMAC_TITLE = "DAMAC Properties"
DAMAC_SUB1 = "Quality Assurance"
DAMAC_SUB2 = "Telesales Division"
APP_NAME = "ATA Audit the Auditor"
LOGO_URL = "https://images.ctfassets.net/zoq5l15g49wj/2qCUAnuJ9WgJiGNxmTXkUa/0505928e3d7060e1bc0c3195d7def448/damac-gold.svg?fm=webp&w=200&h=202&fit=pad&q=60"
LOGIN_LOGO_URL = "https://vectorseek.com/wp-content/uploads/2023/09/DAMAC-Properties-Logo-Vector.svg-.png"
COOKIE_AUTH_KEY = "ata_auth"
COOKIE_THEME_KEY = "ata_theme"
COOKIE_REMEMBER_DAYS = 30


def cookie_expiry(days: int = COOKIE_REMEMBER_DAYS) -> datetime:
    return datetime.utcnow() + pd.Timedelta(days=days)


def apply_theme_css() -> None:
    dark = st.session_state.get("theme_mode", "light") == "dark"
    if dark:
        vars_css = """
        --primary:#CEAE72;
        --secondary:#0b1f3a;
        --accent-gold:#CEAE72;
        --bg-main:#0b1f3a;
        --bg-card:#111827;
        --text-main:#ffffff;
        --text-muted:#cbd5e1;
        --border:#1f2937;
        --grid:#334155;
        --chart-bg:#111827;
        --chart-primary:#CEAE72;
        --chart-accent:#60a5fa;
        --hero-bg:linear-gradient(135deg, var(--secondary) 0%, #1f2937 100%);
        --sidebar-box:linear-gradient(135deg, var(--secondary) 0%, #1f2937 100%);
        """
    else:
        vars_css = """
        --primary:#0b1f3a;
        --secondary:#1e3a8a;
        --accent-gold:#CEAE72;
        --bg-main:#f8fafc;
        --bg-card:#ffffff;
        --text-main:#0b1f3a;
        --text-muted:#64748b;
        --border:#e2e8f0;
        --grid:#0b1f3a;
        --chart-bg:#ffffff;
        --chart-primary:#1e3a8a;
        --chart-accent:#CEAE72;
        --hero-bg:linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
        --sidebar-box:linear-gradient(135deg, var(--primary) 0%, var(--secondary) 100%);
        """
    st.markdown(
        f"""
        <style>
        :root {{{vars_css}}}
        .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
            background: var(--bg-main) !important;
            color: var(--text-main) !important;
        }}
        .block-container, .stMarkdown, p, span, label, div {{ color: var(--text-main); }}
        .ata-hero {{ background: var(--hero-bg) !important; color: #fff !important; border:1px solid var(--border) !important; }}
        .logo-box, .credit-box {{ background: var(--sidebar-box) !important; border-color: var(--border) !important; }}
        .credit-line {{ color: var(--accent-gold) !important; }}
        .stat-card, .ata-card {{ background: var(--bg-card) !important; border:1px solid var(--border) !important; color: var(--text-main) !important; }}
        .stat-val, .view-header h2, .page-title h2 {{ color: var(--primary) !important; }}
        .stat-label, .login-note, .login-extra {{ color: var(--text-muted) !important; }}
        .styled-table th {{ background-color: var(--primary) !important; color: #fff !important; }}
        .styled-table td {{ color: var(--text-main) !important; border-bottom: 1px solid var(--border) !important; }}
        .login-wrap {{ background: var(--bg-card) !important; border-color: var(--border) !important; }}
        .login-title {{ color: var(--text-main) !important; }}
        .stButton>button, .stDownloadButton>button {{
            background: var(--secondary) !important;
            color: #fff !important;
            border: 1px solid var(--border) !important;
        }}
        .stButton>button:hover, .stDownloadButton>button:hover {{
            background: var(--primary) !important;
            color: #fff !important;
        }}
        .stExpander, .streamlit-expanderHeader {{ background: var(--bg-card) !important; color: var(--text-main) !important; border-color: var(--border) !important; }}
        [data-testid="stSidebar"] {{ background: var(--bg-main) !important; border-right: 1px solid var(--border); }}
        [data-testid="stSidebar"] * {{ color: var(--text-main); }}
        [data-testid="stDataFrame"], .stDataFrame, .stTable {{ background: var(--bg-card) !important; color: var(--text-main) !important; }}
        hr {{ border-color: var(--border) !important; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def get_chart_theme() -> dict:
    dark = st.session_state.get("theme_mode", "light") == "dark"
    if dark:
        return {
            "bg": "#111827",
            "grid": "#334155",
            "text": "#ffffff",
            "primary": "#CEAE72",
            "accent": "#60a5fa",
            "fail": "#ef4444",
            "pass": "#10b981",
        }
    return {
        "bg": "#ffffff",
        "grid": "#e5e7eb",
        "text": "#0b1f3a",
        "primary": "#1e3a8a",
        "accent": "#CEAE72",
        "fail": "#ef4444",
        "pass": "#10b981",
    }


def style_chart(ax, theme: dict) -> None:
    ax.set_facecolor(theme["bg"])
    for spine in ax.spines.values():
        spine.set_color(theme["grid"])
    ax.tick_params(colors=theme["text"])
    ax.title.set_color(theme["text"])
    if ax.xaxis.label:
        ax.xaxis.label.set_color(theme["text"])
    if ax.yaxis.label:
        ax.yaxis.label.set_color(theme["text"])
# -------------------- UI THEME --------------------
st.markdown(
    """
<style>
.block-container { padding-top: 1.0rem; font-family: "Candara", "Segoe UI", sans-serif; }
.stApp, .stMarkdown, .stTextInput, .stSelectbox, .stDataEditor, .stButton, .stTable, .stDataFrame {
  font-family: "Candara", "Segoe UI", sans-serif;
}
.ata-hero{ padding:22px; border-radius:20px; box-shadow:0 10px 25px -5px rgba(0,0,0,0.3); margin-bottom:25px; }
.ata-hero.left-align { text-align:left; }
.ata-hero .t1 { font-size:28px; font-weight:900; margin:0; letter-spacing:1px; }
.ata-hero .t2 { font-size:16px; opacity:.85; margin-top:8px; }
.stat-card { padding:20px; border-radius:15px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.05); text-align:center; transition:transform 0.2s; }
.stat-card:hover { transform: translateY(-5px); }
.stat-val { font-size:24px; font-weight:800; }
.stat-label { font-size:14px; margin-top:5px; }
.ata-card { border-radius:16px; padding:20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); }
.styled-table { width:100%; border-collapse:collapse; margin:10px 0; font-size:14px; }
.styled-table th { text-align:left; padding:12px 15px; }
.styled-table td { padding:10px 15px; }
.credit-line { text-align:left; font-size:12px; margin-top:5px; font-style:italic; font-weight:700; }
.sidebar-credit { margin-top:12px; }
.logo-box { border-radius:14px; padding:8px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.08); display:flex; justify-content:center; align-items:center; margin-bottom:12px; }
.credit-box { border-radius:14px; padding:10px 12px; box-shadow:0 4px 6px -1px rgba(0,0,0,0.08); margin-top:10px; width:100%; }
.view-header h2 { font-weight:800; margin-bottom:12px; }
.page-title { margin-top:10px; margin-bottom:12px; }
.login-wrap {max-width:460px; margin:20px auto 10px auto; padding:26px; border:1px solid var(--border); border-radius:16px; box-shadow:0 10px 30px rgba(0,0,0,0.08);} 
.login-logo {display:flex; justify-content:center; margin-bottom:12px;} 
.login-title {text-align:center; font-weight:800; margin-bottom:0;} 
.login-note {text-align:center; font-size:12px; margin-top:8px;} 
.login-extra {text-align:center; font-size:12px; margin-top:4px;} 
</style>
""",
    unsafe_allow_html=True,
)
# -------------------- RUBRIC (SHEET-ALIGNED) --------------------
ACCURACY_HEADER = "Accuracy of Scoring"
ACCURACY_SUBPARAMS = [
    "Call Opening (Readiness / energy)",
    "Call Opening 2 (Confirming lead source / meeting focused)",
    "Effective Probing / Qualifying client",
    "Accurate / Complete info",
    "Objection / Call Handling",
    "Soft Skills (active listening - building rapport)",
    "Positivity / professionality / politeness",
    "Call closure / meeting summarized",
    "Accurate Disposition",
    "Comment / notes",
    "Accurate Data inputs / shows",
    "WhatsApp message sent",
    "Took lead ownership",
    "Follow-up made properly",
]
EVALUATION_QUALITY_PARAMS = [
    ("Adherence to QA Guidelines", "Followed QA process and aligned with calibration standards"),
    ("Evidence & Notes", "Left a clear, specific and improvement-focused comment"),
    ("Objectivity & Fairness", "Evaluation is unbiased and fact-based"),
    ("Critical Error Identification", "Correct identification of fatal errors"),
    (
        "Evaluation Variety & Sample Coverage",
        "Evaluations cover a balanced mix of call durations and call types",
    ),
    ("Feedback Actionability", "Conducted coaching session on the call topic (If required)"),
    ("Timeliness & Completeness", "On track with the evaluations target SLA"),
]
DEFAULT_PARAMETERS = {
    "form_name": APP_NAME,
    "parameters": [
        {
            "Parameter": ACCURACY_HEADER,
            "Description": "Header (not scored) | Category total = 14 points",
            "Points": 0,
            "Group": "HEADER",
        },
        *[
            {
                "Parameter": p,
                "Description": "Accuracy of Scoring – Sub Parameter",
                "Points": 1,
                "Group": "ACCURACY_SUB",
            }
            for p in ACCURACY_SUBPARAMS
        ],
        *[
            {"Parameter": p, "Description": d, "Points": 1, "Group": "EVAL_QUALITY"}
            for (p, d) in EVALUATION_QUALITY_PARAMS
        ],
    ],
}
def ensure_parameters_file() -> None:
    if not os.path.exists(PARAMETERS_JSON):
        with open(PARAMETERS_JSON, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_PARAMETERS, f, indent=2, ensure_ascii=False)
def load_parameters_df() -> pd.DataFrame:
    ensure_parameters_file()
    with open(PARAMETERS_JSON, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    df = pd.DataFrame(cfg.get("parameters", []))
    if df.empty:
        df = pd.DataFrame(DEFAULT_PARAMETERS["parameters"])
    df["Result"] = "Pass"
    df["Comment"] = ""
    for c in ["Parameter", "Points", "Description", "Result", "Comment", "Group"]:
        if c not in df.columns:
            df[c] = "" if c in ("Description", "Comment", "Group") else 1
    df["Points"] = pd.to_numeric(df["Points"], errors="coerce").fillna(0).astype(int)
    return df[["Group", "Parameter", "Points", "Description", "Result", "Comment"]].copy()
def normalize_details_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty or "Parameter" not in df.columns:
        return load_parameters_df()
    for col in ["Group", "Parameter", "Points", "Description", "Result", "Comment"]:
        if col not in df.columns:
            if col in ("Description", "Comment", "Group"):
                df[col] = ""
            elif col == "Result":
                df[col] = "Pass"
            else:
                df[col] = 1
    return df[["Group", "Parameter", "Points", "Description", "Result", "Comment"]].copy()
def format_date(value) -> str:
    if value in ("", None):
        return ""
    try:
        parsed = pd.to_datetime(value, dayfirst=True)
        if pd.isna(parsed):
            return ""
        return parsed.strftime("%d/%m/%Y")
    except Exception:
        return str(value)
def norm_id(x) -> str:
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = str(x).strip()
    if s.lower() in ("nan", "none"):
        return ""
    return s
def safe_read_excel(path: str, sheet: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()
def next_evaluation_id(evaluation_date_str: str) -> str:
    df = read_google_summary()
    yyyymmdd = evaluation_date_str.replace("-", "")
    prefix = f"ATA-{yyyymmdd}-"
    if df.empty or "Evaluation ID" not in df.columns:
        return f"{prefix}0001"
    existing = df["Evaluation ID"].astype(str)
    existing = existing[existing.str.startswith(prefix)]
    if existing.empty:
        return f"{prefix}0001"

    def _seq(x: str) -> int:
        try:
            return int(str(x).split("-")[-1])
        except Exception:
            return 0

    max_seq = max(existing.apply(_seq).tolist() + [0])
    return f"{prefix}{max_seq + 1:04d}"
def write_formatted_report(
    record: dict,
    filename: str,
    summary_df: pd.DataFrame | None = None,
    details_df: pd.DataFrame | None = None,
) -> None:
    if not os.path.exists(filename):
        return
    wb = load_workbook(filename)
    sheet_name = "Formatted Report"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    header_dark_blue = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_gray = PatternFill(start_color="C9C9C9", end_color="C9C9C9", fill_type="solid")
    header_light_gray = PatternFill(start_color="E5E5E5", end_color="E5E5E5", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    dark_font = Font(color="000000", bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    details_cols = [
        "Evaluation ID",
        "Evaluation Date",
        "Audit Date",
        "Reaudit",
        "QA Name",
        "Auditor",
        "Call ID",
        "Overall Score %",
    ]
    acc_cols = ACCURACY_SUBPARAMS
    eval_cols = [p for p, _ in EVALUATION_QUALITY_PARAMS]
    all_cols = details_cols + acc_cols + eval_cols
    details_start = 1
    details_end = details_start + len(details_cols) - 1
    acc_start = details_end + 1
    acc_end = acc_start + len(acc_cols) - 1
    eval_start = acc_end + 1
    eval_end = eval_start + len(eval_cols) - 1
    ws.row_dimensions[1].height = 10
    header_row = 2
    label_row = 3
    data_start_row = 4
    ws.merge_cells(start_row=header_row, start_column=details_start, end_row=header_row, end_column=details_end)
    ws.merge_cells(start_row=header_row, start_column=acc_start, end_row=header_row, end_column=acc_end)
    ws.merge_cells(start_row=header_row, start_column=eval_start, end_row=header_row, end_column=eval_end)
    ws.cell(row=header_row, column=details_start, value="Details").fill = header_dark_blue
    ws.cell(row=header_row, column=details_start).font = header_font
    ws.cell(row=header_row, column=details_start).alignment = center
    ws.cell(row=header_row, column=acc_start, value="ACCURACY_SUB").fill = header_gray
    ws.cell(row=header_row, column=acc_start).font = dark_font
    ws.cell(row=header_row, column=acc_start).alignment = center
    ws.cell(row=header_row, column=eval_start, value="EVAL_QUALITY").fill = header_light_gray
    ws.cell(row=header_row, column=eval_start).font = dark_font
    ws.cell(row=header_row, column=eval_start).alignment = center
    for idx, col_name in enumerate(all_cols, start=1):
        cell = ws.cell(row=label_row, column=idx, value=col_name)
        cell.alignment = center
        if idx <= details_end:
            cell.fill = header_dark_blue
            cell.font = header_font
        elif idx <= acc_end:
            cell.fill = header_gray
            cell.font = dark_font
        else:
            cell.fill = header_light_gray
            cell.font = dark_font
    summary = summary_df if summary_df is not None else safe_read_excel(filename, "Summary")
    details = details_df if details_df is not None else safe_read_excel(filename, "Details")
    if not summary.empty:
        summary = summary.dropna(axis=1, how="all")
    if summary.empty or details.empty:
        wb.save(filename)
        return
    if "Parameter" not in details.columns:
        wb.save(filename)
        return
    details_lookup = {
        eval_id: details[details["Evaluation ID"] == eval_id]
        for eval_id in summary["Evaluation ID"].dropna().unique().tolist()
    }
    for row_idx, row_data in enumerate(summary.to_dict(orient="records"), start=data_start_row):
        eval_id = row_data.get("Evaluation ID", "")
        details_values = [
            row_data.get("Evaluation ID", ""),
            format_date(row_data.get("Evaluation Date", "")),
            format_date(row_data.get("Audit Date", "")),
            row_data.get("Reaudit", ""),
            row_data.get("QA Name", ""),
            row_data.get("Auditor", ""),
            row_data.get("Call ID", ""),
            f"{row_data.get('Overall Score %', 0):.2f}%",
        ]
        detail_rows = details_lookup.get(eval_id, pd.DataFrame())
        if detail_rows.empty or "Parameter" not in detail_rows.columns:
            acc_values = ["" for _ in acc_cols]
            eval_values = ["" for _ in eval_cols]
        else:
            acc_values = []
            for param in acc_cols:
                match = detail_rows[detail_rows["Parameter"] == param]
                acc_values.append(match.iloc[0]["Result"] if not match.empty else "")
            eval_values = []
            for param in eval_cols:
                match = detail_rows[detail_rows["Parameter"] == param]
                eval_values.append(match.iloc[0]["Result"] if not match.empty else "")
        values = details_values + acc_values + eval_values
        for idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=idx, value=value)
            cell.fill = white_fill
            cell.alignment = center
        ws.row_dimensions[row_idx].height = 20
    ws.row_dimensions[header_row].height = 20
    ws.row_dimensions[label_row].height = 45
    for idx in range(1, len(all_cols) + 1):
        ws.column_dimensions[ws.cell(row=label_row, column=idx).column_letter].width = 18
    wb.save(filename)
def upsert_google_sheet(record: dict):
    summary_existing = read_google_summary()
    details_existing = read_google_details()
    rid = norm_id(record["evaluation_id"])

    summary_existing["_rid"] = summary_existing["Evaluation ID"].apply(norm_id)
    summary_existing = summary_existing[summary_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")

    details_existing["_rid"] = details_existing["Evaluation ID"].apply(norm_id)
    details_existing = details_existing[details_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")

    summary_row = pd.DataFrame(
        [
            {
                "Evaluation ID": record["evaluation_id"],
                "Evaluation Date": record["evaluation_date"],
                "Audit Date": record["audit_date"],
                "Reaudit": record["reaudit"],
                "QA Name": record["qa_name"],
                "Auditor": record["auditor"],
                "Call ID": record["call_id"],
                "Call Duration": record["call_duration"],
                "Call Disposition": record["call_disposition"],
                "Overall Score %": record["overall_score"],
                "Passed Points": record["passed_points"],
                "Failed Points": record["failed_points"],
                "Total Points": record["total_points"],
                "Last Updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]
    )
    details_df = normalize_details_df(record["details"])
    details_map = {
        "Evaluation ID": record["evaluation_id"],
        "Evaluation Date": record["evaluation_date"],
        "Audit Date": record["audit_date"],
        "Reaudit": record["reaudit"],
        "QA Name": record["qa_name"],
        "Auditor": record["auditor"],
        "Call ID": record["call_id"],
        "Overall Score %": record["overall_score"],
        "Points": details_df.get("Points", 1),
        "Description": details_df.get("Description", ""),
    }
    details_df = details_df.drop(columns=list(details_map.keys()), errors="ignore")
    for i, col in enumerate(details_map):
        details_df.insert(i, col, details_map[col])

    out_summary = pd.concat([summary_existing, summary_row], ignore_index=True)
    out_details = pd.concat([details_existing, details_df], ignore_index=True)

    sheet = connect_google_sheet()
    _rewrite_google_worksheet(sheet.worksheet("Summary"), out_summary, SUMMARY_COLUMNS)
    _rewrite_google_worksheet(sheet.worksheet("Details"), out_details, DETAILS_COLUMNS)
    st.cache_data.clear()

def delete_evaluation(eval_id: str) -> bool:
    rid = norm_id(eval_id)
    if not rid:
        return False
    summary_existing = read_google_summary()
    details_existing = read_google_details()
    if summary_existing.empty and details_existing.empty:
        return False

    summary_existing["_rid"] = summary_existing["Evaluation ID"].apply(norm_id)
    details_existing["_rid"] = details_existing["Evaluation ID"].apply(norm_id)

    before_summary = len(summary_existing)
    before_details = len(details_existing)

    summary_existing = summary_existing[summary_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")
    details_existing = details_existing[details_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")

    changed = (len(summary_existing) != before_summary) or (len(details_existing) != before_details)
    if not changed:
        return False

    sheet = connect_google_sheet()
    _rewrite_google_worksheet(sheet.worksheet("Summary"), summary_existing, SUMMARY_COLUMNS)
    _rewrite_google_worksheet(sheet.worksheet("Details"), details_existing, DETAILS_COLUMNS)
    st.cache_data.clear()
    return True

def compute_weighted_score(df: pd.DataFrame) -> dict:
    df = df.copy()
    df["Result"] = df["Result"].fillna("Pass")
    accuracy_rows = df[df["Group"] == "ACCURACY_SUB"]
    eval_quality_rows = df[df["Group"] == "EVAL_QUALITY"]
    accuracy_failed = (accuracy_rows["Result"] == "Fail").any()
    accuracy_points = 1
    accuracy_passed = 0 if accuracy_failed else 1
    eval_quality_total = int(len(eval_quality_rows))
    eval_quality_passed = int((eval_quality_rows["Result"] == "Pass").sum())
    total_points = accuracy_points + eval_quality_total
    passed_points = accuracy_passed + eval_quality_passed
    failed_points = total_points - passed_points
    score = round((passed_points / total_points) * 100, 2) if total_points else 0.0
    return {
        "score": score,
        "passed_points": passed_points,
        "failed_points": failed_points,
        "total_points": total_points,
    }
# -------------------- EXPORT HELPERS --------------------
def copy_to_clipboard_button(label: str, text_to_copy: str, key: str) -> None:
    safe_text = text_to_copy.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
    html = f"""
    <button id="btn-{key}" style="width:100%;padding:10px;background:#0b1f3a;color:white;border:none;border-radius:8px;cursor:pointer;font-weight:bold;">{label}</button>
    <script>
      document.getElementById("btn-{key}").onclick = () => {{
        const text = `{safe_text}`;
        const htmlBlob = new Blob([text], {{ type: "text/html" }});
        const plainBlob = new Blob([text], {{ type: "text/plain" }});
        navigator.clipboard.write([
          new ClipboardItem({{ "text/html": htmlBlob, "text/plain": plainBlob }})
        ]).then(() => {{
          const el = document.getElementById("copystatus-{key}");
          if(el) el.innerText = "Email copied to clipboard!";
        }});
      }};
    </script>
    <div id="copystatus-{key}" style="font-family:sans-serif;color:#10b981;font-size:12px;margin-top:5px;text-align:center;"></div>
    """
    components.html(html, height=70)
def email_subject_text(record: dict) -> str:
    return f"ATA Evaluation | {record['evaluation_id']} | {record['qa_name']} | {format_date(record['audit_date'])}"

def email_html_inline(record: dict) -> str:
    def make_table(df):
        rows = []
        for _, r in df.iterrows():
            status = r["Result"]
            badge = (
                "background:#e9f7ef;color:#1f8f4a"
                if status == "Pass"
                else "#fde8e6;color:#d93025"
            )
            badge = f"{badge};padding:4px 8px;border-radius:12px;font-weight:bold;"
            comm = str(r["Comment"]) if str(r["Comment"]).lower() != "nan" else ""
            rows.append(
                "<tr>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;'>{r['Parameter']}</td>"
                "<td style='padding:8px;border-bottom:1px solid #eee;text-align:center;'>"
                f"<span style='{badge}'>{status}</span>"
                "</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;'>{comm}</td>"
                "</tr>"
            )
        return (
            "<table style='width:100%;border-collapse:collapse;font-size:13px;'>"
            "<tr style='background:#f8fafc;'>"
            "<th style='text-align:left;padding:8px;'>Parameter</th>"
            "<th style='width:80px;padding:8px;'>Result</th>"
            "<th style='padding:8px;'>Comment</th>"
            "</tr>"
            f"{''.join(rows)}"
            "</table>"
        )
    det = record["details"]
    return f"""
    <div style="font-family:sans-serif;max-width:800px;border:1px solid #eee;padding:20px;border-radius:15px;">
        <div style="background:#0b1f3a;color:white;padding:15px;border-radius:10px;margin-bottom:20px;">
            <h2 style="margin:0;">{DAMAC_TITLE} | ATA Evaluation</h2>
            <p style="margin:5px 0 0 0;opacity:0.8;">{DAMAC_SUB1} | {DAMAC_SUB2}</p>
        </div>
        <table style="width:100%;margin-bottom:20px;font-size:14px;">
            <tr><td><b>Evaluation ID:</b> {record['evaluation_id']}</td><td><b>Evaluation Date:</b> {record['evaluation_date']}</td></tr>
            <tr><td><b>QA Name:</b> {record['qa_name']}</td><td><b>Auditor Name:</b> {record['auditor']}</td></tr>
            <tr><td><b>Audit Date:</b> {record['audit_date']}</td><td><b>Call ID:</b> {record['call_id']}</td></tr>
            <tr><td><b>Call Duration:</b> {record['call_duration']}</td><td><b>Call Disposition:</b> {record['call_disposition']}</td></tr>
            <tr><td colspan="2" style="padding-top:10px;"><div style="background:#f1f5f9;padding:10px;border-radius:8px;text-align:center;"><b>Overall Score: <span style="font-size:20px;color:#0b1f3a;">{record['overall_score']:.2f}%</span></b></div></td></tr>
        </table>
        <h3 style="color:#0b1f3a;border-left:4px solid #0b1f3a;padding-left:10px;">Accuracy of Scoring</h3>
        {make_table(det[det["Group"] == "ACCURACY_SUB"])}
        <h3 style="color:#0b1f3a;border-left:4px solid #0b1f3a;padding-left:10px;margin-top:20px;">Evaluation Quality</h3>
        {make_table(det[det["Group"] == "EVAL_QUALITY"])}
        <div style="margin-top:20px;padding:10px;background:#f8fafc;border-radius:8px;"><b>Reaudit Status:</b> {record['reaudit']}</div>
    </div>
    """
class PDFReport(FPDF):
    def header(self):
        self.set_font("Arial", "B", 10)
        self.set_text_color(11, 31, 58)
        self.cell(0, 8, f"{DAMAC_TITLE} | {DAMAC_SUB1} | {DAMAC_SUB2}", ln=True, align="C")
        self.ln(2)
    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")
def pdf_evaluation(record: dict) -> bytes:
    pdf = PDFReport()
    pdf.add_page()
    pdf.set_auto_page_break(auto=False, margin=12)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 12, "ATA Evaluation Report", ln=True, align="L")
    pdf.ln(2)
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(241, 245, 249)
    data = [
        ["Evaluation ID", record["evaluation_id"], "Evaluation Date", format_date(record["evaluation_date"])],
        ["QA Name", record["qa_name"], "Auditor Name", record["auditor"]],
        ["Audit Date", format_date(record["audit_date"]), "Call ID", record["call_id"]],
        ["Call Duration", record["call_duration"], "Call Disposition", record["call_disposition"]],
    ]
    def line_count(text: str, width: float) -> int:
        words = str(text).split()
        if not words:
            return 1
        lines = 1
        line = ""
        for word in words:
            test = f"{line} {word}".strip()
            if pdf.get_string_width(test) <= width:
                line = test
            else:
                lines += 1
                line = word
        return lines
    col_widths = [42, 42, 42, 64]
    def truncate_text(text: str, width: float) -> str:
        text = str(text)
        if pdf.get_string_width(text) <= width:
            return text
        while text and pdf.get_string_width(f"{text}...") > width:
            text = text[:-1]
        return f"{text}..." if text else ""
    row_height = 10
    for row in data:
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.set_font("Arial", "B", 9)
        pdf.cell(col_widths[0], row_height, truncate_text(row[0], col_widths[0]), border=1, fill=True)
        pdf.set_font("Arial", "", 9)
        pdf.cell(col_widths[1], row_height, truncate_text(row[1], col_widths[1]), border=1)
        pdf.set_font("Arial", "B", 9)
        pdf.cell(col_widths[2], row_height, truncate_text(row[2], col_widths[2]), border=1, fill=True)
        pdf.set_font("Arial", "", 9)
        pdf.cell(col_widths[3], row_height, truncate_text(row[3], col_widths[3]), border=1)
        pdf.ln(row_height)
    pdf.ln(4)
    pdf.set_font("Arial", "B", 11)
    pdf.cell(
        0,
        10,
        f"Overall Score: {record['overall_score']:.2f}%",
        ln=True,
        align="C",
        border=1,
        fill=True,
    )
    pdf.ln(5)
    def draw_section(title, df):
        pdf.set_font("Arial", "B", 10)
        pdf.set_text_color(11, 31, 58)
        pdf.cell(0, 8, title, ln=True)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "B", 8)
        pdf.set_fill_color(11, 31, 58)
        pdf.set_text_color(255, 255, 255)
        param_w, result_w, comment_w = 90, 20, 80
        pdf.cell(param_w, 8, " Parameter", border=1, fill=True)
        pdf.cell(result_w, 8, " Result", border=1, fill=True, align="C")
        pdf.cell(comment_w, 8, " Comment", border=1, fill=True)
        pdf.ln()
        pdf.set_font("Arial", "", 7)
        pdf.set_text_color(0, 0, 0)
        def truncate_section(text: str, width: float) -> str:
            text = str(text)
            if pdf.get_string_width(text) <= width:
                return text
            while text and pdf.get_string_width(f"{text}...") > width:
                text = text[:-1]
            return f"{text}..." if text else ""
        def wrap_lines(text: str, width: float) -> list[str]:
            words = str(text).split()
            if not words:
                return [""]
            lines = []
            line = ""
            for word in words:
                test = f"{line} {word}".strip()
                if pdf.get_string_width(test) <= width:
                    line = test
                else:
                    lines.append(line)
                    line = word
            if line:
                lines.append(line)
            return lines
        def ensure_space(height: float) -> None:
            if pdf.get_y() + height > (pdf.h - 12):
                pdf.add_page()
                pdf.set_font("Arial", "B", 10)
                pdf.set_text_color(11, 31, 58)
                pdf.cell(0, 8, title, ln=True)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Arial", "B", 8)
                pdf.set_fill_color(11, 31, 58)
                pdf.set_text_color(255, 255, 255)
                pdf.cell(param_w, 8, " Parameter", border=1, fill=True)
                pdf.cell(result_w, 8, " Result", border=1, fill=True, align="C")
                pdf.cell(comment_w, 8, " Comment", border=1, fill=True)
                pdf.ln()
                pdf.set_font("Arial", "", 7)
                pdf.set_text_color(0, 0, 0)
        for _, r in df.iterrows():
            param, res, comm = str(r["Parameter"]), str(r["Result"]), str(r["Comment"])
            if comm.lower() == "nan":
                comm = ""
            comment_lines = wrap_lines(comm, comment_w - 2) if comm.strip() else [""]
            if len(comment_lines) > 2:
                comment_lines = comment_lines[:2]
                comment_lines[-1] = f"{comment_lines[-1]}..."
            row_height = max(6, 4 * len(comment_lines))
            line_height = 4
            ensure_space(row_height + 2)
            start_x, start_y = pdf.get_x(), pdf.get_y()
            pdf.rect(start_x, start_y, param_w, row_height)
            pdf.rect(start_x + param_w, start_y, result_w, row_height)
            pdf.rect(start_x + param_w + result_w, start_y, comment_w, row_height)
            y_offset = (row_height - line_height) / 2
            pdf.set_xy(start_x + 1, start_y + y_offset)
            pdf.cell(param_w - 2, line_height, truncate_section(param, param_w - 2), border=0)
            pdf.set_xy(start_x + param_w, start_y + y_offset)
            if res == "Pass":
                pdf.set_text_color(31, 143, 74)
            elif res == "Fail":
                pdf.set_text_color(217, 48, 37)
            else:
                pdf.set_text_color(0, 0, 0)
            pdf.cell(result_w, line_height, truncate_section(res, result_w - 2), border=0, align="C")
            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(start_x + param_w + result_w + 1, start_y + 1)
            pdf.multi_cell(comment_w - 2, line_height, "\n".join(comment_lines), border=0)
            pdf.set_xy(start_x, start_y + row_height)
        pdf.ln(5)
    det = record["details"]
    draw_section("Accuracy of Scoring", det[det["Group"] == "ACCURACY_SUB"])
    draw_section("Evaluation Quality", det[det["Group"] == "EVAL_QUALITY"])
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 8, f"Reaudit Status: {record['reaudit']}", ln=True)
    out = pdf.output(dest="S")
    return bytes(out) if isinstance(out, (bytes, bytearray)) else out.encode("latin-1")
# -------------------- DASHBOARD LOGIC --------------------
def build_dashboard_figs(summary: pd.DataFrame | None = None, details: pd.DataFrame | None = None):
    if summary is None or details is None:
        summary = read_google_summary()
        details = read_google_details()
    if summary.empty or details.empty:
        return (None,) * 10 + (summary, details)
    for col in ["Failed Points", "Total Points", "Overall Score %"]:
        if col in summary.columns:
            summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0)
    summary["Evaluation Date"] = pd.to_datetime(summary.get("Evaluation Date"), errors="coerce")
    summary["Month"] = summary["Evaluation Date"].dt.to_period("M").astype(str)
    summary["Failure Rate"] = summary.apply(
        lambda r: (r["Failed Points"] / r["Total Points"]) if r["Total Points"] else 0,
        axis=1,
    )
    theme = get_chart_theme()

    def add_bar_labels(ax):
        heights = [patch.get_height() for patch in ax.patches]
        if heights:
            ax.set_ylim(0, max(heights) * 1.2)
        for patch in ax.patches:
            value = patch.get_height()
            ax.annotate(
                f"{value:.1f}" if isinstance(value, float) else f"{value}",
                (patch.get_x() + patch.get_width() / 2, value),
                ha="center",
                va="bottom",
                xytext=(0, 6),
                textcoords="offset points",
                fontsize=9,
                color=theme["text"],
            )
    # 1. Trend Chart (Failure Rate)
    trend = summary.groupby("Month")["Failure Rate"].mean().sort_index()
    fig_trend, ax = plt.subplots(figsize=(6, 4))
    ax.plot(trend.index, trend.values * 100, marker="o", color=theme["primary"], linewidth=2, markersize=6)
    ax.fill_between(trend.index, trend.values * 100, color=theme["primary"], alpha=0.15)
    ax.set_title("Failure Rate Trend (%)", fontweight="bold", fontsize=11)
    ax.grid(True, alpha=0.25, color=theme["grid"])
    for x, y in zip(trend.index, trend.values * 100):
        ax.annotate(f"{y:.1f}%", (x, y), textcoords="offset points", xytext=(0, 8), ha="center", color=theme["text"])
    style_chart(ax, theme)
    fig_trend.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 2. Heatmap
    fail_rows = details[details.get("Result", "") == "Fail"].copy()
    heat = pd.DataFrame()
    if not fail_rows.empty:
        fail_rows["Date Label"] = pd.to_datetime(
            fail_rows.get("Evaluation Date"), errors="coerce"
        ).dt.strftime("%d-%b")
        heat = pd.pivot_table(
            fail_rows,
            index="Parameter",
            columns="Date Label",
            values="Result",
            aggfunc="count",
            fill_value=0,
        )
    fig_heat, axh = plt.subplots(figsize=(6, 4))
    if heat.empty:
        axh.text(0.5, 0.5, "No failures recorded", ha="center", va="center", color=theme["text"])
        axh.axis("off")
    else:
        im = axh.imshow(heat.values, cmap="YlOrRd", aspect="auto")
        axh.set_xticks(range(len(heat.columns)))
        axh.set_xticklabels(heat.columns, color=theme["text"])
        axh.set_yticks(range(len(heat.index)))
        axh.set_yticklabels(heat.index, fontsize=8, color=theme["text"])
        for i in range(len(heat.index)):
            for j in range(len(heat.columns)):
                axh.text(
                    j,
                    i,
                    str(int(heat.iloc[i, j])),
                    ha="center",
                    va="center",
                    color=theme["text"],
                    fontsize=8,
                )
        axh.set_title("Failure Distribution by Parameter", fontweight="bold", fontsize=11)
        cbar = plt.colorbar(im, ax=axh)
        cbar.ax.yaxis.set_tick_params(color=theme["text"])
        plt.setp(cbar.ax.get_yticklabels(), color=theme["text"])
        style_chart(axh, theme)
    fig_heat.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 3. Pass vs Fail Pie Chart
    pass_points = summary["Passed Points"].sum() if "Passed Points" in summary.columns else 0
    fail_points = summary["Failed Points"].sum() if "Failed Points" in summary.columns else 0
    fig_pie, axp = plt.subplots(figsize=(6, 4))
    axp.pie(
        [pass_points, fail_points],
        labels=["Pass", "Fail"],
        autopct="%1.1f%%",
        colors=[theme["pass"], theme["fail"]],
        startangle=90,
        textprops={"fontsize": 10, "color": theme["text"]},
    )
    axp.set_title("Pass vs Fail Points", fontweight="bold", fontsize=11)
    axp.axis("equal")
    style_chart(axp, theme)
    fig_pie.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 4. QA Average Scores
    fig_qa, axq = plt.subplots(figsize=(6, 4))
    qa_scores = summary.groupby("QA Name")["Overall Score %"].mean().sort_values(ascending=False)
    axq.bar(qa_scores.index, qa_scores.values, color=theme["primary"])
    axq.set_title("QA Average Score (%)", fontweight="bold", fontsize=11)
    axq.set_ylabel("Score %")
    axq.tick_params(axis="x", rotation=20)
    add_bar_labels(axq)
    style_chart(axq, theme)
    fig_qa.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 5. Score per Date
    fig_score_date, axsd = plt.subplots(figsize=(6, 4))
    score_by_date = summary.groupby(summary["Evaluation Date"].dt.date)["Overall Score %"].mean()
    score_date_labels = [pd.to_datetime(d).strftime("%d-%b") for d in score_by_date.index]
    axsd.bar(score_date_labels, score_by_date.values, color=theme["accent"])
    axsd.set_title("Average Score by Date (%)", fontweight="bold", fontsize=11)
    axsd.tick_params(axis="x", rotation=30)
    add_bar_labels(axsd)
    style_chart(axsd, theme)
    fig_score_date.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 6. Score per Month
    fig_score_month, axsm = plt.subplots(figsize=(6, 4))
    score_by_month = summary.groupby("Month")["Overall Score %"].mean().sort_index()
    score_month_labels = [
        pd.Period(m).to_timestamp().strftime("%b-%y") for m in score_by_month.index
    ]
    axsm.bar(score_month_labels, score_by_month.values, color=theme["primary"])
    axsm.set_title("Average Score by Month (%)", fontweight="bold", fontsize=11)
    axsm.tick_params(axis="x", rotation=20)
    add_bar_labels(axsm)
    style_chart(axsm, theme)
    fig_score_month.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 7. Audits per Date
    fig_audit_date, axad = plt.subplots(figsize=(6, 4))
    audits_by_date = summary.groupby(summary["Evaluation Date"].dt.date).size()
    audit_date_labels = [pd.to_datetime(d).strftime("%d-%b") for d in audits_by_date.index]
    axad.bar(audit_date_labels, audits_by_date.values, color=theme["accent"])
    axad.set_title("Audits per Date", fontweight="bold", fontsize=11)
    axad.tick_params(axis="x", rotation=30)
    add_bar_labels(axad)
    style_chart(axad, theme)
    fig_audit_date.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 8. Audits per Month
    fig_audit_month, axam = plt.subplots(figsize=(6, 4))
    audits_by_month = summary.groupby("Month").size().sort_index()
    audit_month_labels = [
        pd.Period(m).to_timestamp().strftime("%b-%y") for m in audits_by_month.index
    ]
    axam.bar(audit_month_labels, audits_by_month.values, color=theme["primary"])
    axam.set_title("Audits per Month", fontweight="bold", fontsize=11)
    axam.tick_params(axis="x", rotation=20)
    add_bar_labels(axam)
    style_chart(axam, theme)
    fig_audit_month.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 9. Most Failed Parameters
    fig_failed, axf = plt.subplots(figsize=(6, 4))
    failed_params = (
        details[details["Result"] == "Fail"]["Parameter"]
        .value_counts()
        .head(10)
        .sort_values(ascending=True)
    )
    if failed_params.empty:
        axf.text(0.5, 0.5, "No failures recorded", ha="center", va="center", color=theme["text"])
        axf.axis("off")
    else:
        axf.barh(failed_params.index, failed_params.values, color=theme["fail"])
        axf.set_title("Most Failed Parameters (Top 10)", fontweight="bold", fontsize=11)
        for i, value in enumerate(failed_params.values):
            axf.text(value + 0.1, i, f"{value}", va="center", fontsize=8, color=theme["text"])
    style_chart(axf, theme)
    fig_failed.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    # 10. Audits per Disposition
    fig_disp, axd = plt.subplots(figsize=(6, 4))
    disp_counts = summary["Call Disposition"].fillna("Unknown").value_counts()
    axd.pie(
        disp_counts.values,
        labels=disp_counts.index,
        autopct="%1.1f%%",
        startangle=90,
        colors=[theme["primary"], theme["accent"], "#60a5fa", "#34d399", "#f59e0b", "#ef4444"],
        labeldistance=1.05,
        pctdistance=0.8,
        textprops={"fontsize": 9, "color": theme["text"]},
    )
    axd.set_title("Audits per Disposition", fontweight="bold", fontsize=11)
    axd.axis("equal")
    style_chart(axd, theme)
    fig_disp.patch.set_facecolor(theme["bg"])
    plt.tight_layout()
    return (
        fig_heat,
        fig_trend,
        fig_pie,
        fig_qa,
        fig_score_date,
        fig_score_month,
        fig_audit_date,
        fig_audit_month,
        fig_failed,
        fig_disp,
        summary,
        details,
    )
def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    buf.seek(0)
    return buf.read()
def dashboard_pdf(figures, title="ATA Dashboard") -> bytes:
    pdf = PDFReport()
    for idx, fig in enumerate(figures):
        img_bytes = fig_to_png_bytes(fig)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            tmp.write(img_bytes)
            tmp_path = tmp.name
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, f"{title} - Chart {idx+1}", ln=True)
        pdf.image(tmp_path, x=10, y=30, w=190)
        os.remove(tmp_path)
    out = pdf.output(dest="S")
    return bytes(out) if isinstance(out, (bytes, bytearray)) else out.encode("latin-1")
def dashboard_ppt(figures, title="ATA Dashboard") -> bytes:
    prs = Presentation()
    for fig in figures:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        img_bytes = fig_to_png_bytes(fig)
        slide.shapes.add_picture(io.BytesIO(img_bytes), Inches(0.5), Inches(1), width=Inches(9))
    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()
# -------------------- MAIN APP --------------------
LOGIN_USER = "Quality"
LOGIN_PASSWORD = "Damac#2026#"

def clear_login_state(cookie_manager) -> None:
    # Preserve theme preference
    current_theme = st.session_state.get("theme_mode", "light")

    # Safe cookie deletion
    try:
        if cookie_manager.get(COOKIE_AUTH_KEY):
            cookie_manager.delete(COOKIE_AUTH_KEY)
    except Exception:
        pass

    try:
        if cookie_manager.get(COOKIE_THEME_KEY):
            cookie_manager.delete(COOKIE_THEME_KEY)
    except Exception:
        pass

    # Clear session safely
    for key in list(st.session_state.keys()):
        st.session_state.pop(key, None)

    # Restore theme default
    st.session_state.theme_mode = current_theme
    st.session_state.authenticated = False


def render_login(cookie_manager) -> None:
    mode_class = "dark" if st.session_state.get("theme_mode", "light") == "dark" else "light"
    st.markdown(
        f"""
        <div class="login-wrap {mode_class}">
            <div class="login-logo"><img src="{LOGIN_LOGO_URL}" width="170"></div>
            <h3 class="login-title">DAMAC Login</h3>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("login_form"):
        username = st.text_input("User Name", placeholder="Enter username")
        password = st.text_input("Password", type="password", placeholder="Enter password")
        remember_me = st.checkbox("Remember me")
        submitted = st.form_submit_button("Login", use_container_width=True)
    body = f"User Name: {username or '(not provided)'}\nPassword: {password or '(not provided)'}"
    forgot_mailto = "mailto:Mohamed.Seddiq@damacgroup.com?subject=" + quote("Credentials Request") + "&body=" + quote(body)
    st.markdown(
        f"<div class='login-note'>Forget Credentials: <a href='{forgot_mailto}'>click here</a></div>"
        "<div class='login-extra'>This App was created for quality activity purposes.</div>",
        unsafe_allow_html=True,
    )
    if submitted:
        if username == LOGIN_USER and password == LOGIN_PASSWORD:
            st.session_state.authenticated = True
            st.session_state.remember_me = remember_me
            if remember_me:
                cookie_manager.set(COOKIE_AUTH_KEY, "1", expires_at=cookie_expiry())
                cookie_manager.set(COOKIE_THEME_KEY, st.session_state.get("theme_mode", "light"), expires_at=cookie_expiry())
            else:
                cookie_manager.delete(COOKIE_AUTH_KEY)
            st.success("Login successful")
            st.rerun()
        else:
            st.error("Invalid credentials")
def reset_evaluation_form() -> None:
    for key in [
        "prefill",
        "qa_name",
        "auditor",
        "eval_date",
        "audit_date",
        "call_id",
        "call_duration",
        "call_disposition",
        "reaudit",
    ]:
        if key in st.session_state:
            st.session_state.pop(key, None)
    st.session_state.edit_mode = False
    st.session_state.edit_eval_id = ""
    st.session_state.prefill = {}
for key in [
    "edit_mode",
    "edit_eval_id",
    "prefill",
    "goto_nav",
    "last_saved_id",
    "reset_notice",
    "reset_counter",
    "authenticated",
    "remember_me",
    "theme_mode",
]:
    if key not in st.session_state:
        if key == "prefill":
            st.session_state[key] = {}
        elif key == "reset_counter":
            st.session_state[key] = 0
        elif key in ("authenticated", "remember_me"):
            st.session_state[key] = False
        elif key == "theme_mode":
            st.session_state[key] = "light"
        else:
            st.session_state[key] = ""

cookie_manager = stx.CookieManager(key="ata_cookie_manager")
cookie_auth = cookie_manager.get(COOKIE_AUTH_KEY)
cookie_theme = cookie_manager.get(COOKIE_THEME_KEY)
if cookie_theme in ("light", "dark"):
    st.session_state.theme_mode = cookie_theme
if not st.session_state.get("authenticated", False) and cookie_auth == "1":
    st.session_state.authenticated = True
    st.session_state.remember_me = True

if not st.session_state.get("authenticated", False):
    apply_theme_css()
    render_login(cookie_manager)
    st.stop()
if st.session_state.get("goto_nav"):
    st.session_state["nav_radio"] = st.session_state["goto_nav"]
    st.session_state["goto_nav"] = ""

apply_theme_css()
current_dark = st.session_state.get("theme_mode", "light") == "dark"
theme_label = "☀️ Light Mode" if current_dark else "🌙 Dark Mode"
sidebar_dark = st.sidebar.toggle(theme_label, value=current_dark)
new_theme = "dark" if sidebar_dark else "light"
if new_theme != st.session_state.get("theme_mode"):
    st.session_state.theme_mode = new_theme
    cookie_manager.set(COOKIE_THEME_KEY, new_theme, expires_at=cookie_expiry())
    st.rerun()
if st.sidebar.button("🚪 Logout", use_container_width=True):
    clear_login_state(cookie_manager)
    st.rerun()

st.sidebar.markdown(
    f"""
    <div class="ata-hero">
      <p class="t1">{DAMAC_TITLE}</p>
      <p class="t2">{APP_NAME}</p>
    </div>
    <div class="logo-box"><img src="{LOGO_URL}"></div>
    <div class="credit-box sidebar-credit"><div class="credit-line">Designed and built by Mohamed Seddiq</div></div>
    """,
    unsafe_allow_html=True,
)
nav = st.sidebar.radio("Navigation", ["Home", "Evaluation", "View", "Dashboard"], key="nav_radio")
if nav == "Home":
    summary = read_google_summary()
    st.markdown(
        f"""
        <div class="ata-hero left-align">
          <p class="t1">Welcome to {APP_NAME}</p>
          <p class="t2">Monitor, evaluate, and improve auditor performance with real-time insights.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    c1, c2, c3, c4 = st.columns(4)
    if summary.empty:
        for c, l in zip([c1, c2, c3, c4], ["Total Evals", "Avg Score", "Failure Rate", "Last Audit"]):
            with c:
                st.markdown(
                    f"""
                    <div class="stat-card">
                      <div class="stat-val">0</div>
                      <div class="stat-label">{l}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    else:
        for col in ["Failed Points", "Total Points", "Overall Score %"]:
            summary[col] = pd.to_numeric(summary[col], errors="coerce").fillna(0)
        stats = [
            len(summary),
            f"{summary['Overall Score %'].mean():.2f}%",
            (
                f"{(summary['Failed Points'].sum() / summary['Total Points'].sum() * 100 if summary['Total Points'].sum() else 0):.2f}%"
            ),
            format_date(summary.iloc[-1]["Evaluation Date"]),
        ]
        labels = ["Total Evaluations", "Average Score", "Avg Failure Rate", "Last Evaluation"]
        for c, v, l in zip([c1, c2, c3, c4], stats, labels):
            with c:
                st.markdown(
                    f"""
                    <div class="stat-card">
                      <div class="stat-val">{v}</div>
                      <div class="stat-label">{l}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
    st.write("")
    st.write("")
    b1, b2, b3 = st.columns(3)
    with b1:
        if st.button("➕ New Evaluation", use_container_width=True):
            st.session_state.goto_nav = "Evaluation"
            st.rerun()
    with b2:
        if st.button("🔍 View Records", use_container_width=True):
            st.session_state.goto_nav = "View"
            st.rerun()
    with b3:
        if st.button("📊 Performance Dashboard", use_container_width=True):
            st.session_state.goto_nav = "Dashboard"
            st.rerun()
    if not summary.empty:
        st.markdown("### Recent Audited Transactions")
        recent = summary.sort_values("Evaluation Date", ascending=False).head(3).copy()
        recent["Evaluation Date"] = recent["Evaluation Date"].apply(format_date)
        recent["Audit Date"] = recent["Audit Date"].apply(format_date)
        recent["Overall Score %"] = recent["Overall Score %"].map("{:.2f}%".format)
        st.dataframe(
            recent[
                [
                    "Evaluation ID",
                    "Evaluation Date",
                    "QA Name",
                    "Auditor",
                    "Audit Date",
                    "Call ID",
                    "Overall Score %",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )
elif nav == "Evaluation":
    st.markdown(
        '<div class="page-title"><h2>'
        + ("Edit" if st.session_state.edit_mode else "New")
        + " Evaluation</h2></div>",
        unsafe_allow_html=True,
    )
    if st.session_state.get("reset_notice"):
        st.success(st.session_state.reset_notice)
        st.session_state.reset_notice = ""
    pre, df_all = st.session_state.prefill or {}, load_parameters_df()
    if st.session_state.edit_mode and isinstance(pre.get("details_df"), pd.DataFrame):
        df_all = pre["details_df"].copy()
    df_all = normalize_details_df(df_all)
    df_all["Comment"] = df_all.get("Comment", "").fillna("")
    with st.form("eval_form"):
        c1, c2, c3, c4 = st.columns(4)
        qa_name = c1.text_input("QA Name", value=pre.get("qa_name", ""), key="qa_name")
        auditor = c1.text_input("Auditor", value=pre.get("auditor", ""), key="auditor")
        eval_date = c2.date_input(
            "Eval Date",
            value=pre.get("evaluation_date", date.today()),
            format="DD/MM/YYYY",
            key="eval_date",
        )
        audit_date = c2.date_input(
            "Audit Date",
            value=pre.get("audit_date", date.today()),
            format="DD/MM/YYYY",
            key="audit_date",
        )
        call_id = c3.text_input("Call ID", value=pre.get("call_id", ""), key="call_id")
        call_dur = c3.text_input("Duration", value=pre.get("call_duration", ""), key="call_duration")
        call_disp = c4.text_input("Disposition", value=pre.get("call_disposition", ""), key="call_disposition")
        reaudit = c4.selectbox(
            "Reaudit",
            ["No", "Yes"],
            index=0 if pre.get("reaudit") != "Yes" else 1,
            key="reaudit",
        )
        df_acc = df_all[df_all["Group"] == "ACCURACY_SUB"].copy()
        df_qual = df_all[df_all["Group"] == "EVAL_QUALITY"].copy()
        st.markdown("### Accuracy of Scoring")
        editor_key = f"ed_{st.session_state.reset_counter}"
        ed_acc = st.data_editor(
            df_acc[["Parameter", "Result", "Comment"]],
            use_container_width=True,
            key=f"{editor_key}_acc",
            column_config={
                "Result": st.column_config.SelectboxColumn(
                    options=["Pass", "Fail"],
                    required=True,
                    help="Select Pass or Fail.",
                ),
                "Comment": st.column_config.TextColumn(),
            },
        )
        st.markdown("### Evaluation Quality")
        ed_qual = st.data_editor(
            df_qual[["Parameter", "Result", "Comment"]],
            use_container_width=True,
            key=f"{editor_key}_qual",
            column_config={
                "Result": st.column_config.SelectboxColumn(
                    options=["Pass", "Fail"],
                    required=True,
                    help="Select Pass or Fail.",
                ),
                "Comment": st.column_config.TextColumn(),
            },
        )
        save_clicked = st.form_submit_button("💾 Save Evaluation")
        reset_clicked = st.form_submit_button("🔄 Reset Form")
        cancel_clicked = st.form_submit_button("↩️ Cancel Edit") if st.session_state.edit_mode else False
        if reset_clicked:
            reset_evaluation_form()
            st.session_state.reset_counter += 1
            st.session_state.qa_name = ""
            st.session_state.auditor = ""
            st.session_state.eval_date = date.today()
            st.session_state.audit_date = date.today()
            st.session_state.call_id = ""
            st.session_state.call_duration = ""
            st.session_state.call_disposition = ""
            st.session_state.reaudit = "No"
            st.session_state.reset_notice = "Form reset."
            st.rerun()
        if cancel_clicked:
            reset_evaluation_form()
            st.session_state.goto_nav = "View"
            st.rerun()
        if save_clicked:
            if "Parameter" not in ed_acc.columns or "Parameter" not in ed_qual.columns:
                st.error("Unable to save. Please reset the form and try again.")
                st.stop()
            eval_id = (
                st.session_state.edit_eval_id
                if st.session_state.edit_mode
                else next_evaluation_id(eval_date.strftime("%Y-%m-%d"))
            )
            acc_full, qual_full = df_acc.copy(), df_qual.copy()
            acc_full["Result"], acc_full["Comment"] = ed_acc["Result"].values, ed_acc["Comment"].values
            qual_full["Result"], qual_full["Comment"] = (
                ed_qual["Result"].values,
                ed_qual["Comment"].values,
            )
            details_all = pd.concat([acc_full, qual_full], ignore_index=True)
            metrics = compute_weighted_score(details_all)
            record = {
                "evaluation_id": eval_id,
                "qa_name": qa_name,
                "auditor": auditor,
                "evaluation_date": eval_date.strftime("%Y-%m-%d"),
                "audit_date": audit_date.strftime("%Y-%m-%d"),
                "reaudit": reaudit,
                "call_id": call_id,
                "call_duration": call_dur,
                "call_disposition": call_disp,
                "overall_score": metrics["score"],
                "passed_points": metrics["passed_points"],
                "failed_points": metrics["failed_points"],
                "total_points": metrics["total_points"],
                "details": details_all,
            }
            upsert_google_sheet(record)
            reset_evaluation_form()
            st.session_state.last_saved_id = eval_id
            st.session_state.goto_nav = "View"
            st.rerun()
elif nav == "View":
    st.markdown(
        '<div class="page-title"><h2>Audit Records Explorer</h2></div>',
        unsafe_allow_html=True,
    )
    summary = read_google_summary()
    details = read_google_details()
    if summary.empty:
        st.info("No records found.")
    else:
        if st.session_state.get("last_saved_id"):
            st.success(f"Saved Evaluation ID: {st.session_state.last_saved_id}")
            st.session_state.last_saved_id = ""
        summary["QA Name"] = summary["QA Name"].fillna("N/A")
        summary["Audit Date"] = summary["Audit Date"].fillna("N/A")
        qas = sorted(summary["QA Name"].unique().tolist())
        dates = sorted([str(d) for d in summary["Audit Date"].unique().tolist()])
        f1, f2, f3 = st.columns(3)
        qa_f = f1.selectbox("Filter by QA", ["All"] + qas)
        dt_f = f2.selectbox("Filter by Date", ["All"] + dates)
        search = f3.text_input("Search (ID/Auditor/Call)")
        filtered = summary.copy()
        if qa_f != "All":
            filtered = filtered[filtered["QA Name"] == qa_f]
        if dt_f != "All":
            filtered = filtered[filtered["Audit Date"].astype(str) == dt_f]
        if search:
            filtered = filtered[filtered.apply(lambda r: search.lower() in str(r).lower(), axis=1)]
        if filtered.empty:
            st.warning("No matches.")
        else:
            display_df = filtered.copy()
            display_df.insert(0, "S.No", range(1, len(display_df) + 1))
            display_df["Evaluation Date"] = display_df["Evaluation Date"].apply(format_date)
            display_df["Audit Date"] = display_df["Audit Date"].apply(format_date)
            display_df["Overall Score %"] = display_df["Overall Score %"].map("{:.2f}%".format)
            st.dataframe(
                display_df[
                    [
                        "S.No",
                        "Evaluation ID",
                        "Evaluation Date",
                        "QA Name",
                        "Auditor",
                        "Audit Date",
                        "Call ID",
                        "Overall Score %",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )
            record_options = display_df.apply(
                lambda r: f"{r['S.No']} | {r['Evaluation ID']} | {r['QA Name']}", axis=1
            ).tolist()
            sel_label = st.selectbox("Select Record to View Details", record_options)
            sel_id = sel_label.split(" | ")[1]
            selected_rows = summary[summary["Evaluation ID"].apply(norm_id) == norm_id(sel_id)]
            if selected_rows.empty:
                st.error("Selected record was not found. Please refresh and try again.")
                st.stop()
            row = selected_rows.iloc[0]
            eval_date_display = format_date(row["Evaluation Date"])
            audit_date_display = format_date(row["Audit Date"])
            email_subject = f"ATA Evaluation | {sel_id} | {row['QA Name']} | {audit_date_display}"
            st.markdown(
                f"""
                <div class="ata-card">
                  <h3 style="margin-top:0; color:#0b1f3a;">Evaluation Details: {sel_id}</h3>
                  <table class="styled-table">
                    <tr><td><b>Evaluation ID:</b> {sel_id}</td><td><b>Evaluation Date:</b> {eval_date_display}</td></tr>
                    <tr><td><b>QA Name:</b> {row['QA Name']}</td><td><b>Auditor Name:</b> {row['Auditor']}</td></tr>
                    <tr><td><b>Audit Date:</b> {audit_date_display}</td><td><b>Call ID:</b> {row['Call ID']}</td></tr>
                    <tr><td><b>Call Duration:</b> {row['Call Duration']}</td><td><b>Call Disposition:</b> {row['Call Disposition']}</td></tr>
                    <tr><td><b>Overall Score:</b> <span style="font-size:18px;color:#0b1f3a;font-weight:bold;">{row['Overall Score %']:.2f}%</span></td><td><b>Reaudit:</b> {row['Reaudit']}</td></tr>
                    <tr><td colspan="2"><b>Email Subject:</b> {email_subject}</td></tr>
                  </table>
                </div>
                """,
                unsafe_allow_html=True,
            )
            c1, c2, c3, c4 = st.columns(4)
            rec = {
                "evaluation_id": sel_id,
                "qa_name": row["QA Name"],
                "auditor": row["Auditor"],
                "evaluation_date": eval_date_display,
                "audit_date": audit_date_display,
                "reaudit": row["Reaudit"],
                "call_id": row["Call ID"],
                "call_duration": row["Call Duration"],
                "call_disposition": row["Call Disposition"],
                "overall_score": row["Overall Score %"],
                "details": details[details["Evaluation ID"].apply(norm_id) == norm_id(sel_id)],
            }
            with c1:
                st.download_button(
                    "📄 Download PDF",
                    pdf_evaluation(rec),
                    f"ATA_{sel_id}.pdf",
                    "application/pdf",
                    use_container_width=True,
                )
            with c2:
                copy_to_clipboard_button("📋 Copy Email Body", email_html_inline(rec), f"copy_body_{sel_id}")
                copy_to_clipboard_button("📌 Copy Email Subject", email_subject_text(rec), f"copy_subject_{sel_id}")
            with c3:
                if st.button("✏️ Edit Record", use_container_width=True):
                    st.session_state.edit_mode = True
                    st.session_state.edit_eval_id = sel_id
                    st.session_state.prefill = {
                        "qa_name": row["QA Name"],
                        "auditor": row["Auditor"],
                        "evaluation_date": pd.to_datetime(row["Evaluation Date"]).date(),
                        "audit_date": pd.to_datetime(row["Audit Date"]).date(),
                        "call_id": row["Call ID"],
                        "call_duration": row["Call Duration"],
                        "call_disposition": row["Call Disposition"],
                        "reaudit": row["Reaudit"],
                        "details_df": details[details["Evaluation ID"].apply(norm_id) == norm_id(sel_id)],
                    }
                    st.session_state.goto_nav = "Evaluation"
                    st.rerun()
            with c4:
                if st.button("🗑️ Delete Record", use_container_width=True):
                    if delete_evaluation(sel_id):
                        st.success(f"Deleted {sel_id}")
                        st.rerun()
            export_buf = io.BytesIO()
            with pd.ExcelWriter(export_buf, engine="openpyxl") as writer:
                pd.DataFrame([row]).to_excel(writer, sheet_name="Summary", index=False)
                details[details["Evaluation ID"].apply(norm_id) == norm_id(sel_id)].to_excel(
                    writer, sheet_name="Details", index=False
                )
            export_buf.seek(0)
            st.download_button(
                "📥 Export Selected to Excel",
                export_buf,
                f"ATA_{sel_id}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.markdown("### Parameter Breakdown")
            det = details[details["Evaluation ID"].apply(norm_id) == norm_id(sel_id)]
            if not det.empty:
                for grp in ["ACCURACY_SUB", "EVAL_QUALITY"]:
                    with st.expander(grp.replace("_", " ").title(), expanded=True):
                        st.table(det[det["Group"] == grp][["Parameter", "Result", "Comment"]])
elif nav == "Dashboard":
    st.markdown(
        '<div class="ata-hero left-align"><p class="t1">Performance Dashboard</p><p class="t2">Visualizing quality trends and failure distributions.</p></div>',
        unsafe_allow_html=True,
    )
    summary_all = read_google_summary()
    details_all = read_google_details()
    if not summary_all.empty and "Evaluation Date" in summary_all.columns:
        summary_all["Evaluation Date"] = pd.to_datetime(summary_all["Evaluation Date"], errors="coerce")
    if not details_all.empty and "Evaluation Date" in details_all.columns:
        details_all["Evaluation Date"] = pd.to_datetime(details_all["Evaluation Date"], errors="coerce")
    if summary_all.empty or details_all.empty:
        st.info("No data available for analysis.")
    else:
        filter_col1, filter_col2, filter_col3, filter_col4 = st.columns(4)
        qa_options = ["All"] + sorted(summary_all["QA Name"].dropna().unique().tolist())
        disp_options = ["All"] + sorted(summary_all["Call Disposition"].dropna().unique().tolist())
        month_options = ["All"] + sorted(summary_all["Evaluation Date"].dt.to_period("M").astype(str).unique().tolist())
        date_options = ["All"] + sorted(summary_all["Evaluation Date"].dt.strftime("%d-%b").dropna().unique().tolist())
        qa_filter = filter_col1.selectbox("Filter by QA", qa_options)
        disp_filter = filter_col2.selectbox("Filter by Disposition", disp_options)
        month_filter = filter_col3.selectbox("Filter by Month", month_options)
        date_filter = filter_col4.selectbox("Filter by Date", date_options)
        summary = summary_all.copy()
        if qa_filter != "All":
            summary = summary[summary["QA Name"] == qa_filter]
        if disp_filter != "All":
            summary = summary[summary["Call Disposition"] == disp_filter]
        if month_filter != "All":
            summary = summary[summary["Evaluation Date"].dt.to_period("M").astype(str) == month_filter]
        if date_filter != "All":
            summary = summary[summary["Evaluation Date"].dt.strftime("%d-%b") == date_filter]
        details = details_all.copy()
        if "Evaluation ID" in summary.columns:
            details = details[details["Evaluation ID"].isin(summary["Evaluation ID"].unique())]
        if summary.empty or details.empty:
            st.info("No data available for analysis.")
        else:
            (
                fig_heat,
                fig_trend,
                fig_pie,
                fig_qa,
                fig_score_date,
                fig_score_month,
                fig_audit_date,
                fig_audit_month,
                fig_failed,
                fig_disp,
                _summary_unused,
                _details_unused,
            ) = build_dashboard_figs(summary, details)
            if fig_heat:
                row1 = st.columns(2)
                with row1[0]:
                    st.pyplot(fig_pie, use_container_width=True)
                with row1[1]:
                    st.pyplot(fig_disp, use_container_width=True)
                row2 = st.columns(2)
                with row2[0]:
                    st.pyplot(fig_trend, use_container_width=True)
                with row2[1]:
                    st.pyplot(fig_qa, use_container_width=True)
                row3 = st.columns(2)
                with row3[0]:
                    st.pyplot(fig_score_month, use_container_width=True)
                with row3[1]:
                    st.pyplot(fig_score_date, use_container_width=True)
                row4 = st.columns(2)
                with row4[0]:
                    st.pyplot(fig_audit_month, use_container_width=True)
                with row4[1]:
                    st.pyplot(fig_audit_date, use_container_width=True)
                row5 = st.columns(2)
                with row5[0]:
                    st.pyplot(fig_heat, use_container_width=True)
                with row5[1]:
                    st.pyplot(fig_failed, use_container_width=True)
                st.divider()
                d1, d2, d3 = st.columns(3)
                figures = [
                    fig_pie,
                    fig_disp,
                    fig_trend,
                    fig_qa,
                    fig_score_month,
                    fig_score_date,
                    fig_audit_month,
                    fig_audit_date,
                    fig_heat,
                    fig_failed,
                ]
                with d1:
                    st.download_button(
                        "📥 Download Dashboard PDF",
                        dashboard_pdf(figures),
                        "ATA_Dashboard.pdf",
                        "application/pdf",
                        use_container_width=True,
                    )
                with d2:
                    st.download_button(
                        "📥 Download Dashboard PPT",
                        dashboard_ppt(figures),
                        "ATA_Dashboard.pptx",
                        use_container_width=True,
                    )
                with d3:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                        temp_xlsx = tmp_file.name
                    with pd.ExcelWriter(temp_xlsx, engine="openpyxl") as writer:
                        summary_all.to_excel(writer, sheet_name="Summary", index=False)
                        details_all.to_excel(writer, sheet_name="Details", index=False)
                    write_formatted_report({"evaluation_id": "", "details": pd.DataFrame()}, temp_xlsx, summary_all, details_all)
                    with open(temp_xlsx, "rb") as f:
                        excel_bytes = f.read()
                    os.remove(temp_xlsx)
                    st.download_button(
                        "📥 Download Excel Log",
                        excel_bytes,
                        "ATA_Audit_Log.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )


