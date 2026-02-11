import io
import json
import os
import tempfile
from datetime import date, datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from fpdf import FPDF
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

# -------------------- APP CONFIG --------------------
st.set_page_config(page_title="DAMAC | ATA Tool", layout="wide")

PARAMETERS_JSON = "parameters.json"
EXPORT_XLSX = "ATA_Audit_Log.xlsx"
LOGO_FILE = "logo.png"  # optional

DAMAC_TITLE = "DAMAC Properties"
DAMAC_SUB1 = "Quality Assurance"
DAMAC_SUB2 = "Telesales Division"
APP_NAME = "ATA Audit the Auditor"

# -------------------- UI THEME --------------------
st.markdown(
    """
<style>
.block-container { padding-top: 1.0rem; }
div[data-testid="stMetric"] { background:#ffffff; border:1px solid #eceff3; padding:14px; border-radius:14px; }
.ata-hero{
  background: linear-gradient(90deg, #0b1f3a, #173a6b);
  padding: 18px 18px;
  border-radius: 18px;
  color: white;
  border: 1px solid rgba(255,255,255,.10);
}
.ata-hero .t1 { font-size: 20px; font-weight: 900; margin: 0; }
.ata-hero .t2 { font-size: 13px; opacity: .92; margin: 0; margin-top: 6px; line-height:1.5; }
.ata-card { background:#fff; border:1px solid #eceff3; border-radius:16px; padding:16px; }
.ata-muted { color:#6b7280; font-size:13px; line-height:1.6; }
.small-note { color:#6b7280; font-size:12px; line-height:1.5; }
.kpi-chip{
  display:inline-block;
  padding:6px 10px;
  border-radius:999px;
  background:#f6f8fb;
  border:1px solid #eceff3;
  margin-right:6px;
  font-size:12px;
  color:#374151;
}
</style>
""",
    unsafe_allow_html=True,
)

# -------------------- PARAMETERS --------------------
DEFAULT_PARAMETERS = {
    "form_name": APP_NAME,
    "parameters": [
        {
            "Parameter": "Accuracy of Scoring",
            "Description": "Applied the scoring correctly with no unjustified or incorrect deductions",
        },
        {
            "Parameter": "Adherence to QA Guidelines",
            "Description": "Followed QA process and aligned with calibration standards",
        },
        {
            "Parameter": "Evidence & Notes",
            "Description": "Left a clear, specific and improvement-focused comment",
        },
        {
            "Parameter": "Objectivity & Fairness",
            "Description": "Evaluation is unbiased and fact-based",
        },
        {
            "Parameter": "Critical Error Identification",
            "Description": "Correct identification of fatal errors",
        },
        {
            "Parameter": "Evaluation Variety & Sample Coverage",
            "Description": "Evaluations cover a balanced mix of call durations and call types",
        },
        {
            "Parameter": "Feedback Actionability",
            "Description": "Conducted coaching session on the call topic (If required)",
        },
        {
            "Parameter": "Timeliness & Completeness",
            "Description": "On track with the evaluations target SLA",
        },
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
    df["Result"] = "Pass"
    df["Comment"] = ""
    return df[["Parameter", "Description", "Result", "Comment"]]


# -------------------- ID NORMALIZATION --------------------
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


# -------------------- EXCEL HELPERS --------------------
def safe_read_excel(path: str, sheet: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def next_evaluation_id(evaluation_date_str: str) -> str:
    """
    Enterprise ID: ATA-YYYYMMDD-0001
    Sequence resets per evaluation date.
    """
    yyyymmdd = evaluation_date_str.replace("-", "")
    summary = safe_read_excel(EXPORT_XLSX, "Summary")
    if summary.empty or "Evaluation ID" not in summary.columns:
        return f"ATA-{yyyymmdd}-0001"

    prefix = f"ATA-{yyyymmdd}-"
    existing = summary["Evaluation ID"].apply(norm_id)
    existing = existing[existing.str.startswith(prefix)]
    if existing.empty:
        return f"ATA-{yyyymmdd}-0001"

    def _seq(x: str) -> int:
        try:
            return int(x.split("-")[-1])
        except Exception:
            return 0

    max_seq = max(existing.apply(_seq).tolist() + [0])
    return f"ATA-{yyyymmdd}-{max_seq + 1:04d}"


def upsert_excel(record: dict) -> None:
    summary_existing = safe_read_excel(EXPORT_XLSX, "Summary")
    details_existing = safe_read_excel(EXPORT_XLSX, "Details")

    rid = norm_id(record["evaluation_id"])

    if not summary_existing.empty and "Evaluation ID" in summary_existing.columns:
        summary_existing["_rid"] = summary_existing["Evaluation ID"].apply(norm_id)
        summary_existing = summary_existing[summary_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")

    if not details_existing.empty and "Evaluation ID" in details_existing.columns:
        details_existing["_rid"] = details_existing["Evaluation ID"].apply(norm_id)
        details_existing = details_existing[details_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")

    summary_row = pd.DataFrame(
        [
            {
                "Evaluation ID": rid,
                "Evaluation Date": record["evaluation_date"],
                "Audit Date": record["audit_date"],
                "QA Name": record["qa_name"],
                "Auditor": record["auditor"],
                "Call ID": record["call_id"],
                "Call Duration": record["call_duration"],
                "Call Disposition": record["call_disposition"],
                "Overall Score %": record["overall_score"],
                "Passed": record["passed"],
                "Failed": record["failed"],
                "Last Updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        ]
    )

    details_df = record["details"].copy()
    details_df.insert(0, "Evaluation ID", rid)
    details_df.insert(1, "Evaluation Date", record["evaluation_date"])
    details_df.insert(2, "Audit Date", record["audit_date"])
    details_df.insert(3, "QA Name", record["qa_name"])
    details_df.insert(4, "Auditor", record["auditor"])
    details_df.insert(5, "Call ID", record["call_id"])
    details_df.insert(6, "Overall Score %", record["overall_score"])

    out_summary = pd.concat([summary_existing, summary_row], ignore_index=True)
    out_details = pd.concat([details_existing, details_df], ignore_index=True)

    with pd.ExcelWriter(EXPORT_XLSX, engine="openpyxl") as writer:
        out_summary.to_excel(writer, sheet_name="Summary", index=False)
        out_details.to_excel(writer, sheet_name="Details", index=False)


def delete_evaluation(eval_id: str) -> bool:
    rid = norm_id(eval_id)
    summary_existing = safe_read_excel(EXPORT_XLSX, "Summary")
    details_existing = safe_read_excel(EXPORT_XLSX, "Details")

    if summary_existing.empty and details_existing.empty:
        return False

    changed = False

    if not summary_existing.empty and "Evaluation ID" in summary_existing.columns:
        summary_existing["_rid"] = summary_existing["Evaluation ID"].apply(norm_id)
        before = len(summary_existing)
        summary_existing = summary_existing[summary_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")
        changed = changed or (len(summary_existing) != before)

    if not details_existing.empty and "Evaluation ID" in details_existing.columns:
        details_existing["_rid"] = details_existing["Evaluation ID"].apply(norm_id)
        before = len(details_existing)
        details_existing = details_existing[details_existing["_rid"] != rid].drop(columns=["_rid"], errors="ignore")
        changed = changed or (len(details_existing) != before)

    with pd.ExcelWriter(EXPORT_XLSX, engine="openpyxl") as writer:
        summary_existing.to_excel(writer, sheet_name="Summary", index=False)
        details_existing.to_excel(writer, sheet_name="Details", index=False)

    return changed


# -------------------- SCORING --------------------
def overall_score(df: pd.DataFrame) -> float:
    total = len(df)
    if total == 0:
        return 0.0
    passed = int((df["Result"] == "Pass").sum())
    return round((passed / total) * 100, 2)


# -------------------- EMAIL HTML --------------------
def email_html_inline(record: dict) -> str:
    rows = []
    for _, r in record["details"].iterrows():
        status = r["Result"]
        badge_style = "display:inline-block;padding:4px 10px;border-radius:999px;font-weight:800;font-size:12px;"
        if status == "Pass":
            badge_style += "background:#e9f7ef;color:#1f8f4a;"
        else:
            badge_style += "background:#fde8e6;color:#d93025;"

        comment = str(r["Comment"]).strip() if str(r["Comment"]).strip() else "-"
        rows.append(
            f"""
            <tr>
              <td style="padding:10px;border-bottom:1px solid #eee;">{r['Parameter']}</td>
              <td style="padding:10px;border-bottom:1px solid #eee;width:120px;">
                <span style="{badge_style}">{status}</span>
              </td>
              <td style="padding:10px;border-bottom:1px solid #eee;">{comment}</td>
            </tr>
            """
        )

    return f"""<!doctype html>
<html>
<head><meta charset="utf-8"></head>
<body>
  <div style="font-family:Calibri, Arial, sans-serif;">
    <div style="padding:14px 16px;background:#0b1f3a;color:#fff;border-radius:12px;">
      <div style="font-size:18px;font-weight:800;">{DAMAC_TITLE} | QA Audit Evaluation</div>
      <div style="opacity:.9;">{DAMAC_SUB1} | {DAMAC_SUB2}</div>
    </div>

    <div style="margin-top:10px;padding:14px 16px;border:1px solid #eceff3;border-radius:12px;">
      <div style="font-size:14px;line-height:1.7;">
        <div><span style="color:#6b7280;">Evaluation ID:</span> {record['evaluation_id']}</div>
        <div><span style="color:#6b7280;">QA Name:</span> {record['qa_name']}</div>
        <div><span style="color:#6b7280;">Auditor:</span> {record['auditor']}</div>
        <div><span style="color:#6b7280;">Evaluation Date:</span> {record['evaluation_date']}</div>
        <div><span style="color:#6b7280;">Audit Date:</span> {record['audit_date']}</div>
        <div><span style="color:#6b7280;">Call ID:</span> {record['call_id']}</div>
        <div><span style="color:#6b7280;">Call Duration:</span> {record['call_duration']}</div>
        <div><span style="color:#6b7280;">Disposition:</span> {record['call_disposition']}</div>
      </div>

      <div style="margin-top:10px;padding:10px 12px;background:#f6f8fb;border-radius:12px;">
        <div style="color:#6b7280;font-size:13px;">Overall Score</div>
        <div style="font-size:26px;font-weight:900;color:#0b1f3a;line-height:1;">{record['overall_score']}%</div>
        <div style="color:#6b7280;font-size:13px;margin-top:4px;">Passed: {record['passed']} | Failed: {record['failed']}</div>
      </div>

      <div style="margin-top:12px;border:1px solid #eceff3;border-radius:12px;overflow:hidden;">
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <thead>
            <tr style="background:#f2f5fb;">
              <th style="text-align:left;padding:10px;">Parameter</th>
              <th style="text-align:left;padding:10px;width:120px;">Result</th>
              <th style="text-align:left;padding:10px;">Comment</th>
            </tr>
          </thead>
          <tbody>
            {''.join(rows)}
          </tbody>
        </table>
      </div>

      <div style="margin-top:10px;color:#6b7280;font-size:12px;">
        Internal use only.
      </div>
    </div>
  </div>
</body>
</html>"""


def build_eml(record: dict, html: str, attach_pdf: bool = True) -> bytes:
    """
    Creates an .eml draft with a rendered HTML body (Outlook-friendly).
    User opens the .eml and Outlook will render the HTML.
    """
    msg = MIMEMultipart("mixed")
    subject = f"ATA Evaluation | {record['evaluation_id']} | {record['qa_name']} | {record['auditor']}"
    msg["Subject"] = subject

    alt = MIMEMultipart("alternative")
    plain = (
        f"ATA Evaluation\n"
        f"Evaluation ID: {record['evaluation_id']}\n"
        f"QA Name: {record['qa_name']}\n"
        f"Auditor: {record['auditor']}\n"
        f"Overall Score: {record['overall_score']}%\n"
    )
    alt.attach(MIMEText(plain, "plain", "utf-8"))
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    if attach_pdf:
        pdf_bytes = pdf_evaluation(record)
        part = MIMEApplication(pdf_bytes, _subtype="pdf")
        part.add_header("Content-Disposition", "attachment", filename=f"ATA_Evaluation_{record['evaluation_id']}.pdf")
        msg.attach(part)

    return msg.as_bytes()


# -------------------- PDF (EVALUATION) --------------------
class PDFReport(FPDF):
    def header(self):
        self.set_font("Arial", "B", 11)
        self.cell(0, 7, f"{DAMAC_TITLE} | {DAMAC_SUB1} | {DAMAC_SUB2}", ln=True, align="C")
        self.ln(1)

    def footer(self):
        self.set_y(-12)
        self.set_font("Arial", "", 9)
        self.cell(0, 8, f"Page {self.page_no()}", align="C")


def pdf_evaluation(record: dict) -> bytes:
    pdf = PDFReport()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=12)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "ATA Evaluation Report", ln=True)

    pdf.set_font("Arial", "", 11)
    meta = [
        f"Evaluation ID: {record['evaluation_id']}",
        f"QA Name: {record['qa_name']}",
        f"Auditor: {record['auditor']}",
        f"Evaluation Date: {record['evaluation_date']}",
        f"Audit Date: {record['audit_date']}",
        f"Call ID: {record['call_id']}",
        f"Call Duration: {record['call_duration']}",
        f"Call Disposition: {record['call_disposition']}",
        f"Overall Score: {record['overall_score']}%",
        f"Passed: {record['passed']} | Failed: {record['failed']}",
    ]
    for m in meta:
        pdf.cell(0, 7, m, ln=True)

    pdf.ln(3)
    col_w = [75, 20, 95]

    pdf.set_font("Arial", "B", 10)
    pdf.cell(col_w[0], 8, "Parameter", border=1)
    pdf.cell(col_w[1], 8, "Result", border=1, align="C")
    pdf.cell(col_w[2], 8, "Comment", border=1)
    pdf.ln()

    pdf.set_font("Arial", "", 10)
    for _, r in record["details"].iterrows():
        param = str(r["Parameter"])
        res = str(r["Result"])
        comment = str(r["Comment"]).strip() if str(r["Comment"]).strip() else "-"

        x = pdf.get_x()
        y = pdf.get_y()

        pdf.multi_cell(col_w[0], 6, param, border=1)
        pdf.set_xy(x + col_w[0], y)
        pdf.multi_cell(col_w[1], 6, res, border=1, align="C")
        pdf.set_xy(x + col_w[0] + col_w[1], y)
        pdf.multi_cell(col_w[2], 6, comment, border=1)

        pdf.set_y(max(pdf.get_y(), y + 6))

    out = io.BytesIO()
    out.write(pdf.output(dest="S").encode("latin-1"))
    out.seek(0)
    return out.read()


# -------------------- DASHBOARD: FIGS + EXPORTS --------------------
def build_dashboard_figs():
    if not os.path.exists(EXPORT_XLSX):
        return None, None, pd.DataFrame(), pd.DataFrame()

    summary = safe_read_excel(EXPORT_XLSX, "Summary")
    details = safe_read_excel(EXPORT_XLSX, "Details")

    if summary.empty or details.empty:
        return None, None, summary, details

    summary["Evaluation Date"] = pd.to_datetime(summary.get("Evaluation Date"), errors="coerce")
    details["Evaluation Date"] = pd.to_datetime(details.get("Evaluation Date"), errors="coerce")
    summary["Month"] = summary["Evaluation Date"].dt.to_period("M").astype(str)
    details["Month"] = details["Evaluation Date"].dt.to_period("M").astype(str)

    summary["Total"] = pd.to_numeric(summary.get("Passed"), errors="coerce").fillna(0) + pd.to_numeric(
        summary.get("Failed"), errors="coerce"
    ).fillna(0)
    summary["Failed"] = pd.to_numeric(summary.get("Failed"), errors="coerce").fillna(0)
    summary["Failure Rate"] = summary.apply(lambda r: (r["Failed"] / r["Total"]) if r["Total"] else 0, axis=1)
    trend = summary.groupby("Month")["Failure Rate"].mean().sort_index()

    fig_trend = plt.figure(figsize=(9, 3.6))
    ax = fig_trend.add_subplot(111)
    ax.plot(trend.index, trend.values * 100)
    ax.set_title("Failure Trend Over Time")
    ax.set_xlabel("Month")
    ax.set_ylabel("Avg Failure Rate (%)")
    ax.grid(True)
    for i, v in enumerate(trend.values):
        ax.text(i, v * 100 + 0.6, f"{v*100:.1f}%", ha="center", fontsize=9)
    fig_trend.tight_layout()

    fail_rows = details[details["Result"] == "Fail"].copy()
    heat = pd.DataFrame()
    if not fail_rows.empty:
        heat = pd.pivot_table(
            fail_rows,
            index="Parameter",
            columns="Month",
            values="Result",
            aggfunc="count",
            fill_value=0,
        )

    fig_heat = plt.figure(figsize=(12, 4.2))
    axh = fig_heat.add_subplot(111)
    if heat.empty:
        axh.text(0.5, 0.5, "No failures to display", ha="center", va="center", fontsize=12)
        axh.axis("off")
    else:
        im = axh.imshow(heat.values, aspect="auto")
        axh.set_title("Failure Heatmap by Parameter")
        axh.set_xticks(range(len(heat.columns)))
        axh.set_xticklabels(heat.columns, rotation=0)
        axh.set_yticks(range(len(heat.index)))
        axh.set_yticklabels(heat.index)
        for i in range(heat.shape[0]):
            for j in range(heat.shape[1]):
                axh.text(j, i, str(int(heat.iloc[i, j])), ha="center", va="center", fontsize=8)
        fig_heat.colorbar(im, ax=axh, fraction=0.02, pad=0.02)
        fig_heat.tight_layout()

    return fig_heat, fig_trend, summary, details


def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=180, bbox_inches="tight")
    buf.seek(0)
    return buf.read()


def dashboard_pdf(figures, title="ATA Dashboard") -> bytes:
    pdf = PDFReport()
    pdf.set_auto_page_break(auto=True, margin=12)

    for idx, fig in enumerate(figures, start=1):
        img_bytes = fig_to_png_bytes(fig)
        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
            tmp.write(img_bytes)
            tmp_path = tmp.name

        pdf.add_page()
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, f"{title} | Chart {idx}", ln=True)
        pdf.image(tmp_path, x=10, y=30, w=190)

        try:
            os.remove(tmp_path)
        except Exception:
            pass

    out = io.BytesIO()
    out.write(pdf.output(dest="S").encode("latin-1"))
    out.seek(0)
    return out.read()


def dashboard_ppt(figures, title="ATA Dashboard") -> bytes:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    slide = prs.slides.add_slide(blank)
    tx = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), Inches(12.1), Inches(1.2))
    tf = tx.text_frame
    tf.text = f"{DAMAC_TITLE} | {DAMAC_SUB1} | {DAMAC_SUB2}"
    tf.paragraphs[0].font.size = Pt(24)
    tf.paragraphs[0].font.bold = True
    p2 = tf.add_paragraph()
    p2.text = title
    p2.font.size = Pt(18)

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    margin = Inches(0.4)

    for fig in figures:
        img_bytes = fig_to_png_bytes(fig)
        img = Image.open(io.BytesIO(img_bytes))
        img_w, img_h = img.size

        slide = prs.slides.add_slide(blank)

        max_w = slide_w - 2 * margin
        max_h = slide_h - 2 * margin

        img_ratio = img_w / img_h
        box_ratio = max_w / max_h

        if img_ratio > box_ratio:
            w = max_w
            h = w / img_ratio
        else:
            h = max_h
            w = h * img_ratio

        left = (slide_w - w) / 2
        top = (slide_h - h) / 2

        slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=w, height=h)

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.read()


# -------------------- SESSION STATE --------------------
if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False
if "edit_eval_id" not in st.session_state:
    st.session_state.edit_eval_id = ""
if "prefill" not in st.session_state:
    st.session_state.prefill = {}
if "goto_nav" not in st.session_state:
    st.session_state.goto_nav = ""
if "show_email_preview" not in st.session_state:
    st.session_state.show_email_preview = False
if "pending_delete" not in st.session_state:
    st.session_state.pending_delete = ""

# Navigation redirect fix: must run before sidebar radio is instantiated
if st.session_state.get("goto_nav"):
    st.session_state["nav_radio"] = st.session_state["goto_nav"]
    st.session_state["goto_nav"] = ""

# -------------------- SIDEBAR --------------------
st.sidebar.markdown(
    f"""
<div class="ata-hero">
  <p class="t1">{DAMAC_TITLE}</p>
  <p class="t2">{DAMAC_SUB1} | {DAMAC_SUB2}<br>{APP_NAME}</p>
</div>
""",
    unsafe_allow_html=True,
)
if os.path.exists(LOGO_FILE):
    st.sidebar.image(LOGO_FILE, use_container_width=True)

nav_options = ["Home", "Evaluation", "View", "Dashboard"]
nav = st.sidebar.radio("Navigation", nav_options, key="nav_radio")

# -------------------- HOME --------------------
if nav == "Home":
    summary = safe_read_excel(EXPORT_XLSX, "Summary")

    st.markdown("## ATA Tool")
    st.caption("DAMAC Properties | Quality Assurance | Telesales Division")

    k1, k2, k3, k4 = st.columns(4)
    if summary.empty:
        k1.metric("Total evaluations", "0")
        k2.metric("Average score", "0%")
        k3.metric("Avg failure rate", "0%")
        k4.metric("Last evaluation", "-")
        last_eval_line = "No evaluations yet."
    else:
        summary["Overall Score %"] = pd.to_numeric(summary.get("Overall Score %"), errors="coerce").fillna(0)
        summary["Passed"] = pd.to_numeric(summary.get("Passed"), errors="coerce").fillna(0)
        summary["Failed"] = pd.to_numeric(summary.get("Failed"), errors="coerce").fillna(0)
        summary["Total"] = summary["Passed"] + summary["Failed"]
        summary["Failure Rate"] = summary.apply(lambda r: (r["Failed"] / r["Total"]) if r["Total"] else 0, axis=1)

        last_dt = "-"
        if "Last Updated" in summary.columns:
            dt = pd.to_datetime(summary["Last Updated"], errors="coerce")
            if dt.notna().any():
                last_dt = dt.max().strftime("%Y-%m-%d %H:%M")

        k1.metric("Total evaluations", f"{len(summary)}")
        k2.metric("Average score", f"{summary['Overall Score %'].mean():.1f}%")
        k3.metric("Avg failure rate", f"{summary['Failure Rate'].mean()*100:.1f}%")

        # metric text doesn't wrap, so keep metric short and show wrapped text below
        k4.metric("Last evaluation", "Available")
        latest = summary.copy()
        if "Last Updated" in latest.columns:
            latest["_sort"] = pd.to_datetime(latest["Last Updated"], errors="coerce")
            latest = latest.sort_values("_sort", ascending=False).drop(columns=["_sort"], errors="ignore")
        row = latest.iloc[0].to_dict()
        last_eval_line = (
            f"Evaluation ID: {norm_id(row.get('Evaluation ID'))} | "
            f"QA: {norm_id(row.get('QA Name'))} | "
            f"Auditor: {norm_id(row.get('Auditor'))} | "
            f"Audit Date: {norm_id(row.get('Audit Date'))} | "
            f"Score: {row.get('Overall Score %', 0)}%"
        )

    st.markdown(
        f"""
<div class="ata-card">
  <div style="font-weight:900;color:#0b1f3a;">Last evaluation</div>
  <div class="ata-muted" style="margin-top:6px;">{last_eval_line}</div>
  <div style="margin-top:10px;">
    <span class="kpi-chip">Pass/Fail scoring</span>
    <span class="kpi-chip">PDF export</span>
    <span class="kpi-chip">Rendered Outlook draft (EML)</span>
    <span class="kpi-chip">Excel log</span>
    <span class="kpi-chip">Heatmap + trend</span>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.divider()

    a1, a2, a3, a4 = st.columns(4)
    with a1:
        if st.button("New Evaluation", use_container_width=True):
            st.session_state.goto_nav = "Evaluation"
            st.rerun()
    with a2:
        if st.button("View", use_container_width=True):
            st.session_state.goto_nav = "View"
            st.rerun()
    with a3:
        if st.button("Dashboard", use_container_width=True):
            st.session_state.goto_nav = "Dashboard"
            st.rerun()
    with a4:
        if os.path.exists(EXPORT_XLSX):
            with open(EXPORT_XLSX, "rb") as f:
                st.download_button(
                    "Download Excel Log",
                    f,
                    file_name=EXPORT_XLSX,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.button("Download Excel Log", disabled=True, use_container_width=True)

    st.divider()

    st.markdown("### Recent evaluations")
    if summary.empty:
        st.info("Submit an evaluation to populate the home overview.")
    else:
        view_df = summary.copy()
        if "Last Updated" in view_df.columns:
            view_df["_sort"] = pd.to_datetime(view_df["Last Updated"], errors="coerce")
            view_df = view_df.sort_values("_sort", ascending=False).drop(columns=["_sort"], errors="ignore")
        st.dataframe(view_df.head(10), width="stretch")

# -------------------- EVALUATION --------------------
elif nav == "Evaluation":
    st.markdown("## New Evaluation" if not st.session_state.edit_mode else "## Edit Evaluation")
    st.caption("Scoring model: Pass/Fail per parameter. Overall score equals pass rate.")

    pre = st.session_state.prefill or {}
    base_df = load_parameters_df()
    if isinstance(pre.get("details_df"), pd.DataFrame) and not pre["details_df"].empty:
        base_df = pre["details_df"][["Parameter", "Description", "Result", "Comment"]].copy()

    with st.form("eval_form", clear_on_submit=True):
        a, b, c = st.columns(3)
        with a:
            qa_name = st.text_input("QA name", value=pre.get("qa_name", ""))
            auditor = st.text_input("Auditor name", value=pre.get("auditor", ""))
        with b:
            evaluation_date = st.date_input("Evaluation date", value=pre.get("evaluation_date", date.today()))
            audit_date = st.date_input("Audit date", value=pre.get("audit_date", date.today()))
        with c:
            call_id = st.text_input("Call ID", value=pre.get("call_id", ""))
            call_duration = st.text_input("Call duration", value=pre.get("call_duration", ""), placeholder="00:10:32")
            call_disposition = st.text_input("Call disposition", value=pre.get("call_disposition", ""), placeholder="Resolved")

        st.markdown("### Parameters")
        edited = st.data_editor(
            base_df,
            width="stretch",
            num_rows="fixed",
            column_config={
                "Parameter": st.column_config.TextColumn(disabled=True, width="medium"),
                "Description": st.column_config.TextColumn(disabled=True, width="large"),
                "Result": st.column_config.SelectboxColumn(options=["Pass", "Fail"], required=True, width="small"),
                "Comment": st.column_config.TextColumn(width="large"),
            },
        )

        submitted = st.form_submit_button("Save evaluation", use_container_width=True)

    if submitted:
        eval_date_str = evaluation_date.strftime("%Y-%m-%d")

        # critical fix: in edit mode ALWAYS keep the same Evaluation ID (prevents duplication)
        if st.session_state.edit_mode:
            eval_id = norm_id(st.session_state.edit_eval_id)
            if not eval_id:
                eval_id = next_evaluation_id(eval_date_str)
        else:
            eval_id = next_evaluation_id(eval_date_str)

        score = overall_score(edited)
        passed = int((edited["Result"] == "Pass").sum())
        failed = int((edited["Result"] == "Fail").sum())

        record = {
            "evaluation_id": eval_id,
            "qa_name": (qa_name or "N/A").strip(),
            "auditor": (auditor or "N/A").strip(),
            "evaluation_date": eval_date_str,
            "audit_date": audit_date.strftime("%Y-%m-%d"),
            "call_id": (call_id or "N/A").strip(),
            "call_duration": (call_duration or "N/A").strip(),
            "call_disposition": (call_disposition or "N/A").strip(),
            "overall_score": score,
            "passed": passed,
            "failed": failed,
            "details": edited.copy(),
        }

        upsert_excel(record)

        # exit edit mode after save
        st.session_state.edit_mode = False
        st.session_state.edit_eval_id = ""
        st.session_state.prefill = {}
        st.session_state.show_email_preview = False

        st.success(f"Evaluation saved. Evaluation ID: {eval_id}")

        m1, m2, m3 = st.columns(3)
        m1.metric("Overall score", f"{score}%")
        m2.metric("Passed", passed)
        m3.metric("Failed", failed)

        st.divider()
        html = email_html_inline(record)

        colA, colB, colC = st.columns(3)
        with colA:
            st.download_button(
                "Download evaluation PDF",
                data=pdf_evaluation(record),
                file_name=f"ATA_Evaluation_{record['evaluation_id']}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        with colB:
            st.download_button(
                "Download rendered email (HTML)",
                data=html.encode("utf-8"),
                file_name=f"ATA_Email_{record['evaluation_id']}.html",
                mime="text/html",
                use_container_width=True,
            )
        with colC:
            eml_bytes = build_eml(record, html, attach_pdf=True)
            st.download_button(
                "Open Outlook draft (EML)",
                data=eml_bytes,
                file_name=f"ATA_Evaluation_{record['evaluation_id']}.eml",
                mime="message/rfc822",
                use_container_width=True,
            )

        st.markdown("### Email preview")
        st.markdown(html, unsafe_allow_html=True)
        st.info("To open the Outlook draft: download the .eml then double-click it. Outlook will render the HTML and attach the PDF automatically.")

        st.divider()
        if st.button("Go to View", use_container_width=True):
            st.session_state.goto_nav = "View"
            st.rerun()

# -------------------- VIEW --------------------
elif nav == "View":
    st.markdown("## View")
    st.caption("Filter by QA name and audit date, then open, edit, delete, or export.")

    summary = safe_read_excel(EXPORT_XLSX, "Summary")
    details = safe_read_excel(EXPORT_XLSX, "Details")

    if summary.empty or "Evaluation ID" not in summary.columns:
        st.info("No evaluations saved yet.")
    else:
        summary["Evaluation ID"] = summary["Evaluation ID"].apply(norm_id)
        summary["QA Name"] = summary.get("QA Name", "").apply(norm_id)
        summary["Auditor"] = summary.get("Auditor", "").apply(norm_id)
        summary["Call ID"] = summary.get("Call ID", "").apply(norm_id)

        summary["Audit Date"] = pd.to_datetime(summary.get("Audit Date"), errors="coerce").dt.date
        summary["Evaluation Date"] = pd.to_datetime(summary.get("Evaluation Date"), errors="coerce").dt.date

        qa_list = sorted([x for x in summary["QA Name"].dropna().unique().tolist() if x])
        audit_dates = sorted([x for x in summary["Audit Date"].dropna().unique().tolist()])

        f1, f2, f3 = st.columns([1, 1, 2])
        with f1:
            qa_filter = st.selectbox("QA name", ["All"] + qa_list)
        with f2:
            audit_filter = st.selectbox("Audit date", ["All"] + audit_dates)
        with f3:
            search = st.text_input("Search (Auditor / Call ID / Evaluation ID)")

        filtered = summary.copy()
        if qa_filter != "All":
            filtered = filtered[filtered["QA Name"] == qa_filter]
        if audit_filter != "All":
            filtered = filtered[filtered["Audit Date"] == audit_filter]
        if search.strip():
            s = search.strip().lower()
            filtered = filtered[
                filtered["Auditor"].astype(str).str.lower().str.contains(s)
                | filtered["Call ID"].astype(str).str.lower().str.contains(s)
                | filtered["Evaluation ID"].astype(str).str.lower().str.contains(s)
            ]

        if filtered.empty:
            st.warning("No matching evaluations found.")
        else:
            if "Last Updated" in filtered.columns:
                filtered["_sort"] = pd.to_datetime(filtered["Last Updated"], errors="coerce")
                filtered = filtered.sort_values("_sort", ascending=False).drop(columns=["_sort"], errors="ignore")

            filtered = filtered.copy()
            filtered["Label"] = filtered.apply(
                lambda r: f"{r['Audit Date']} | {r['QA Name']} | {r['Auditor']} | Call {r['Call ID']} | Score {r['Overall Score %']}% | {r['Evaluation ID']}",
                axis=1,
            )

            st.dataframe(filtered.drop(columns=["Label"], errors="ignore"), width="stretch")

            sel = st.selectbox("Select evaluation", filtered["Label"].tolist())
            row = filtered[filtered["Label"] == sel].iloc[0].to_dict()
            eval_id = norm_id(row.get("Evaluation ID"))

            det = pd.DataFrame()
            if not details.empty and "Evaluation ID" in details.columns:
                details["Evaluation ID"] = details["Evaluation ID"].apply(norm_id)
                det = details[details["Evaluation ID"] == eval_id].copy()
                if not det.empty:
                    det = det[["Parameter", "Description", "Result", "Comment"]].copy()

            record = {
                "evaluation_id": eval_id,
                "qa_name": norm_id(row.get("QA Name")) or "N/A",
                "auditor": norm_id(row.get("Auditor")) or "N/A",
                "evaluation_date": str(row.get("Evaluation Date", "")),
                "audit_date": str(row.get("Audit Date", "")),
                "call_id": norm_id(row.get("Call ID")) or "N/A",
                "call_duration": norm_id(row.get("Call Duration")) or "N/A",
                "call_disposition": norm_id(row.get("Call Disposition")) or "N/A",
                "overall_score": float(row.get("Overall Score %", 0) or 0),
                "passed": int(row.get("Passed", 0) or 0),
                "failed": int(row.get("Failed", 0) or 0),
                "details": det if not det.empty else load_parameters_df(),
            }

            c1, c2, c3 = st.columns(3)
            c1.metric("Overall score", f"{record['overall_score']}%")
            c2.metric("Passed", record["passed"])
            c3.metric("Failed", record["failed"])

            st.divider()
            html = email_html_inline(record)

            colA, colB, colC, colD, colE = st.columns(5)
            with colA:
                st.download_button(
                    "PDF",
                    data=pdf_evaluation(record),
                    file_name=f"ATA_Evaluation_{eval_id}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            with colB:
                st.download_button(
                    "HTML",
                    data=html.encode("utf-8"),
                    file_name=f"ATA_Email_{eval_id}.html",
                    mime="text/html",
                    use_container_width=True,
                )
            with colC:
                eml_bytes = build_eml(record, html, attach_pdf=True)
                st.download_button(
                    "Outlook draft (EML)",
                    data=eml_bytes,
                    file_name=f"ATA_Evaluation_{eval_id}.eml",
                    mime="message/rfc822",
                    use_container_width=True,
                )
            with colD:
                if st.button("Edit", use_container_width=True, key=f"edit_{eval_id}"):
                    def _to_date(v):
                        try:
                            return pd.to_datetime(v).date()
                        except Exception:
                            return date.today()

                    st.session_state.edit_mode = True
                    st.session_state.edit_eval_id = eval_id
                    st.session_state.prefill = {
                        "qa_name": record["qa_name"],
                        "auditor": record["auditor"],
                        "evaluation_date": _to_date(row.get("Evaluation Date")),
                        "audit_date": _to_date(row.get("Audit Date")),
                        "call_id": record["call_id"],
                        "call_duration": record["call_duration"],
                        "call_disposition": record["call_disposition"],
                        "details_df": record["details"],
                    }
                    st.session_state.goto_nav = "Evaluation"
                    st.rerun()

            with colE:
                if st.button("Delete", use_container_width=True, key=f"del_{eval_id}"):
                    st.session_state.pending_delete = eval_id
                    st.rerun()

            # delete confirmation (reliable, prevents accidental delete)
            if st.session_state.pending_delete == eval_id:
                st.warning(f"Confirm delete for Evaluation ID: {eval_id}")
                cc1, cc2, cc3 = st.columns([1, 1, 3])
                with cc1:
                    confirm = st.checkbox("Confirm", key=f"confirm_{eval_id}")
                with cc2:
                    if st.button("Delete now", key=f"delete_now_{eval_id}", use_container_width=True, disabled=not confirm):
                        ok = delete_evaluation(eval_id)
                        st.session_state.pending_delete = ""
                        if ok:
                            st.success("Deleted successfully.")
                        else:
                            st.error("Delete failed. Evaluation ID was not found in the Excel log.")
                        st.rerun()
                with cc3:
                    if st.button("Cancel", key=f"cancel_del_{eval_id}", use_container_width=False):
                        st.session_state.pending_delete = ""
                        st.rerun()

            st.divider()
            st.markdown("### Email preview")
            st.markdown(html, unsafe_allow_html=True)
            st.info("Outlook: download the .eml then double-click it. It opens a draft with rendered HTML and PDF attached.")

            st.markdown("### Parameters")
            st.dataframe(record["details"], width="stretch")

# -------------------- DASHBOARD --------------------
else:
    st.markdown("## Dashboard")
    st.caption("Heatmap and failure trend extracted from the Excel log.")

    fig_heat, fig_trend, summary, details = build_dashboard_figs()
    if fig_heat is None and fig_trend is None:
        st.info("No dashboard data yet. Submit evaluations first.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.pyplot(fig_heat)
        with c2:
            st.pyplot(fig_trend)

        st.divider()
        figs = [f for f in [fig_heat, fig_trend] if f is not None]

        d1, d2, d3 = st.columns(3)
        with d1:
            st.download_button(
                "Download Dashboard PDF",
                data=dashboard_pdf(figs, title="ATA Dashboard"),
                file_name="ATA_Dashboard.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        with d2:
            st.download_button(
                "Download Dashboard PPT",
                data=dashboard_ppt(figs, title="ATA Dashboard"),
                file_name="ATA_Dashboard.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        with d3:
            if os.path.exists(EXPORT_XLSX):
                with open(EXPORT_XLSX, "rb") as f:
                    st.download_button(
                        "Download Excel Log",
                        f,
                        file_name=EXPORT_XLSX,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
            else:
                st.button("Download Excel Log", disabled=True, use_container_width=True)
