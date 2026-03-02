import streamlit as st
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials

APP_VERSION = "SLEI-v2.0-pilot"

ITEMS = [
    (1, "Pause long enough to respond thoughtfully in high-pressure situations", "Sight"),
    (2, "Take deliberate actions to influence long-term outcomes", "Sight"),
    (3, "Test your assumptions before responding to a situation", "Sight"),
    (4, "Recognize predictable patterns in your reactions and adjust your response accordingly", "Sight"),
    (5, "Align major commitments with what you identify as most important in your role", "Tenacity"),
    (6, "Adjust workload or boundaries to sustain performance", "Tenacity"),
    (7, "Remove or delegate commitments that do not require your direct involvement", "Tenacity"),
    (8, "Intentionally focus your time on the highest-level responsibilities your role is designed to perform", "Tenacity"),
    (9, "Take timely, considered action on important issues even when outcomes are uncertain", "Ability"),
    (10, "When challenges arise, focus first on actions within your control", "Ability"),
    (11, "Renegotiate commitments before they become risks", "Ability"),
    (12, "Intentionally develop others’ capability so they can operate with greater ownership and independence", "Ability"),
    (13, "Before launching work, ensure clarity about outcomes, ownership, resources, and how the team will operate", "Results"),
    (14, "Hold check-ins to ensure expectations remain clear", "Results"),
    (15, "Delegate work with clear expectations rather than retaining it", "Results"),
    (16, "Effectively influence stakeholders beyond your formal authority to move important work forward", "Results"),
]

DOMAINS = {
    "Sight": [1, 2, 3, 4],
    "Tenacity": [5, 6, 7, 8],
    "Ability": [9, 10, 11, 12],
    "Results": [13, 14, 15, 16],
}

FREQ_OPTIONS = ["Rarely", "Occasionally", "Sometimes", "Often", "Consistently", "Not applicable to my role"]
CHANGE_OPTIONS = ["Much less often", "Slightly less often", "About the same", "Slightly more often", "Much more often"]

FREQ_MAP = {
    "Rarely": 1,
    "Occasionally": 2,
    "Sometimes": 3,
    "Often": 4,
    "Consistently": 5,
    "Not applicable to my role": None,
}
CHANGE_MAP = {
    "Much less often": -2,
    "Slightly less often": -1,
    "About the same": 0,
    "Slightly more often": 1,
    "Much more often": 2,
}


def safe_mean(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return sum(vals) / len(vals) if vals else None


def round1(x):
    return None if x is None else round(x, 1)


def overall_descriptor(score):
    if score is None:
        return "Not scored"
    if score >= 4.5:
        return "Consistently / Automatic"
    if score >= 4.0:
        return "Often"
    if score >= 3.0:
        return "Sometimes"
    if score >= 2.0:
        return "Inconsistent"
    return "Rarely"


def open_sheet():
    creds_info = st.secrets["gcp_service_account"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    return gc.open(st.secrets["sheet_name"]).worksheet(st.secrets["worksheet_name"])


def append_row_to_sheet(ws, row):
    ws.append_row(row, value_input_option="RAW")


# --- UI ---
st.set_page_config(page_title="SLEI v2.0", layout="wide")

st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")
st.caption("This dashboard will be non-anonymous if you include identifying information (e.g., name/email).")

st.markdown(
    "**Instructions**\n\n"
    "Respond based on how you typically operate now, at the conclusion of this program.\n\n"
    "For each behavior:\n"
    "1) Indicate how frequently you currently perform it.\n"
    "2) Indicate how the frequency has changed compared to before this course.\n"
)

with st.form("slei_form"):
    role_anchor = st.selectbox(
        "Which single leadership role will you use as your reference point for this assessment? (required)",
        [
            "My primary professional/employer role",
            "A volunteer or board leadership role",
            "A family or community leadership role",
            "Other",
        ],
        index=None,
        placeholder="Select one…",
    )
    if role_anchor == "Other":
        role_anchor = st.text_input("If Other, specify (required)").strip()

    profession = st.selectbox(
        "Profession type (required)",
        ["Student", "Resident", "Pharmacy Technician", "Pharmacist", "Other"],
        index=None,
        placeholder="Select one…",
    )

    years = st.selectbox(
        "Years of experience (required)",
        ["0–2", "3–5", "6–10", "11–15", "16–20", "21+"],
        index=None,
        placeholder="Select one…",
    )

    scope = st.selectbox(
        "Leadership scope (required)",
        [
            "Individual contributor (no direct reports)",
            "Supervisor / Manager of individuals",
            "Manager of managers",
            "Senior leader / executive",
            "Other",
        ],
        index=None,
        placeholder="Select one…",
    )

    st.subheader("Current frequency")
    freq_sel = {}
    for qid, text, dom in ITEMS:
        freq_sel[qid] = st.radio(
            f"Q{qid}. {text}",
            FREQ_OPTIONS,
            horizontal=True,
            key=f"freq_{qid}",
        )

    st.subheader("Change compared to before the course")
    st.caption("If you selected “Not applicable” above for an item, choose “About the same” below.")
    chg_sel = {}
    for qid, text, dom in ITEMS:
        chg_sel[qid] = st.radio(
            f"Q{qid}. {text}",
            CHANGE_OPTIONS,
            horizontal=True,
            key=f"chg_{qid}",
        )

    submitted = st.form_submit_button("Submit")

if submitted:
    missing = []
    for name, val in [
        ("Role anchor", role_anchor),
        ("Profession type", profession),
        ("Years of experience", years),
        ("Leadership scope", scope),
    ]:
        if val is None or (isinstance(val, str) and val.strip() == ""):
            missing.append(name)

    if missing:
        st.error("Missing required fields: " + ", ".join(missing))
        st.stop()

    freq_num = {qid: FREQ_MAP[freq_sel[qid]] for qid, _, _ in ITEMS}
    chg_num = {qid: CHANGE_MAP[chg_sel[qid]] for qid, _, _ in ITEMS}

    freq_vals = [v for v in freq_num.values() if isinstance(v, (int, float))]
    overall = round1(safe_mean(freq_vals))
    overall_desc = overall_descriptor(overall)

    st.success("Submitted.")
    st.write(f"**Overall score**: {overall} / 5 — {overall_desc}")

    try:
        ws = open_sheet()
        row = [
            pd.Timestamp.utcnow().isoformat(),
            APP_VERSION,
            role_anchor,
            profession,
            years,
            scope,
            str(overall),
            overall_desc,
        ] + [
            str(freq_num[qid]) if isinstance(freq_num[qid], (int, float)) else "" for qid in range(1, 17)
        ] + [str(chg_num[qid]) for qid in range(1, 17)]
        append_row_to_sheet(ws, row)
        st.info("Saved to Google Sheets.")
    except Exception as e:
        st.error(
            "Could not save to Google Sheets (secrets / sheet setup may be missing or incomplete)."
        )
        st.exception(e)
