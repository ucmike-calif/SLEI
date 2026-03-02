import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

APP_VERSION = "SLEI-v2.0-pilot"

ITEMS = [
    (1, "Pause long enough to respond thoughtfully in high-pressure situations"),
    (2, "Take deliberate actions to influence long-term outcomes"),
    (3, "Test your assumptions before responding to a situation"),
    (4, "Recognize predictable patterns in your reactions and adjust your response accordingly"),
    (5, "Align major commitments with what you identify as most important in your role"),
    (6, "Adjust workload or boundaries to sustain performance"),
    (7, "Remove or delegate commitments that do not require your direct involvement"),
    (8, "Intentionally focus your time on the highest-level responsibilities your role is designed to perform"),
    (9, "Take timely, considered action on important issues even when outcomes are uncertain"),
    (10, "When challenges arise, focus first on actions within your control"),
    (11, "Renegotiate commitments before they become risks"),
    (12, "Intentionally develop others’ capability so they can operate with greater ownership and independence"),
    (13, "Before launching work, ensure clarity about outcomes, ownership, resources, and how the team will operate"),
    (14, "Hold check-ins to ensure expectations remain clear"),
    (15, "Delegate work with clear expectations rather than retaining it"),
    (16, "Effectively influence stakeholders beyond your formal authority to move important work forward"),
]

FREQ_OPTIONS = [
    "Rarely",
    "Occasionally",
    "Sometimes",
    "Often",
    "Consistently",
    "Not applicable to my role",
]

CHANGE_OPTIONS = [
    "Much less often",
    "Slightly less often",
    "About the same",
    "Slightly more often",
    "Much more often",
]

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


# -----------------------
# Utility Functions
# -----------------------

def safe_mean(values):
    nums = [v for v in values if isinstance(v, (int, float))]
    return sum(nums) / len(nums) if nums else None


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


# -----------------------
# UI Setup
# -----------------------

st.set_page_config(page_title="SLEI v2.0", layout="wide")

st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")

st.caption("Not a grade — a development tool.")

st.markdown("""
**Purpose**

This assessment supports your reflection and helps us evaluate and improve the program.

**Structure**

You will answer each leadership behavior in two sections:
1. Current frequency (how often you do it now)
2. Change (how that frequency compares to before the course)

Select one leadership role and use it consistently throughout.
""")


# -----------------------
# Form
# -----------------------

with st.form("slei_form"):

    role_anchor = st.selectbox(
        "Which leadership role are you using as your reference point?",
        [
            "Primary professional role",
            "Volunteer or board role",
            "Family/community role",
            "Other",
        ],
        index=None,
        placeholder="Select one…",
    )

    st.subheader("Current frequency")

    freq_responses = {}
    for qid, text in ITEMS:
        freq_responses[qid] = st.radio(
            f"Q{qid}. {text}",
            FREQ_OPTIONS,
            index=None,
            horizontal=True,
        )

    st.subheader("Change compared to before the course")

    change_responses = {}
    for qid, text in ITEMS:
        if freq_responses.get(qid) == "Not applicable to my role":
            continue
        change_responses[qid] = st.radio(
            f"Q{qid}. {text}",
            CHANGE_OPTIONS,
            index=None,
            horizontal=True,
        )

    submitted = st.form_submit_button("Submit")


if submitted:

    if role_anchor is None:
        st.error("Please select a leadership role.")
        st.stop()

    missing_freq = [qid for qid in freq_responses if freq_responses[qid] is None]
    if missing_freq:
        st.error("Please answer all Current Frequency items.")
        st.stop()

    missing_change = [qid for qid in change_responses if change_responses[qid] is None]
    if missing_change:
        st.error("Please answer all Change items shown.")
        st.stop()

    freq_numeric = {qid: FREQ_MAP[freq_responses[qid]] for qid in freq_responses}
    change_numeric = {
        qid: CHANGE_MAP[change_responses[qid]]
        if qid in change_responses
        else None
        for qid in range(1, 17)
    }

    overall = safe_mean(freq_numeric.values())
    descriptor = overall_descriptor(overall)

    st.success("Submitted successfully.")
    st.write(f"Overall score: {round(overall,1) if overall else 'N/A'} — {descriptor}")

    try:
        ws = open_sheet()
        row = [
            pd.Timestamp.utcnow().isoformat(),
            APP_VERSION,
            role_anchor,
            str(round(overall,1) if overall else ""),
            descriptor,
        ]

        row += [
            str(freq_numeric[qid]) if isinstance(freq_numeric[qid], (int, float)) else ""
            for qid in range(1, 17)
        ]

        row += [
            str(change_numeric[qid]) if isinstance(change_numeric[qid], (int, float)) else ""
            for qid in range(1, 17)
        ]

        ws.append_row(row)
        st.info("Saved to Google Sheets.")

    except Exception as e:
        st.error("Error saving to Google Sheets.")
        st.exception(e)
