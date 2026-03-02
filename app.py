import streamlit as st
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials

# -----------------------
# App Config
# -----------------------
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
# Helpers
# -----------------------

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


def require_text_if_other(selected_value: str | None, other_text: str, other_label: str):
    if selected_value == "Other" and not other_text.strip():
        return f"{other_label} (Other)"
    return None


# -----------------------
# UI Setup
# -----------------------

st.set_page_config(page_title="SLEI v2.0", layout="wide")

st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")

st.markdown(
    """
**Purpose**

This assessment is designed to support your continued development and help us improve the program.

**Structure**

You will answer 16 questions about key leadership behaviors in two sections:
1. The Current Frequency section asks you to indicate how often you perform these behaviors now (after completing the course)
2. The Change in Frequency section asks how your current frequency compares to the frequency before the course

Because leadership looks different across contexts, select one role and use it consistently so your responses are accurate and comparable.
"""
)

# Session state
if "step" not in st.session_state:
    st.session_state.step = 1

if "context" not in st.session_state:
    st.session_state.context = {}


# -----------------------
# Step 1 — Context (NOT in a form, so conditional fields appear immediately)
# -----------------------

if st.session_state.step == 1:
    st.header("Step 1 of 2 — Context")

    role_anchor = st.selectbox(
        "Which single leadership role will you use as your reference point? (required)",
        [
            "My primary professional/employer role",
            "A volunteer or board leadership role",
            "A family or community leadership role",
            "Other",
        ],
        index=None,
        placeholder="Select one…",
        key="role_anchor",
    )

    role_anchor_other = ""
    if role_anchor == "Other":
        role_anchor_other = st.text_input("If Other, specify (required)", key="role_anchor_other").strip()

    profession = st.selectbox(
        "Profession type (required)",
        ["Student", "Resident", "Pharmacy Technician", "Pharmacist", "Other"],
        index=None,
        placeholder="Select one…",
        key="profession",
    )

    profession_other = ""
    if profession == "Other":
        profession_other = st.text_input("If Other, specify (required)", key="profession_other").strip()

    years = st.selectbox(
        "Years of experience (required)",
        ["0–2", "3–5", "6–10", "11–15", "16–20", "21+"],
        index=None,
        placeholder="Select one…",
        key="years",
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
        key="scope",
    )

    scope_other = ""
    if scope == "Other":
        scope_other = st.text_input("If Other, specify (required)", key="scope_other").strip()

    st.caption("You can proceed once all required fields are complete.")

    col1, col2 = st.columns([1, 4])
    with col1:
        next_clicked = st.button("Next →")

    if next_clicked:
        missing = []

        # Required base fields
        if role_anchor is None:
            missing.append("Leadership role")
        if profession is None:
            missing.append("Profession type")
        if years is None:
            missing.append("Years of experience")
        if scope is None:
            missing.append("Leadership scope")

        # Required other-text fields
        other_missing = [
            require_text_if_other(role_anchor, role_anchor_other, "Leadership role"),
            require_text_if_other(profession, profession_other, "Profession type"),
            require_text_if_other(scope, scope_other, "Leadership scope"),
        ]
        missing.extend([m for m in other_missing if m])

        if missing:
            st.error("Missing required fields: " + ", ".join(missing))
        else:
            # Normalize stored values
            role_value = role_anchor_other if role_anchor == "Other" else role_anchor
            prof_value = profession_other if profession == "Other" else profession
            scope_value = scope_other if scope == "Other" else scope

            st.session_state.context = {
                "role_anchor": role_value,
                "profession": prof_value,
                "years": years,
                "scope": scope_value,
            }
            st.session_state.step = 2
            st.rerun()


# -----------------------
# Step 2 — Assessment + Optional Feedback
# -----------------------

if st.session_state.step == 2:
    st.header("Step 2 of 2 — Assessment")

    with st.form("slei_form_step2"):
        st.subheader("Current frequency")
        st.caption("Leave items blank if you’re unsure. Use ‘Not applicable’ only when the behavior truly does not apply to your selected role.")

        freq_sel = {}
        for qid, text, _dom in ITEMS:
            freq_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                FREQ_OPTIONS,
                index=None,
                horizontal=True,
                key=f"freq_{qid}",
            )

        applicable_qids = [qid for qid, _t, _d in ITEMS if freq_sel.get(qid) not in (None, "Not applicable to my role")]
        na_qids = [qid for qid, _t, _d in ITEMS if freq_sel.get(qid) == "Not applicable to my role"]

        st.subheader("Change in frequency")
        st.caption("You’ll only see change questions for items you did not mark as ‘Not applicable.’")

        chg_sel = {}
        for qid, text, _dom in ITEMS:
            if qid in applicable_qids:
                chg_sel[qid] = st.radio(
                    f"Q{qid}. {text}",
                    CHANGE_OPTIONS,
                    index=None,
                    horizontal=True,
                    key=f"chg_{qid}",
                )
            else:
                chg_sel[qid] = None

        st.subheader("Optional feedback")
        st.caption("This section is optional, but extremely helpful.")

        testimonial_ok = st.checkbox(
            "I’m open to being contacted about using my feedback/testimonial (optional)",
            value=False,
            key="testimonial_ok",
        )

        testimonial_text = st.text_area(
            "If you’d like, share a short testimonial or comment about the program (optional)",
            value="",
            height=120,
            key="testimonial_text",
        )

        contact_name = ""
        contact_email = ""
        if testimonial_ok:
            c1, c2 = st.columns(2)
            with c1:
                contact_name = st.text_input("Name (optional)", value="", key="contact_name").strip()
            with c2:
                contact_email = st.text_input("Email (optional)", value="", key="contact_email").strip()

        submitted = st.form_submit_button("Submit")

    if submitted:
        # Validate Step 2 required inputs: require frequency for all 16, and change for all applicable
        missing = []

        # Frequency must be answered for all 16 (including N/A when appropriate)
        for qid, _text, _dom in ITEMS:
            if freq_sel.get(qid) is None:
                missing.append(f"Q{qid} (Current frequency)")

        # Change must be answered for applicable items
        for qid in applicable_qids:
            if chg_sel.get(qid) is None:
                missing.append(f"Q{qid} (Change)")

        if missing:
            st.error("Please complete: " + "; ".join(missing))
            st.stop()

        # Compute scores
        freq_num = {qid: FREQ_MAP[freq_sel[qid]] for qid, _, _ in ITEMS}
        chg_num = {qid: (CHANGE_MAP[chg_sel[qid]] if chg_sel[qid] is not None else None) for qid, _, _ in ITEMS}

        freq_vals = [v for v in freq_num.values() if isinstance(v, (int, float))]
        overall = round1(safe_mean(freq_vals))
        overall_desc = overall_descriptor(overall)

        st.success("Submitted.")
        st.write(f"**Overall score**: {overall} / 5 — {overall_desc}")

        # Save to Google Sheets
        try:
            ws = open_sheet()

            ctx = st.session_state.context
            row = [
                pd.Timestamp.utcnow().isoformat(),
                APP_VERSION,
                ctx.get("role_anchor", ""),
                ctx.get("profession", ""),
                ctx.get("years", ""),
                ctx.get("scope", ""),
                str(overall) if overall is not None else "",
                overall_desc,
            ]

            # Fixed 16 freq columns
            row += [
                str(freq_num[qid]) if isinstance(freq_num[qid], (int, float)) else ""
                for qid in range(1, 17)
            ]

            # Fixed 16 change columns (blank for N/A)
            row += [
                str(chg_num[qid]) if isinstance(chg_num[qid], (int, float)) else ""
                for qid in range(1, 17)
            ]

            # Optional feedback fields (fixed columns)
            row += [
                "YES" if testimonial_ok else "",
                testimonial_text.strip(),
                contact_name,
                contact_email,
            ]

            append_row_to_sheet(ws, row)
            st.info("Saved to Google Sheets.")
        except Exception as e:
            st.error("Could not save to Google Sheets (secrets / sheet setup may be missing or incomplete).")
            st.exception(e)

        # Reset for next respondent
        st.session_state.step = 1
        st.session_state.context = {}
        # Also clear widget state keys for radios/selects to avoid any persistence
        for k in list(st.session_state.keys()):
            if k.startswith("freq_") or k.startswith("chg_") or k in {
                "role_anchor",
                "role_anchor_other",
                "profession",
                "profession_other",
                "years",
                "scope",
                "scope_other",
                "testimonial_ok",
                "testimonial_text",
                "contact_name",
                "contact_email",
            }:
                del st.session_state[k]

        st.rerun()
