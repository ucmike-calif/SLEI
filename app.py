import streamlit as st
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials

APP_VERSION = "SLEI-v2.0-pilot"

# -----------------------
# Assessment items
# -----------------------
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
# Helpers
# -----------------------

def safe_mean(values):
    nums = [v for v in values if isinstance(v, (int, float))]
    return (sum(nums) / len(nums)) if nums else None


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


def init_state():
    if "step" not in st.session_state:
        st.session_state.step = 1  # 1 = freq, 2 = change
    if "meta" not in st.session_state:
        st.session_state.meta = {}
    if "freq_sel" not in st.session_state:
        st.session_state.freq_sel = {}
    if "chg_sel" not in st.session_state:
        st.session_state.chg_sel = {}
    if "just_reset" not in st.session_state:
        st.session_state.just_reset = True


def clear_widget_keys():
    """Remove widget keys so nothing is pre-selected for the next respondent."""
    prefixes = ("freq_", "chg_")
    exact = {
        "role_anchor",
        "role_other",
        "profession",
        "profession_other",
        "years",
        "scope",
        "wants_testimonial",
        "testimonial_text",
        "testimonial_ok_public",
        "testimonial_attrib",
        "testimonial_name",
    }
    for k in list(st.session_state.keys()):
        if k in exact or k.startswith(prefixes):
            del st.session_state[k]


def reset_for_next():
    st.session_state.step = 1
    st.session_state.meta = {}
    st.session_state.freq_sel = {}
    st.session_state.chg_sel = {}
    st.session_state.just_reset = True


# -----------------------
# UI Setup
# -----------------------

st.set_page_config(page_title="SLEI v2.0", layout="wide")
init_state()

# Clear any prior widget state at the start of a fresh respondent
if st.session_state.step == 1 and st.session_state.just_reset:
    clear_widget_keys()
    st.session_state.just_reset = False

st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")

st.markdown(
    """
**Purpose**

This assessment is designed to support your continued development and help us improve the program.

**Structure**

You will answer 16 questions about key leadership behaviors in two sections:
1. **Current Frequency** – How often you intentionally demonstrate each behavior now, at the conclusion of the course.
2. **Change in Frequency** – How your current approach compares to how you typically operated before the course.

Because leadership looks different across contexts, select one role and use it consistently so your responses are accurate and comparable.
"""
)


# -----------------------
# Step 1 — Context + Current Frequency
# -----------------------

if st.session_state.step == 1:
    with st.form("slei_step1"):
        st.subheader("Step 1 of 2 — Context")

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

        role_other = ""
        if role_anchor == "Other":
            role_other = st.text_input("If Other, specify (required)", key="role_other").strip()

        profession = st.selectbox(
            "Profession type (required)",
            ["Student", "Resident", "Pharmacy Technician", "Pharmacist", "Other"],
            index=None,
            placeholder="Select one…",
            key="profession",
        )

        profession_other = ""
        if profession == "Other":
            profession_other = st.text_input(
                "If Other, specify (required)",
                key="profession_other",
            ).strip()

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

        st.subheader("Step 1 of 2 — Current Frequency")
        st.caption("No answers are pre-selected. Please choose one option per item.")

        freq_sel = {}
        for qid, text, _dom in ITEMS:
            freq_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                FREQ_OPTIONS,
                horizontal=True,
                index=None,
                key=f"freq_{qid}",
            )

        next_btn = st.form_submit_button("Next: Change questions")

    if next_btn:
        role_final = role_other if role_anchor == "Other" else role_anchor
        profession_final = profession_other if profession == "Other" else profession

        missing = []
        for name, val in [
            ("Role anchor", role_final),
            ("Profession type", profession_final),
            ("Years of experience", years),
            ("Leadership scope", scope),
        ]:
            if val is None or (isinstance(val, str) and val.strip() == ""):
                missing.append(name)

        unanswered = [qid for qid, _t, _d in ITEMS if freq_sel.get(qid) is None]

        if missing:
            st.error("Missing required fields: " + ", ".join(missing))
            st.stop()

        if unanswered:
            st.error(
                "Please answer all Current Frequency items (missing: "
                + ", ".join([str(q) for q in unanswered])
                + ")."
            )
            st.stop()

        st.session_state.meta = {
            "role_anchor": role_final,
            "profession": profession_final,
            "years": years,
            "scope": scope,
        }
        st.session_state.freq_sel = freq_sel
        st.session_state.step = 2
        st.rerun()


# -----------------------
# Step 2 — Change Compared to Before Course
# -----------------------

if st.session_state.step == 2:
    freq_sel = st.session_state.freq_sel

    # Only include items that were NOT marked N/A in Step 1
    applicable_qids = [
        qid for qid, _t, _d in ITEMS if FREQ_MAP.get(freq_sel.get(qid)) is not None
    ]

    with st.form("slei_step2"):
        st.subheader("Step 2 of 2 — Change in Frequency")
        st.caption("Only behaviors marked applicable in Step 1 are shown below.")

        chg_sel = {}
        for qid, text, _dom in ITEMS:
            if qid not in applicable_qids:
                continue
            chg_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                CHANGE_OPTIONS,
                horizontal=True,
                index=None,
                key=f"chg_{qid}",
            )

        submitted = st.form_submit_button("Submit")

    if submitted:
        unanswered = [qid for qid in applicable_qids if chg_sel.get(qid) is None]
        if unanswered:
            st.error(
                "Please answer all Change items shown (missing: "
                + ", ".join([str(q) for q in unanswered])
                + ")."
            )
            st.stop()

        # Convert to numeric
        freq_num = {qid: FREQ_MAP[freq_sel[qid]] for qid, _t, _d in ITEMS}
        chg_num = {qid: None for qid, _t, _d in ITEMS}
        for qid in applicable_qids:
            chg_num[qid] = CHANGE_MAP[chg_sel[qid]]

        overall = round1(safe_mean(freq_num.values()))
        overall_desc = overall_descriptor(overall)

        increased_count = sum(1 for qid in applicable_qids if (chg_num[qid] or 0) > 0)
        decreased_count = sum(1 for qid in applicable_qids if (chg_num[qid] or 0) < 0)

        st.success("Submitted.")
        st.write(f"**Overall score**: {overall} / 5 — {overall_desc}")
        st.write(
            f"**Growth summary**: Increased in **{increased_count}** behaviors (decreased in {decreased_count})."
        )

        # -----------------
        # Dynamic testimonial section (gated)
        # -----------------
        testimonial_text = ""
        testimonial_ok_public = ""
        testimonial_attrib = ""
        testimonial_name = ""

        if increased_count >= 6:
            st.divider()
            st.subheader("Optional: Share a short testimonial")
            st.caption(
                "We only ask this when your responses suggest meaningful growth. "
                "A testimonial is optional, and you can choose whether it’s anonymous."
            )

            wants_testimonial = st.radio(
                "Would you like to share a brief testimonial about the value of the program?",
                ["No", "Yes"],
                index=None,
                horizontal=True,
                key="wants_testimonial",
            )

            if wants_testimonial == "Yes":
                testimonial_text = st.text_area(
                    "Your testimonial (1–3 sentences is great)",
                    value="",
                    placeholder="Example: The course helped me...",
                    key="testimonial_text",
                ).strip()

                if testimonial_text:
                    testimonial_ok_public = st.radio(
                        "May we use this testimonial in marketing materials?",
                        ["No", "Yes"],
                        index=None,
                        horizontal=True,
                        key="testimonial_ok_public",
                    )

                    if testimonial_ok_public == "Yes":
                        st.caption(
                            "Choosing to attach your name makes your responses non-anonymous. "
                            "If you prefer, select Anonymous."
                        )
                        testimonial_attrib = st.radio(
                            "Attribution preference:",
                            ["Anonymous", "Use my initials", "Use my name"],
                            index=None,
                            horizontal=True,
                            key="testimonial_attrib",
                        )

                        if testimonial_attrib in ("Use my initials", "Use my name"):
                            label = (
                                "Enter your initials"
                                if testimonial_attrib == "Use my initials"
                                else "Enter your name"
                            )
                            testimonial_name = st.text_input(
                                label, value="", key="testimonial_name"
                            ).strip()

        # -----------------
        # Persist to Sheets
        # -----------------
        try:
            ws = open_sheet()
            meta = st.session_state.meta

            row = [
                pd.Timestamp.utcnow().isoformat(),
                APP_VERSION,
                meta.get("role_anchor", ""),
                meta.get("profession", ""),
                meta.get("years", ""),
                meta.get("scope", ""),
                str(overall) if overall is not None else "",
                overall_desc,
                str(increased_count),
                str(decreased_count),
                testimonial_text,
                testimonial_ok_public,
                testimonial_attrib,
                testimonial_name,
            ]

            # Always export 16 freq + 16 change columns in Q order
            row += [
                str(freq_num[qid]) if isinstance(freq_num[qid], (int, float)) else ""
                for qid in range(1, 17)
            ]
            row += [
                str(chg_num[qid]) if isinstance(chg_num[qid], (int, float)) else ""
                for qid in range(1, 17)
            ]

            append_row_to_sheet(ws, row)
            st.info("Saved to Google Sheets.")

            # Prepare clean slate for next respondent
            reset_for_next()

        except Exception as e:
            st.error("Could not save to Google Sheets.")
            st.exception(e)
