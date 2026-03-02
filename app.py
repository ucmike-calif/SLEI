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


# -----------------
# Helpers
# -----------------

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


def init_state():
    if "step" not in st.session_state:
        st.session_state.step = 1  # 1=freq, 2=change
    if "freq_sel" not in st.session_state:
        st.session_state.freq_sel = {}
    if "chg_sel" not in st.session_state:
        st.session_state.chg_sel = {}
    if "meta" not in st.session_state:
        st.session_state.meta = {}
    # Used to clear widget keys so nothing is pre-selected on a fresh respondent
    if "just_reset" not in st.session_state:
        st.session_state.just_reset = True


def clear_widget_keys():
    """Remove widget keys so Streamlit doesn't carry forward prior selections."""
    prefixes = ("freq_", "chg_")
    exact = {
        "wants_testimonial",
        "testimonial_text",
        "testimonial_ok_public",
        "testimonial_attrib",
        "testimonial_name",
    }
    for k in list(st.session_state.keys()):
        if k in exact or k.startswith(prefixes):
            del st.session_state[k]


# -----------------
# UI
# -----------------

st.set_page_config(page_title="SLEI v2.0", layout="wide")
init_state()

st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")
st.caption("Not a grade — a development tool. Your responses support reflection and program improvement.")

st.markdown("""
**What this is for**

This brief post-course self-assessment helps you reflect on your growth and helps us improve the program over time. Aggregate (group-level) results may also be shared with sponsoring organizations to demonstrate impact.

**How to complete it**

You’ll answer the same set of leadership behaviors in two sections:
- Current frequency: how often you do each behavior now
- Change: how your current frequency compares to before the course (only for behaviors that apply to your chosen role)

Choose one leadership role as your reference point for the entire assessment. If a behavior isn’t relevant to that role, select Not applicable.
""")
 results may also be shared with sponsoring organizations to demonstrate impact.

"
    "**How to complete it**

"
    "You’ll answer the same set of leadership behaviors in **two sections**:
"
    "- **Current frequency**: how often you do each behavior *now*
"
    "- **Change**: how your current frequency compares to *before* the course (only for behaviors that apply to your chosen role)

"
    "Choose **one** leadership role as your reference point for the entire assessment. If a behavior isn’t relevant to that role, select **Not applicable**."
)

# -----------------
# Step 1: Current Frequency
# -----------------

if st.session_state.step == 1:
    # Clear any prior widget state so radios/selectors do not pre-populate.
    if st.session_state.just_reset:
        clear_widget_keys()
        st.session_state.just_reset = False
    with st.form("slei_step1"):
        st.subheader("Step 1 of 2 — Context")

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
        role_other = ""
        if role_anchor == "Other":
            role_other = st.text_input("If Other, specify (required)").strip()

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

        st.subheader("Step 1 of 2 — Current frequency")
        st.caption("No answers are pre-selected. Please choose one option per item.")

        freq_sel = {}
        for qid, text, dom in ITEMS:
            freq_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                FREQ_OPTIONS,
                horizontal=True,
                index=None,
                key=f"freq_{qid}",
            )

        next_btn = st.form_submit_button("Next: Change questions")

    if next_btn:
        # Validate required metadata
        role_final = role_other if role_anchor == "Other" else role_anchor
        missing = []
        for name, val in [
            ("Role anchor", role_final),
            ("Profession type", profession),
            ("Years of experience", years),
            ("Leadership scope", scope),
        ]:
            if val is None or (isinstance(val, str) and val.strip() == ""):
                missing.append(name)

        # Validate that every freq item is answered
        unanswered = [qid for qid, _, _ in ITEMS if freq_sel.get(qid) is None]
        if unanswered:
            st.error(f"Please answer all Current Frequency items (missing: {', '.join([str(q) for q in unanswered])}).")
            st.stop()

        if missing:
            st.error("Missing required fields: " + ", ".join(missing))
            st.stop()

        st.session_state.meta = {
            "role_anchor": role_final,
            "profession": profession,
            "years": years,
            "scope": scope,
        }
        st.session_state.freq_sel = freq_sel
        st.session_state.step = 2
        st.rerun()


# -----------------
# Step 2: Change Compared to Before Course
# -----------------

if st.session_state.step == 2:
    freq_sel = st.session_state.freq_sel

    applicable_qids = [
        qid for qid, _, _ in ITEMS if FREQ_MAP.get(freq_sel.get(qid)) is not None
    ]
    na_qids = [qid for qid, _, _ in ITEMS if qid not in applicable_qids]

    with st.form("slei_step2"):
        st.subheader("Step 2 of 2 — Change compared to before the course")
        st.caption("Only behaviors marked applicable in Step 1 are shown below.")

        chg_sel = {}
        for qid, text, dom in ITEMS:
            if qid not in applicable_qids:
                continue
            chg_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                CHANGE_OPTIONS,
                horizontal=True,
                index=None,
                key=f"chg_{qid}",
            )

        submit_btn = st.form_submit_button("Submit")

    if submit_btn:
        # Validate all shown change items answered
        unanswered = [qid for qid in applicable_qids if chg_sel.get(qid) is None]
        if unanswered:
            st.error(f"Please answer all Change items shown (missing: {', '.join([str(q) for q in unanswered])}).")
            st.stop()

        # Convert values for scoring
        freq_num = {qid: FREQ_MAP[freq_sel[qid]] for qid, _, _ in ITEMS}
        # Ensure stable export: create chg_num for ALL items; N/A items get None
        chg_num = {qid: None for qid, _, _ in ITEMS}
        for qid in applicable_qids:
            chg_num[qid] = CHANGE_MAP[chg_sel[qid]]

        # Scoring
        freq_vals = [v for v in freq_num.values() if isinstance(v, (int, float))]
        overall = round1(safe_mean(freq_vals))
        overall_desc = overall_descriptor(overall)

        # Growth summary
        increased_count = sum(1 for qid in applicable_qids if (chg_num[qid] or 0) > 0)
        decreased_count = sum(1 for qid in applicable_qids if (chg_num[qid] or 0) < 0)

        st.success("Submitted.")
        st.write(f"**Overall score**: {overall} / 5 — {overall_desc}")
        st.write(f"**Growth summary**: Increased in **{increased_count}** behaviors (decreased in {decreased_count}).")

        # -----------------
        # Dynamic testimonial section (gated by growth)
        # -----------------
        testimonial_text = ""
        testimonial_ok_public = ""
        testimonial_attrib = ""
        testimonial_name = ""

        # Gate: at least 6 increases (you can tune this later)
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
                            label = "Enter your initials" if testimonial_attrib == "Use my initials" else "Enter your name"
                            testimonial_name = st.text_input(label, value="", key="testimonial_name").strip()

        # -----------------
        # Persist to Sheets
        # -----------------
        try:
            ws = open_sheet()

            meta = st.session_state.meta
            role_final = meta.get("role_anchor", "")
            profession = meta.get("profession", "")
            years = meta.get("years", "")
            scope = meta.get("scope", "")

            row = [
                pd.Timestamp.utcnow().isoformat(),
                APP_VERSION,
                role_final,
                profession,
                years,
                scope,
                str(overall) if overall is not None else "",
                overall_desc,
                str(increased_count),
                str(decreased_count),
                testimonial_text,
                testimonial_ok_public,
                testimonial_attrib,
                testimonial_name,
            ]

            # Keep stable schema for analytics: always write 16 freq columns then 16 change columns
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

            # Reset for next respondent
            st.session_state.step = 1
            st.session_state.freq_sel = {}
            st.session_state.chg_sel = {}
            st.session_state.meta = {}
            st.session_state.just_reset = True

        except Exception as e:
            st.error("Could not save to Google Sheets.")
            st.exception(e)
