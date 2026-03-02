import streamlit as st
import pandas as pd

import gspread
from google.oauth2.service_account import Credentials

# -----------------------
# Config / Constants
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


def init_state():
    st.session_state.setdefault("step", 1)
    st.session_state.setdefault("role_anchor", None)
    st.session_state.setdefault("role_anchor_other", "")
    st.session_state.setdefault("profession", None)
    st.session_state.setdefault("profession_other", "")
    st.session_state.setdefault("years", None)
    st.session_state.setdefault("scope", None)
    st.session_state.setdefault("scope_other", "")

    # Per-item answers
    st.session_state.setdefault("freq_sel", {qid: None for qid, _, _ in ITEMS})
    st.session_state.setdefault("chg_sel", {qid: None for qid, _, _ in ITEMS})

    # Feedback / contact
    st.session_state.setdefault("improve_feedback", "")
    st.session_state.setdefault("willing_contact", False)
    st.session_state.setdefault("testimonial", "")
    st.session_state.setdefault("contact_name", "")
    st.session_state.setdefault("contact_email", "")


def go_next():
    st.session_state.step += 1


def go_prev():
    st.session_state.step -= 1


def compute_scores():
    freq_num = {qid: FREQ_MAP.get(st.session_state.freq_sel.get(qid)) for qid, _, _ in ITEMS}
    chg_num = {qid: CHANGE_MAP.get(st.session_state.chg_sel.get(qid)) for qid, _, _ in ITEMS}

    freq_vals = [v for v in freq_num.values() if isinstance(v, (int, float))]
    overall = round1(safe_mean(freq_vals))
    overall_desc = overall_descriptor(overall)

    # Growth flag: positive average change across non-NA items
    chg_vals = [
        chg_num[qid]
        for qid, _, _ in ITEMS
        if freq_num.get(qid) is not None and isinstance(chg_num.get(qid), (int, float))
    ]
    avg_change = safe_mean(chg_vals)
    growth = (avg_change is not None) and (avg_change > 0)

    return freq_num, chg_num, overall, overall_desc, avg_change, growth


def required_missing_step1():
    missing = []

    ra = st.session_state.role_anchor
    if not ra:
        missing.append("Leadership role")
    elif ra == "Other" and not st.session_state.role_anchor_other.strip():
        missing.append("Leadership role (Other)")

    prof = st.session_state.profession
    if not prof:
        missing.append("Profession type")
    elif prof == "Other" and not st.session_state.profession_other.strip():
        missing.append("Profession type (Other)")

    if not st.session_state.years:
        missing.append("Years of experience")

    sc = st.session_state.scope
    if not sc:
        missing.append("Leadership scope")
    elif sc == "Other" and not st.session_state.scope_other.strip():
        missing.append("Leadership scope (Other)")

    return missing


def required_missing_freq():
    return [qid for qid, _, _ in ITEMS if not st.session_state.freq_sel.get(qid)]


def required_missing_change(non_na_ids):
    return [qid for qid in non_na_ids if not st.session_state.chg_sel.get(qid)]


# -----------------------
# UI Setup
# -----------------------
init_state()

st.set_page_config(page_title="SLEI v2.0", layout="wide")
st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")

st.markdown(
    """**Purpose**

This assessment is designed to support your continued development and help us improve the program.

**Structure**

You will answer 16 questions about key leadership behaviors in two sections:
1. The **Current Frequency** section asks you to indicate how often you perform these behaviors now (after completing the course).
2. The **Change in Frequency** section asks how your current frequency compares to the frequency before the course.

Because leadership looks different across contexts, select one role and use it consistently so your responses are accurate and comparable.
"""
)

# -----------------------
# Step 1 of 5 — Context
# -----------------------
if st.session_state.step == 1:
    st.header("Step 1 of 5 — Context")

    st.session_state.role_anchor = st.selectbox(
        "Which single leadership role will you use as your reference point? (required)",
        [
            "My primary professional/employer role",
            "A volunteer or board leadership role",
            "A family or community leadership role",
            "Other",
        ],
        index=None,
        placeholder="Select one…",
    )
    if st.session_state.role_anchor == "Other":
        st.session_state.role_anchor_other = st.text_input(
            "If Other, please specify (required)",
            value=st.session_state.role_anchor_other,
        )

    st.session_state.profession = st.selectbox(
        "Profession type (required)",
        ["Student", "Resident", "Pharmacy Technician", "Pharmacist", "Other"],
        index=None,
        placeholder="Select one…",
    )
    if st.session_state.profession == "Other":
        st.session_state.profession_other = st.text_input(
            "If Other, please specify (required)",
            value=st.session_state.profession_other,
        )

    st.session_state.years = st.selectbox(
        "Years of experience (required)",
        ["0–2", "3–5", "6–10", "11–15", "16–20", "21+"],
        index=None,
        placeholder="Select one…",
    )

    st.session_state.scope = st.selectbox(
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
    if st.session_state.scope == "Other":
        st.session_state.scope_other = st.text_input(
            "If Other, please specify (required)",
            value=st.session_state.scope_other,
        )

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("Next →", type="primary", on_click=go_next)

    missing = required_missing_step1()
    if missing:
        st.info("To continue, complete: " + ", ".join(missing))

    # Guard: keep step at 1 if they clicked next without required fields
    if st.session_state.step == 2 and missing:
        st.session_state.step = 1


# -----------------------
# Step 2 of 5 — Current frequency
# -----------------------
elif st.session_state.step == 2:
    st.header("Step 2 of 5 — Current frequency")

    st.caption(
        "For each item, select how often you perform it now. Choose ‘Not applicable’ if it does not apply to the role you selected."
    )

    for qid, text, _ in ITEMS:
        st.session_state.freq_sel[qid] = st.radio(
            f"Q{qid}. {text}",
            FREQ_OPTIONS,
            index=None,
            horizontal=True,
            key=f"freq_{qid}",
        )

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next)

    missing_qs = required_missing_freq()
    if missing_qs:
        st.info(f"To continue, answer all current-frequency items (missing: {len(missing_qs)}).")

    # Guard
    if st.session_state.step == 3 and missing_qs:
        st.session_state.step = 2


# -----------------------
# Step 3 of 5 — Change in frequency (hidden for N/A)
# -----------------------
elif st.session_state.step == 3:
    st.header("Step 3 of 5 — Change in frequency")

    st.caption("You’ll only see change questions for items you did not mark as ‘Not applicable.’")

    non_na_ids = [qid for qid, _, _ in ITEMS if st.session_state.freq_sel.get(qid) != "Not applicable to my role"]

    if not non_na_ids:
        st.warning("You marked all items as ‘Not applicable.’ You can still provide optional feedback.")
    else:
        for qid, text, _ in ITEMS:
            if qid not in non_na_ids:
                continue
            st.session_state.chg_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                CHANGE_OPTIONS,
                index=None,
                horizontal=True,
                key=f"chg_{qid}",
            )

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next)

    missing_chg = required_missing_change(non_na_ids)
    if non_na_ids and missing_chg:
        st.info(f"To continue, answer all change items shown (missing: {len(missing_chg)}).")

    # Guard
    if st.session_state.step == 4 and non_na_ids and missing_chg:
        st.session_state.step = 3


# -----------------------
# Step 4 of 5 — Optional feedback
# -----------------------
elif st.session_state.step == 4:
    st.header("Step 4 of 5 — Optional feedback")

    freq_num, chg_num, overall, overall_desc, avg_change, growth = compute_scores()

    st.markdown(f"**Overall score (current frequency):** {overall} / 5 — {overall_desc}")
    st.caption(
        "This isn’t a grade. It’s a snapshot of how consistently these leadership behaviors show up in your day-to-day application."
    )

    st.markdown("**How to interpret your current score:**")
    st.markdown("- **Consistently / Automatic (4.5–5.0):** behaviors are reliable defaults, even under pressure.")
    st.markdown("- **Often (4.0–4.4):** behaviors show up most of the time; a solid strength.")
    st.markdown("- **Sometimes (3.0–3.9):** behaviors are present but inconsistent; strong opportunity for reinforcement.")
    st.markdown("- **Inconsistent (2.0–2.9):** behaviors show up occasionally; may require clearer systems or support.")
    st.markdown("- **Rarely (≤1.9):** behaviors are not yet habitual; focus on small, repeatable practice.")

    if avg_change is not None:
        st.markdown(
            f"**Average change in application (vs. before the course):** {round1(avg_change)} on a -2 to +2 scale"
        )
        st.caption(
            "This reflects how your *application of these behaviors* has shifted over time, not your potential."
        )
        st.markdown("**Change scale reference:**")
        st.markdown("- **-2:** Much less often applying the behaviors")
        st.markdown("- **-1:** Slightly less often applying the behaviors")
        st.markdown("- **0:** About the same level of application")
        st.markdown("- **+1:** Slightly more often applying the behaviors")
        st.markdown("- **+2:** Much more often applying the behaviors")

    st.session_state.improve_feedback = st.text_area(
        "Any suggestions to improve the course structure, processes, systems, or curriculum? (optional)",
        value=st.session_state.improve_feedback,
        height=140,
    )

    st.markdown("---")
    st.subheader("Testimonial")
    st.caption(
        "If you’re willing, a helpful testimonial often includes: what changed for you, a concrete example, "
        "and what you’d say to someone considering the program."
    )

    if growth:
        st.session_state.testimonial = st.text_area(
            "If you’d like, share a short testimonial or comment about the program (optional)",
            value=st.session_state.testimonial,
            height=140,
        )

        st.session_state.willing_contact = st.checkbox(
            "I’m open to being contacted about using my feedback/testimonial (optional)",
            value=st.session_state.willing_contact,
        )
    else:
        st.session_state.willing_contact = False
        st.session_state.testimonial = ""

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)

    next_label = "Next →" if st.session_state.willing_contact else "Review & submit →"
    with cols[1]:
        st.button(next_label, type="primary", on_click=go_next)


# -----------------------
# Step 5 of 5 — Contact info (only if willing_contact)


# -----------------------
elif st.session_state.step == 5:
    st.header("Step 5 of 5 — Contact information (optional)")

    if not st.session_state.willing_contact:
        st.info("You did not indicate you’re open to being contacted. You can submit now.")
    else:
        st.caption("Provide contact details only if you’re comfortable being contacted about your testimonial.")
        st.session_state.contact_name = st.text_input(
            "Name (optional)",
            value=st.session_state.contact_name,
        )
        st.session_state.contact_email = st.text_input(
            "Email (optional)",
            value=st.session_state.contact_email,
        )

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        submitted = st.button("Submit", type="primary")

    if submitted:
        # Build final context values
        role_anchor = st.session_state.role_anchor
        if role_anchor == "Other":
            role_anchor = st.session_state.role_anchor_other.strip()

        profession = st.session_state.profession
        if profession == "Other":
            profession = st.session_state.profession_other.strip()

        scope = st.session_state.scope
        if scope == "Other":
            scope = st.session_state.scope_other.strip()

        # Final required checks (should already be satisfied, but keep as safety)
        missing = required_missing_step1()
        if missing:
            st.error("Missing required fields: " + ", ".join(missing))
            st.stop()

        missing_freq = required_missing_freq()
        if missing_freq:
            st.error("Missing current-frequency answers.")
            st.stop()

        # Recompute
        freq_num, chg_num, overall, overall_desc, avg_change, growth = compute_scores()

        # Fill change values for NA items as blank
        non_na_ids = [qid for qid, _, _ in ITEMS if freq_num.get(qid) is not None]
        missing_chg = required_missing_change(non_na_ids)
        if non_na_ids and missing_chg:
            st.error("Missing change answers for one or more items.")
            st.stop()

        # Save
        try:
            ws = open_sheet()
            row = [
                pd.Timestamp.utcnow().isoformat(),
                APP_VERSION,
                role_anchor,
                profession,
                st.session_state.years,
                scope,
                str(overall) if overall is not None else "",
                overall_desc,
                str(round1(avg_change)) if avg_change is not None else "",
                "Yes" if growth else "No",
                "Yes" if st.session_state.willing_contact else "No",
                st.session_state.testimonial.strip(),
                st.session_state.improve_feedback.strip(),
                st.session_state.contact_name.strip(),
                st.session_state.contact_email.strip(),
            ]

            # Append frequency (1..16)
            row += [
                str(freq_num[qid]) if isinstance(freq_num.get(qid), (int, float)) else ""
                for qid in range(1, 17)
            ]

            # Append change (1..16) - blank for NA
            row += [
                str(chg_num[qid]) if (qid in non_na_ids and isinstance(chg_num.get(qid), (int, float))) else ""
                for qid in range(1, 17)
            ]

            append_row_to_sheet(ws, row)
            st.success("Submitted and saved to Google Sheets.")
        except Exception as e:
            st.error("Could not save to Google Sheets (secrets / sheet setup may be missing or incomplete).")
            st.exception(e)

        # Reset for next respondent
        st.session_state.step = 1

else:
    st.session_state.step = 1
