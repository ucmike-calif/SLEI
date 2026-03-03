import streamlit as st
import pandas as pd

import io
import re
from datetime import datetime, timezone

import gspread
from google.oauth2.service_account import Credentials
from pptx import Presentation

# =======================
# Config / Constants
# =======================
APP_VERSION = "SLEI-v2.0-pilot"
TEMPLATE_PATH = "SLEI_Dashboard_TEMPLATE.pptx"  # tokenized pptx in repo root

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

# =======================
# Helpers
# =======================

def safe_mean(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return sum(vals) / len(vals) if vals else None


def round1(x):
    return None if x is None else round(float(x), 1)


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


def utc_now_iso():
    return datetime.now(timezone.utc).isoformat()


def slugify_filename(s: str) -> str:
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s[:64] if s else "dashboard"


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
    st.session_state.setdefault("submitted_once", False)  # prevents duplicate dashboard gen in-session

    # Context
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

    # Feedback / testimonial / contact
    st.session_state.setdefault("improve_feedback", "")
    st.session_state.setdefault("testimonial", "")
    st.session_state.setdefault("willing_contact", False)
    st.session_state.setdefault("contact_name", "")
    st.session_state.setdefault("contact_email", "")

    # Dashboard
    st.session_state.setdefault("dashboard_bytes", None)
    st.session_state.setdefault("dashboard_filename", "")
    st.session_state.setdefault("dashboard_generated_at", "")


def go_next():
    st.session_state.step += 1


def go_prev():
    st.session_state.step -= 1


def role_anchor_value():
    ra = st.session_state.role_anchor
    if ra == "Other":
        return st.session_state.role_anchor_other.strip()
    return ra


def profession_value():
    prof = st.session_state.profession
    if prof == "Other":
        return st.session_state.profession_other.strip()
    return prof


def scope_value():
    sc = st.session_state.scope
    if sc == "Other":
        return st.session_state.scope_other.strip()
    return sc


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


def non_na_ids_from_freq():
    return [qid for qid, _, _ in ITEMS if st.session_state.freq_sel.get(qid) != "Not applicable to my role"]


def required_missing_change(non_na_ids):
    return [qid for qid in non_na_ids if not st.session_state.chg_sel.get(qid)]


def compute_scores():
    freq_num = {qid: FREQ_MAP.get(st.session_state.freq_sel.get(qid)) for qid, _, _ in ITEMS}
    chg_num = {qid: CHANGE_MAP.get(st.session_state.chg_sel.get(qid)) for qid, _, _ in ITEMS}

    # Overall (current consistency)
    freq_vals = [v for v in freq_num.values() if isinstance(v, (int, float))]
    overall = round1(safe_mean(freq_vals))
    overall_desc = overall_descriptor(overall)

    # Domain means
    domain_scores = {}
    for dom, ids in DOMAINS.items():
        dom_vals = [freq_num[i] for i in ids if isinstance(freq_num.get(i), (int, float))]
        domain_scores[dom] = round1(safe_mean(dom_vals))

    # Average change (only for non-NA)
    chg_vals = [
        chg_num[qid]
        for qid, _, _ in ITEMS
        if freq_num.get(qid) is not None and isinstance(chg_num.get(qid), (int, float))
    ]
    avg_change = round1(safe_mean(chg_vals))
    growth = (avg_change is not None) and (avg_change > 0)

    # Strongest growth items (top 2 positive change among non-NA)
    growth_items = []
    for qid, text, dom in ITEMS:
        if freq_num.get(qid) is None:
            continue
        cv = chg_num.get(qid)
        if isinstance(cv, (int, float)):
            growth_items.append((cv, qid, text, dom))
    growth_items.sort(reverse=True, key=lambda t: t[0])
    top_growth = [gi for gi in growth_items if gi[0] > 0][:2]

    # Opportunities = lowest current frequency items (top 2) among non-NA
    opp_items = []
    for qid, text, dom in ITEMS:
        fv = freq_num.get(qid)
        if isinstance(fv, (int, float)):
            opp_items.append((fv, qid, text, dom))
    opp_items.sort(key=lambda t: t[0])
    top_opp = opp_items[:2]

    return {
        "freq_num": freq_num,
        "chg_num": chg_num,
        "overall": overall,
        "overall_desc": overall_desc,
        "domain_scores": domain_scores,
        "avg_change": avg_change,
        "growth": growth,
        "top_growth": top_growth,
        "top_opp": top_opp,
    }


def band_sentence_for_descriptor(desc: str) -> str:
    # short, non-judgmental coaching line
    mapping = {
        "Consistently / Automatic": "Very strong consistency — these behaviors are close to default settings.",
        "Often": "Strong consistency — these behaviors show up reliably in most situations.",
        "Sometimes": "Solid foundation — these behaviors show up, with room to make them more consistent.",
        "Inconsistent": "Early momentum — consistency may depend on context, systems, or support.",
        "Rarely": "Starting point — focus on one small practice to build consistency over time.",
        "Not scored": "Not enough applicable items to score this section.",
    }
    return mapping.get(desc, "")


def growth_summary_from_avg(avg_change):
    if avg_change is None:
        return "No change score (insufficient applicable items)"
    sign = "+" if avg_change > 0 else ""  # show + for positive
    return f"{sign}{avg_change:.1f} avg"


def growth_interpretation(avg_change):
    if avg_change is None:
        return "No change score available."
    if avg_change >= 1.25:
        return "Large positive shift in how often you’re applying the behaviors."
    if avg_change >= 0.5:
        return "Clear positive shift in how often you’re applying the behaviors."
    if avg_change > 0:
        return "Modest positive shift in how often you’re applying the behaviors."
    if avg_change == 0:
        return "No overall change in reported application (some items may still have improved)."
    if avg_change <= -1.25:
        return "Large decrease in reported application (consider context, workload, or role changes)."
    if avg_change <= -0.5:
        return "Decrease in reported application (consider context, workload, or role changes)."
    return "Slight decrease in reported application (consider context, workload, or role changes)."


def format_top_items(items):
    # items: list[(val, qid, text, dom)]
    if not items:
        return ""
    lines = []
    for _, qid, text, _dom in items:
        lines.append(f"• Q{qid}: {text}")
    return "
".join(lines)


def pick_strongest_domain_by_score(domain_scores: dict):
    # Use current frequency (not growth) for the domain profile label.
    best_dom = None
    best_val = None
    for dom, val in domain_scores.items():
        if val is None:
            continue
        if best_val is None or val > best_val:
            best_val = val
            best_dom = dom
    return best_dom, best_val


def replace_tokens_in_ppt(prs: Presentation, token_map: dict) -> None:
    """Replaces {{TOKEN}} text in all text frames in-place."""
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            tf = shape.text_frame
            for p in tf.paragraphs:
                for r in p.runs:
                    if not r.text:
                        continue
                    for k, v in token_map.items():
                        if k in r.text:
                            r.text = r.text.replace(k, v)


def build_token_map(scores: dict) -> dict:
    domain_scores = scores["domain_scores"]
    strongest_dom, _ = pick_strongest_domain_by_score(domain_scores)

    # Growth text areas
    top_growth_text = format_top_items(scores["top_growth"]) or "• (No positive-growth items identified)"
    opp_text = format_top_items(scores["top_opp"]) or "• (No opportunities identified)"

    # Optional: strongest detail line can be your narrative or a short phrase
    strongest_detail = "Most gains in alignment & execution habits" if strongest_dom == "Results" else "Most gains in key leadership habits"

    overall = scores["overall"]
    overall_desc = scores["overall_desc"]

    token_map = {
        "{{DASHBOARD_TITLE}}": "SLEI Participant Dashboard",
        "{{DASHBOARD_SUBTITLE}}": "STAR Leadership Effectiveness Index • v2 Pilot",
        "{{DASHBOARD_INTRO}}": (
            "This dashboard is not a course grade. It is a development tool that summarizes your self-perceived growth "
            "during the program and your current consistency in performing critical leadership behaviors across four domains "
            "(Sight, Tenacity, Ability, and Results). Use it to recognize progress, identify 1–2 priorities, and choose a specific "
            "60–90 day practice focus."
        ),

        # Consistency KPI
        "{{CONSISTENCY_SCORE}}": "" if overall is None else f"{overall:.1f} / 5",
        "{{CONSISTENCY_BAND}}": overall_desc,
        "{{CONSISTENCY_BAND_SENTENCE}}": band_sentence_for_descriptor(overall_desc),

        # Growth KPI
        "{{GROWTH_AVG}}": growth_summary_from_avg(scores["avg_change"]),
        "{{GROWTH_SUMMARY}}": (
            "Increased in "
            + str(
                sum(
                    1
                    for qid in range(1, 17)
                    if scores["freq_num"].get(qid) is not None and isinstance(scores["chg_num"].get(qid), (int, float)) and scores["chg_num"].get(qid) > 0
                )
            )
            + " of "
            + str(sum(1 for qid in range(1, 17) if scores["freq_num"].get(qid) is not None))
            + " behaviors"
        ),
        "{{GROWTH_SENTENCE}}": growth_interpretation(scores["avg_change"]),

        # Strongest domain tile
        "{{STRONGEST_DOMAIN}}": strongest_dom or "—",
        "{{STRONGEST_DETAIL}}": strongest_detail,

        # Domain scores
        "{{SIGHT_SCORE}}": "—" if domain_scores.get("Sight") is None else f"{domain_scores['Sight']:.1f}",
        "{{TENACITY_SCORE}}": "—" if domain_scores.get("Tenacity") is None else f"{domain_scores['Tenacity']:.1f}",
        "{{ABILITY_SCORE}}": "—" if domain_scores.get("Ability") is None else f"{domain_scores['Ability']:.1f}",
        "{{RESULTS_SCORE}}": "—" if domain_scores.get("Results") is None else f"{domain_scores['Results']:.1f}",

        # Left column narrative (optional)
        "{{CONSISTENCY_INTERPRETATION}}": (
            "Overall: "
            + ("—" if overall is None else f"{overall:.1f}/5")
            + f" ({overall_desc}) — "
            + (band_sentence_for_descriptor(overall_desc) or "")
        ),

        # Growth / Opportunities blocks
        "{{TOP_GROWTH_ITEMS}}": top_growth_text,
        "{{TOP_OPPORTUNITIES_ITEMS}}": opp_text,

        # Action plan prompt
        "{{ACTIONPLAN_PROMPT}}": "Complete this 60–90 day action plan (progress > perfection):",
    }

    # Ensure all values are strings
    return {k: ("" if v is None else str(v)) for k, v in token_map.items()}


def build_dashboard_bytes(template_path: str, token_map: dict) -> bytes:
    prs = Presentation(template_path)
    replace_tokens_in_ppt(prs, token_map)
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.read()


def reset_for_next_respondent():
    # keep submitted_once True so they can't spam within same session
    keep = {"submitted_once": st.session_state.get("submitted_once", False)}
    st.session_state.clear()
    for k, v in keep.items():
        st.session_state[k] = v
    init_state()


# =======================
# UI
# =======================
init_state()

st.set_page_config(page_title="SLEI v2.0", layout="wide")
st.title("STAR Leadership Effectiveness Index (SLEI) – v2 Pilot")

st.markdown(
    """**Purpose**

This assessment supports your continued development and helps us improve the program.

**Structure**

You will answer 16 questions about key leadership behaviors in two sections:
1. The **Current Frequency** section asks you to indicate how often you perform these behaviors now (after completing the course).
2. The **Change in Frequency** section asks how your current frequency compares to the frequency before the course.

Because leadership looks different across contexts, we ask you to select **one role** and use it consistently so your responses are accurate and comparable.
"""
)

# -----------------------
# Step 1 — Context
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

    missing = required_missing_step1()
    if missing:
        st.info("To continue, complete: " + ", ".join(missing))

    cols = st.columns([1, 8])
    with cols[0]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(missing))

# -----------------------
# Step 2 — Current frequency
# -----------------------
elif st.session_state.step == 2:
    st.header("Step 2 of 5 — Current frequency")
    st.caption("For each item, select how often you perform it now. Choose ‘Not applicable’ if it does not apply to the role you selected.")

    for qid, text, _ in ITEMS:
        st.session_state.freq_sel[qid] = st.radio(
            f"Q{qid}. {text}",
            FREQ_OPTIONS,
            index=None,
            horizontal=True,
            key=f"freq_{qid}",
        )

    missing_qs = required_missing_freq()
    if missing_qs:
        st.info(f"To continue, answer all current-frequency items (missing: {len(missing_qs)}).")

    cols = st.columns([1, 1, 8])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(missing_qs))

# -----------------------
# Step 3 — Change in frequency (hide N/A)
# -----------------------
elif st.session_state.step == 3:
    st.header("Step 3 of 5 — Change in frequency")
    st.caption("You’ll only see change questions for items you did not mark as ‘Not applicable.’")

    non_na_ids = non_na_ids_from_freq()
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

    missing_chg = required_missing_change(non_na_ids) if non_na_ids else []
    if non_na_ids and missing_chg:
        st.info(f"To continue, answer all change items shown (missing: {len(missing_chg)}).")

    cols = st.columns([1, 1, 8])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(missing_chg))

# -----------------------
# Step 4 — Feedback & testimonial
# -----------------------
elif st.session_state.step == 4:
    st.header("Step 4 of 5 — Optional feedback")

    st.session_state.improve_feedback = st.text_area(
        "Any suggestions to improve the course structure, processes, systems, or curriculum? (optional)",
        value=st.session_state.improve_feedback,
        height=140,
    )

    st.markdown("---")
    st.subheader("Testimonial (optional)")
    st.caption(
        "If you’re willing, a helpful testimonial often includes: what changed for you, a concrete example, and what you’d say to someone considering the program."
    )

    # Ask for testimonial only if their results indicate positive growth
    scores = compute_scores()
    if scores["growth"]:
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
        st.session_state.testimonial = ""
        st.session_state.willing_contact = False

    cols = st.columns([1, 1, 8])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next)

# -----------------------
# Step 5 — Review & submit (auto-generate dashboard)
# -----------------------
elif st.session_state.step == 5:
    st.header("Step 5 of 5 — Submit & download your dashboard")

    if st.session_state.willing_contact:
        st.caption("Provide contact details only if you’re comfortable being contacted about your testimonial.")
        st.session_state.contact_name = st.text_input("Name (optional)", value=st.session_state.contact_name)
        st.session_state.contact_email = st.text_input("Email (optional)", value=st.session_state.contact_email)

    # Keep this page clean: no scoring narrative here.
    st.info("When you click Submit, your responses will be saved and your dashboard will be generated for download.")

    cols = st.columns([1, 1, 8])
    with cols[0]:
        st.button("← Back", on_click=go_prev)

    # Prevent duplicate generation in-session
    disable_submit = bool(st.session_state.submitted_once)
    with cols[1]:
        do_submit = st.button("Submit", type="primary", disabled=disable_submit)

    if do_submit:
        # Safety checks
        missing = required_missing_step1()
        if missing:
            st.error("Missing required fields: " + ", ".join(missing))
            st.stop()

        missing_freq = required_missing_freq()
        if missing_freq:
            st.error("Missing current-frequency answers.")
            st.stop()

        non_na_ids = non_na_ids_from_freq()
        missing_chg = required_missing_change(non_na_ids) if non_na_ids else []
        if non_na_ids and missing_chg:
            st.error("Missing change answers for one or more items shown.")
            st.stop()

        # Compute
        scores = compute_scores()

        # Generate dashboard
        try:
            token_map = build_token_map(scores)
            pptx_bytes = build_dashboard_bytes(TEMPLATE_PATH, token_map)
        except Exception as e:
            st.error("Dashboard generation failed. Check that the PPTX template exists in the repo and is tokenized.")
            st.exception(e)
            st.stop()

        # Filename + timestamps
        generated_at = utc_now_iso()
        base = "SLEI_Dashboard"
        fname = f"{base}_{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}.pptx"

        # Save to Google Sheet
        try:
            ws = open_sheet()

            # Prepare values for row
            ra = role_anchor_value()
            prof = profession_value()
            sc = scope_value()

            freq_num = scores["freq_num"]
            chg_num = scores["chg_num"]

            row = [
                utc_now_iso(),
                APP_VERSION,
                ra,
                prof,
                st.session_state.years,
                sc,
                "" if scores["overall"] is None else str(scores["overall"]),
                scores["overall_desc"],
                "" if scores["avg_change"] is None else str(scores["avg_change"]),
                "Yes" if scores["growth"] else "No",
                "Yes" if st.session_state.willing_contact else "No",
                st.session_state.testimonial.strip(),
                st.session_state.improve_feedback.strip(),
                st.session_state.contact_name.strip(),
                st.session_state.contact_email.strip(),
            ]

            # freq_1..freq_16
            row += [
                str(freq_num[qid]) if isinstance(freq_num.get(qid), (int, float)) else ""
                for qid in range(1, 17)
            ]

            # chg_1..chg_16 (blank for NA)
            non_na_ids_num = [qid for qid in range(1, 17) if freq_num.get(qid) is not None]
            row += [
                str(chg_num[qid]) if (qid in non_na_ids_num and isinstance(chg_num.get(qid), (int, float))) else ""
                for qid in range(1, 17)
            ]

            # dashboard fields
            row += [generated_at, fname]

            append_row_to_sheet(ws, row)
        except Exception as e:
            st.error("Could not save to Google Sheets (secrets / sheet setup may be missing or incomplete).")
            st.exception(e)
            st.stop()

        # Mark as submitted (prevents duplicates)
        st.session_state.submitted_once = True
        st.session_state.dashboard_bytes = pptx_bytes
        st.session_state.dashboard_filename = fname
        st.session_state.dashboard_generated_at = generated_at

        st.success("Submitted. Your dashboard is ready to download below.")

    # Show download button if we have a dashboard in session
    if st.session_state.dashboard_bytes:
        st.download_button(
            label="Download your dashboard (PowerPoint)",
            data=st.session_state.dashboard_bytes,
            file_name=st.session_state.dashboard_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        st.caption("Tip: Save this file somewhere you can find it later (e.g., Downloads → move to a folder).")

else:
    st.session_state.step = 1
