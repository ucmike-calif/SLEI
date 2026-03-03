import io
import re
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st
import pandas as pd

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

DOMAINS = ["Sight", "Tenacity", "Ability", "Results"]

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


# =======================
# Helpers
# =======================

def safe_mean(vals):
    vals = [v for v in vals if isinstance(v, (int, float))]
    return (sum(vals) / len(vals)) if vals else None


def round1(x):
    return None if x is None else round(float(x), 1)


def overall_descriptor(score):
    if score is None:
        return "Not scored"
    if score >= 4.5:
        return "Consistently"
    if score >= 4.0:
        return "Often"
    if score >= 3.0:
        return "Sometimes"
    if score >= 2.0:
        return "Occasionally"
    return "Rarely"


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
    """Format top-items list for dashboard text blocks.

    items: list of tuples like (metric_value, qid, text, domain)
    Returns a bullet list string.
    """
    if not items:
        return ""

    lines = []
    for _val, qid, text, _dom in items:
        lines.append(f"• Q{qid}: {text}")

    return "\n".join(lines)



def domain_scores_from_freq(freq_num: dict[int, int | None]):
    scores = {}
    for dom in DOMAINS:
        dom_ids = [qid for qid, _txt, d in ITEMS if d == dom]
        scores[dom] = round1(safe_mean([freq_num.get(qid) for qid in dom_ids]))
    return scores


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


def ensure_headers(ws, headers):
    existing = ws.row_values(1)
    if existing == headers:
        return
    # If sheet is empty, write headers.
    if not any(existing):
        ws.insert_row(headers, 1)
        return
    # If sheet has a different header, do nothing (avoid destructive changes).


def generate_report_code(ts: datetime, role_anchor: str, profession: str):
    base = f"{ts.isoformat()}|{role_anchor}|{profession}"
    # short stable-ish code; not a security feature
    import hashlib

    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:8].upper()


def replace_tokens_in_ppt(prs: Presentation, token_map: dict[str, str]) -> None:
    """Replace {{TOKEN}} placeholders across the deck.

    Robust to tokens being split across multiple runs.
    Any leftover {{...}} is blanked so participants never see raw tokens.
    """
    token_keys = list(token_map.keys())

    def apply(text: str) -> str:
        if not text:
            return text
        for k in token_keys:
            if k in text:
                text = text.replace(k, token_map[k])
        # Safety: remove any remaining tokens
        return re.sub(r"\{\{[^}]+\}\}", "", text)

    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue

            tf = shape.text_frame
            for p in tf.paragraphs:
                # Capture original text across runs
                original = "".join(r.text for r in p.runs) if p.runs else p.text
                replaced = apply(original)
                if replaced == original:
                    continue

                # Preserve paragraph alignment
                alignment = p.alignment

                # Best-effort: preserve first run styling
                size = bold = italic = name = color = None
                if p.runs:
                    f = p.runs[0].font
                    size = f.size
                    bold = f.bold
                    italic = f.italic
                    name = f.name
                    if f.color and f.color.rgb:
                        color = f.color.rgb

                p.text = replaced
                p.alignment = alignment

                if p.runs:
                    r0 = p.runs[0]
                    if size is not None:
                        r0.font.size = size
                    if bold is not None:
                        r0.font.bold = bold
                    if italic is not None:
                        r0.font.italic = italic
                    if name is not None:
                        r0.font.name = name
                    if color is not None:
                        r0.font.color.rgb = color


def build_dashboard_pptx_bytes(
    *,
    token_map: dict[str, str],
    template_path: str,
) -> bytes:
    template_file = Path(template_path)
    if not template_file.exists():
        raise FileNotFoundError(
            f"Dashboard template not found at '{template_path}'. Ensure it is committed to the repo root."
        )

    prs = Presentation(str(template_file))
    replace_tokens_in_ppt(prs, token_map)

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.read()


def init_state():
    st.session_state.setdefault("step", 1)

    # Step 1 context
    st.session_state.setdefault("role_anchor", None)
    st.session_state.setdefault("role_anchor_other", "")
    st.session_state.setdefault("profession", None)
    st.session_state.setdefault("profession_other", "")
    st.session_state.setdefault("years", None)
    st.session_state.setdefault("scope", None)
    st.session_state.setdefault("scope_other", "")

    # Step 2 / 3 answers
    st.session_state.setdefault("freq_sel", {qid: None for qid, _t, _d in ITEMS})
    st.session_state.setdefault("chg_sel", {qid: None for qid, _t, _d in ITEMS})

    # Step 4 feedback
    st.session_state.setdefault("improve_feedback", "")
    st.session_state.setdefault("testimonial", "")
    st.session_state.setdefault("willing_contact", False)

    # Step 5 contact (optional)
    st.session_state.setdefault("contact_name", "")
    st.session_state.setdefault("contact_email", "")

    # Dashboard generation state
    st.session_state.setdefault("dashboard_bytes", None)
    st.session_state.setdefault("dashboard_filename", "")
    st.session_state.setdefault("dashboard_generated_at", "")
    st.session_state.setdefault("dashboard_generated", False)


def go_next():
    st.session_state.step += 1


def go_prev():
    st.session_state.step -= 1


def compute_scores():
    freq_num = {qid: FREQ_MAP.get(st.session_state.freq_sel.get(qid)) for qid, _t, _d in ITEMS}
    chg_num = {qid: CHANGE_MAP.get(st.session_state.chg_sel.get(qid)) for qid, _t, _d in ITEMS}

    overall = round1(safe_mean([v for v in freq_num.values() if isinstance(v, (int, float))]))
    overall_desc = overall_descriptor(overall)

    # only include change items for non-NA freq answers
    chg_vals = [
        chg_num[qid]
        for qid, _t, _d in ITEMS
        if freq_num.get(qid) is not None and isinstance(chg_num.get(qid), (int, float))
    ]
    avg_change = safe_mean(chg_vals)

    # Growth boolean (any positive average)
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
    return [qid for qid, _t, _d in ITEMS if not st.session_state.freq_sel.get(qid)]


def required_missing_change(non_na_ids):
    return [qid for qid in non_na_ids if not st.session_state.chg_sel.get(qid)]


# =======================
# UI Setup
# =======================
st.set_page_config(page_title="SLEI v2.0", layout="wide")
init_state()

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


# =======================
# Step 1 of 5 — Context
# =======================
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

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(missing))


# =======================
# Step 2 of 5 — Current frequency
# =======================
elif st.session_state.step == 2:
    st.header("Step 2 of 5 — Current frequency")

    st.caption(
        "For each item, select how often you perform it now. Choose ‘Not applicable’ if it does not apply to the role you selected."
    )

    for qid, text, _dom in ITEMS:
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

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(missing_qs))


# =======================
# Step 3 of 5 — Change in frequency (hidden for N/A)
# =======================
elif st.session_state.step == 3:
    st.header("Step 3 of 5 — Change in frequency")
    st.caption("You’ll only see change questions for items you did not mark as ‘Not applicable.’")

    non_na_ids = [
        qid
        for qid, _t, _d in ITEMS
        if st.session_state.freq_sel.get(qid) != "Not applicable to my role"
    ]

    if not non_na_ids:
        st.warning("You marked all items as ‘Not applicable.’ You can still provide optional feedback.")
    else:
        for qid, text, _dom in ITEMS:
            if qid not in non_na_ids:
                continue
            st.session_state.chg_sel[qid] = st.radio(
                f"Q{qid}. {text}",
                CHANGE_OPTIONS,
                index=None,
                horizontal=True,
                key=f"chg_{qid}",
            )

    missing_chg = required_missing_change(non_na_ids)
    if non_na_ids and missing_chg:
        st.info(f"To continue, answer all change items shown (missing: {len(missing_chg)}).")

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next, disabled=bool(non_na_ids and missing_chg))


# =======================
# Step 4 of 5 — Testimonial (optional)
# =======================
elif st.session_state.step == 4:
    st.header("Step 4 of 5 — Testimonial (optional)")

    st.caption(
        "This section is optional. If you choose to share a testimonial, it helps others understand the value of the program."
    )

    st.markdown(
        "**Guidance:** A helpful testimonial often includes (1) what changed for you, (2) one concrete example, and (3) what you’d say to someone considering the program."
    )

    st.session_state.testimonial = st.text_area(
        "If you’d like, share a short testimonial or comment about the program (optional)",
        value=st.session_state.testimonial,
        height=160,
    )

    st.session_state.willing_contact = st.checkbox(
        "I’m open to being contacted about using my feedback/testimonial (optional)",
        value=st.session_state.willing_contact,
    )

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)
    with cols[1]:
        st.button("Next →", type="primary", on_click=go_next)


# =======================
# Step 5 of 5 — Submit + optional feedback + dashboard download
# =======================
elif st.session_state.step == 5:
    st.header("Step 5 of 5 — Optional feedback & submit")

    st.session_state.improve_feedback = st.text_area(
        "Any suggestions to improve the course structure, processes, systems, or curriculum? (optional)",
        value=st.session_state.improve_feedback,
        height=160,
    )

    if st.session_state.willing_contact:
        st.subheader("Contact information (optional)")
        st.caption("Provide contact details only if you’re comfortable being contacted about your testimonial.")
        st.session_state.contact_name = st.text_input("Name (optional)", value=st.session_state.contact_name)
        st.session_state.contact_email = st.text_input("Email (optional)", value=st.session_state.contact_email)

    cols = st.columns([1, 1, 6])
    with cols[0]:
        st.button("← Back", on_click=go_prev)

    with cols[1]:
        submitted = st.button("Submit", type="primary")

    if submitted:
        # Prevent duplicate generation on reruns
        if st.session_state.dashboard_generated:
            st.info("Your dashboard is already generated below.")
        else:
            # Required checks
            missing = required_missing_step1()
            if missing:
                st.error("Missing required fields: " + ", ".join(missing))
                st.stop()

            missing_freq = required_missing_freq()
            if missing_freq:
                st.error("Missing current-frequency answers.")
                st.stop()

            non_na_ids = [
                qid
                for qid, _t, _d in ITEMS
                if st.session_state.freq_sel.get(qid) != "Not applicable to my role"
            ]
            missing_chg = required_missing_change(non_na_ids)
            if non_na_ids and missing_chg:
                st.error("Missing change answers for one or more items.")
                st.stop()

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

            # Compute scores
            freq_num, chg_num, overall, overall_desc, avg_change, growth = compute_scores()
            dom_scores = domain_scores_from_freq(freq_num)

            # Strongest reported growth: top positive change items (exclude NA)
            pos_growth = []
            for qid, text, dom in ITEMS:
                if freq_num.get(qid) is None:
                    continue
                v = chg_num.get(qid)
                if isinstance(v, (int, float)) and v > 0:
                    pos_growth.append((v, qid, text, dom))
            pos_growth.sort(key=lambda x: (x[0], x[1]), reverse=True)
            strongest_growth_items = pos_growth[:2]

            # Opportunities: lowest current frequency among non-NA
            lows = []
            for qid, text, dom in ITEMS:
                v = freq_num.get(qid)
                if isinstance(v, (int, float)):
                    lows.append((v, qid, text, dom))
            lows.sort(key=lambda x: (x[0], x[1]))
            opp_items = lows[:2]

            # Growth summary
            improved_count = sum(
                1
                for qid in non_na_ids
                if isinstance(chg_num.get(qid), (int, float)) and chg_num[qid] > 0
            )

                        # Dashboard token map
            ts = datetime.now(timezone.utc)
            report_code = generate_report_code(ts, role_anchor, profession)
            filename = f"SLEI_Dashboard_{report_code}.pptx"

            # Convenience helpers for token text (template already includes bullets)
            def as_item_text(items, idx):
                if len(items) > idx:
                    _v, qid, text, _dom = items[idx]
                    return f"Q{qid}: {text}"
                return ""

            # Strongest growth domain should reflect CHANGE (not current score)
            domain_growth = {}
            for dom in DOMAINS:
                dom_ids = [qid for qid, _txt, d in ITEMS if d == dom]
                vals = [chg_num.get(qid) for qid in dom_ids if freq_num.get(qid) is not None and isinstance(chg_num.get(qid), (int, float))]
                domain_growth[dom] = safe_mean(vals)

            strongest_growth_dom = None
            strongest_growth_val = None
            for dom, val in domain_growth.items():
                if val is None:
                    continue
                if strongest_growth_val is None or val > strongest_growth_val:
                    strongest_growth_val = val
                    strongest_growth_dom = dom

            strongest_detail = (
                f"Most gains in {strongest_growth_dom.lower()} behaviors"
                if strongest_growth_dom
                else ""
            )

            # Consistency interpretation line (kept short for the tile)
            consistency_line = (
                f"Overall: {overall:.1f}/5 ({overall_desc}) — a snapshot of day-to-day application."
                if overall is not None
                else "Overall: Not scored"
            )

            # Growth avg text for tile
            growth_avg_text = "" if avg_change is None else f"{avg_change:+.1f} avg"
            growth_summary = (
                f"Increased in {improved_count} of {len(non_na_ids)} behaviors" if non_na_ids else "No applicable behaviors"
            )

            # Narrative tokens for the lower tiles
            top_growth_why = (
                "Why this matters: small shifts in application compound — especially when you build repeatable habits in real situations."
            )
            top_growth_prompt = (
                "Use this insight as a prompt: pick one real situation in the next 60–90 days where you will practice intentionally."
            )
            opp_nextstep = (
                "Next step: Identify one upcoming situation to practice each behavior intentionally."
            )

            token_map = {
                # Header / intro
                "{{DASHBOARD_TITLE}}": "SLEI Participant Dashboard",
                "{{DASHBOARD_SUBTITLE}}": f"STAR Leadership Effectiveness Index • {APP_VERSION}",
                "{{DASHBOARD_INTRO}}": (
                    "This dashboard is not a grade. It summarizes your self-reported application of key leadership behaviors and the change you reported since before the course."
                ),

                # Consistency tile
                "{{CONSISTENCY_SCORE}}": ("" if overall is None else f"{overall:.1f} / 5"),
                "{{CONSISTENCY_BAND}}": overall_desc,
                "{{CONSISTENCY_INTERPRETATION}}": consistency_line,

                # Growth tile
                "{{GROWTH_AVG}}": growth_avg_text,
                "{{GROWTH_SUMMARY}}": growth_summary,
                "{{GROWTH_INTERPRETATION}}": growth_interpretation(avg_change),

                # Domain scores
                "{{SIGHT_SCORE}}": ("" if dom_scores.get("Sight") is None else f"{dom_scores['Sight']:.1f}"),
                "{{TENACITY_SCORE}}": ("" if dom_scores.get("Tenacity") is None else f"{dom_scores['Tenacity']:.1f}"),
                "{{ABILITY_SCORE}}": ("" if dom_scores.get("Ability") is None else f"{dom_scores['Ability']:.1f}"),
                "{{RESULTS_SCORE}}": ("" if dom_scores.get("Results") is None else f"{dom_scores['Results']:.1f}"),

                # Strongest area of growth (CHANGE-based)
                "{{STRONGEST_GROWTH_DOMAIN}}": strongest_growth_dom or "",
                "{{STRONGEST_DOMAIN}}": strongest_growth_dom or "",  # backward-compat
                "{{STRONGEST_DETAIL}}": strongest_detail,

                # Strongest reported growth (top positive-change items)
                "{{TOP_GROWTH_ITEM_1}}": as_item_text(strongest_growth_items, 0),
                "{{TOP_GROWTH_ITEM_2}}": as_item_text(strongest_growth_items, 1),
                "{{TOP_GROWTH_WHY}}": top_growth_why,
                "{{TOP_GROWTH_PROMPT}}": top_growth_prompt,

                # Opportunities (lowest current-frequency items)
                "{{TOP_OPPORTUNITY_ITEM_1}}": as_item_text(opp_items, 0),
                "{{TOP_OPPORTUNITY_ITEM_2}}": as_item_text(opp_items, 1),
                "{{TOP_OPPORTUNITY_NEXTSTEP}}": opp_nextstep,

                # Action plan
                "{{ACTIONPLAN_PROMPT}}": "Complete this 60–90 day action plan",

                # Optional tracking
                "{{REPORT_CODE}}": report_code,

                # Backward-compat (if older token template variants exist)
                "{{STRONGEST_REPORTED_GROWTH_ITEMS}}": format_top_items(strongest_growth_items),
                "{{OPPORTUNITIES_ITEMS}}": format_top_items(opp_items),
            }

            # Generate dashboard bytes
            try:
                ppt_bytes = build_dashboard_pptx_bytes(token_map=token_map, template_path=TEMPLATE_PATH)
            except Exception as e:
                st.error("Dashboard generation failed.")
                st.exception(e)
                st.stop()

            # Write to Google Sheet
            try:
                ws = open_sheet()

                headers = [
                    "timestamp",
                    "app_version",
                    "role_anchor",
                    "profession",
                    "years_experience",
                    "leadership_scope",
                    "overall_score",
                    "overall_descriptor",
                    "avg_change",
                    "growth",
                    "willing_contact",
                    "testimonial",
                    "improve_feedback",
                    "contact_name",
                    "contact_email",
                ]
                headers += [f"freq_{i}" for i in range(1, 17)]
                headers += [f"chg_{i}" for i in range(1, 17)]
                headers += ["dashboard_generated_at", "dashboard_filename"]

                ensure_headers(ws, headers)

                row = [
                    ts.isoformat(),
                    APP_VERSION,
                    role_anchor,
                    profession,
                    st.session_state.years,
                    scope,
                    ("" if overall is None else str(overall)),
                    overall_desc,
                    ("" if avg_change is None else str(round1(avg_change))),
                    "Yes" if growth else "No",
                    "Yes" if st.session_state.willing_contact else "No",
                    st.session_state.testimonial.strip(),
                    st.session_state.improve_feedback.strip(),
                    st.session_state.contact_name.strip(),
                    st.session_state.contact_email.strip(),
                ]

                row += [
                    ("" if freq_num.get(i) is None else str(freq_num[i]))
                    for i in range(1, 17)
                ]

                non_na_for_save = [i for i in range(1, 17) if freq_num.get(i) is not None]
                row += [
                    ("" if (i not in non_na_for_save or not isinstance(chg_num.get(i), (int, float))) else str(chg_num[i]))
                    for i in range(1, 17)
                ]

                row += [ts.isoformat(), filename]

                append_row_to_sheet(ws, row)
            except Exception as e:
                st.error("Could not save to Google Sheets (check Streamlit secrets and sheet permissions).")
                st.exception(e)
                st.stop()

            # Cache dashboard in session_state (prevents duplicate generation on rerun)
            st.session_state.dashboard_bytes = ppt_bytes
            st.session_state.dashboard_filename = filename
            st.session_state.dashboard_generated_at = ts.isoformat()
            st.session_state.dashboard_generated = True

            st.success("Submitted. Your dashboard is ready below.")

    # Download button (shows after successful submit)
    if st.session_state.dashboard_generated and st.session_state.dashboard_bytes:
        st.markdown("---")
        st.subheader("Download your dashboard")
        st.download_button(
            label="Download PowerPoint",
            data=st.session_state.dashboard_bytes,
            file_name=st.session_state.dashboard_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        st.caption(
            "If you misplace the file, you can re-download during this session using the button above."
        )

else:
    # Safety: reset to first step if state is ever out of range
    st.session_state.step = 1
