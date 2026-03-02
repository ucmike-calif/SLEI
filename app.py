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
    return gc.open(st.se
