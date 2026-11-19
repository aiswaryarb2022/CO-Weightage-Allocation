import streamlit as st
import pandas as pd
import numpy as np
import re
import json
import requests

from io import BytesIO

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)
from reportlab.lib.units import mm


# ==========================================================
# 1. PAGE CONFIGURATION
# ==========================================================

st.set_page_config(
    page_title="CO Assessment Dashboard",
    page_icon="📊",
    layout="wide"
)

ASSESSMENT_ORDER = ["PA1", "PA2", "LE", "LM"]

ALL_ASSESSMENTS = ASSESSMENT_ORDER + ["ND"]

CO_LIST = ["CO1", "CO2", "CO3", "CO4"]

RESCALING_OPTIONS = [
    "Absolute",
    "Relative",
    "Moderate"
]

CLUBBING_OPTIONS = [
    "Average of All",
    "Average of Best",
    "Weighted Average"
]

ND_REPLACEMENT_OPTIONS = [
    "Do not change",
    "PA1",
    "PA2",
    "LE",
    "LM"
]


# ============================================================
# TEAL THEME – COMPLETE DASHBOARD
# ============================================================

st.markdown("""
<style>

/* -----------------------------------------
   MAIN COLOURS
----------------------------------------- */

:root {
    --teal-dark: #064E4A;
    --teal-primary: #0F766E;
    --teal-medium: #0D9488;
    --teal-light: #CCFBF1;
    --teal-pale: #F0FDFA;
    --teal-border: #99F6E4;
    --text-dark: #134E4A;
    --white: #FFFFFF;
}

/* -----------------------------------------
   MAIN BACKGROUND
----------------------------------------- */

.stApp {
    background-color: #F0FDFA;
    color: #134E4A;
}

/* Main content */

[data-testid="stAppViewContainer"] {
    background-color: #F0FDFA;
}

/* -----------------------------------------
   SIDEBAR
----------------------------------------- */

[data-testid="stSidebar"] {
    background-color: #064E4A;
}

[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {
    color: #FFFFFF;
}

/* -----------------------------------------
   MAIN HEADINGS
----------------------------------------- */

h1 {
    color: #064E4A !important;
    font-weight: 800 !important;
}

h2 {
    color: #0F766E !important;
    font-weight: 750 !important;
}

h3 {
    color: #0F766E !important;
    font-weight: 700 !important;
}

/* -----------------------------------------
   NORMAL TEXT
----------------------------------------- */

p, label, span {
    color: #134E4A;
}

/* -----------------------------------------
   BUTTONS
----------------------------------------- */

.stButton > button,
.stDownloadButton > button {
    background-color: #0F766E !important;
    color: #FFFFFF !important;
    border: 1px solid #0F766E !important;
    border-radius: 9px !important;
    font-weight: 700 !important;
    padding: 0.55rem 1.1rem !important;
    transition: 0.2s ease-in-out;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #064E4A !important;
    border-color: #064E4A !important;
    color: #FFFFFF !important;
}

.stButton > button:focus,
.stDownloadButton > button:focus {
    box-shadow: 0 0 0 0.2rem #99F6E4 !important;
}

/* -----------------------------------------
   INPUT BOXES
----------------------------------------- */

.stTextInput input,
.stNumberInput input,
.stDateInput input,
.stTextArea textarea {
    background-color: #FFFFFF !important;
    color: #134E4A !important;
    border: 1px solid #99F6E4 !important;
    border-radius: 7px !important;
}

/* Input focus */

.stTextInput input:focus,
.stNumberInput input:focus,
.stDateInput input:focus,
.stTextArea textarea:focus {
    border: 2px solid #0D9488 !important;
    box-shadow: 0 0 0 0.1rem #CCFBF1 !important;
}

/* -----------------------------------------
   SELECTBOX AND MULTISELECT
----------------------------------------- */

[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    border-color: #99F6E4 !important;
    color: #134E4A !important;
    border-radius: 7px !important;
}

[data-baseweb="popover"] {
    background-color: #FFFFFF !important;
}

[data-baseweb="menu"] {
    background-color: #FFFFFF !important;
}

[data-baseweb="option"] {
    color: #134E4A !important;
}

[data-baseweb="option"]:hover {
    background-color: #CCFBF1 !important;
}

/* -----------------------------------------
   DATAFRAME
----------------------------------------- */

[data-testid="stDataFrame"] {
    border: 1px solid #99F6E4;
    border-radius: 10px;
    overflow: hidden;
}

/* -----------------------------------------
   DATA EDITOR
----------------------------------------- */

[data-testid="stDataEditor"] {
    border: 1px solid #99F6E4 !important;
    border-radius: 10px !important;
    overflow: hidden;
}

/* -----------------------------------------
   FILE UPLOADER
----------------------------------------- */

[data-testid="stFileUploader"] {
    background-color: #FFFFFF;
    border: 1.5px dashed #0D9488;
    border-radius: 12px;
    padding: 15px;
}

/* -----------------------------------------
   ALERT MESSAGES
----------------------------------------- */

[data-testid="stAlert"] {
    border-radius: 9px !important;
}

/* -----------------------------------------
   INFORMATION BOXES
----------------------------------------- */

[data-testid="stAlert"][kind="info"] {
    background-color: #CCFBF1 !important;
    border-left: 5px solid #0F766E !important;
}

/* -----------------------------------------
   SUCCESS BOXES
----------------------------------------- */

[data-testid="stAlert"][kind="success"] {
    background-color: #D1FAE5 !important;
    border-left: 5px solid #059669 !important;
}

/* -----------------------------------------
   WARNING BOXES
----------------------------------------- */

[data-testid="stAlert"][kind="warning"] {
    background-color: #FEF3C7 !important;
    border-left: 5px solid #D97706 !important;
}

/* -----------------------------------------
   ERROR BOXES
----------------------------------------- */

[data-testid="stAlert"][kind="error"] {
    background-color: #FEE2E2 !important;
    border-left: 5px solid #DC2626 !important;
}

/* -----------------------------------------
   DIVIDERS
----------------------------------------- */

hr {
    border-color: #99F6E4 !important;
}

/* -----------------------------------------
   EXPANDERS
----------------------------------------- */

[data-testid="stExpander"] {
    background-color: #FFFFFF;
    border: 1px solid #99F6E4;
    border-radius: 9px;
}

/* -----------------------------------------
   TABS
----------------------------------------- */

button[data-baseweb="tab"] {
    color: #0F766E !important;
    font-weight: 650 !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    color: #064E4A !important;
    border-bottom: 3px solid #0D9488 !important;
}

/* -----------------------------------------
   METRICS
----------------------------------------- */

[data-testid="stMetric"] {
    background-color: #FFFFFF;
    border: 1px solid #99F6E4;
    border-radius: 10px;
    padding: 15px;
}

[data-testid="stMetricLabel"] {
    color: #0F766E !important;
}

[data-testid="stMetricValue"] {
    color: #064E4A !important;
    font-weight: 800 !important;
}

/* -----------------------------------------
   CAPTIONS
----------------------------------------- */

.stCaption {
    color: #527C77 !important;
}

/* -----------------------------------------
   SCROLLBAR
----------------------------------------- */

::-webkit-scrollbar {
    width: 8px;
}

::-webkit-scrollbar-thumb {
    background-color: #0D9488;
    border-radius: 10px;
}

::-webkit-scrollbar-track {
    background-color: #CCFBF1;
}

</style>
""", unsafe_allow_html=True)


# ==========================================================
# 3. HELPER FUNCTIONS
# ==========================================================

def normalize_column_name(name):

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(name).lower()
    )


def find_column(df, aliases):

    lookup = {
        normalize_column_name(c): c
        for c in df.columns
    }

    for alias in aliases:

        found = lookup.get(
            normalize_column_name(alias)
        )

        if found is not None:
            return found

    return None


def standardize_columns(df):

    aliases = {

        "Course Name": [
            "Course Name",
            "Course",
            "Subject Name"
        ],

        "Course Code": [
            "Course Code",
            "Subject Code"
        ],

        "Course Type": [
            "Course Type",
            "Course Category",
            "Category"
        ],

        "Assessment": [
            "Assessment",
            "Assessment Type",
            "Exam"
        ],

        "CO": [
            "CO",
            "Course Outcome",
            "Course Outcomes"
        ],

        "Mark Obtained": [
            "Mark Obtained",
            "Marks Obtained",
            "Mark",
            "Marks"
        ],

        "Max Marks": [
            "Max Marks",
            "Maximum Marks",
            "Total Marks",
            "Maximum Mark"
        ],

        "Date": [
            "Date",
            "Assessment Date",
            "Exam Date"
        ],

        "Reg No": [
            "Reg No",
            "Register Number",
            "Registration Number",
            "Register No",
            "Roll No"
        ],

        "Admission No": [
            "Admission No",
            "Admission Number"
        ],

        "Student ID": [
            "Student ID",
            "Student Number"
        ],

        "Student Name": [
            "Student Name",
            "Name"
        ],

        "Student Prog": [
            "Student Prog",
            "Student Program",
            "Programme"
        ]
    }

    rename = {}

    for target, names in aliases.items():

        if target not in df.columns:

            source = find_column(df, names)

            if source is not None:
                rename[source] = target

    return df.rename(columns=rename)


def normalize_assessment(value):

    value = (
        str(value)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
        .replace("_", "")
    )

    mapping = {

        "PA01": "PA1",
        "PA1": "PA1",

        "PA02": "PA2",
        "PA2": "PA2",

        "LE1": "LE",
        "LE": "LE",
        "LABEXAM": "LE",

        "LM1": "LM",
        "LM": "LM",

        "ND": "ND"
    }

    return mapping.get(value, value)


def normalize_co(value):

    match = re.search(
        r"(\d+)",
        str(value).strip().upper()
    )

    if match:
        return f"CO{int(match.group(1))}"

    return str(value).strip().upper()


def date_key(value):

    if pd.isna(value):
        return "Not specified"

    parsed = pd.to_datetime(
        value,
        errors="coerce",
        dayfirst=True
    )

    if pd.isna(parsed):
        return "Not specified"

    return parsed.strftime("%d-%m-%Y")


def safe_key(value):

    return re.sub(
        r"[^a-zA-Z0-9_]",
        "_",
        str(value)
    )


def clean_identifier(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in [
        "",
        "nan",
        "none",
        "null",
        "unknown",
        "na",
        "n/a"
    ]:
        return ""

    return value


def read_uploaded_file(uploaded_file):

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


# ==========================================================
# 4. PREPARE DATA
# ==========================================================

def prepare_data(raw_df):

    df = standardize_columns(raw_df.copy())

    required = [
        "Course Name",
        "Assessment",
        "CO",
        "Mark Obtained"
    ]

    missing = [
        c for c in required
        if c not in df.columns
    ]

    if missing:

        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing)
        )

    for col, default in [
        ("Course Code", "Not specified"),
        ("Course Type", "Not specified")
    ]:

        if col not in df.columns:
            df[col] = default

        df[col] = (
            df[col]
            .fillna(default)
            .astype(str)
            .str.strip()
        )

    if "Max Marks" not in df.columns:
        df["Max Marks"] = 100

    if "Date" not in df.columns:
        df["Date"] = pd.NaT

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce",
        dayfirst=True
    ).dt.normalize()

    df["Mark Obtained"] = pd.to_numeric(
        df["Mark Obtained"],
        errors="coerce"
    )

    df["Max Marks"] = pd.to_numeric(
        df["Max Marks"],
        errors="coerce"
    ).fillna(100)

    df.loc[
        df["Max Marks"] <= 0,
        "Max Marks"
    ] = 100

    df["Assessment"] = (
        df["Assessment"].apply(normalize_assessment)
    )

    df["CO"] = df["CO"].apply(normalize_co)

    preferred_ids = [
        "Reg No",
        "Admission No",
        "Student ID",
        "Student Name"
    ]

    available_ids = [
        c for c in preferred_ids
        if c in df.columns
    ]

    if not available_ids:

        raise ValueError(
            "A student identifier is required. "
            "Include Reg No, Admission No, Student ID "
            "or Student Name."
        )

    def get_student_identifier(row):

        for col in available_ids:

            identifier = clean_identifier(row[col])

            if identifier:
                return identifier

        return f"Missing_ID_{row.name}"

    df["__StudentKey"] = df.apply(
        get_student_identifier,
        axis=1
    )

    identity_columns = ["__StudentKey"]

    df = df.dropna(
        subset=["Mark Obtained", "Course Name"]
    ).copy()

    df = df[
        df["Assessment"].isin(ALL_ASSESSMENTS)
        & df["CO"].isin(CO_LIST)
    ].copy()

    if df.empty:

        raise ValueError(
            "No valid assessment records were found. "
            "Check the Assessment, CO and Mark Obtained columns."
        )

    return df, identity_columns


# ==========================================================
# 5. ND REPLACEMENT
# ==========================================================

def build_nd_replacement_table(course_data):

    nd_data = course_data[
        course_data["Assessment"] == "ND"
    ].copy()

    if nd_data.empty:
        return pd.DataFrame()

    nd_rows = []

    for row_index, row in nd_data.iterrows():

        nd_rows.append({

            "Record Index": row_index,

            "Date": date_key(row["Date"]),

            "ND CO": row["CO"],

            "Mark Obtained": row["Mark Obtained"],

            "Replace With": "Do not change"
        })

    return pd.DataFrame(nd_rows)


def apply_nd_replacements(
    course_data,
    edited_nd_config
):

    updated_data = course_data.copy()

    for _, row in edited_nd_config.iterrows():

        row_index = int(row["Record Index"])

        updated_data.loc[
            row_index,
            "Assessment"
        ] = row["Replace With"]

    return updated_data


# ==========================================================
# 6. DETECT REPEATED CO COMBINATIONS
# ==========================================================

def get_repeated_co_summary(course_data):

    """
    Clubbing is required only if the same CO occurs
    on more than one date within the same assessment.

    Student records are not counted as separate entries.
    """

    temp = course_data.copy()

    temp["Date Label"] = temp["Date"].apply(date_key)

    counts = (
        temp.groupby(
            ["Assessment", "CO"],
            dropna=False
        )["Date Label"]
        .nunique()
    )

    repeated_summary = {}

    for key, count in counts.items():

        if count > 1:
            repeated_summary[key] = int(count)

    return repeated_summary


# ==========================================================
# 7. CREATE ORDERED ASSESSMENT-CO ROWS
# ==========================================================

def get_settings_rows(course_data):

    settings = course_data[
        ["Assessment", "Date", "CO"]
    ].copy()

    settings["Date Label"] = (
        settings["Date"].apply(date_key)
    )

    settings = settings.drop_duplicates(
        ["Assessment", "Date Label", "CO"]
    ).copy()

    settings["_assessment_order"] = (
        settings["Assessment"].map({
            a: i for i, a in enumerate(ASSESSMENT_ORDER)
        })
    )

    settings["_date_order"] = pd.to_datetime(
        settings["Date Label"],
        format="%d-%m-%Y",
        errors="coerce"
    )

    settings["_co_order"] = (
        settings["CO"].map({
            co: i for i, co in enumerate(CO_LIST)
        })
    )

    settings = settings.sort_values(
        [
            "_assessment_order",
            "_date_order",
            "_co_order"
        ],
        na_position="last"
    )

    return settings[
        ["Assessment", "Date Label", "CO"]
    ].reset_index(drop=True)

# ==========================================================
# 8. CALCULATE RESCALED RESULTS
# ==========================================================

def calculate_results(
    data,
    identity_columns,
    config_map
):

    class_average = {}

    for (
        assessment_name,
        assessment_date,
        co_name
    ), class_group in data.groupby(
        ["Assessment", "Date", "CO"],
        dropna=False
    ):

        class_average[
            (
                assessment_name,
                date_key(assessment_date),
                co_name
            )
        ] = class_group["Mark Obtained"].mean()

    group_cols = (
        identity_columns
        + [
            "Course Name",
            "Course Code",
            "Course Type",
            "Assessment",
            "Date",
            "CO"
        ]
    )

    rows = []

    for key, group in data.groupby(
        group_cols,
        dropna=False,
        sort=False
    ):

        values = dict(
            zip(
                group_cols,
                key if isinstance(key, tuple) else (key,)
            )
        )

        assessment = values["Assessment"]
        dlabel = date_key(values["Date"])
        co = values["CO"]

        settings = config_map.get(
            (assessment, dlabel, co)
        )

        if settings is None:

            raise ValueError(
                f"Calculation settings are missing for "
                f"{assessment}, {dlabel}, {co}."
            )

        avg_mark = group["Mark Obtained"].mean()
        avg_max = group["Max Marks"].mean()

        class_avg = class_average.get(
            (assessment, dlabel, co),
            np.nan
        )

        method = settings["Rescaling Method"]

        # --------------------------------------------------
        # RELATIVE RESCALING
        # --------------------------------------------------

        if method == "Relative":

            if pd.notna(class_avg) and class_avg != 0:

                rescaled = (
                    avg_mark / class_avg
                ) * 20

            else:
                rescaled = np.nan

        # --------------------------------------------------
        # ABSOLUTE RESCALING
        # --------------------------------------------------

        elif method == "Absolute":

            if pd.notna(avg_max) and avg_max > 0:

                rescaled = (
                    avg_mark / avg_max
                ) * 100

            else:
                rescaled = np.nan

        # --------------------------------------------------
        # MODERATE RESCALING
        # --------------------------------------------------

        elif method == "Moderate":

            minimum_mark = settings.get(
                "Minimum Mark",
                0.0
            )

            adjusted_mark = max(
                avg_mark,
                minimum_mark
            )

            if pd.notna(avg_max) and avg_max > 0:

                rescaled = (
                    adjusted_mark / avg_max
                ) * 100

            else:
                rescaled = np.nan

        else:

            rescaled = np.nan

        # Keep scores within 0 to 100
        if pd.notna(rescaled):

            rescaled = float(
                np.clip(rescaled, 0, 100)
            )

        result_row = {
            c: values[c]
            for c in identity_columns
        }

        result_row.update({

            "Course Name": values["Course Name"],
            "Course Code": values["Course Code"],
            "Course Type": values["Course Type"],

            "Assessment": assessment,
            "Date Label": dlabel,
            "CO": co,

            "Original Mark": avg_mark,
            "Rescaling Method": method,
            "Rescaling Reference": class_avg,
            "Rescaled Mark": rescaled

        })

        rows.append(result_row)

    return pd.DataFrame(rows)


# ==========================================================
# 9. CLUBBING SAME CO ACROSS DIFFERENT DATES
# ==========================================================

def aggregate_co_score(
    group,
    assessment,
    co,
    repeated_pairs,
    clubbing_map,
    best_n_map,
    co_weights
):

    co_rows = group[
        (group["Assessment"] == assessment)
        & (group["CO"] == co)
    ].dropna(
        subset=["Rescaled Mark"]
    )

    if co_rows.empty:
        return np.nan

    date_values = (
        co_rows.groupby("Date Label")["Rescaled Mark"]
        .mean()
    )

    combination = (assessment, co)

    if combination not in repeated_pairs:

        return float(date_values.mean())

    method = clubbing_map.get(combination)

    if method == "Average of All":

        return float(date_values.mean())

    elif method == "Average of Best":

        best_n = best_n_map.get(
            combination,
            1
        )

        best_values = (
            date_values
            .sort_values(ascending=False)
            .head(best_n)
        )

        return float(best_values.mean())

    elif method == "Weighted Average":

        weights = pd.Series(
            {
                dlabel: co_weights.get(
                    (assessment, co, dlabel),
                    0.0
                )
                for dlabel in date_values.index
            },
            dtype=float
        )

        weights = weights.reindex(
            date_values.index
        ).fillna(0)

        if weights.sum() <= 0:
            return np.nan

        return float(
            np.average(
                date_values.values,
                weights=weights.values
            )
        )

    return np.nan


# ==========================================================
# 10. CREATE ASSESSMENT-CO MATRIX
# ==========================================================

def make_output_table(
    results,
    identity_columns,
    repeated_pairs,
    clubbing_map,
    best_n_map,
    co_weights
):

    keys = (
        identity_columns
        + [
            "Course Name",
            "Course Code",
            "Course Type"
        ]
    )

    output_rows = []

    for key, group in results.groupby(
        keys,
        dropna=False,
        sort=False
    ):

        key_values = (
            key if isinstance(key, tuple)
            else (key,)
        )

        base = dict(zip(keys, key_values))

        row = {
            col: base[col]
            for col in identity_columns
        }

        row.update({
            "Course Name": base["Course Name"],
            "Course Code": base["Course Code"],
            "Course Type": base["Course Type"]
        })

        # --------------------------------------------------
        # INDIVIDUAL ASSESSMENT-CO MARKS
        # --------------------------------------------------

        for assessment in ASSESSMENT_ORDER:

            for co in CO_LIST:

                score = aggregate_co_score(
                    group,
                    assessment,
                    co,
                    repeated_pairs,
                    clubbing_map,
                    best_n_map,
                    co_weights
                )

                row[f"{assessment}_{co}"] = score

        # --------------------------------------------------
        # ASSESSMENT-WISE AVERAGES
        # --------------------------------------------------

        for assessment in ASSESSMENT_ORDER:

            values = [
                row.get(f"{assessment}_{co}", np.nan)
                for co in CO_LIST
            ]

            valid = [
                value for value in values
                if pd.notna(value)
            ]

            row[assessment] = (
                float(np.mean(valid))
                if valid
                else np.nan
            )

        # --------------------------------------------------
        # CO-WISE AVERAGES
        # --------------------------------------------------

        for co in CO_LIST:

            values = [
                row.get(f"{assessment}_{co}", np.nan)
                for assessment in ASSESSMENT_ORDER
            ]

            valid = [
                value for value in values
                if pd.notna(value)
            ]

            row[co] = (
                float(np.mean(valid))
                if valid
                else np.nan
            )

        output_rows.append(row)

    output = pd.DataFrame(output_rows)

    ordered = (
        identity_columns
        + [
            "Course Name",
            "Course Code",
            "Course Type"
        ]
        + [
            f"{assessment}_{co}"
            for assessment in ASSESSMENT_ORDER
            for co in CO_LIST
        ]
        + ASSESSMENT_ORDER
        + CO_LIST
    )

    for col in ordered:

        if col not in output.columns:
            output[col] = np.nan

    return output[ordered].replace(
        {None: np.nan}
    )


# ==========================================================
# 11. CREATE EXCEL FILE
# ==========================================================

def create_excel_file(export_data):

    buffer = BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        export_data.to_excel(
            writer,
            index=False,
            sheet_name="CO Assessment"
        )

        ws = writer.sheets["CO Assessment"]

        fill = PatternFill(
            start_color="185C37",
            end_color="185C37",
            fill_type="solid"
        )

        font = Font(
            color="FFFFFF",
            bold=True
        )

        for cell in ws[1]:

            cell.fill = fill
            cell.font = font

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        for cells in ws.columns:

            letter = get_column_letter(
                cells[0].column
            )

            width = max(
                (
                    len(str(c.value))
                    for c in cells
                    if c.value is not None
                ),
                default=10
            )

            ws.column_dimensions[letter].width = min(
                width + 3,
                28
            )

        for row in ws.iter_rows(min_row=2):

            for cell in row:

                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 28

    buffer.seek(0)

    return buffer.getvalue()


# ==========================================================
# 12. FETCH PHASE 0 ALLOCATION FROM GOOGLE SHEETS
# ==========================================================

def fetch_phase0_allocation(course_code):

    try:

        # --------------------------------------------------
        # GOOGLE SHEETS CONNECTION
        # --------------------------------------------------

        script_url = st.secrets["GOOGLE_SCRIPT_URL"]

        secret_token = st.secrets["GOOGLE_SECRET_TOKEN"]

        response = requests.get(
            script_url,
            params={
                "token": secret_token,
                "course_code": str(course_code).strip()
            },
            timeout=30
        )

        response.raise_for_status()

        data = response.json()

        # --------------------------------------------------
        # CHECK RESPONSE
        # --------------------------------------------------

        if data.get("status") == "error":

            return None, data.get(
                "message",
                "Allocation retrieval failed."
            )

        # --------------------------------------------------
        # READ HEADERS AND ROWS
        # --------------------------------------------------

        headers = data.get("headers")

        rows = data.get("rows")

        if not headers:

            return None, (
                "Google Sheets did not return column headers."
            )

        if not rows:

            return None, (
                "Google Sheets did not return allocation rows."
            )

        # --------------------------------------------------
        # CREATE DATAFRAME USING CORRECT HEADERS
        # --------------------------------------------------

        allocation_df = pd.DataFrame(
            rows,
            columns=headers
        )

        # --------------------------------------------------
        # CLEAN COLUMN NAMES
        # --------------------------------------------------

        allocation_df.columns = (
            allocation_df.columns.astype(str).str.strip()
        )

        if allocation_df.empty:

            return None, (
                "No Phase 0 allocation was found for this course."
            )

        return allocation_df, None

    except Exception as e:

        return None, str(e)
# ==========================================================
# 13. PROCESS PHASE 0 ALLOCATION
# ==========================================================

def process_phase0_allocation(allocation):

    allocation = allocation.copy()

    # ------------------------------------------------------
    # DETECT ACTUAL HEADER ROW
    # ------------------------------------------------------

    allocation.columns = (
        allocation.columns.astype(str).str.strip()
    )

    if "Assessment" not in allocation.columns and \
       "Assessment Type" not in allocation.columns:

        header_index = None

        for i in range(len(allocation)):

            row_values = [
                str(value).strip().lower()
                for value in allocation.iloc[i].values
            ]

            if "assessment" in row_values or \
               "assessment type" in row_values:

                header_index = i
                break

        if header_index is None:

            raise ValueError(
                "Could not find the Assessment header row "
                "in the Google Sheets data."
            )

        allocation.columns = allocation.iloc[header_index]

        allocation = allocation.iloc[header_index + 1:].copy()

    # ------------------------------------------------------
    # CLEAN COLUMN NAMES
    # ------------------------------------------------------

    allocation.columns = [
        str(col).strip()
        for col in allocation.columns
    ]

    if "Assessment" in allocation.columns:

        allocation = allocation.rename(
            columns={"Assessment": "Assessment Type"}
        )

    if "Assessment Type" not in allocation.columns:

        raise ValueError(
            "Assessment column could not be identified."
        )

    # ------------------------------------------------------
    # IDENTIFY CO COLUMNS
    # ------------------------------------------------------

    co_columns = {}

    for co in CO_LIST:

        found = find_column(allocation, [co])

        if found is None:

            raise ValueError(
                f"{co} column was not found in Phase 0."
            )

        co_columns[co] = found

    # ------------------------------------------------------
    # MAP ASSESSMENTS
    # ------------------------------------------------------

    mapping = {
        "PA1": "PA1",
        "PA2": "PA2",
        "LE1": "LE",
        "LE": "LE",
        "LM": "LM"
    }

    weights = {}

    for _, row in allocation.iterrows():

        assessment_value = str(
            row["Assessment Type"]
        ).strip().upper()

        if assessment_value not in mapping:
            continue

        assessment = mapping[assessment_value]

        weights[assessment] = {}

        for co in CO_LIST:

            value = pd.to_numeric(
                row[co_columns[co]],
                errors="coerce"
            )

            weights[assessment][co] = (
                float(value)
                if pd.notna(value)
                else 0.0
            )

    # ------------------------------------------------------
    # VALIDATE REQUIRED ASSESSMENTS
    # ------------------------------------------------------

    required_assessments = [
        "PA1",
        "PA2",
        "LE",
        "LM"
    ]

    missing = [
        a for a in required_assessments
        if a not in weights
    ]

    if missing:

        raise ValueError(
            "Missing allocation rows: "
            + ", ".join(missing)
        )

    return weights

# ==========================================================
# 14. DISPLAY PHASE 0 ALLOCATION
# ==========================================================

def create_allocation_display(weights):

    rows = []

    for assessment in ASSESSMENT_ORDER:

        row = {
            "Assessment Type": assessment
        }

        for co in CO_LIST:

            row[co] = weights[assessment][co]

        row["Total"] = sum(
            weights[assessment][co]
            for co in CO_LIST
        )

        rows.append(row)

    return pd.DataFrame(rows)



# ==========================================================
# 15. PHASE 0 × PHASE 2 WEIGHTED CO CALCULATION
# ==========================================================

def calculate_weighted_cca(
    assessment_matrix,
    allocation_weights
):

    result = assessment_matrix.copy()

    # ------------------------------------------------------
    # CALCULATE EACH ASSESSMENT CONTRIBUTION
    # Formula: Rescaled Mark × Phase 0 Weight / 100
    # ------------------------------------------------------

    for assessment in ASSESSMENT_ORDER:

        for co in CO_LIST:

            source_column = f"{assessment}_{co}"

            contribution_column = (
                f"{assessment}_{co}_Weighted"
            )

            if source_column not in result.columns:
                result[source_column] = np.nan

            phase2_marks = pd.to_numeric(
                result[source_column],
                errors="coerce"
            )

            phase0_weight = float(
                allocation_weights[assessment][co]
            )

            result[contribution_column] = (
                phase2_marks * phase0_weight / 100
            )

    # ------------------------------------------------------
    # FINAL CO TOTALS
    # ------------------------------------------------------

    for co in CO_LIST:

        contribution_columns = [
            f"{assessment}_{co}_Weighted"
            for assessment in ASSESSMENT_ORDER
        ]

        result[co] = result[
            contribution_columns
        ].fillna(0).sum(axis=1)

    return result
# ==========================================================
# 16. CREATE FINAL WEIGHTED OUTPUT
# ==========================================================

def create_final_weighted_output(
    weighted_matrix,
    identity_columns
):

    columns = (
        identity_columns
        + [
            "Course Name",
            "Course Code",
            "Course Type"
        ]
        + CO_LIST
    )

    columns = [
        col for col in columns
        if col in weighted_matrix.columns
    ]

    return weighted_matrix[columns].copy()


# ==========================================================
# 17. FINAL WEIGHTED CO MATRIX PDF
# ==========================================================

def create_weighted_cca_pdf(course_details, final_matrix):

    import io

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import (
        SimpleDocTemplate,
        Table,
        TableStyle,
        Paragraph,
        Spacer
    )

    buffer = io.BytesIO()

    # ------------------------------------------------------
    # COURSE DETAILS
    # ------------------------------------------------------

    course_name = course_details.get("course_name", "")
    course_code = course_details.get("course_code", "")
    course_type = course_details.get("course_type", "")

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    elements = []

    styles = getSampleStyleSheet()

    title_style = styles["Title"]

    elements.append(
        Paragraph(
            "FINAL WEIGHTED CO MATRIX",
            title_style
        )
    )

    elements.append(Spacer(1, 15))

    # ------------------------------------------------------
    # COURSE INFORMATION
    # ------------------------------------------------------

    course_data = [
        ["COURSE NAME", course_name],
        ["COURSE CODE", course_code],
        ["COURSE TYPE", course_type]
    ]

    course_table = Table(
        course_data,
        colWidths=[120, 350]
    )

    course_table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (0, -1),
             colors.HexColor("#064E4A")),

            ("TEXTCOLOR", (0, 0), (0, -1),
             colors.white),

            ("BACKGROUND", (1, 0), (1, -1),
             colors.HexColor("#CCFBF1")),

            ("TEXTCOLOR", (1, 0), (1, -1),
             colors.HexColor("#064E4A")),

            ("FONTNAME", (0, 0), (-1, -1),
             "Helvetica-Bold"),

            ("GRID", (0, 0), (-1, -1), 0.5,
             colors.grey),

            ("PADDING", (0, 0), (-1, -1), 7),
        ])
    )

    elements.append(course_table)
    elements.append(Spacer(1, 20))

    # ------------------------------------------------------
    # FINAL MATRIX
    # IMPORTANT: DO NOT GENERATE EXTRA TOTAL ROWS
    # ------------------------------------------------------

    columns = [
        "Student Name",
        "Admission Number",
        "Course Code",
        "Assessment",
        "CO1",
        "CO2",
        "CO3",
        "CO4",
        "TOTAL"
    ]

    pdf_data = [columns]

    # Use the existing final matrix directly
    for _, row in final_matrix.iterrows():

        pdf_data.append([
            str(row.get("Student Name", "")),
            str(row.get("Admission Number", "")),
            str(row.get("Course Code", "")),
            str(row.get("Assessment", "")),
            f'{float(row.get("CO1", 0) or 0):.2f}',
            f'{float(row.get("CO2", 0) or 0):.2f}',
            f'{float(row.get("CO3", 0) or 0):.2f}',
            f'{float(row.get("CO4", 0) or 0):.2f}',
            f'{float(row.get("TOTAL", 0) or 0):.2f}'
        ])

    # ------------------------------------------------------
    # TABLE
    # ------------------------------------------------------

    matrix_table = Table(
        pdf_data,
        repeatRows=1,
        colWidths=[
            105, 85, 105, 65,
            55, 55, 55, 55, 65
        ]
    )

    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0),
         colors.HexColor("#0F766E")),

        ("TEXTCOLOR", (0, 0), (-1, 0),
         colors.white),

        ("FONTNAME", (0, 0), (-1, 0),
         "Helvetica-Bold"),

        ("ALIGN", (0, 0), (-1, -1), "CENTER"),

        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        ("GRID", (0, 0), (-1, -1), 0.5,
         colors.HexColor("#99F6E4")),

        ("FONTSIZE", (0, 0), (-1, -1), 8),

        ("TOPPADDING", (0, 0), (-1, -1), 7),

        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]

    # ------------------------------------------------------
    # HIGHLIGHT TOTAL ROWS
    # ------------------------------------------------------

    for row_index in range(1, len(pdf_data)):

        if pdf_data[row_index][3] == "TOTAL":

            table_style.append(
                (
                    "BACKGROUND",
                    (0, row_index),
                    (-1, row_index),
                    colors.HexColor("#CCFBF1")
                )
            )

            table_style.append(
                (
                    "FONTNAME",
                    (0, row_index),
                    (-1, row_index),
                    "Helvetica-Bold"
                )
            )

    # ------------------------------------------------------
    # MERGE STUDENT DETAILS ACROSS FOUR ASSESSMENT ROWS
    # ------------------------------------------------------

    # Each student has five rows:
    # PA1, PA2, LE1, LM, TOTAL

    start_row = 1

    while start_row + 4 < len(pdf_data):

        # Merge only if this is a PA1 row
        if pdf_data[start_row][3] == "PA1":

            for column in range(3):

                table_style.append(
                    (
                        "SPAN",
                        (column, start_row),
                        (column, start_row + 3)
                    )
                )

        start_row += 5

    matrix_table.setStyle(
        TableStyle(table_style)
    )

    elements.append(matrix_table)

    # ------------------------------------------------------
    # BUILD PDF
    # ------------------------------------------------------

    doc.build(elements)

    buffer.seek(0)

    return buffer.getvalue()
# ==========================================================
# 18. DASHBOARD
# ==========================================================

st.title("📊 CO Assessment Dashboard")

st.markdown(
    "Upload assessment data, configure rescaling and clubbing, "
    "retrieve Phase 0 allocation, and generate the final "
    "weighted CO assessment matrix."
)

st.divider()


# ==========================================================
# 19. UPLOAD DATA
# ==========================================================

st.subheader("01 | Upload Assessment Data")

uploaded_file = st.file_uploader(
    "Upload your Phase 1 Excel or CSV file",
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is None:

    st.info("Upload your assessment dataset to begin.")

    st.stop()

try:

    raw_data = read_uploaded_file(uploaded_file)

    df, identity_columns = prepare_data(raw_data)

except Exception as exc:

    st.error(
        f"Unable to prepare the uploaded data: {exc}"
    )

    st.stop()

st.success(
    "Dataset uploaded and checked successfully."
)

st.divider()


# ==========================================================
# 20. COURSE SELECTION
# ==========================================================

st.subheader("02 | Course Selection")

course_lookup = (
    df[
        ["Course Name", "Course Code", "Course Type"]
    ]
    .drop_duplicates()
    .copy()
)

course_lookup["_label"] = course_lookup.apply(
    lambda r: (
        f"{r['Course Name']} | "
        f"{r['Course Code']} | "
        f"{r['Course Type']}"
    ),
    axis=1
)

selected_label = st.selectbox(
    "Select Course",
    course_lookup["_label"].tolist(),
    key="selected_course_label"
)

selected = course_lookup.loc[
    course_lookup["_label"] == selected_label
].iloc[0]

course_name = selected["Course Name"]
course_code = selected["Course Code"]
course_type = selected["Course Type"]

course_data = df[
    (df["Course Name"] == course_name)
    & (df["Course Code"] == course_code)
    & (df["Course Type"] == course_type)
].copy().reset_index(drop=True)


# ==========================================================
# 21. COURSE DETAILS
# ==========================================================

detail_col1, detail_col2, detail_col3 = st.columns(3)

detail_col1.markdown(
    f"**Course Name:** {course_name}"
)

detail_col2.markdown(
    f"**Course Code:** {course_code}"
)

detail_col3.markdown(
    f"**Course Type:** {course_type}"
)

st.divider()


# ==========================================================
# 22. FETCH PHASE 0 ALLOCATION
# ==========================================================

st.subheader("03 | Phase 0 Allocation")

allocation_weights = None

allocation_df, allocation_error = fetch_phase0_allocation(
    str(course_code).strip()
)

if allocation_error:

    st.warning(
        "Phase 0 allocation could not be retrieved."
    )

    st.caption(allocation_error)

    st.info(
        "You can continue configuring Phase 2, but the "
        "final weighted CO calculation and PDF download "
        "will require a valid Phase 0 allocation."
    )

else:

    try:

        allocation_weights = process_phase0_allocation(
            allocation_df
        )

        st.success(
            "Phase 0 allocation successfully retrieved."
        )

       
    except Exception as exc:

        st.error(
            f"Unable to process Phase 0 allocation: {exc}"
        )

st.divider()


# ==========================================================
# 23. ND REPLACEMENT
# ==========================================================

st.subheader("04 | ND Replacement")

nd_config = build_nd_replacement_table(course_data)

nd_replacement_signature = []

if not nd_config.empty:

    st.caption(
        "Select the replacement assessment for every ND record. "
        "The original CO is retained automatically."
    )

    edited_nd_config = st.data_editor(
        nd_config,
        use_container_width=True,
        hide_index=True,
        disabled=[
            "Record Index",
            "Date",
            "ND CO",
            "Mark Obtained"
        ],
        column_config={
            "Record Index": None,

            "Replace With": st.column_config.SelectboxColumn(
                "Replace Assessment",
                options=ND_REPLACEMENT_OPTIONS,
                required=True
            )
        },
        key=f"nd_replacement_{safe_key(course_code)}"
    )

    nd_replacement_signature = (
        edited_nd_config[
            ["Record Index", "Replace With"]
        ].to_dict("records")
    )

    not_replaced = edited_nd_config[
        edited_nd_config["Replace With"] == "Do not change"
    ]

    if not_replaced.shape[0] > 0:

        st.warning(
            f"{len(not_replaced)} ND record(s) are still not replaced. "
            "Complete every ND replacement to continue."
        )

        st.stop()

    course_data = apply_nd_replacements(
        course_data,
        edited_nd_config
    )

    st.success(
        "All ND records have been replaced successfully."
    )

else:

    st.info(
        "No ND records were found for this course."
    )


course_data = course_data[
    course_data["Assessment"].isin(ASSESSMENT_ORDER)
].copy()

if course_data.empty:

    st.warning(
        "No PA1, PA2, LE or LM records are available "
        "for calculation."
    )

    st.stop()

st.divider()


# ==========================================================
# 24. RESCALING AND CO CALCULATION SETTINGS
# ==========================================================

st.subheader(
    "05 | Rescaling and CO Calculation Settings"
)

st.caption(
    "Rescaling is configured separately for each assessment, "
    "date and CO. Clubbing is configured only when the same CO "
    "occurs on different dates within the same assessment."
)

settings_frame = get_settings_rows(course_data)

if settings_frame.empty:

    st.warning(
        "No assessment and CO combinations are available."
    )

    st.stop()


# ==========================================================
# 25. REPEATED CO DETECTION
# ==========================================================

repeated_summary = get_repeated_co_summary(course_data)

repeated_pairs = set(repeated_summary.keys())

repeated_pair_dates = {}

for assessment, co in repeated_pairs:

    pair_rows = settings_frame[
        (settings_frame["Assessment"] == assessment)
        & (settings_frame["CO"] == co)
    ]

    repeated_pair_dates[
        (assessment, co)
    ] = pair_rows["Date Label"].tolist()


# ==========================================================
# 26. SETTINGS DICTIONARIES
# ==========================================================

config_map = {}
clubbing_map = {}
best_n_map = {}
co_weights = {}

shown_clubbing = set()


# ==========================================================
# 27. SETTINGS TABLE HEADER
# ==========================================================

header_cols = st.columns(
    [0.85, 1.15, 0.55, 1.15, 1.55, 2.4]
)

header_cols[0].markdown("**Assessment**")
header_cols[1].markdown("**Date**")
header_cols[2].markdown("**CO**")
header_cols[3].markdown("**Rescaling**")

st.markdown("<hr>", unsafe_allow_html=True)


# ==========================================================
# 28. SETTINGS ROWS
# ==========================================================

for _, setting_row in settings_frame.iterrows():

    assessment = setting_row["Assessment"]
    dlabel = setting_row["Date Label"]
    co = setting_row["CO"]

    combination = (assessment, co)

    row_key = safe_key(
        f"{course_code}_{assessment}_{dlabel}_{co}"
    )

    cols = st.columns(
        [0.85, 1.15, 0.55, 1.15, 1.55, 2.4]
    )

    cols[0].write(assessment)
    cols[1].write(dlabel)
    cols[2].write(co)

    # ------------------------------------------------------
    # RESCALING
    # ------------------------------------------------------

    rescaling = cols[3].selectbox(
        "Rescaling",
        RESCALING_OPTIONS,
        index=0,
        key=f"rescale_{row_key}",
        label_visibility="collapsed"
    )

    minimum_mark = 0.0

    # ------------------------------------------------------
    # MODERATION MINIMUM MARK
    # ------------------------------------------------------

    if rescaling == "Moderate":

        minimum_mark = cols[4].number_input(
            "Minimum Mark",
            min_value=0.0,
            value=0.0,
            step=0.5,
            key=f"minimum_mark_{row_key}",
            help=(
                "Marks below this value will be replaced "
                "by this minimum mark before rescaling."
            )
        )

    config_map[
        (assessment, dlabel, co)
    ] = {
        "Rescaling Method": rescaling,
        "Minimum Mark": float(minimum_mark)
    }

    # ------------------------------------------------------
    # CLUBBING FOR REPEATED CO
    # ------------------------------------------------------

    if combination in repeated_pairs:

        if combination not in shown_clubbing:

            pair_dates = repeated_pair_dates[combination]

            pair_key = safe_key(
                f"{course_code}_{assessment}_{co}_"
                + "_".join(pair_dates)
            )

            clubbing = cols[4].selectbox(
                "Clubbing",
                CLUBBING_OPTIONS,
                index=0,
                key=f"clubbing_{pair_key}",
                label_visibility="collapsed"
            )

            clubbing_map[combination] = clubbing

            shown_clubbing.add(combination)

            # ----------------------------------------------
            # AVERAGE OF BEST
            # ----------------------------------------------

            if clubbing == "Average of Best":

                best_n = cols[5].number_input(
                    "Best N",
                    min_value=1,
                    max_value=len(pair_dates),
                    value=1,
                    step=1,
                    key=f"best_n_{pair_key}",
                    label_visibility="collapsed"
                )

                best_n_map[combination] = int(best_n)

            # ----------------------------------------------
            # WEIGHTED AVERAGE
            # ----------------------------------------------

            elif clubbing == "Weighted Average":

                with cols[5]:

                    weight_cols = st.columns(
                        len(pair_dates)
                    )

                    for i, entry_date in enumerate(pair_dates):

                        with weight_cols[i]:

                            st.caption(
                                f"Entry {i + 1}"
                            )

                            weight = st.number_input(
                                "Weight",
                                min_value=0.0,
                                value=1.0,
                                step=1.0,
                                key=f"weight_{pair_key}_{i}",
                                label_visibility="collapsed",
                                help=f"Weight for {co} on {entry_date}"
                            )

                        co_weights[
                            (assessment, co, entry_date)
                        ] = weight

    st.markdown("<hr>", unsafe_allow_html=True)


# ==========================================================
# 29. VALIDATE CLUBBING
# ==========================================================

missing_clubbing = []
invalid_best_n = []
invalid_weights = []

for assessment, co in repeated_pairs:

    combination = (assessment, co)

    method = clubbing_map.get(combination)

    if method is None:

        missing_clubbing.append(
            f"{assessment} - {co}"
        )

    elif method == "Average of Best":

        best_n = best_n_map.get(
            combination,
            0
        )

        maximum = len(
            repeated_pair_dates[combination]
        )

        if best_n < 1 or best_n > maximum:

            invalid_best_n.append(
                f"{assessment} - {co}"
            )

    elif method == "Weighted Average":

        pair_dates = repeated_pair_dates[combination]

        total_weight = sum(
            co_weights.get(
                (assessment, co, dlabel),
                0.0
            )
            for dlabel in pair_dates
        )

        if total_weight <= 0:

            invalid_weights.append(
                f"{assessment} - {co}"
            )


# ==========================================================
# 30. GENERATE ASSESSMENT TABLE
# ==========================================================

st.divider()

st.subheader("06 | Generate Assessment Table")


# ==========================================================
# 31. SETTINGS SIGNATURE
# ==========================================================

settings_signature = [
    {
        "key": list(k),
        **v
    }
    for k, v in config_map.items()
]

signature = json.dumps(
    {
        "course": selected_label,

        "settings": settings_signature,

        "clubbing": {
            str(k): v
            for k, v in clubbing_map.items()
        },

        "best_n": {
            str(k): v
            for k, v in best_n_map.items()
        },

        "weights": {
            str(k): v
            for k, v in co_weights.items()
        },

        "nd_choices": nd_replacement_signature

    },
    sort_keys=True,
    default=str
)


if (
    st.session_state.get("generated_signature")
    != signature
):

    st.session_state.pop(
        "generated_table",
        None
    )

    st.session_state.pop(
        "generated_excel",
        None
    )

    st.session_state.pop(
        "generated_signature",
        None
    )

    st.session_state.pop(
        "edited_assessment_matrix",
        None
    )


# ==========================================================
# 32. GENERATE BUTTON
# ==========================================================

if st.button(
    "Generate Assessment Table",
    type="primary",
    use_container_width=True
):

    if missing_clubbing:

        st.error(
            "Select a clubbing method for: "
            + ", ".join(missing_clubbing)
        )

    elif invalid_best_n:

        st.error(
            "Check Best N for: "
            + ", ".join(invalid_best_n)
        )

    elif invalid_weights:

        st.error(
            "The total weight must be greater than zero for: "
            + ", ".join(invalid_weights)
        )

    else:

        try:

            with st.spinner(
                "Calculating assessment values..."
            ):

                results = calculate_results(
                    course_data,
                    identity_columns,
                    config_map
                )

                if results.empty:

                    st.error(
                        "No calculation results were produced."
                    )

                    st.stop()

                output = make_output_table(
                    results,
                    identity_columns,
                    repeated_pairs,
                    clubbing_map,
                    best_n_map,
                    co_weights
                )

                st.session_state["generated_table"] = (
                    output.copy()
                )

                st.session_state["generated_excel"] = (
                    create_excel_file(output)
                )

                st.session_state["generated_signature"] = (
                    signature
                )

                st.success(
                    "Assessment table generated successfully."
                )

        except Exception as exc:

            st.error(
                f"Calculation error: {exc}"
            )


# ==========================================================
# ==========================================================
# 33. STUDENT ASSESSMENT DISPLAY MATRIX
# ==========================================================

if (
    "generated_table" in st.session_state
    and st.session_state.get("generated_signature") == signature
):

    st.divider()

    generated_table = st.session_state[
        "generated_table"
    ].copy()

    # ------------------------------------------------------
    # STUDENT DETAILS
    # ------------------------------------------------------

    student_key = "__StudentKey"

    student_name_candidates = [
        "Student Name",
        "Name",
        "Student"
    ]

    admission_candidates = [
        "Admission Number",
        "Admission No",
        "Admission No.",
        "Admission Number."
    ]

    component_candidates = [
        "Component (T/P)",
        "Component",
        "Component Type"
    ]

    programme_candidates = [
        "Student Prog",
        "Student Program",
        "Programme",
        "Program"
    ]

    student_name_column = next(
        (
            col for col in student_name_candidates
            if col in course_data.columns
        ),
        None
    )

    admission_column = next(
        (
            col for col in admission_candidates
            if col in course_data.columns
        ),
        None
    )

    component_column = next(
        (
            col for col in component_candidates
            if col in course_data.columns
        ),
        None
    )

    programme_column = next(
        (
            col for col in programme_candidates
            if col in course_data.columns
        ),
        None
    )

    # ------------------------------------------------------
    # CREATE STUDENT DETAILS TABLE
    # ------------------------------------------------------

    detail_columns = [student_key]

    for col in [
        student_name_column,
        admission_column,
        component_column,
        programme_column
    ]:

        if col and col not in detail_columns:
            detail_columns.append(col)

    student_details = (
        course_data[detail_columns]
        .drop_duplicates(subset=[student_key])
        .copy()
    )

    rename_columns = {}

    if student_name_column:
        rename_columns[student_name_column] = "Name of Student"

    if admission_column:
        rename_columns[admission_column] = "Admission Number"

    if component_column:
        rename_columns[component_column] = "Component (T/P)"

    if programme_column:
        rename_columns[programme_column] = "Programme"

    student_details = student_details.rename(
        columns=rename_columns
    )

    # ------------------------------------------------------
    # MERGE STUDENT DETAILS WITH GENERATED MATRIX
    # ------------------------------------------------------

    if student_key in generated_table.columns:

        display_matrix = generated_table.merge(
            student_details,
            on=student_key,
            how="left"
        )

    else:

        display_matrix = generated_table.copy()

    # ------------------------------------------------------
    # ADD MISSING STUDENT DETAILS
    # ------------------------------------------------------

    for col in [
        "Name of Student",
        "Admission Number",
        "Component (T/P)",
        "Programme"
    ]:

        if col not in display_matrix.columns:
            display_matrix[col] = ""

    # ------------------------------------------------------
    # ADD COURSE DETAILS
    # ------------------------------------------------------

    display_matrix["Course Name"] = course_name
    display_matrix["Course Code"] = course_code

    if "Programme" not in display_matrix.columns:
        display_matrix["Programme"] = course_type

    # ------------------------------------------------------
    # REMOVE UNWANTED COLUMNS
    # ------------------------------------------------------

    display_matrix = display_matrix.drop(
        columns=[
            "__StudentKey",
            "Student ID",
            "Student Id",
            "StudentID",
            "Student Key",
            "Student Name",
            "Course Type",
            "PA1",
            "PA2",
            "LE",
            "LM",
            "CO1",
            "CO2",
            "CO3",
            "CO4"
        ],
        errors="ignore"
    )

    # ------------------------------------------------------
    # REQUIRED ASSESSMENT COLUMNS
    # ------------------------------------------------------

    assessment_co_columns = [
        f"{assessment}_{co}"
        for assessment in ASSESSMENT_ORDER
        for co in CO_LIST
    ]

    for col in assessment_co_columns:

        if col not in display_matrix.columns:
            display_matrix[col] = 0.0

    # ------------------------------------------------------
    # FINAL COLUMN ORDER
    # ------------------------------------------------------

    identity_display_columns = [
        "Name of Student",
        "Admission Number",
        "Component (T/P)",
        "Programme",
        "Course Name",
        "Course Code"
    ]

    display_matrix = display_matrix[
        identity_display_columns
        + assessment_co_columns
    ].copy()

    display_matrix = display_matrix.reset_index(
        drop=True
    )

    # ------------------------------------------------------
    # KEEP GENERATED MATRIX FOR CALCULATIONS
    # ------------------------------------------------------

    edited_matrix = generated_table.copy()

    # Add student details to calculation matrix
    if student_key in edited_matrix.columns:

        edited_matrix = edited_matrix.merge(
            student_details,
            on=student_key,
            how="left"
        )

    edited_matrix = edited_matrix.rename(
        columns={
            "Name of Student": "Student Name"
        }
    )

    # Ensure required student columns exist
    if "Student Name" not in edited_matrix.columns:

        if "Student Name" in student_details.columns:
            edited_matrix["Student Name"] = ""

        elif "Name of Student" in edited_matrix.columns:
            edited_matrix["Student Name"] = (
                edited_matrix["Name of Student"]
            )

        else:
            edited_matrix["Student Name"] = ""

    if "Admission Number" not in edited_matrix.columns:
        edited_matrix["Admission Number"] = ""

    # Store matrix in session state
    st.session_state["edited_matrix"] = (
        edited_matrix.copy()
    )

    # ======================================================
    # 34. ORIGINAL ASSESSMENT EXCEL EXPORT
    # ======================================================

    import io

    from openpyxl.styles import (
        Font,
        PatternFill,
        Alignment,
        Border,
        Side
    )

    from openpyxl.utils import get_column_letter

    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:

        display_matrix.to_excel(
            writer,
            sheet_name="Assessment Matrix",
            index=False,
            startrow=5
        )

        worksheet = writer.sheets["Assessment Matrix"]

        # --------------------------------------------------
        # COURSE INFORMATION
        # --------------------------------------------------

        worksheet["A1"] = "COURSE NAME"
        worksheet["B1"] = course_name

        worksheet["A2"] = "COURSE CODE"
        worksheet["B2"] = course_code

        worksheet["A3"] = "COURSE TYPE"
        worksheet["B3"] = course_type

        teal_dark = "064E4A"
        teal_primary = "0F766E"
        teal_light = "CCFBF1"
        white = "FFFFFF"

        header_fill = PatternFill(
            "solid",
            fgColor=teal_primary
        )

        label_fill = PatternFill(
            "solid",
            fgColor=teal_dark
        )

        value_fill = PatternFill(
            "solid",
            fgColor=teal_light
        )

        thin_border = Border(
            left=Side(style="thin", color="99F6E4"),
            right=Side(style="thin", color="99F6E4"),
            top=Side(style="thin", color="99F6E4"),
            bottom=Side(style="thin", color="99F6E4")
        )

        for row in range(1, 4):

            label_cell = worksheet.cell(
                row=row,
                column=1
            )

            value_cell = worksheet.cell(
                row=row,
                column=2
            )

            label_cell.font = Font(
                bold=True,
                color=white
            )

            label_cell.fill = label_fill

            value_cell.font = Font(
                bold=True,
                color=teal_dark
            )

            value_cell.fill = value_fill

            label_cell.border = thin_border
            value_cell.border = thin_border

        # --------------------------------------------------
        # HEADER FORMATTING
        # --------------------------------------------------

        for cell in worksheet[6]:

            cell.font = Font(
                bold=True,
                color=white
            )

            cell.fill = header_fill

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = thin_border

        worksheet.row_dimensions[6].height = 35

        # --------------------------------------------------
        # DATA FORMATTING
        # --------------------------------------------------

        for row in worksheet.iter_rows(
            min_row=7,
            max_row=worksheet.max_row
        ):

            for cell in row:

                cell.font = Font(
                    color=teal_dark
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

                cell.border = thin_border

                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"

        # --------------------------------------------------
        # COLUMN WIDTHS
        # --------------------------------------------------

        worksheet.column_dimensions["A"].width = 25
        worksheet.column_dimensions["B"].width = 22
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 18
        worksheet.column_dimensions["E"].width = 30
        worksheet.column_dimensions["F"].width = 22

        for col_index in range(
            7,
            len(display_matrix.columns) + 1
        ):

            worksheet.column_dimensions[
                get_column_letter(col_index)
            ].width = 15

        # --------------------------------------------------
        # FREEZE AND FILTER
        # --------------------------------------------------

        worksheet.freeze_panes = "G7"

        last_column = get_column_letter(
            len(display_matrix.columns)
        )

        last_row = len(display_matrix) + 6

        worksheet.auto_filter.ref = (
            f"A6:{last_column}{last_row}"
        )

        worksheet.sheet_view.showGridLines = False

        worksheet.page_setup.orientation = "landscape"
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0

    excel_buffer.seek(0)

    st.download_button(
        "📥 Download Excel File",
        data=excel_buffer.getvalue(),
        file_name=f"CO_Assessment_{course_code}.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    # ======================================================
    # 35. DISPLAY COMBINED ASSESSMENT MATRIX
    # ======================================================

    st.divider()

    st.subheader(
        "07 | Student Assessment and Final Assessment Matrix"
    )

    # ------------------------------------------------------
    # CHECK PHASE 0 ALLOCATION
    # ------------------------------------------------------

    if allocation_weights is None:

        st.warning(
            "Final Assessment cannot be calculated because "
            "the Phase 0 allocation is unavailable. "
            "Check the Google Sheets connection and allocation data."
        )

    else:

        try:

            # --------------------------------------------------
            # WEIGHTED CO CALCULATION
            # --------------------------------------------------

            weighted_matrix = calculate_weighted_cca(
                edited_matrix,
                allocation_weights
            )

            # --------------------------------------------------
            # FINAL WEIGHTED CO VALUES
            # --------------------------------------------------

            final_totals = pd.DataFrame({
                "Final CO1": weighted_matrix["CO1"].fillna(0).round(2),
                "Final CO2": weighted_matrix["CO2"].fillna(0).round(2),
                "Final CO3": weighted_matrix["CO3"].fillna(0).round(2),
                "Final CO4": weighted_matrix["CO4"].fillna(0).round(2)
            })

            display_matrix = display_matrix.reset_index(
                drop=True
            )

            final_totals = final_totals.reset_index(
                drop=True
            )

            display_matrix = pd.concat(
                [
                    display_matrix,
                    final_totals
                ],
                axis=1
            )

            # --------------------------------------------------
            # DISPLAY
            # --------------------------------------------------

            st.dataframe(
                display_matrix,
                use_container_width=True,
                hide_index=True
            )

            # --------------------------------------------------
            # CREATE FINAL WEIGHTED OUTPUT
            # --------------------------------------------------

            final_rows = []

            for _, student in weighted_matrix.iterrows():

                student_co_totals = {
                    co: 0.0 for co in CO_LIST
                }

                for assessment_index, assessment in enumerate(
                    ASSESSMENT_ORDER
                ):

                    row = {}

                    if assessment_index == 0:

                        row["Student Name"] = student.get(
                            "Student Name", ""
                        )

                        row["Admission Number"] = student.get(
                            "Admission Number", ""
                        )

                        row["Course Code"] = course_code

                    else:

                        row["Student Name"] = ""
                        row["Admission Number"] = ""
                        row["Course Code"] = ""

                    row["Assessment"] = (
                        "LE1" if assessment == "LE" else assessment
                    )

                    assessment_total = 0.0

                    for co in CO_LIST:

                        weighted_column = (
                            f"{assessment}_{co}_Weighted"
                        )

                        value = pd.to_numeric(
                            student.get(weighted_column, 0),
                            errors="coerce"
                        )

                        if pd.isna(value):
                            value = 0.0

                        value = float(value)

                        row[co] = round(value, 2)

                        assessment_total += value

                        student_co_totals[co] += value

                    row["TOTAL"] = round(
                        assessment_total,
                        2
                    )

                    final_rows.append(row)

                # --------------------------------------------------
                # ONE TOTAL ROW PER STUDENT
                # --------------------------------------------------

                total_row = {
                    "Student Name": "",
                    "Admission Number": "",
                    "Course Code": "",
                    "Assessment": "TOTAL"
                }

                for co in CO_LIST:

                    total_row[co] = round(
                        student_co_totals[co],
                        2
                    )

                total_row["TOTAL"] = round(
                    sum(student_co_totals.values()),
                    2
                )

                final_rows.append(total_row)

            final_weighted_output = pd.DataFrame(
                final_rows,
                columns=[
                    "Student Name",
                    "Admission Number",
                    "Course Code",
                    "Assessment",
                    "CO1",
                    "CO2",
                    "CO3",
                    "CO4",
                    "TOTAL"
                ]
            )

            st.session_state[
                "final_weighted_output"
            ] = final_weighted_output.copy()

            # ==================================================
            # 36. FINAL WEIGHTED CO MATRIX PDF
            # ==================================================

            pdf_file = create_weighted_cca_pdf(
                {
                    "course_name": course_name,
                    "course_code": course_code,
                    "course_type": course_type
                },
                final_weighted_output
            )

            st.download_button(
                label="📄 Download Final Weighted CO Matrix PDF",
                data=pdf_file,
                file_name=f"Weighted_CO_Matrix_{course_code}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        except Exception as exc:

            st.error(
                f"Unable to generate final weighted assessment: {exc}"
            )
