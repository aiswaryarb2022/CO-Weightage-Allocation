import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


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


# ==========================================================
# 2. GREEN THEME – TEXT VISIBILITY FIX
# ==========================================================

st.markdown("""
<style>

/* Main application background */
.stApp {
    background-color: #F5F9F5;
    color: #222222 !important;
}

/* General text visibility */
html, body, [class*="css"] {
    color: #222222;
}

/* Headings */
h1, h2, h3, h4, h5, h6 {
    color: #185C37 !important;
}

/* Paragraphs, labels and captions */
p, label, span, div,
.stMarkdown,
.stCaption,
[data-testid="stWidgetLabel"] {
    color: #222222;
}

/* Input fields */
.stTextInput input,
.stNumberInput input,
.stDateInput input,
.stSelectbox div[data-baseweb="select"] {
    color: #222222 !important;
    background-color: #FFFFFF !important;
}

/* Dropdown text */
[data-baseweb="select"] * {
    color: #222222 !important;
}

/* Buttons */
.stButton > button,
.stDownloadButton > button {
    background-color: #21864B !important;
    color: #FFFFFF !important;
    border-radius: 8px;
    border: none;
    font-weight: 600;
}

/* Button hover */
.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #176B39 !important;
    color: #FFFFFF !important;
}

/* Data tables */
[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
    color: #222222 !important;
}

/* Horizontal line */
hr {
    border: none;
    border-top: 1px solid #DCE8DE;
    margin-top: 5px;
    margin-bottom: 5px;
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
            "Course Name", "Course", "Subject Name"
        ],
        "Course Code": [
            "Course Code", "Subject Code"
        ],
        "Course Type": [
            "Course Type", "Course Category", "Category"
        ],
        "Assessment": [
            "Assessment", "Assessment Type", "Exam"
        ],
        "CO": [
            "CO", "Course Outcome", "Course Outcomes"
        ],
        "Mark Obtained": [
            "Mark Obtained", "Marks Obtained", "Mark", "Marks"
        ],
        "Max Marks": [
            "Max Marks", "Maximum Marks",
            "Total Marks", "Maximum Mark"
        ],
        "Date": [
            "Date", "Assessment Date", "Exam Date"
        ],
        "Reg No": [
            "Reg No", "Register Number",
            "Registration Number", "Register No", "Roll No"
        ],
        "Admission No": [
            "Admission No", "Admission Number"
        ],
        "Student ID": [
            "Student ID", "Student Number"
        ],
        "Student Name": [
            "Student Name", "Name"
        ],
        "Student Prog": [
            "Student Prog", "Student Program", "Programme"
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
        "", "nan", "none", "null",
        "unknown", "na", "n/a"
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

                if rescaled < 1:
                    rescaled = 1

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
        # The minimum mark is used when the obtained mark
        # is below the teacher-entered minimum mark.
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

    # ------------------------------------------------------
    # NON-REPEATED CO
    # No clubbing is needed.
    # ------------------------------------------------------

    if combination not in repeated_pairs:

        return float(date_values.mean())

    # ------------------------------------------------------
    # REPEATED CO
    # Use the single selected clubbing method.
    # ------------------------------------------------------

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
# 10. CREATE FINAL ASSESSMENT MATRIX
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
            "Course Name": base["Course Name"],
            "Course Code": base["Course Code"]
        }

        # --------------------------------------------------
        # ASSESSMENT-WISE VALUES
        # --------------------------------------------------

        for assessment in ASSESSMENT_ORDER:

            co_scores = []

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

                if pd.notna(score):
                    co_scores.append(score)

            row[assessment] = (
                float(np.mean(co_scores))
                if co_scores
                else np.nan
            )

        # --------------------------------------------------
        # CO-WISE VALUES
        # --------------------------------------------------

        for co in CO_LIST:

            assessment_scores = []

            for assessment in ASSESSMENT_ORDER:

                score = aggregate_co_score(
                    group,
                    assessment,
                    co,
                    repeated_pairs,
                    clubbing_map,
                    best_n_map,
                    co_weights
                )

                if pd.notna(score):
                    assessment_scores.append(score)

            row[co] = (
                float(np.mean(assessment_scores))
                if assessment_scores
                else np.nan
            )

        output_rows.append(row)

    output = pd.DataFrame(output_rows)

    ordered = (
        ["Course Name", "Course Code"]
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
# 12. DASHBOARD
# ==========================================================

st.title("📊 CO Assessment Dashboard")

st.markdown(
    "Upload assessment data, replace ND assessments, "
    "select rescaling and clubbing methods, and generate "
    "the final CO assessment matrix."
)

st.divider()


# ==========================================================
# 13. UPLOAD DATA
# ==========================================================

st.subheader("01 | Upload Assessment Data")

uploaded_file = st.file_uploader(
    "Upload your Excel or CSV file",
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
# 14. COURSE SELECTION
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
    course_lookup["_label"].tolist()
)

selected = course_lookup.loc[
    course_lookup["_label"] == selected_label
].iloc[0]

course_name = selected["Course Name"]
course_code = selected["Course Code"]
course_type = selected["Course Type"]

col1, col2 = st.columns(2)

col1.markdown(
    f"**Course Type:** {course_type}"
)

col2.markdown(
    f"**Course Code:** {course_code}"
)

course_data = df[
    (df["Course Name"] == course_name)
    & (df["Course Code"] == course_code)
    & (df["Course Type"] == course_type)
].copy().reset_index(drop=True)

st.divider()


# ==========================================================
# 15. ND REPLACEMENT
# ==========================================================

st.subheader("03 | ND Replacement")

nd_config = build_nd_replacement_table(course_data)

nd_replacement_signature = []

if not nd_config.empty:

    st.caption(
        "Select the replacement assessment for every ND record. "
        "The original CO is retained automatically."
    )

    st.markdown("### Available ND Assessments")

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

    st.info("No ND records were found for this course.")

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
# 16. RESCALING AND CO CALCULATION SETTINGS
# ==========================================================

st.subheader(
    "04 | Rescaling and CO Calculation Settings"
)

st.caption(
    "Rescaling is selected separately for each assessment, date and CO. "
    "Clubbing is shown only when the same CO occurs on different dates "
    "within the same assessment. One clubbing selection applies to all "
    "those dates."
)

settings_frame = get_settings_rows(course_data)

if settings_frame.empty:

    st.warning(
        "No assessment and CO combinations are available."
    )

    st.stop()


# ----------------------------------------------------------
# REPEATED CO DETECTION
# ----------------------------------------------------------

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


# ----------------------------------------------------------
# SETTINGS DICTIONARIES
# ----------------------------------------------------------

config_map = {}
clubbing_map = {}
best_n_map = {}
co_weights = {}

shown_clubbing = set()


# ----------------------------------------------------------
# TABLE HEADER
# ----------------------------------------------------------

header_cols = st.columns(
    [0.85, 1.15, 0.55, 1.15, 1.55, 2.4]
)

header_cols[0].markdown("**Assessment**")
header_cols[1].markdown("**Date**")
header_cols[2].markdown("**CO**")
header_cols[3].markdown("**Rescaling**")

st.markdown("<hr>", unsafe_allow_html=True)


# ----------------------------------------------------------
# SETTINGS ROWS
# ----------------------------------------------------------

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
    # RESCALING METHOD
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
    # MODERATE MINIMUM MARK
    # Appears only when Moderate is selected.
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
                "by this minimum mark for Moderate rescaling."
            )
        )

    config_map[
        (assessment, dlabel, co)
    ] = {
        "Rescaling Method": rescaling,
        "Minimum Mark": float(minimum_mark)
    }

    # ------------------------------------------------------
    # CLUBBING ONLY FOR REPEATED ASSESSMENT + CO
    # ------------------------------------------------------

    if combination in repeated_pairs:

        # Only one dropdown is displayed for this pair.
        # It applies to all dates for this Assessment + CO.

        if combination not in shown_clubbing:

            pair_dates = repeated_pair_dates[combination]

            pair_key = safe_key(
                f"{course_code}_{assessment}_{co}_"
                + "_".join(pair_dates)
            )

            clubbing = cols[4].selectbox(
                "Clubbing",
                CLUBBING_OPTIONS,
                index=None,
                placeholder="Select clubbing method",
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
                    label_visibility="collapsed",
                    help=(
                        f"Choose how many of the best "
                        f"{len(pair_dates)} date entries "
                        "should be averaged."
                    )
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

                        entry_label = f"Entry {i + 1}"

                        with weight_cols[i]:

                            st.caption(entry_label)

                            weight = st.number_input(
                                "Weight",
                                min_value=0.0,
                                max_value=100.0,
                                value=1.0,
                                step=1.0,
                                key=f"weight_{pair_key}_{i}",
                                label_visibility="collapsed",
                                help=(
                                    f"Weight for {co} on "
                                    f"{entry_date}"
                                )
                            )

                        co_weights[
                            (assessment, co, entry_date)
                        ] = weight

    # No clubbing dropdown for non-repeated CO combinations.
    # No duplicate dropdown for subsequent dates of the same pair.

    st.markdown("<hr>", unsafe_allow_html=True)


# ==========================================================
# 17. VALIDATE CLUBBING AND WEIGHTS
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

        best_n = best_n_map.get(combination, 0)

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
# 18. GENERATE ASSESSMENT TABLE
# ==========================================================

st.divider()

st.subheader("05 | Generate Assessment Table")


# ----------------------------------------------------------
# SETTINGS SIGNATURE
# ----------------------------------------------------------

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
            str(k): v for k, v in clubbing_map.items()
        },
        "best_n": {
            str(k): v for k, v in best_n_map.items()
        },
        "weights": {
            str(k): v for k, v in co_weights.items()
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

    st.session_state.pop("generated_table", None)
    st.session_state.pop("generated_excel", None)
    st.session_state.pop("generated_signature", None)


# ----------------------------------------------------------
# GENERATE BUTTON
# ----------------------------------------------------------

if st.button(
    "Generate Assessment Table",
    type="primary",
    use_container_width=True
):

    if missing_clubbing:

        st.error(
            "Select a clubbing method for these repeated "
            "assessment-CO combinations: "
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
                "Calculating rescaled assessment values and CO scores..."
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

                st.session_state["generated_table"] = output

                st.session_state["generated_excel"] = (
                    create_excel_file(output)
                )

                st.session_state["generated_signature"] = signature

                st.success(
                    "Assessment table generated successfully."
                )

        except Exception as exc:

            st.error(
                f"Calculation error: {exc}"
            )


# ==========================================================
# 19. PREVIEW AND DOWNLOAD
# ==========================================================

if (
    "generated_table" in st.session_state
    and st.session_state.get("generated_signature") == signature
):

    st.divider()

    st.subheader("06 | Generated Assessment Matrix")

    st.caption(
        "PA1, PA2, LE and LM contain assessment-wise values. "
        "CO1, CO2, CO3 and CO4 contain the corresponding CO-wise "
        "averages. Student names and identifiers are not displayed."
    )

    display_table = st.session_state[
        "generated_table"
    ].copy()

    st.dataframe(
        display_table,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "📥 Download Excel File",
        data=st.session_state["generated_excel"],
        file_name="CO_Assessment_Output.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )
