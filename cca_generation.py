import streamlit as st
import pandas as pd
import numpy as np
import re
import json
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

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
    "PA1",
    "PA2",
    "LE",
    "LM"
]

st.markdown("""
<style>
.stApp, [data-testid="stHeader"] {
    background-color: #F5F9F5;
    color: #222222 !important;
}

html, body, [class*="css"] {
    color: #222222 !important;
}

h1, h2, h3, h4, h5, h6 {
    color: #185C37 !important;
}

p, label, span, [data-testid="stWidgetLabel"] {
    color: #222222;
}

.stTextInput input,
.stNumberInput input,
.stDateInput input,
[data-baseweb="select"] * {
    color: #222222 !important;
    background-color: #FFFFFF !important;
}

.stButton > button,
.stDownloadButton > button {
    background: #21864B;
    color: white;
    border-radius: 8px;
    border: 0;
    font-weight: 600;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #176B39;
    color: white;
}

[data-testid="stDataFrame"] {
    border: 1px solid #D5E6D9;
    border-radius: 8px;
}

hr {
    border: none;
    border-top: 1px solid #DCE8DE;
    margin-top: 5px;
    margin-bottom: 5px;
}
</style>
""", unsafe_allow_html=True)


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
        "Course Component": [
            "Course Component", "Component",
            "Theory/Practical", "Theory or Practical"
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

    if "Course Component" not in df.columns:

        course_type_text = (
            df["Course Type"].astype(str).str.lower()
        )

        df["Course Component"] = np.where(
            course_type_text.str.contains(
                "practical",
                na=False
            ),
            "Practical",
            np.where(
                course_type_text.str.contains(
                    "theory",
                    na=False
                ),
                "Theory",
                "Not specified"
            )
        )

    else:

        df["Course Component"] = (
            df["Course Component"]
            .fillna("Not specified")
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

    identity_columns = ["__StudentKey"] + [
        col for col in [
            "Student Name",
            "Reg No",
            "Admission No"
        ]
        if col in df.columns
    ]

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
            "Replace With": "PA1"
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


def get_repeated_co_summary(course_data):

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
            "Course Component",
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

        if method == "Relative":

            if pd.notna(class_avg) and class_avg != 0:

                rescaled = (
                    avg_mark / class_avg
                ) * 20

                if rescaled < 1:
                    rescaled = 1

            else:
                rescaled = np.nan

        elif method == "Absolute":

            if pd.notna(avg_max) and avg_max > 0:

                rescaled = (
                    avg_mark / avg_max
                ) * 100

            else:
                rescaled = np.nan

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
            "Course Component": values["Course Component"],
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
    # ------------------------------------------------------

    if combination not in repeated_pairs:
        return float(date_values.mean())

    # ------------------------------------------------------
    # REPEATED CO
    # ------------------------------------------------------

    method = clubbing_map.get(combination)

    # Average of All
    if method == "Average of All":

        return float(date_values.mean())

    # Average of Best
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

    # Weighted Average
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
# 10. CREATE GENERATED ASSESSMENT MATRIX
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
            "Student Name",
            "Reg No",
            "Admission No",
            "Course Name",
            "Course Code",
            "Course Type",
            "Course Component"
        ]
    )

    keys = [
        col for col in keys
        if col in results.columns
    ]

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

        row = {}

        # --------------------------------------------------
        # STUDENT DETAILS
        # --------------------------------------------------

        row["Student Name"] = base.get(
            "Student Name",
            ""
        )

        row["Reg No"] = base.get(
            "Reg No",
            ""
        )

        row["Admission No"] = base.get(
            "Admission No",
            ""
        )

        # --------------------------------------------------
        # COURSE DETAILS
        # --------------------------------------------------

        row["Course Code"] = base.get(
            "Course Code",
            ""
        )

        row["Course Component"] = base.get(
            "Course Component",
            base.get("Course Type", "")
        )

        row["Course Name"] = base.get(
            "Course Name",
            ""
        )

        # --------------------------------------------------
        # ASSESSMENT-WISE CO VALUES
        # Each assessment has its own CO columns.
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

                column_name = (
                    f"{assessment}_{co}"
                )

                row[column_name] = score

        output_rows.append(row)

    output = pd.DataFrame(output_rows)

    # ------------------------------------------------------
    # COLUMN ORDER
    # ------------------------------------------------------

    ordered_columns = [
        "Student Name",
        "Reg No",
        "Admission No",
        "Course Code",
        "Course Component",
        "Course Name"
    ]

    for assessment in ASSESSMENT_ORDER:

        for co in CO_LIST:

            column_name = f"{assessment}_{co}"

            if column_name in output.columns:
                ordered_columns.append(column_name)

    for col in ordered_columns:

        if col not in output.columns:
            output[col] = np.nan

    output = output[ordered_columns]

    return output


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
            sheet_name="Generated Matrix"
        )

        ws = writer.sheets["Generated Matrix"]

        # --------------------------------------------------
        # HEADER FORMATTING
        # --------------------------------------------------

        header_fill = PatternFill(
            start_color="185C37",
            end_color="185C37",
            fill_type="solid"
        )

        header_font = Font(
            color="FFFFFF",
            bold=True
        )

        for cell in ws[1]:

            cell.fill = header_fill
            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        # --------------------------------------------------
        # ASSESSMENT HEADER COLOURS
        # --------------------------------------------------

        assessment_fills = {
            "PA1": "DCEFE2",
            "PA2": "DCE7F5",
            "LE": "FCE9D5",
            "LM": "E9E0F5"
        }

        for col_index, cell in enumerate(
            ws[1],
            start=1
        ):

            column_name = str(cell.value)

            for assessment, colour in assessment_fills.items():

                if column_name.startswith(
                    assessment + "_"
                ):

                    cell.fill = PatternFill(
                        start_color=colour,
                        end_color=colour,
                        fill_type="solid"
                    )

                    cell.font = Font(
                        color="222222",
                        bold=True
                    )

        # --------------------------------------------------
        # COLUMN WIDTH
        # --------------------------------------------------

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

        # --------------------------------------------------
        # NUMBER FORMATTING
        # --------------------------------------------------

        for row in ws.iter_rows(min_row=2):

            for cell in row:

                if isinstance(
                    cell.value,
                    (int, float)
                ):

                    cell.number_format = "0.00"

        ws.freeze_panes = "G2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 35

    buffer.seek(0)

    return buffer.getvalue()


# ==========================================================
# 12. DASHBOARD TITLE
# ==========================================================

st.title("📊 CO Assessment Dashboard")

st.markdown(
    """
    Upload assessment data, replace ND assessments,
    select rescaling and clubbing methods, and generate
    the final student-wise assessment matrix.
    """
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

    st.info(
        "Upload your assessment dataset to begin."
    )

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
        [
            "Course Name",
            "Course Code",
            "Course Type",
            "Course Component"
        ]
    ]
    .drop_duplicates()
    .copy()
)

course_lookup["_label"] = course_lookup.apply(
    lambda r: (
        f"{r['Course Name']} | "
        f"{r['Course Code']} | "
        f"{r['Course Component']}"
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
course_component = selected["Course Component"]

# ----------------------------------------------------------
# DISPLAY AUTOMATIC COURSE DETAILS
# ----------------------------------------------------------

course_col1, course_col2, course_col3 = st.columns(3)

with course_col1:
    st.text_input(
        "Course Name",
        value=str(course_name),
        disabled=True
    )

with course_col2:
    st.text_input(
        "Course Code",
        value=str(course_code),
        disabled=True
    )

with course_col3:
    st.text_input(
        "Course Component",
        value=str(course_component),
        disabled=True
    )

course_data = df[
    (df["Course Name"] == course_name)
    & (df["Course Code"] == course_code)
    & (df["Course Type"] == course_type)
    & (df["Course Component"] == course_component)
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
        "Select the assessment that should replace each ND record. "
        "PA1 is selected by default. You can change it to PA2, LE or LM."
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

    course_data = apply_nd_replacements(
        course_data,
        edited_nd_config
    )

    st.success(
        "ND replacements applied successfully."
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
# 16. RESCALING AND CLUBBING CONFIGURATION
# ==========================================================

st.subheader("04 | Rescaling and Clubbing Configuration")

st.write(
    "Configure the rescaling and clubbing methods separately "
    "for each assessment and CO."
)

settings_df = get_settings_rows(course_data)

if settings_df.empty:
    st.error("No assessment settings could be generated.")
    st.stop()

repeated_summary = get_repeated_co_summary(course_data)

repeated_pairs = set(repeated_summary.keys())

config_map = {}

clubbing_map = {}
best_n_map = {}
co_weights = {}

st.markdown("### Assessment-wise Configuration")

for assessment in ASSESSMENT_ORDER:

    assessment_settings = settings_df[
        settings_df["Assessment"] == assessment
    ]

    if assessment_settings.empty:
        continue

    with st.expander(
        f"{assessment} Configuration",
        expanded=True
    ):

        for _, setting_row in assessment_settings.iterrows():

            current_assessment = setting_row["Assessment"]
            current_date = setting_row["Date Label"]
            current_co = setting_row["CO"]

            pair = (
                current_assessment,
                current_co
            )

            widget_key = (
                f"{safe_key(current_assessment)}_"
                f"{safe_key(current_date)}_"
                f"{safe_key(current_co)}"
            )

            st.markdown(
                f"#### {current_assessment} | "
                f"{current_co} | {current_date}"
            )

            col1, col2 = st.columns(2)

            with col1:

                rescaling_method = st.selectbox(
                    "Rescaling Method",
                    [
                        "Absolute",
                        "Relative",
                        "Moderation"
                    ],
                    index=0,
                    key=f"rescaling_{widget_key}"
                )

            with col2:

                minimum_mark = 0.0

                if rescaling_method == "Moderation":

                    minimum_mark = st.number_input(
                        "Minimum Pass Mark",
                        min_value=0.0,
                        max_value=100.0,
                        value=40.0,
                        step=1.0,
                        key=f"minimum_{widget_key}",
                        help=(
                            "Marks below this percentage "
                            "will be raised to the minimum."
                        )
                    )

            config_map[
                (
                    current_assessment,
                    current_date,
                    current_co
                )
            ] = {
                "Rescaling Method": rescaling_method,
                "Minimum Mark": minimum_mark
            }

            # ------------------------------------------------
            # CLUBBING CONTROLS
            # ------------------------------------------------

            if pair in repeated_pairs:

                st.markdown(
                    f"**Clubbing method for {current_assessment} "
                    f"{current_co}**"
                )

                clubbing_method = st.selectbox(
                    "Select Clubbing Method",
                    CLUBBING_OPTIONS,
                    index=0,
                    key=f"clubbing_{widget_key}"
                )

                clubbing_map[pair] = clubbing_method

                if clubbing_method == "Average of Best":

                    total_entries = repeated_summary[pair]

                    best_n = st.number_input(
                        "How many best entries?",
                        min_value=1,
                        max_value=total_entries,
                        value=1,
                        step=1,
                        key=f"bestn_{widget_key}"
                    )

                    best_n_map[pair] = int(best_n)

                elif clubbing_method == "Weighted Average":

                    st.caption(
                        "Enter a weight for each date/entry. "
                        "Weights will be normalized automatically."
                    )

                    pair_settings = assessment_settings[
                        assessment_settings["CO"] == current_co
                    ]

                    weight_total = 0.0

                    for _, weight_row in pair_settings.iterrows():

                        weight_date = weight_row["Date Label"]

                        weight_key = (
                            f"weight_{safe_key(current_assessment)}_"
                            f"{safe_key(weight_date)}_"
                            f"{safe_key(current_co)}"
                        )

                        weight_value = st.number_input(
                            f"Weight for {weight_date}",
                            min_value=0.0,
                            value=1.0,
                            step=0.5,
                            key=weight_key
                        )

                        co_weights[
                            (
                                current_assessment,
                                current_co,
                                weight_date
                            )
                        ] = float(weight_value)

                        weight_total += float(weight_value)

                    if weight_total <= 0:

                        st.warning(
                            "At least one weight must be greater than zero."
                        )

                st.markdown("---")


# ==========================================================
# 17. VALIDATE CONFIGURATION
# ==========================================================

st.markdown("### Configuration Summary")

configuration_rows = []

for (
    assessment,
    date_label,
    co
), settings in config_map.items():

    pair = (assessment, co)

    configuration_rows.append({
        "Assessment": assessment,
        "Date": date_label,
        "CO": co,
        "Rescaling Method": settings["Rescaling Method"],
        "Minimum Pass Mark": settings["Minimum Mark"],
        "Clubbing Method": clubbing_map.get(
            pair,
            "Not Required"
        ),
        "Best N": best_n_map.get(
            pair,
            "-"
        )
    })

configuration_table = pd.DataFrame(
    configuration_rows
)

st.dataframe(
    configuration_table,
    use_container_width=True,
    hide_index=True
)


# ==========================================================
# 18. VALIDATE WEIGHTED AVERAGE
# ==========================================================

weight_validation_passed = True

for pair, method in clubbing_map.items():

    if method != "Weighted Average":
        continue

    assessment, co = pair

    pair_settings = settings_df[
        (settings_df["Assessment"] == assessment)
        & (settings_df["CO"] == co)
    ]

    total_weight = 0.0

    for _, row in pair_settings.iterrows():

        weight = co_weights.get(
            (
                assessment,
                co,
                row["Date Label"]
            ),
            0.0
        )

        total_weight += weight

    if total_weight <= 0:

        weight_validation_passed = False

        st.error(
            f"Enter at least one positive weight for "
            f"{assessment} - {co}."
        )


# ==========================================================
# 19. GENERATE RESULTS
# ==========================================================

st.divider()

st.subheader("05 | Generate Assessment Results")

if st.button(
    "Generate Assessment Matrix",
    type="primary",
    use_container_width=True
):

    if not weight_validation_passed:

        st.error(
            "Please correct the weighted average settings."
        )

    else:

        try:

            # ----------------------------------------------
            # CALCULATE RESCALED RESULTS
            # ----------------------------------------------

            results = calculate_results(
                course_data,
                identity_columns,
                config_map
            )

            # ----------------------------------------------
            # CREATE FINAL MATRIX
            # ----------------------------------------------

            final_matrix = make_output_table(
                results,
                identity_columns,
                repeated_pairs,
                clubbing_map,
                best_n_map,
                co_weights
            )

            st.session_state["generated_results"] = results

            st.session_state["final_matrix"] = final_matrix

            st.session_state["generated_course"] = {
                "Course Name": course_name,
                "Course Code": course_code,
                "Course Component": course_component
            }

            st.success(
                "Assessment matrix generated successfully!"
            )

        except Exception as exc:

            st.error(
                f"Unable to generate the assessment matrix: {exc}"
            )


# ==========================================================
# 20. SECTION 06 – GENERATED ASSESSMENT MATRIX
# ==========================================================

st.divider()

st.subheader("06 | Generated Assessment Matrix")

if "final_matrix" in st.session_state:

    final_matrix = st.session_state["final_matrix"]

    course_details = st.session_state.get(
        "generated_course",
        {}
    )

    # ------------------------------------------------------
    # COURSE DETAILS BESIDE THE MATRIX
    # ------------------------------------------------------

    left, right = st.columns([2, 1])

    with left:

        st.markdown("### Course Information")

        st.text_input(
            "Course Name",
            value=str(
                course_details.get("Course Name", "")
            ),
            disabled=True,
            key="final_course_name"
        )

        st.text_input(
            "Course Code",
            value=str(
                course_details.get("Course Code", "")
            ),
            disabled=True,
            key="final_course_code"
        )

        st.text_input(
            "Course Component",
            value=str(
                course_details.get("Course Component", "")
            ),
            disabled=True,
            key="final_course_component"
        )

    with right:

        st.markdown("### Matrix Information")

        st.metric(
            "Total Students",
            len(final_matrix)
        )

        st.metric(
            "Assessment Types",
            len(ASSESSMENT_ORDER)
        )

        st.metric(
            "CO Columns",
            len(ASSESSMENT_ORDER) * len(CO_LIST)
        )

    # ------------------------------------------------------
    # ASSESSMENT GROUP HEADERS
    # ------------------------------------------------------

    st.markdown("### Generated Student-wise Matrix")

    st.caption(
        "Each assessment has separate CO columns. "
        "Repeated CO entries are clubbed using the selected method."
    )

    # ------------------------------------------------------
    # DISPLAY MATRIX
    # ------------------------------------------------------

    st.dataframe(
        final_matrix,
        use_container_width=True,
        hide_index=True,
        height=550
    )

    # ------------------------------------------------------
    # DOWNLOAD EXCEL
    # ------------------------------------------------------

    excel_data = create_excel_file(
        final_matrix
    )

    st.download_button(
        label="Download Generated Matrix (Excel)",
        data=excel_data,
        file_name=(
            f"{safe_key(course_code)}_"
            f"Generated_Assessment_Matrix.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    # ------------------------------------------------------
    # DOWNLOAD CSV
    # ------------------------------------------------------

    csv_data = final_matrix.to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        label="Download Generated Matrix (CSV)",
        data=csv_data,
        file_name=(
            f"{safe_key(course_code)}_"
            f"Generated_Assessment_Matrix.csv"
        ),
        mime="text/csv",
        use_container_width=True
    )

else:

    st.info(
        "Click Generate Assessment Matrix to display "
        "the final student-wise assessment table."
    )


# ==========================================================
# 21. FOOTER
# ==========================================================

st.divider()

st.caption(
    "CO Assessment Dashboard | Assessment Rescaling "
    "and CO Clubbing"
)
