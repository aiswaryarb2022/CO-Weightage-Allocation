import streamlit as st
import pandas as pd
import numpy as np
import re
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
    "Moderation"
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

# ==========================================================
# 2. CSS
# ==========================================================

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

[data-testid="stDataFrame"],
[data-testid="stDataEditor"] {
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


# ==========================================================
# 3. COLUMN STANDARDIZATION
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

        "Course Component": [
            "Course Component",
            "Component",
            "Theory/Practical",
            "Theory or Practical"
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


# ==========================================================
# 4. NORMALIZATION FUNCTIONS
# ==========================================================

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


# ==========================================================
# 5. READ UPLOADED FILE
# ==========================================================

def read_uploaded_file(uploaded_file):

    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


# ==========================================================
# 6. PREPARE DATA
# ==========================================================

def prepare_data(raw_df):

    df = standardize_columns(
        raw_df.copy()
    )

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

    # ------------------------------------------------------
    # COURSE DETAILS
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # COURSE COMPONENT
    # ------------------------------------------------------

    if "Course Component" not in df.columns:

        course_type_text = (
            df["Course Type"]
            .astype(str)
            .str.lower()
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

    # ------------------------------------------------------
    # MARKS
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # NORMALIZE ASSESSMENT AND CO
    # ------------------------------------------------------

    df["Assessment"] = (
        df["Assessment"]
        .apply(normalize_assessment)
    )

    df["CO"] = (
        df["CO"]
        .apply(normalize_co)
    )

    # ------------------------------------------------------
    # STUDENT IDENTIFIERS
    # ------------------------------------------------------

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

            identifier = clean_identifier(
                row[col]
            )

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

    # ------------------------------------------------------
    # REMOVE INVALID RECORDS
    # ------------------------------------------------------

    df = df.dropna(
        subset=[
            "Mark Obtained",
            "Course Name"
        ]
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
# 7. ND REPLACEMENT
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

            "Date": date_key(
                row["Date"]
            ),

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

        row_index = int(
            row["Record Index"]
        )

        updated_data.loc[
            row_index,
            "Assessment"
        ] = row["Replace With"]

    return updated_data


# ==========================================================
# 8. GET REPEATED CO SUMMARY
# ==========================================================

def get_repeated_co_summary(course_data):

    temp = course_data.copy()

    temp["Date Label"] = (
        temp["Date"].apply(date_key)
    )

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
# 9. GET CONFIGURATION ROWS
# ==========================================================

def get_settings_rows(course_data):

    settings = course_data[
        [
            "Assessment",
            "Date",
            "CO"
        ]
    ].copy()

    settings["Date Label"] = (
        settings["Date"].apply(date_key)
    )

    settings = settings.drop_duplicates(
        [
            "Assessment",
            "Date Label",
            "CO"
        ]
    ).copy()

    settings["_assessment_order"] = (
        settings["Assessment"].map({
            a: i
            for i, a in enumerate(ASSESSMENT_ORDER)
        })
    )

    settings["_date_order"] = pd.to_datetime(
        settings["Date Label"],
        format="%d-%m-%Y",
        errors="coerce"
    )

    settings["_co_order"] = (
        settings["CO"].map({
            co: i
            for i, co in enumerate(CO_LIST)
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
        [
            "Assessment",
            "Date Label",
            "CO"
        ]
    ].reset_index(drop=True)


# ==========================================================
# 10. CALCULATE RESCALED RESULTS
# ==========================================================

def calculate_results(
    data,
    identity_columns,
    config_map
):

    class_average = {}

    # ------------------------------------------------------
    # CLASS AVERAGE FOR RELATIVE RESCALING
    # ------------------------------------------------------

    for (
        assessment_name,
        assessment_date,
        co_name
    ), class_group in data.groupby(
        [
            "Assessment",
            "Date",
            "CO"
        ],
        dropna=False
    ):

        class_average[
            (
                assessment_name,
                date_key(assessment_date),
                co_name
            )
        ] = class_group["Mark Obtained"].mean()

    # ------------------------------------------------------
    # STUDENT-WISE GROUPING
    # ------------------------------------------------------

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
                key if isinstance(key, tuple)
                else (key,)
            )
        )

        assessment = values["Assessment"]

        dlabel = date_key(
            values["Date"]
        )

        co = values["CO"]

        settings = config_map.get(
            (
                assessment,
                dlabel,
                co
            )
        )

        if settings is None:

            raise ValueError(
                f"Calculation settings are missing for "
                f"{assessment}, {dlabel}, {co}."
            )

        avg_mark = group["Mark Obtained"].mean()

        avg_max = group["Max Marks"].mean()

        class_avg = class_average.get(
            (
                assessment,
                dlabel,
                co
            ),
            np.nan
        )

        method = settings["Rescaling Method"]

        minimum_mark = float(
            settings.get("Minimum Mark", 0.0)
        )

        # --------------------------------------------------
        # ABSOLUTE RESCALING
        # --------------------------------------------------

        if method == "Absolute":

            if pd.notna(avg_max) and avg_max > 0:

                rescaled = (
                    avg_mark / avg_max
                ) * 100

            else:
                rescaled = np.nan

        # --------------------------------------------------
        # RELATIVE RESCALING
        # --------------------------------------------------

        elif method == "Relative":

            if (
                pd.notna(class_avg)
                and class_avg > 0
            ):

                rescaled = (
                    avg_mark / class_avg
                ) * 100

            else:
                rescaled = np.nan

        # --------------------------------------------------
        # MODERATION
        # --------------------------------------------------

        elif method == "Moderation":

            if pd.notna(avg_max) and avg_max > 0:

                original_percentage = (
                    avg_mark / avg_max
                ) * 100

                rescaled = max(
                    original_percentage,
                    minimum_mark
                )

            else:
                rescaled = np.nan

        else:

            rescaled = np.nan

        # --------------------------------------------------
        # KEEP ALL VALUES WITHIN 0–100
        # --------------------------------------------------

        if pd.notna(rescaled):

            rescaled = float(
                np.clip(
                    rescaled,
                    0,
                    100
                )
            )

            rescaled = round(
                rescaled,
                2
            )

        # --------------------------------------------------
        # CREATE RESULT ROW
        # --------------------------------------------------

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
# 11. CLUBBING SAME CO ACROSS DIFFERENT DATES
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

    combination = (
        assessment,
        co
    )

    # ------------------------------------------------------
    # NON-REPEATED CO
    # ------------------------------------------------------

    if combination not in repeated_pairs:

        return float(
            date_values.mean()
        )

    method = clubbing_map.get(
        combination
    )

    # ------------------------------------------------------
    # AVERAGE OF ALL
    # ------------------------------------------------------

    if method == "Average of All":

        return float(
            date_values.mean()
        )

    # ------------------------------------------------------
    # AVERAGE OF BEST
    # ------------------------------------------------------

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

        return float(
            best_values.mean()
        )

    # ------------------------------------------------------
    # WEIGHTED AVERAGE
    # ------------------------------------------------------

    elif method == "Weighted Average":

        weights = pd.Series(
            {
                dlabel: co_weights.get(
                    (
                        assessment,
                        co,
                        dlabel
                    ),
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
# 12. CREATE GENERATED STUDENT-WISE MATRIX
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

    # Remove duplicate column names
    keys = list(
        dict.fromkeys(
            col for col in keys
            if col in results.columns
        )
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

        base = dict(
            zip(keys, key_values)
        )

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

    output = pd.DataFrame(
        output_rows
    )

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

            ordered_columns.append(
                f"{assessment}_{co}"
            )

    for col in ordered_columns:

        if col not in output.columns:
            output[col] = np.nan

    output = output[
        ordered_columns
    ]

    # Keep generated results between 0 and 100
    for col in ordered_columns[6:]:

        output[col] = pd.to_numeric(
            output[col],
            errors="coerce"
        ).clip(
            lower=0,
            upper=100
        )

    return output


# ==========================================================
# 13. CREATE EXCEL FILE
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

        ws = writer.sheets[
            "Generated Matrix"
        ]

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

        for cell in ws[1]:

            column_name = str(
                cell.value
            )

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

        for row in ws.iter_rows(
            min_row=2
        ):

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

# ============================================================
# 13. DASHBOARD TITLE
# ============================================================

st.title("📊 CO Assessment Dashboard")

st.markdown(
    """
    Configure assessment-wise rescaling, manage repeated COs,
    generate the final CCA matrix and edit the marks before export.
    """
)

st.divider()


# ============================================================
# 14. SECTION 01 — UPLOAD PHASE 1 FILE
# ============================================================

st.subheader("01. Upload Phase 1 File")

uploaded_file = st.file_uploader(
    "Upload Phase 1 Excel or CSV file",
    type=["xlsx", "xls", "csv"],
    key="phase1_upload"
)

if uploaded_file is None:
    st.info("Please upload the Phase 1 file to continue.")
    st.stop()

try:
    raw_data = read_uploaded_file(uploaded_file)
    prepared_data = prepare_data(raw_data)

except Exception as e:
    st.error(f"Unable to read the uploaded file: {e}")
    st.stop()

if prepared_data.empty:
    st.warning("The uploaded file does not contain valid assessment records.")
    st.stop()

st.success(
    f"File uploaded successfully. {len(prepared_data)} records found."
)


# ============================================================
# 15. SECTION 02 — COURSE SELECTION
# ============================================================

st.divider()
st.subheader("02. Course Details")

course_columns = [
    "Course Name",
    "Course Code",
    "Course Type",
    "Course Component"
]

for column in course_columns:
    if column not in prepared_data.columns:
        prepared_data[column] = ""

course_lookup = (
    prepared_data[course_columns]
    .fillna("")
    .drop_duplicates()
    .reset_index(drop=True)
)

course_lookup["Course Name"] = (
    course_lookup["Course Name"].astype(str)
)

course_lookup["Course Code"] = (
    course_lookup["Course Code"].astype(str)
)

course_lookup["Course Component"] = (
    course_lookup["Course Component"].astype(str)
)

course_lookup["Course Type"] = (
    course_lookup["Course Type"].astype(str)
)

if course_lookup.empty:
    st.error("No course details were found in the uploaded file.")
    st.stop()


def course_label(row):
    return (
        f"{row['Course Name']} | "
        f"{row['Course Code']} | "
        f"{row['Course Component']}"
    )


course_lookup["Display"] = course_lookup.apply(
    course_label,
    axis=1
)

selected_course_label = st.selectbox(
    "Select Course",
    course_lookup["Display"].tolist(),
    key="selected_course"
)

selected_course = course_lookup[
    course_lookup["Display"] == selected_course_label
].iloc[0]

course_name = selected_course["Course Name"]
course_code = selected_course["Course Code"]
course_component = selected_course["Course Component"]
course_type = selected_course["Course Type"]

# Filter records belonging to selected course
course_data = prepared_data[
    (prepared_data["Course Name"].astype(str) == str(course_name))
    &
    (prepared_data["Course Code"].astype(str) == str(course_code))
].copy()

course_data = course_data.reset_index(drop=True)

# If Course Component is not available, use Course Type
if not course_component.strip():
    if "practical" in course_type.lower():
        course_component = "Practical"

    elif "theory" in course_type.lower():
        course_component = "Theory"

    else:
        course_component = "Not specified"

c1, c2, c3 = st.columns(3)

with c1:
    st.text_input(
        "Course Name",
        value=course_name,
        disabled=True
    )

with c2:
    st.text_input(
        "Course Code",
        value=course_code,
        disabled=True
    )

with c3:
    st.text_input(
        "Course Component",
        value=course_component,
        disabled=True
    )

st.caption(
    f"Total records for selected course: {len(course_data)}"
)


# ============================================================
# 16. SECTION 03 — ND REPLACEMENT
# ============================================================

st.divider()
st.subheader("03. Replace ND Assessment")

st.caption(
    "Select the assessment that should replace each ND entry. "
    "PA1 is selected by default."
)

nd_table = build_nd_replacement_table(course_data)

if not nd_table.empty:

    nd_column_config = {}

    if "Replace With" in nd_table.columns:

        nd_column_config["Replace With"] = (
            st.column_config.SelectboxColumn(
                "Replace With",
                options=ND_REPLACEMENT_OPTIONS,
                required=True
            )
        )

    disabled_nd_columns = [
        col for col in nd_table.columns
        if col != "Replace With"
    ]

    edited_nd_table = st.data_editor(
        nd_table,
        column_config=nd_column_config,
        disabled=disabled_nd_columns,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="nd_replacement_editor"
    )

    course_data = apply_nd_replacement(
        course_data,
        edited_nd_table
    )

else:
    st.info("No ND assessments were found.")

# Keep only valid assessments
course_data = course_data[
    course_data["Assessment"].isin(ASSESSMENT_ORDER)
].copy()

course_data = course_data.reset_index(drop=True)

if course_data.empty:
    st.warning(
        "No PA1, PA2, LE or LM assessment records are available."
    )
    st.stop()


# ============================================================
# 17. SECTION 04 — ASSESSMENT CONFIGURATION
# ============================================================

st.divider()
st.subheader("04. Assessment Configuration")

st.caption(
    "Configure rescaling separately for each assessment, date and CO. "
    "Repeated COs can be combined using the selected clubbing method."
)

settings_df = get_settings_rows(course_data)

if settings_df.empty:
    st.warning("No assessment configuration records were found.")
    st.stop()

repeated_summary = get_repeated_co_summary(course_data)

# Configuration storage
config_map = {}
clubbing_map = {}
best_n_map = {}
co_weights = {}

# Separate assessment layout
left_column, right_column = st.columns(2)

assessment_layout = [
    (left_column, ["PA1", "LE"]),
    (right_column, ["PA2", "LM"])
]


# ------------------------------------------------------------
# Render each assessment
# ------------------------------------------------------------

for display_column, assessments in assessment_layout:

    with display_column:

        for assessment in assessments:

            assessment_rows = settings_df[
                settings_df["Assessment"] == assessment
            ].copy()

            if assessment_rows.empty:
                continue

            with st.expander(
                f"⚙️ {assessment} Configuration",
                expanded=True
            ):

                st.markdown(f"### {assessment}")

                # --------------------------------------------
                # Date-wise and CO-wise rescaling
                # --------------------------------------------

                for _, row in assessment_rows.iterrows():

                    date_label = row["Date Label"]
                    co = row["CO"]

                    config_key = (
                        assessment,
                        date_label,
                        co
                    )

                    st.markdown(
                        f"**{date_label} — {co}**"
                    )

                    method_column, moderation_column = st.columns(2)

                    with method_column:

                        method = st.selectbox(
                            "Rescaling Method",
                            options=RESCALING_OPTIONS,
                            index=0,
                            key=(
                                f"method_{assessment}_"
                                f"{date_label}_{co}"
                            )
                        )

                    minimum_pass_mark = 40.0

                    with moderation_column:

                        if method == "Moderation":

                            minimum_pass_mark = st.number_input(
                                "Minimum Pass Mark (%)",
                                min_value=0.0,
                                max_value=100.0,
                                value=40.0,
                                step=1.0,
                                key=(
                                    f"minimum_{assessment}_"
                                    f"{date_label}_{co}"
                                )
                            )

                        else:

                            st.text_input(
                                "Minimum Pass Mark",
                                value="Not applicable",
                                disabled=True,
                                key=(
                                    f"minimum_display_{assessment}_"
                                    f"{date_label}_{co}"
                                )
                            )

                    config_map[config_key] = {
                        "method": method,
                        "minimum_pass_mark": minimum_pass_mark
                    }

                    st.markdown("---")

                # --------------------------------------------
                # Clubbing configuration
                # --------------------------------------------

                assessment_repeated = {
                    co: count
                    for (assess, co), count in repeated_summary.items()
                    if assess == assessment and count > 1
                }

                if assessment_repeated:

                    st.markdown("### Clubbing of Repeated COs")

                    for co, count in assessment_repeated.items():

                        st.markdown(
                            f"**{co} — {count} entries**"
                        )

                        clubbing_method = st.selectbox(
                            f"Clubbing Method for {co}",
                            options=CLUBBING_OPTIONS,
                            index=0,
                            key=f"clubbing_{assessment}_{co}"
                        )

                        clubbing_map[(assessment, co)] = (
                            clubbing_method
                        )

                        if clubbing_method == "Average of Best":

                            best_n = st.number_input(
                                f"Number of Best Entries for {co}",
                                min_value=1,
                                max_value=count,
                                value=count,
                                step=1,
                                key=f"best_n_{assessment}_{co}"
                            )

                            best_n_map[(assessment, co)] = int(best_n)

                        elif clubbing_method == "Weighted Average":

                            repeated_dates = (
                                settings_df[
                                    (settings_df["Assessment"] == assessment)
                                    &
                                    (settings_df["CO"] == co)
                                ]["Date Label"]
                                .drop_duplicates()
                                .tolist()
                            )

                            st.caption(
                                "Enter a weight for each repeated entry."
                            )

                            weight_columns = st.columns(
                                min(len(repeated_dates), 3)
                            )

                            for index, date_label in enumerate(
                                repeated_dates
                            ):

                                with weight_columns[
                                    index % len(weight_columns)
                                ]:

                                    weight = st.number_input(
                                        f"{date_label} Weight",
                                        min_value=0.0,
                                        value=1.0,
                                        step=0.5,
                                        key=(
                                            f"weight_{assessment}_"
                                            f"{co}_{date_label}"
                                        )
                                    )

                                co_weights[
                                    (assessment, co, date_label)
                                ] = weight

                        st.markdown("---")


# ============================================================
# 18. SECTION 05 — CONFIGURATION SUMMARY
# ============================================================

st.divider()
st.subheader("05. Configuration Summary")

summary_records = []

for key, value in config_map.items():

    assessment, date_label, co = key

    summary_records.append({
        "Assessment": assessment,
        "Date": date_label,
        "CO": co,
        "Rescaling Method": value["method"],
        "Minimum Pass Mark (%)": (
            value["minimum_pass_mark"]
            if value["method"] == "Moderation"
            else "-"
        )
    })

configuration_summary = pd.DataFrame(summary_records)

st.dataframe(
    configuration_summary,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# 19. SECTION 05 — GENERATE RESULTS
# ============================================================

st.markdown("### Generate CCA Results")

generate_button = st.button(
    "🚀 Generate Final CCA Matrix",
    use_container_width=True,
    type="primary"
)

if generate_button:

    # --------------------------------------------
    # Validate weighted average settings
    # --------------------------------------------

    weight_errors = []

    for (assessment, co), method in clubbing_map.items():

        if method != "Weighted Average":
            continue

        weight_values = [
            value
            for (a, c, date), value in co_weights.items()
            if a == assessment and c == co
        ]

        if not weight_values or sum(weight_values) <= 0:

            weight_errors.append(
                f"{assessment} - {co}: "
                "The total weight must be greater than zero."
            )

    if weight_errors:

        for error in weight_errors:
            st.error(error)

    else:

        try:

            with st.spinner("Generating CCA results..."):

                # Calculate assessment-wise results
                results = calculate_results(
                    course_data,
                    config_map,
                    clubbing_map,
                    best_n_map,
                    co_weights
                )

                # Create final CO matrix
                final_matrix = make_output_table(results)

                if final_matrix.empty:

                    st.error(
                        "No results were generated. "
                        "Please check the uploaded assessment data."
                    )

                else:

                    # Store generated data
                    st.session_state["generated_results"] = results

                    st.session_state["final_matrix"] = (
                        final_matrix.copy()
                    )

                    st.session_state["generated_course"] = {
                        "Course Name": course_name,
                        "Course Code": course_code,
                        "Course Component": course_component
                    }

                    st.success(
                        "CCA matrix generated successfully!"
                    )

        except Exception as e:

            st.error(
                f"Error while generating results: {e}"
            )


# ============================================================
# 20. SECTION 06 — FINAL CCA MATRIX
# ============================================================

st.divider()
st.subheader("06. Final CCA Matrix")

if "final_matrix" in st.session_state:

    final_matrix = st.session_state["final_matrix"].copy()

    generated_course = st.session_state.get(
        "generated_course",
        {}
    )

    # --------------------------------------------
    # Display generated course details
    # --------------------------------------------

    st.markdown("### Generated Course Details")

    detail_col1, detail_col2, detail_col3 = st.columns(3)

    with detail_col1:

        st.text_input(
            "Course Name",
            value=generated_course.get("Course Name", ""),
            disabled=True,
            key="generated_course_name"
        )

    with detail_col2:

        st.text_input(
            "Course Code",
            value=generated_course.get("Course Code", ""),
            disabled=True,
            key="generated_course_code"
        )

    with detail_col3:

        st.text_input(
            "Course Component",
            value=generated_course.get("Course Component", ""),
            disabled=True,
            key="generated_course_component"
        )

    st.markdown("---")

    st.markdown("### Editable Assessment Marks")

    st.caption(
        "You can edit the marks in the matrix. "
        "Course and student details are locked. "
        "All CO marks must remain between 0 and 100."
    )

    # --------------------------------------------
    # Identify columns
    # --------------------------------------------

    identity_columns = [
        "Student Name",
        "Reg No",
        "Admission No",
        "Course Code",
        "Course Component",
        "Course Name"
    ]

    disabled_columns = [
        col
        for col in identity_columns
        if col in final_matrix.columns
    ]

    mark_columns = [
        col
        for col in final_matrix.columns
        if col not in disabled_columns
    ]

    # --------------------------------------------
    # Column configuration
    # --------------------------------------------

    matrix_column_config = {}

    for col in mark_columns:

        matrix_column_config[col] = (
            st.column_config.NumberColumn(
                col,
                min_value=0.0,
                max_value=100.0,
                step=0.01,
                format="%.2f"
            )
        )

    # --------------------------------------------
    # Editable final matrix
    # --------------------------------------------

    edited_matrix = st.data_editor(
        final_matrix,
        column_config=matrix_column_config,
        disabled=disabled_columns,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        height=550,
        key="generated_matrix_editor"
    )

    # --------------------------------------------
    # Validate edited values
    # --------------------------------------------

    edited_matrix = edited_matrix.copy()

    for col in mark_columns:

        edited_matrix[col] = pd.to_numeric(
            edited_matrix[col],
            errors="coerce"
        )

        edited_matrix[col] = (
            edited_matrix[col]
            .fillna(0)
            .clip(lower=0, upper=100)
            .round(2)
        )

    # Save edited matrix
    st.session_state["final_matrix"] = (
        edited_matrix.copy()
    )

    st.success(
        "The final CCA matrix is ready. "
        "Your edits will be included in the downloaded files."
    )

    # --------------------------------------------
    # Download options
    # --------------------------------------------

    st.markdown("### Download Final Matrix")

    download_col1, download_col2 = st.columns(2)

    # Excel
    excel_file = create_excel_file(edited_matrix)

    safe_course_code = (
        str(generated_course.get("Course Code", "Course"))
        .replace("/", "_")
        .replace("\\", "_")
        .replace(" ", "_")
    )

    with download_col1:

        st.download_button(
            label="📥 Download Excel",
            data=excel_file,
            file_name=f"CCA_Matrix_{safe_course_code}.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    # CSV
    csv_file = edited_matrix.to_csv(
        index=False
    ).encode("utf-8-sig")

    with download_col2:

        st.download_button(
            label="📥 Download CSV",
            data=csv_file,
            file_name=f"CCA_Matrix_{safe_course_code}.csv",
            mime="text/csv",
            use_container_width=True
        )

else:

    st.info(
        "The final CCA matrix will appear here after you click "
        "'Generate Final CCA Matrix'."
    )


# ============================================================
# 21. FOOTER
# ============================================================

st.divider()

st.markdown(
    """
    <div style="text-align:center; color:#777; padding:12px;">
        CO Assessment Dashboard
    </div>
    """,
    unsafe_allow_html=True
)
