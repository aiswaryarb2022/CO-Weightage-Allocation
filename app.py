import streamlit as st
import pandas as pd
import numpy as np
import requests
from pathlib import Path

from io import BytesIO
from scipy.optimize import linprog, minimize
from openpyxl.styles import Font, PatternFill, Alignment


# ==================================================
# PAGE CONFIGURATION
# ==================================================

st.set_page_config(
    page_title="CO Weightage Allocation Dashboard",
    page_icon="📊",
    layout="wide"
)

CONVERSION_TARGET = 20.0
ASSESSMENT_NAMES = ["PA1", "PA2", "LE1", "LM"]

from pathlib import Path
import re

# ==================================================
# COMBINED COURSE DATABASE
# ==================================================

BASE_DIR = Path(__file__).resolve().parent

COURSE_FILE = BASE_DIR / "All_Courses_Consolidated.xlsx"
COURSE_GROUP_FILE = BASE_DIR / "Course Group(1).xlsx"


# Check Excel files
if not COURSE_FILE.exists():
    st.error(f"File not found: {COURSE_FILE}")
    st.stop()

if not COURSE_GROUP_FILE.exists():
    st.error(f"File not found: {COURSE_GROUP_FILE}")
    st.stop()


@st.cache_data
def load_all_courses():

    # Load first Excel file
    df1 = pd.read_excel(
        COURSE_FILE,
        sheet_name="All Courses"
    )

    # Load second Excel file
    df2 = pd.read_excel(
        COURSE_GROUP_FILE,
        sheet_name="Sheet1",
        header=1
    )

    # Standardise column names
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Prepare first file
    df1 = df1.rename(columns={
        "Program": "Course Name",
        "Program Code": "Course Code",
        "Program Type": "Program Type",
        "Handling Department": "Department"
    })

    # Prepare second file
    df2 = df2.rename(columns={
        "Program": "Course Name",
        "Program Code": "Course Code",
        "Program Type": "Program Type",
        "Handling Department": "Department"
    })

    # Extract semester from course code
    def get_semester(code):

        match = re.search(r"(\d{3})$", str(code))

        if match:
            number = int(match.group(1))

            if 101 <= number <= 500:
                return (number - 101) // 50 + 1

        return None

    # Infer UG / PG from recognised course patterns
    def get_level(code, program_type):

        code = str(code).upper()
        program_type = str(program_type).upper()

        if "PG" in code or "PG" in program_type:
            return "PG"

        if any(x in code for x in ["DSC", "AEC", "MDC"]):
            return "UG"

        return None

    # Prepare semester in first file
    df1["Semester Number"] = (
        df1["Semester"]
        .astype(str)
        .str.extract(r"(\d+)")
        .iloc[:, 0]
    )

    df1["Semester Number"] = pd.to_numeric(
        df1["Semester Number"],
        errors="coerce"
    )

    # Prepare second file
    df2["Semester Number"] = df2["Course Code"].apply(
        get_semester
    )

    df2["Level"] = df2.apply(
        lambda row: get_level(
            row["Course Code"],
            row["Program Type"]
        ),
        axis=1
    )

    # First file already has its level
    df1["Level"] = df1["Level"].astype(str).str.upper().str.strip()

    # Combine both files
    combined = pd.concat(
        [df1, df2],
        ignore_index=True,
        sort=False
    )

    # Clean data
    combined["Course Code"] = (
        combined["Course Code"].astype(str).str.strip()
    )

    combined["Department"] = (
        combined["Department"].astype(str).str.strip()
    )

    combined["Program Type"] = (
        combined["Program Type"].astype(str).str.strip()
    )

    combined["Course Name"] = (
        combined["Course Name"].astype(str).str.strip()
    )

    combined["Level"] = (
        combined["Level"].astype(str).str.upper().str.strip()
    )

    # Remove invalid records
    combined = combined[
        combined["Course Code"].notna()
        & combined["Course Name"].notna()
        & combined["Department"].notna()
        & combined["Course Code"].ne("nan")
        & combined["Department"].ne("nan")
    ].copy()

    # Calculate academic year from semester
    def get_year(row):

        semester = row["Semester Number"]
        level = row["Level"]

        if pd.isna(semester):
            return None

        semester = int(semester)

        if level == "UG" and 1 <= semester <= 6:
            return (semester + 1) // 2

        if level == "PG" and 1 <= semester <= 4:
            return (semester + 1) // 2

        return None

    combined["Year"] = combined.apply(
        get_year,
        axis=1
    )

    # Remove duplicate course records
    combined = combined.drop_duplicates(
        subset=["Course Code", "Department"],
        keep="first"
    )

    return combined.reset_index(drop=True)


COURSE_DATA = load_all_courses()
# ==================================================
# CUSTOM CSS
# ==================================================

st.markdown("""
<style>

.stApp {
    background-color: #F3F6F1;
}

.stApp p,
.stApp li,
.stApp label,
.stApp [data-testid="stMarkdownContainer"] p,
.stApp [data-testid="stCaptionContainer"] {
    color: #171717 !important;
}

.dashboard-header {
    background: linear-gradient(120deg, #183D32, #285C49);
    padding: 22px 28px;
    border-radius: 12px;
    margin-bottom: 15px;
    border-bottom: 5px solid #C8B98A;
}

.dashboard-header h1,
.dashboard-header .dashboard-subtitle {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

.dashboard-header h1 {
    font-size: 28px;
    font-weight: 750;
    margin: 0;
}

.dashboard-header .dashboard-subtitle {
    font-size: 14px !important;
    font-weight: 700 !important;
    margin-top: 6px;
    margin-bottom: 0;
}

.section-heading {
    color: #183D32 !important;
    font-size: 20px;
    font-weight: 750;
    margin-top: 14px;
    margin-bottom: 9px;
    border-left: 5px solid #C8B98A;
    padding-left: 12px;
}

[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #FFFFFF;
    border: 1px solid #DCE5DA;
    border-top: 4px solid #285C49;
    border-radius: 12px;
    padding: 10px;
    margin-bottom: 8px;
}

.assessment-title {
    color: #183D32 !important;
    font-size: 18px;
    font-weight: 750;
    margin-bottom: 8px;
}

.stTextInput input,
.stNumberInput input {
    color: #171717 !important;
    -webkit-text-fill-color: #171717 !important;
    background-color: #FFFFFF !important;
}

.stTextInput label,
.stNumberInput label,
.stSelectbox label,
.stCheckbox label {
    color: #171717 !important;
    font-weight: 600 !important;
}

.stSelectbox div[data-baseweb="select"],
.stSelectbox div[data-baseweb="select"] span {
    color: #171717 !important;
}

.stButton > button,
.stDownloadButton > button {
    background-color: #183D32 !important;
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
    border: 2px solid #C8B98A !important;
    border-radius: 8px;
    font-weight: 700 !important;
    padding: 9px 22px;
}

.stButton > button *,
.stDownloadButton > button * {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
    font-weight: 700 !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background-color: #285C49 !important;
    border-color: #D8CBA5 !important;
}

[data-testid="stMetric"] {
    background-color: #FFFFFF;
    border: 1px solid #DCE5DA;
    border-left: 4px solid #C8B98A;
    border-radius: 10px;
    padding: 12px;
}

[data-testid="stMetricLabel"],
[data-testid="stMetricValue"] {
    color: #183D32 !important;
    font-weight: 700 !important;
}

[data-testid="stDataFrame"] {
    border: 1px solid #DCE5DA;
    border-radius: 8px;
}

[data-testid="stExpander"] summary {
    color: #183D32 !important;
    font-weight: 700 !important;
}

</style>
""", unsafe_allow_html=True)


# ==================================================
# DASHBOARD HEADER
# ==================================================

st.markdown("""
<div class="dashboard-header">
    <h1>CO Weightage Allocation Dashboard</h1>
    <p class="dashboard-subtitle">
        Course Outcome Assessment Weightage Management System
    </p>
</div>
""", unsafe_allow_html=True)

# ==================================================
# COURSE DETAILS
# ==================================================

st.markdown(
    '<div class="section-heading">Course Details</div>',
    unsafe_allow_html=True
)

# Programme Level
col1, col2, col3 = st.columns(3)

with col1:

    levels = sorted(
        COURSE_DATA["Level"]
        .dropna()
        .loc[lambda x: x.isin(["UG", "PG"])]
        .unique()
        .tolist()
    )

    programme_level = st.selectbox(
        "Programme Level",
        levels,
        key="programme_level"
    )


# Filter by level
level_data = COURSE_DATA[
    COURSE_DATA["Level"] == programme_level
].copy()


# Year dropdown
with col2:

    years = sorted(
        level_data["Year"].dropna().unique().tolist()
    )

    selected_year = st.selectbox(
        "Year",
        years,
        format_func=lambda x: (
            f"{int(x)} Year"
        ),
        key="selected_year"
    )


year_data = level_data[
    level_data["Year"] == selected_year
].copy()


# Department dropdown
with col3:

    departments = sorted(
        year_data["Department"].dropna().unique().tolist()
    )

    department = st.selectbox(
        "Department",
        departments,
        key="selected_department"
    )


department_data = year_data[
    year_data["Department"] == department
].copy()


# Programme Type and Semester
col4, col5 = st.columns(2)

with col4:

    program_types = sorted(
        department_data["Program Type"]
        .dropna()
        .unique()
        .tolist()
    )

    program_type = st.selectbox(
        "Program Type",
        program_types,
        key="selected_program_type"
    )


type_data = department_data[
    department_data["Program Type"] == program_type
].copy()


with col5:

    semesters = sorted(
        type_data["Semester Number"]
        .dropna()
        .unique()
        .astype(int)
        .tolist()
    )

    semester = st.selectbox(
        "Semester",
        semesters,
        format_func=lambda x: f"Semester {x}",
        key="selected_semester"
    )


# Filter by semester
semester_data = type_data[
    type_data["Semester Number"] == semester
].copy()


# Course Code dropdown
course_options = semester_data.apply(
    lambda row: (
        f"{row['Course Code']} - {row['Course Name']}"
    ),
    axis=1
).tolist()

selected_course_option = st.selectbox(
    "Course Code",
    course_options,
    key="selected_course_code"
)


# Get selected course
selected_course = semester_data.iloc[
    course_options.index(selected_course_option)
]

program = selected_course["Course Name"]

course_code = selected_course["Course Code"]

semester_label = f"Semester {semester}"


# Display automatically selected details
col6, col7 = st.columns(2)

with col6:

    st.text_input(
        "Selected Course Code",
        value=str(course_code),
        disabled=True
    )

with col7:

    st.text_input(
        "Selected Semester",
        value=semester_label,
        disabled=True
    )

# ==================================================
# COURSE OUTCOMES
# ==================================================

st.markdown(
    '<div class="section-heading">Course Outcome Details</div>',
    unsafe_allow_html=True
)

number_of_cos = st.number_input(
    "Number of Course Outcomes",
    min_value=1,
    max_value=20,
    value=4,
    step=1
)

co_list = [
    f"CO{i}"
    for i in range(1, int(number_of_cos) + 1)
]


# ==================================================
# CO WEIGHT DETAILS
# ==================================================

st.markdown(
    '<div class="section-heading">CO Weight Details</div>',
    unsafe_allow_html=True
)

st.caption(
    "Enter the weight of each Course Outcome. "
    "The CO-wise total marks will be calculated proportionally out of 20."
)

co_weight_columns = st.columns(min(len(co_list), 4))

co_weightages = {}

for i, co in enumerate(co_list):

    with co_weight_columns[i % len(co_weight_columns)]:

        co_weightages[co] = st.number_input(
            f"{co} Weight",
            min_value=0.0,
            value=1.0,
            step=0.5,
            key=f"{course_code}_{co}_weightage"
        )


# ==================================================
# ASSESSMENT INPUT FUNCTION
# ==================================================

def assessment_input(assessment_name, key_prefix):

    with st.container(border=True):

        st.markdown(
            f'<div class="assessment-title">{assessment_name}</div>',
            unsafe_allow_html=True
        )

        exam_max = st.number_input(
            f"{assessment_name} Total Marks",
            min_value=0.0,
            value=0.0,
            step=0.5,
            key=f"{course_code}_{key_prefix}_total_weight"
        )

        st.markdown("**Select Course Outcomes**")

        checkbox_columns = st.columns(
            min(len(co_list), 5)
        )

        selected_cos = []

        for i, co in enumerate(co_list):

            with checkbox_columns[i % len(checkbox_columns)]:

                selected = st.checkbox(
                    co,
                    key=f"{course_code}_{key_prefix}_{co}_checkbox"
                )

                if selected:
                    selected_cos.append(co)

        co_weights = {}
        co_pcts = {}

        if selected_cos:

            st.markdown("**Enter weight for selected COs**")

            weight_columns = st.columns(
                min(len(selected_cos), 4)
            )

            for i, co in enumerate(selected_cos):

                with weight_columns[i % len(weight_columns)]:

                    co_weights[co] = st.number_input(
                        f"{co} Weight",
                        min_value=0.0,
                        value=1.0,
                        step=0.5,
                        key=f"{course_code}_{key_prefix}_{co}_weight"
                    )

            total_weight = sum(co_weights.values())

            if total_weight > 0:

                for co in selected_cos:

                    co_pcts[co] = (
                        co_weights[co] / total_weight
                    ) * 100

                st.markdown("**Calculated CO Allocation**")

                percentage_df = pd.DataFrame({

                    "Course Outcome": selected_cos,

                    "Weight": [
                        co_weights[co]
                        for co in selected_cos
                    ],

                    "Allocation Percentage": [
                        f"{co_pcts[co]:.2f}%"
                        for co in selected_cos
                    ]

                })

                st.dataframe(
                    percentage_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.warning(
                    "Enter a positive weight for at least one CO."
                )

        else:

            st.caption("No Course Outcomes selected.")

    return exam_max, selected_cos, co_pcts


# ==================================================
# ASSESSMENT COMPONENTS
# ==================================================

st.markdown(
    '<div class="section-heading">Assessment Components</div>',
    unsafe_allow_html=True
)

st.caption(
    "Enter the total marks for PA1, PA2, LE1 and LM. "
    "Select the COs and enter their weights. "
    "Allocation percentages are calculated automatically. "
    "All four assessment totals must add up to 20."
)

all_assessments = []

assessment_columns = st.columns(4)

for i, name in enumerate(ASSESSMENT_NAMES):

    with assessment_columns[i]:

        exam_max, selected_cos, co_pcts = assessment_input(
            name,
            name.lower()
        )

        all_assessments.append(
            (name, exam_max, selected_cos, co_pcts)
        )


# ==================================================
# PREPARE ASSESSMENT TOTALS
# ==================================================

exam_max_map = {
    item[0]: item[1]
    for item in all_assessments
}

exam_max_sum = sum(exam_max_map.values())


# ==================================================
# VALIDATION
# ==================================================

errors = []

if not np.isclose(
    exam_max_sum,
    CONVERSION_TARGET,
    atol=0.01
):

    errors.append(
        f"Assessment totals must add up to 20. "
        f"Current total: {exam_max_sum:.2f}."
    )


for name, exam_max, selected_cos, co_pcts in all_assessments:

    if exam_max > 0 and not selected_cos:

        errors.append(
            f"{name}: Select at least one CO."
        )

    elif exam_max > 0 and abs(sum(co_pcts.values()) - 100) > 0.01:

        errors.append(
            f"{name}: CO weights must be positive. "
            "Please enter valid weights for the selected COs."
        )


# ==================================================
# CO TARGET CALCULATION
# ==================================================

total_co_weight = sum(co_weightages.values())

if total_co_weight <= 0:

    errors.append(
        "Total CO weight must be greater than zero."
    )

co_targets = {}

if total_co_weight > 0:

    for co in co_list:

        co_targets[co] = (
            co_weightages[co] / total_co_weight
        ) * CONVERSION_TARGET


# ==================================================
# INITIAL ALLOCATION AND ALLOWED CELLS
# ==================================================

n_co = len(co_list)

n_assessments = len(ASSESSMENT_NAMES)

initial_matrix = np.zeros(
    (n_co, n_assessments)
)

allowed = np.zeros(
    (n_co, n_assessments),
    dtype=bool
)

for j, (name, exam_max, selected_cos, co_pcts) in enumerate(
    all_assessments
):

    for co in selected_cos:

        i = co_list.index(co)

        allowed[i, j] = True

        initial_matrix[i, j] = (
            exam_max * co_pcts.get(co, 0.0) / 100
        )


# ==================================================
# ALLOCATION FUNCTION - LEAST SQUARES OPTIMISATION
# ==================================================

def calculate_final_allocation():

    assessment_targets = np.array([
        exam_max_map[a]
        for a in ASSESSMENT_NAMES
    ], dtype=float)

    co_target_array = np.array([
        co_targets[co]
        for co in co_list
    ], dtype=float)

    n = n_co * n_assessments

    seed = initial_matrix.flatten()

    # Check total consistency
    if not np.isclose(
        assessment_targets.sum(),
        co_target_array.sum(),
        atol=0.01
    ):
        return None, "Assessment and CO totals do not match."

    # Check whether every positive CO target is selected
    for i, co in enumerate(co_list):

        if co_target_array[i] > 0 and not allowed[i, :].any():

            return None, (
                f"{co} has a positive target but is not selected "
                "in any assessment."
            )

    # Check whether every positive assessment has a selected CO
    for j, assessment in enumerate(ASSESSMENT_NAMES):

        if assessment_targets[j] > 0 and not allowed[:, j].any():

            return None, (
                f"{assessment} has positive marks but no CO is selected."
            )

    # --------------------------------------------------
    # CREATE EQUALITY CONSTRAINTS
    # --------------------------------------------------

    A_eq = []
    b_eq = []

    # Assessment row totals
    for j in range(n_assessments):

        row = np.zeros(n)

        for i in range(n_co):
            row[i * n_assessments + j] = 1

        A_eq.append(row)
        b_eq.append(assessment_targets[j])

    # CO column totals
    # Last CO constraint is omitted because it is redundant.
    for i in range(n_co - 1):

        row = np.zeros(n)

        for j in range(n_assessments):
            row[i * n_assessments + j] = 1

        A_eq.append(row)
        b_eq.append(co_target_array[i])

    A_eq = np.array(A_eq)
    b_eq = np.array(b_eq)

    # --------------------------------------------------
    # BOUNDS
    # --------------------------------------------------

    bounds = []

    for i in range(n_co):

        for j in range(n_assessments):

            if allowed[i, j]:
                bounds.append((0, None))
            else:
                bounds.append((0, 0))

    # --------------------------------------------------
    # FIND A FEASIBLE STARTING SOLUTION
    # --------------------------------------------------

    feasible = linprog(
        c=np.zeros(n),
        A_eq=A_eq,
        b_eq=b_eq,
        bounds=bounds,
        method="highs"
    )

    if not feasible.success:

        return None, (
            "The selected CO combinations cannot satisfy both "
            "assessment totals and CO target totals. "
            "Please review your CO selections."
        )

    # --------------------------------------------------
    # LEAST SQUARES OBJECTIVE
    # --------------------------------------------------

    def objective(x):

        difference = x - seed

        return 0.5 * np.sum(difference ** 2)

    def gradient(x):

        return x - seed

    # --------------------------------------------------
    # OPTIMISE WHILE PRESERVING BOTH TOTALS
    # --------------------------------------------------

    result = minimize(
        objective,
        feasible.x,
        jac=gradient,
        method="SLSQP",
        bounds=bounds,
        constraints=[
            {
                "type": "eq",
                "fun": lambda x: A_eq @ x - b_eq,
                "jac": lambda x: A_eq
            }
        ],
        options={
            "maxiter": 2000,
            "ftol": 1e-10,
            "disp": False
        }
    )

    if not result.success:

        return None, (
            "The allocation optimisation could not converge. "
            "Please review the assessment CO weights and try again."
        )

    allocation = result.x.reshape(
        n_co,
        n_assessments
    )

    allocation[np.abs(allocation) < 1e-9] = 0

    return allocation, None


# ==================================================
# CALCULATE ALLOCATION
# ==================================================

allocation = None

if not errors:

    allocation, allocation_error = calculate_final_allocation()

    if allocation_error:
        errors.append(allocation_error)


# ==================================================
# BALANCED ROUNDING
# ==================================================

def balanced_round_matrix(matrix):

    n_co, n_assessments = matrix.shape

    # Convert to cents
    scaled = matrix * 100

    base = np.floor(
        scaled + 1e-9
    ).astype(int)

    fractions = scaled - base

    # Assessment targets in cents
    row_targets = np.rint(
        matrix.sum(axis=0) * 100
    ).astype(int)

    # CO targets in cents using largest remainder
    raw_co_targets = np.array([
        co_targets[co] * 100
        for co in co_list
    ])

    col_targets = np.floor(
        raw_co_targets
    ).astype(int)

    remaining = (
        int(round(CONVERSION_TARGET * 100))
        - col_targets.sum()
    )

    order = np.argsort(
        -(raw_co_targets - col_targets)
    )

    for i in order[:remaining]:
        col_targets[i] += 1

    # Deficits
    row_deficits = (
        row_targets - base.sum(axis=0)
    )

    col_deficits = (
        col_targets - base.sum(axis=1)
    )

    n = n_co * n_assessments

    A_eq = []
    b_eq = []

    # CO constraints
    for i in range(n_co):

        constraint = np.zeros(n)

        for j in range(n_assessments):

            constraint[i * n_assessments + j] = 1

        A_eq.append(constraint)

        b_eq.append(col_deficits[i])

    # Assessment constraints
    for j in range(n_assessments):

        constraint = np.zeros(n)

        for i in range(n_co):

            constraint[i * n_assessments + j] = 1

        A_eq.append(constraint)

        b_eq.append(row_deficits[j])

    # Bounds
    bounds = []

    for i in range(n_co):

        for j in range(n_assessments):

            if allowed[i, j] and fractions[i, j] > 1e-8:
                bounds.append((0, 1))
            else:
                bounds.append((0, 0))

    objective = -fractions.flatten()

    result = linprog(
        c=objective,
        A_eq=np.array(A_eq),
        b_eq=np.array(b_eq),
        bounds=bounds,
        method="highs"
    )

    if not result.success:

        raise ValueError(
            "Balanced rounding is not possible for this matrix. "
            "Please check the assessment weights and CO weights."
        )

    adjustments = np.rint(
        result.x
    ).astype(int).reshape(
        n_co,
        n_assessments
    )

    rounded_matrix = (
        base + adjustments
    ) / 100

    return rounded_matrix


# ==================================================
# BUILD FINAL MATRIX
# ==================================================

def build_matrix_df(matrix):

    rows = []

    for j, assessment in enumerate(ASSESSMENT_NAMES):

        row = {
            "Assessment": assessment
        }

        for i, co in enumerate(co_list):

            row[co] = matrix[i, j]

        row["TOTAL"] = matrix[:, j].sum()

        rows.append(row)

    total = {
        "Assessment": "TOTAL"
    }

    for i, co in enumerate(co_list):

        total[co] = matrix[i, :].sum()

    total["TOTAL"] = matrix.sum()

    rows.append(total)

    return pd.DataFrame(rows)


# ==================================================
# BUILD PERCENTAGE DATAFRAME
# ==================================================

def build_percentage_df():

    rows = []

    for name, exam_max, selected_cos, co_pcts in all_assessments:

        row = {
            "Assessment": name
        }

        for co in co_list:

            row[co] = co_pcts.get(co, 0.0)

        row["TOTAL"] = sum(co_pcts.values())

        rows.append(row)

    return pd.DataFrame(rows)


# ==================================================
# FINAL MATRIX DISPLAY
# ==================================================

st.markdown(
    '<div class="section-heading">'
    'FINAL CO ASSESSMENT WEIGHTAGE ALLOCATION MATRIX'
    '</div>',
    unsafe_allow_html=True
)

st.caption(
    "Each assessment row matches its entered total marks. "
    "Each CO column matches its proportional target weight."
)

result_df = None
display_allocation = None

if errors:

    for error in errors:
        st.error(error)

else:

    unused_cos = [
        co for co in co_list
        if not allowed[co_list.index(co), :].any()
    ]

    if unused_cos:

        st.info(
            f"{', '.join(unused_cos)} are not used in any assessment. "
            "Their weights are included in the proportional CO targets."
        )

    try:

        # Balanced rounding
        display_allocation = balanced_round_matrix(allocation)

        # Final matrix
        result_df = build_matrix_df(display_allocation)

        # Display matrix
        st.dataframe(
            result_df,
            use_container_width=True,
            hide_index=True
        )

        # Summary metrics
        m1, m2, m3 = st.columns(3)

        m1.metric(
            "Grand Total",
            f"{display_allocation.sum():.2f}"
        )

        m2.metric(
            "Assessment Total",
            f"{display_allocation.sum(axis=0).sum():.2f}"
        )

        m3.metric(
            "CO Total",
            f"{display_allocation.sum(axis=1).sum():.2f}"
        )

        st.success(
            "The final matrix has been calculated successfully."
        )

        # Verify assessment totals
        expected_rows = np.rint(
            np.array([
                exam_max_map[a]
                for a in ASSESSMENT_NAMES
            ]) * 100
        ).astype(int)

        actual_rows = np.rint(
            display_allocation.sum(axis=0) * 100
        ).astype(int)

        # Verify CO totals
        expected_cols = np.rint(
            np.array([
                co_targets[co]
                for co in co_list
            ]) * 100
        ).astype(int)

        actual_cols = np.rint(
            display_allocation.sum(axis=1) * 100
        ).astype(int)

        if (
            np.array_equal(expected_rows, actual_rows)
            and np.array_equal(expected_cols, actual_cols)
            and round(display_allocation.sum(), 2) == 20.00
        ):

            st.success(
                "All assessment totals, CO totals and grand total are correct."
            )

        # Percentage adjustment information
        max_shift = 0.0

        for j in range(n_assessments):

            col_total = allocation[:, j].sum()

            if col_total > 0:

                actual_pct = (
                    allocation[:, j] / col_total
                ) * 100

                original_pct = np.array([
                    next(
                        item[3].get(co, 0.0)
                        for item in all_assessments
                        if item[0] == ASSESSMENT_NAMES[j]
                    )
                    for co in co_list
                ])

                max_shift = max(
                    max_shift,
                    np.max(
                        np.abs(actual_pct - original_pct)
                    )
                )

        if max_shift > 0.01:

            st.info(
                "The entered CO weights were adjusted to satisfy "
                "both sets of totals. Maximum adjustment: "
                f"{max_shift:.2f} percentage points."
            )

        else:

            st.success(
                "The entered CO weights were maintained."
            )

    except ValueError as e:

        st.error(str(e))


# ==========================================================
# EXCEL EXPORT - ONLY TWO SHEETS
# ==========================================================

def create_excel(final_matrix, course_code):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # --------------------------------------------------
        # SHEET 1: FINAL ALLOCATION MATRIX
        # --------------------------------------------------

        final_matrix.to_excel(
            writer,
            sheet_name="Final Allocation Matrix",
            index=False
        )

        

        # --------------------------------------------------
        # EXCEL FORMATTING
        # --------------------------------------------------

        for sheet_name in writer.sheets:

            ws = writer.sheets[sheet_name]

            # Header formatting
            for cell in ws[1]:

                cell.font = Font(
                    bold=True,
                    color="FFFFFF"
                )

                cell.fill = PatternFill(
                    start_color="183D32",
                    end_color="183D32",
                    fill_type="solid"
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            # Body formatting
            for row in ws.iter_rows(min_row=2):

                for cell in row:

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

            # Automatic column width
            for column in ws.columns:

                max_length = 0

                column_letter = column[0].column_letter

                for cell in column:

                    if cell.value is not None:

                        max_length = max(
                            max_length,
                            len(str(cell.value))
                        )

                ws.column_dimensions[column_letter].width = max_length + 4

    output.seek(0)

    return output.getvalue()


# ==========================================================
# SAVE FINAL MATRIX TO GOOGLE SHEETS
# ==========================================================


def save_to_google_sheets(final_matrix, course_code, department, course_name):

    try:
        script_url = st.secrets["GOOGLE_SCRIPT_URL"]
        secret_token = st.secrets["GOOGLE_SECRET_TOKEN"]
    except (KeyError, FileNotFoundError):
        st.error("Google Sheets connection details are missing.")
        return

    df = final_matrix.copy()

    # round to 2 decimals so cells show clean numbers
    num_cols = df.columns.drop("Assessment")
    df[num_cols] = df[num_cols].astype(float).round(2)

    # add Course Code as the first column (blank on the TOTAL row)
    df.insert(
        0, "Course Code",
        [course_code if a != "TOTAL" else "" for a in df["Assessment"]]
    )

    payload = {
        "token": secret_token,
        "department": department,
        "course_name": course_name,
        "course_code": course_code,
        "headers": df.columns.tolist(),
        "rows": df.astype(object).where(pd.notna(df), None).values.tolist()
    }

    try:
        response = requests.post(script_url, json=payload, timeout=30)
        response.raise_for_status()
        result = response.json()

        if result.get("status") == "success":
            st.success("Saved to Google Sheets successfully.")
        else:
            st.error("Google Sheets could not save the data: "
                     + str(result.get("message", "Unknown error")))

    except requests.exceptions.RequestException as error:
        st.error(f"Could not connect to Google Sheets: {error}")
    except ValueError:
        st.error("Apps Script returned an unexpected response.")

# ==========================================================
# SAVE TO GOOGLE SHEETS AND DOWNLOAD EXCEL
# ==========================================================

if result_df is not None:

    st.write("DEBUG: Final matrix exists:", result_df is not None)

    # Save to Google Sheets
    if st.button("Save to Google Sheets", key="save_to_google_sheets"):
        save_to_google_sheets(result_df, course_code, department, program)
    # Download Excel
    excel_data = create_excel(result_df, course_code)

    st.download_button(
        label="Download Excel",
        data=excel_data,
        file_name="CO_Allocation_Matrix.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )