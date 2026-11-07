import streamlit as st
import pandas as pd
import numpy as np

from io import BytesIO
from scipy.optimize import linprog
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


# ==================================================
# COURSE DETAILS
# ==================================================

COURSE_DETAILS = {

    "Department of Botany": [
        ("FUNDAMENTALS OF BOTANY", "BCM26DSCBO101", 5),
        ("PLANTS AND HUMAN WELL-BEING", "BCM26DSCBO102", 0),
        ("ECOTOURISM", "BCM26MDCBO101", 4)
    ],

    "Department of Chemistry": [
        ("Fundamentals of Chemistry-1", "BCM26DSCCH101", 5),
        ("Foundations of Food Biochemistry-1", "BCM26DSCCH102", 0),
        ("Food Chemistry and Nutrition", "BCM26MDCCH101", 4)
    ],

    "Department of Commerce": [
        ("Business Regulatory Framework", "BCM26DSCBL101", 5),
        ("Financial Accounting", "BCM26DSCCO101", 5),
        ("Basic Finance for Daily Life", "BCM26DSCCO102", 0),
        ("Business Statistics", "BCM26DSCQT101", 5),
        ("Event Management", "BCM26MDCCO101", 4),
        ("Personal Finance and Planning", "BCM26MDCCO102", 0),
        ("Business Mathematics", "BCM26MDCCO103", 0),
        ("Fundamentals of Business Management", "BCM26MDCCO104", 4)
    ],

    "Department of Commerce(SF)": [
        ("Fundamentals of Business Management", "BCM26MDCCO104", 4)
    ],

    "Department of Computer Science": [
        ("Programming using C", "BCM26DSCCS101", 5),
        ("Basics of Computer Science and AI", "BCM26DSCCS102", 0),
        ("Introduction to IT", "BCM26MDCCS101", 4)
    ],

    "Department of Economics": [
        ("Fundamentals of Economics I", "BCM26DSCEC101", 5),
        ("Economics of Health", "BCM26DSCEC102", 0),
        ("Economics and Finance in Everyday Life", "BCM26MDCEC101", 0)
    ],

    "Department of English": [
        ("English for Commerce Part I", "BCM26AECEN101", 4),
        ("English for Arts and Humanities Part I", "BCM26AECEN102", 4),
        ("English for Science Part I", "BCM26AECEN103", 4),
        ("AEC French Part I", "BCM26AECFR101", 4),
        ("AEC German Part I", "BCM26AECGE101", 4),
        ("Literary Genres:Poetry, Fiction and Folk Tales", "BCM26DSCEN101", 5),
        ("Malayalam Literature in English Translation", "BCM26DSCEN102", 0),
        ("Anglo-Saxon Aesthetics to Renaissance Rhetoric: Makers of English Literature", "BCM26DSCLH101", 5),
        ("Narratives of Humour", "BCM26MDCEN101", 0),
        ("Content Writing", "BCM26MDCEN102", 4)
    ],

    "Department of Food Science and Quality Control": [
        ("Fundamentals of Nutrition & Food Chemistry", "BCM26DSCFQ101", 5),
        ("Introduction to Food Science And Processing", "BCM26DSCFQ102", 0),
        ("Food and Culture", "BCM26MDCFQ101", 4)
    ],

    "Department of Hindi": [
        ("Theatre and Dialogue Presentation", "BCM26AECHN101", 0),
        ("The Role and Representation of Women in Hindi Cinema", "BCM26AECHN102", 4),
        ("Eminent Women Poets of Contemporary Hindi", "BCM26AECHN103", 4)
    ],

    "Department of History": [
        ("Ancient Indian History: From Prehistoric Cultures to Early Empires", "BCM26DSCHI101", 5),
        ("Modern Revolutions and Social Change", "BCM26DSCHI102", 0)
    ],

    "Department of Home Science": [
        ("Essentials of Home Science and Research Methodology", "BCM26DSCHS101", 5),
        ("Creative Arts and Crafts", "BCM26MDCHS101", 4)
    ],

    "Department of Malayalam": [
        ("Bhasha Sahithi", "BCM26AECML101", 4),
        ("Sarga Sahithi", "BCM26AECML102", 4),
        ("Sahitheemalika", "BCM26AECML103", 4),
        ("Gadhya Padangal: Kadha, Aathmakadha, Vimarsanam", "BCM26DSCML101", 5),
        ("Kadhanuroopanam", "BCM26DSCML102", 0),
        ("Keralathile Nadanpattukal", "BCM26MDCML101", 4)
    ],

    "Department of Mathematics": [
        ("Basic Mathematical Structures and Calculus", "BCM26DSCMA101", 5),
        ("Basic Concepts in Graph Theory", "BCM26DSCMA102", 0)
    ],

    "Department of Physical Education": [
        ("First Aid, Safety and Emergency Response", "BCM26MDCPE101", 4)
    ],

    "Department of Physics": [
        ("Introductory Physics", "BCM26DSCPH101", 5),
        ("Physics of Renewable Energy Sources", "BCM26DSCPH102", 0),
        ("Physics Around you", "BCM26MDCPH101", 0)
    ],

    "Department of Psychology": [
        ("Psychology for Personal Growth", "BCM26DSCPS101", 5),
        ("Introduction to Psychology", "BCM26DSCPY101", 0),
        ("Communicate with confidence", "BCM26MDCPS101", 4)
    ],

    "Department of Sociology": [
        ("Basics of Sociological Theory and Method", "BCM26DSCSO101", 5),
        ("Basic Concepts in Sociology", "BCM26DSCSO102", 5),
        ("Social Psychology: Exploring Self, Personality, Behaviour", "BCM26MDCSO101", 0)
    ],

    "Department of Statistics": [
        ("Fundamentals of Statistics and Probability", "BCM26DSCST101", 5),
        ("Statistical Data Analysis using Spreadsheets", "BCM26MDCST101", 4)
    ],

    "Department of Zoology": [
        ("Fascinating World of Animals", "BCM26DSCZO101", 5),
        ("Health, Nutrition and Human Well-Being", "BCM26DSCZO102", 0),
        ("Reproductive Health and Sexual Well-being", "BCM26MDCZO101", 4)
    ]
}


# ==================================================
# PROGRAM TYPES
# ==================================================

PROGRAM_TYPES = [
    "AEC",
    "AEC - OL",
    "Minor - B",
    "Major",
    "Minor - C",
    "Major Elective",
    "MDC"
]


# ==================================================
# SEMESTER CALCULATION
# ==================================================

def semester_from_code(code):

    digits = str(code)[-3:]

    if not digits.isdigit():
        return None

    number = int(digits)

    if number < 101 or number > 500:
        return None

    return (number - 101) // 50 + 1


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

col1, col2, col3 = st.columns(3)

with col1:

    departments = sorted(COURSE_DETAILS.keys())

    department = st.selectbox(
        "Department",
        departments
    )

with col2:

    available_courses = COURSE_DETAILS[department]

    course_names = [
        item[0] for item in available_courses
    ]

    program = st.selectbox(
        "Course Name",
        course_names
    )

with col3:

    program_type = st.selectbox(
        "Program Type",
        PROGRAM_TYPES
    )


selected_course = next(
    item for item in available_courses
    if item[0] == program
)

course_code = selected_course[1]

semester = semester_from_code(course_code)

semester_label = (
    f"Semester {semester}"
    if semester
    else "Not available"
)

col4, col5 = st.columns(2)

with col4:

    st.text_input(
        "Course Code",
        value=course_code,
        disabled=True
    )

with col5:

    st.text_input(
        "Semester",
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
# ALLOCATION FUNCTION
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

    # Check that every CO with positive target is selected
    for i, co in enumerate(co_list):

        if co_target_array[i] > 0 and not allowed[i, :].any():

            return None, (
                f"{co} has a positive target weight but is not "
                "selected in any assessment."
            )

    # Check that each positive assessment has a selected CO
    for j, assessment in enumerate(ASSESSMENT_NAMES):

        if assessment_targets[j] > 0 and not allowed[:, j].any():

            return None, (
                f"{assessment} has a positive weight but no CO is selected."
            )

    # Variables:
    # First n variables = allocation values
    # Next n variables = absolute deviations from entered allocation

    n = n_co * n_assessments

    seed = initial_matrix.flatten()

    objective = np.concatenate([
        np.zeros(n),
        np.ones(n)
    ])

    A_eq = []
    b_eq = []

    # Assessment row constraints
    for j in range(n_assessments):

        constraint = np.zeros(2 * n)

        for i in range(n_co):

            constraint[i * n_assessments + j] = 1

        A_eq.append(constraint)

        b_eq.append(assessment_targets[j])

    # CO column constraints
    for i in range(n_co):

        constraint = np.zeros(2 * n)

        for j in range(n_assessments):

            constraint[i * n_assessments + j] = 1

        A_eq.append(constraint)

        b_eq.append(co_target_array[i])

    A_ub = []
    b_ub = []

    # Absolute deviation constraints
    for k in range(n):

        row1 = np.zeros(2 * n)
        row1[k] = 1
        row1[n + k] = -1

        A_ub.append(row1)
        b_ub.append(seed[k])

        row2 = np.zeros(2 * n)
        row2[k] = -1
        row2[n + k] = -1

        A_ub.append(row2)
        b_ub.append(-seed[k])

    # Bounds
    bounds = []

    for i in range(n_co):

        for j in range(n_assessments):

            if allowed[i, j]:
                bounds.append((0, None))
            else:
                bounds.append((0, 0))

    # Deviation variables
    bounds += [(0, None)] * n

    result = linprog(
        c=objective,
        A_eq=np.array(A_eq),
        b_eq=np.array(b_eq),
        A_ub=np.array(A_ub),
        b_ub=np.array(b_ub),
        bounds=bounds,
        method="highs"
    )

    if not result.success:

        return None, (
            "The selected CO combinations cannot satisfy both "
            "assessment-wise and CO-wise totals. "
            "Please adjust the CO selections or CO weights."
        )

    allocation = result.x[:n].reshape(
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
# DOWNLOAD EXCEL
# ==========================================================

# Download is available only when the final matrix exists
if result_df is not None:

    excel_data = create_excel(
        result_df,
        course_code
    )

    st.download_button(
        label="Download Excel",
        data=excel_data,
        file_name="CO_Allocation_Matrix.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )