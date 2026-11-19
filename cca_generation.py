import io
import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)
from reportlab.lib.styles import getSampleStyleSheet


# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="CCA Generation",
    page_icon="📊",
    layout="wide"
)

st.title("CCA Generation")
st.caption("Phase 2 – Assessment Rescaling and CO Consolidation")


# ============================================================
# 2. THEME
# ============================================================

st.markdown("""
<style>

/* Main background */
.stApp {
    background-color: #F3F6F1;
    color: #183D32;
}

/* Main content */
.main .block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

/* Headings */
h1, h2, h3, h4 {
    color: #183D32 !important;
    font-weight: 700 !important;
}

/* Normal text and labels */
p, label, span {
    color: #263C32;
}

/* Text input boxes */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stDateInput"] input,
div[data-testid="stTextArea"] textarea,
div[data-baseweb="input"] input,
div[data-baseweb="textarea"] textarea {

    background-color: #FFFFFF !important;
    color: #183D32 !important;
    -webkit-text-fill-color: #183D32 !important;
    border: 1.5px solid #789887 !important;
    border-radius: 7px !important;
    font-weight: 600 !important;
}

/* Placeholder text */
input::placeholder,
textarea::placeholder {
    color: #718477 !important;
    -webkit-text-fill-color: #718477 !important;
    opacity: 1 !important;
}

/* Dropdown boxes */
div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    color: #183D32 !important;
    border: 1.5px solid #789887 !important;
    border-radius: 7px !important;
}

/* Selected dropdown text */
div[data-baseweb="select"] span {
    color: #183D32 !important;
}

/* Dropdown menu */
div[data-baseweb="popover"],
div[data-baseweb="popover"] ul {
    background-color: #FFFFFF !important;
}

div[data-baseweb="popover"] li {
    color: #183D32 !important;
}

/* Data tables */
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"] {
    background-color: #FFFFFF !important;
    border: 1px solid #A6B9AA !important;
    border-radius: 10px !important;
}

/* Buttons */
.stButton button,
.stDownloadButton button {
    background-color: #183D32 !important;
    color: #FFFFFF !important;
    border: 1px solid #183D32 !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}

/* Button hover */
.stButton button:hover,
.stDownloadButton button:hover {
    background-color: #285C49 !important;
    color: #FFFFFF !important;
}

/* Radio buttons */
div[data-testid="stRadio"] label {
    color: #183D32 !important;
}

/* File uploader */
div[data-testid="stFileUploader"] {
    background-color: #FFFFFF !important;
    border: 1px solid #789887 !important;
    border-radius: 10px !important;
    padding: 12px;
}

/* Metric cards */
div[data-testid="stMetric"] {
    background-color: #FFFFFF !important;
    border: 1px solid #A6B9AA !important;
    padding: 15px;
    border-radius: 10px;
}

/* Metric values */
div[data-testid="stMetricValue"] {
    color: #183D32 !important;
}

/* Captions */
.stCaption {
    color: #526B5D !important;
}

/* Alerts */
div[data-testid="stAlert"] {
    border-radius: 8px;
}

/* Divider */
hr {
    border-color: #A6B9AA !important;
}

</style>
""", unsafe_allow_html=True)

# ============================================================
# 3. CONSTANTS
# ============================================================

ASSESSMENTS = ["PA1", "PA2", "LM", "LE", "ND"]

ND_REPLACEMENT_OPTIONS = [
    "ND",
    "PA1",
    "PA2",
    "LM",
    "LE"
]

RESCALING_OPTIONS = [
    "Absolute",
    "Relative"
]

CLUBBING_OPTIONS = [
    "Average of All",
    "Average of Best",
    "Weighted Average"
]

REQUIRED_COLUMNS = [
    "Reg No",
    "Admission No",
    "Student Name",
    "Student Prog",
    "Course Name",
    "Course Code",
    "Course Type",
    "Assessment",
    "CO",
    "Mark Obtained",
    "Max Marks"
]


# ============================================================
# 4. SESSION STATE
# ============================================================

defaults = {
    "nd_confirmed": False,
    "nd_mapping": None,
    "edited_config": None,
    "generated": False,
    "finalized": False,
    "generation_results": None,
    "rescaled_results": None,
    "clubbed_results": None,
    "final_matrix": None
}

for key, value in defaults.items():

    if key not in st.session_state:
        st.session_state[key] = value


# ============================================================
# 5. SAMPLE DATA
# ============================================================

def sample_data():

    students = [
        ("UG26016", "AD20261105", "Adithya Raj", "BSc Mathematics"),
        ("UG26017", "AD20261106", "Anjali S", "BSc Mathematics"),
        ("UG26018", "AD20261107", "Biju K", "BSc Mathematics"),
        ("UG26019", "AD20261108", "Devika R", "BSc Mathematics"),
        ("UG26020", "AD20261109", "Farhan M", "BSc Mathematics"),
        ("UG26021", "AD20261110", "Gopika P", "BSc Mathematics")
    ]

    exams = [
        ("PA1", "PA1-01", "2026-08-10", ["CO1", "CO2", "CO3"], 20),
        ("PA1", "PA1-02", "2026-08-17", ["CO1", "CO2", "CO3"], 20),
        ("PA2", "PA2-01", "2026-09-10", ["CO1", "CO2", "CO3"], 25),
        ("LM", "LM-01", "2026-09-20", ["CO2", "CO3", "CO4"], 30),
        ("LE", "LE-01", "2026-10-10", ["CO1", "CO2", "CO3", "CO4"], 25),
        ("ND", "ND-01", "2026-10-20", ["CO1", "CO2"], 20)
    ]

    rows = []

    for i, student in enumerate(students):

        reg, adm, name, prog = student

        for j, exam in enumerate(exams):

            assessment, exam_id, exam_date, cos, maximum = exam

            for k, co in enumerate(cos):

                fraction = (
                    0.50 + ((i * 7 + j * 5 + k * 3) % 41) / 100
                )

                mark = round(maximum * fraction, 1)

                rows.append([
                    reg,
                    adm,
                    name,
                    prog,
                    "Research Methodology",
                    "ST040303",
                    "Major",
                    assessment,
                    co,
                    mark,
                    maximum,
                    exam_date,
                    exam_id
                ])

    columns = REQUIRED_COLUMNS + [
        "Date",
        "Assessment ID"
    ]

    return pd.DataFrame(rows, columns=columns)


# ============================================================
# 6. DATA CLEANING
# ============================================================

def clean_data(df):

    df = df.copy()

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    missing = [
        c for c in REQUIRED_COLUMNS
        if c not in df.columns
    ]

    if missing:

        raise ValueError(
            "Missing columns: " + ", ".join(missing)
        )

    if "Date" not in df.columns:
        df["Date"] = pd.NaT

    if "Assessment ID" not in df.columns:
        df["Assessment ID"] = ""

    df = df[
        REQUIRED_COLUMNS + ["Date", "Assessment ID"]
    ].copy()

    text_columns = [
        "Reg No",
        "Admission No",
        "Student Name",
        "Student Prog",
        "Course Name",
        "Course Code",
        "Course Type",
        "Assessment",
        "CO",
        "Assessment ID"
    ]

    for col in text_columns:

        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df["Assessment"] = df["Assessment"].str.upper()
    df["CO"] = df["CO"].str.upper()

    df["Mark Obtained"] = pd.to_numeric(
        df["Mark Obtained"],
        errors="coerce"
    )

    df["Max Marks"] = pd.to_numeric(
        df["Max Marks"],
        errors="coerce"
    )

    df["Date"] = pd.to_datetime(
        df["Date"],
        errors="coerce"
    )

    df = df.dropna(
        subset=["Mark Obtained", "Max Marks"]
    ).copy()

    df = df[
        (df["Max Marks"] > 0)
        & (df["Mark Obtained"] >= 0)
        & (df["Mark Obtained"] <= df["Max Marks"])
    ].copy()

    invalid = sorted(
        set(df["Assessment"]) - set(ASSESSMENTS)
    )

    if invalid:

        raise ValueError(
            "Invalid assessments: " + ", ".join(invalid)
        )

    if df.empty:
        raise ValueError("No valid marks found.")

    missing_id = df["Assessment ID"].eq("")

    if missing_id.any():

        df.loc[missing_id, "Assessment ID"] = (
            df.loc[missing_id, "Assessment"]
            + "-01"
        )

    return df.reset_index(drop=True)


# ============================================================
# 7. EXCEL EXPORT
# ============================================================

def create_excel(sheets):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet_name, df in sheets.items():

            sheet = sheet_name[:31]

            df.to_excel(
                writer,
                sheet_name=sheet,
                index=False
            )

            ws = writer.sheets[sheet]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            for column in ws.columns:

                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column
                )

                ws.column_dimensions[
                    column[0].column_letter
                ].width = min(max_length + 3, 30)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 8. PDF EXPORT
# ============================================================

def create_pdf(matrix, department, course, course_code):

    output = io.BytesIO()

    styles = getSampleStyleSheet()

    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=25,
        rightMargin=25,
        topMargin=25,
        bottomMargin=25
    )

    rows = [
        [str(col) for col in matrix.columns]
    ] + matrix.fillna("").astype(str).values.tolist()

    table = Table(
        rows,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#183D32")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE")
        ])
    )

    story = [
        Paragraph("FINAL CCA MATRIX", styles["Title"]),
        Spacer(1, 10),
        Paragraph(f"Department: {department}", styles["Normal"]),
        Paragraph(f"Course: {course}", styles["Normal"]),
        Paragraph(f"Course Code: {course_code}", styles["Normal"]),
        Spacer(1, 15),
        table
    ]

    document.build(story)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 9. STEP 1 – UPLOAD DATA
# ============================================================

st.header("Step 1 – Upload Assessment Data")

source = st.radio(
    "Select Data Source",
    ["Upload Phase 1 File", "Use Sample Data"],
    horizontal=True
)

if source == "Upload Phase 1 File":

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV file",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file is None:

        st.info("Please upload your Phase 1 marks file.")

        st.stop()

    try:

        if uploaded_file.name.lower().endswith(".csv"):
            raw_data = pd.read_csv(uploaded_file)
        else:
            raw_data = pd.read_excel(uploaded_file)

        data = clean_data(raw_data)

    except Exception as error:

        st.error(f"File error: {error}")

        st.stop()

else:

    data = clean_data(sample_data())


# ============================================================
# 10. STEP 2 – COURSE DETAILS
# ============================================================

st.divider()

st.header("Step 2 – Course Details")

first = data.iloc[0]

c1, c2, c3 = st.columns(3)

department = c1.text_input(
    "Department",
    "Mathematics"
)

course = c2.text_input(
    "Course Name",
    first["Course Name"]
)

course_code = c3.text_input(
    "Course Code",
    first["Course Code"]
)

c1, c2 = st.columns(2)

c1.metric(
    "Total Students",
    data["Reg No"].nunique()
)

c2.metric(
    "Total Records",
    len(data)
)


# ============================================================
# 11. STEP 3 – ND REPLACEMENT
# ============================================================

st.divider()

st.header("Step 3 – ND Assessment Replacement")

st.write(
    "Select a replacement assessment for ND, if required. "
    "If no replacement is needed, leave ND selected."
)

course_details = (
    data[
        ["Course Code", "Course Name", "Course Type"]
    ]
    .drop_duplicates(subset=["Course Code"])
    .reset_index(drop=True)
)

course_details["Replace ND"] = "ND"

if st.session_state.nd_confirmed:

    edited_nd = st.session_state["nd_mapping"]

    st.dataframe(
        edited_nd,
        use_container_width=True,
        hide_index=True
    )

else:

    edited_nd = st.data_editor(
        course_details,
        use_container_width=True,
        hide_index=True,

        disabled=[
            "Course Code",
            "Course Name",
            "Course Type"
        ],

        column_config={

            "Replace ND":
                st.column_config.SelectboxColumn(
                    "Replace ND",
                    options=ND_REPLACEMENT_OPTIONS,
                    required=True
                )
        },

        key="nd_editor"
    )

st.info(
    "Please review the ND replacement table. "
    "ND will remain unchanged for any course where you do not select "
    "another assessment."
)

if not st.session_state.nd_confirmed:

    if st.button(
        "Confirm ND Replacement and Continue",
        type="primary",
        use_container_width=True
    ):

        st.session_state["nd_mapping"] = edited_nd.copy()

        st.session_state["nd_confirmed"] = True

        st.session_state["generated"] = False

        st.session_state["finalized"] = False

        st.success("ND replacement confirmed.")

        st.rerun()


# ============================================================
# 12. APPLY ND REPLACEMENT
# ============================================================

if not st.session_state.nd_confirmed:

    st.warning(
        "Complete Step 3 and confirm the ND selection to proceed to Step 4."
    )

    st.stop()

nd_mapping = dict(
    zip(
        st.session_state["nd_mapping"]["Course Code"],
        st.session_state["nd_mapping"]["Replace ND"]
    )
)

data["Assessment"] = data.apply(

    lambda row: (
        nd_mapping.get(row["Course Code"], "ND")
        if row["Assessment"] == "ND"
        else row["Assessment"]
    ),

    axis=1
)


# ============================================================
# 13. STEP 4 – STUDENT-WISE MARKS
# ============================================================

st.divider()

st.header("Step 4 – Assessment Generation")

st.subheader("Table 1 – Student-wise Assessment Marks")

student_columns = [
    "Student Name",
    "Reg No",
    "Admission No",
    "Course Name",
    "Course Code",
    "Course Type",
    "Assessment",
    "Assessment ID",
    "Date",
    "CO",
    "Mark Obtained",
    "Max Marks"
]

student_marks = data[student_columns].copy()

student_marks = student_marks.sort_values(
    [
        "Student Name",
        "Assessment",
        "Assessment ID",
        "CO"
    ]
)

st.dataframe(
    student_marks,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# 14. CREATE COMBINED CONFIGURATION
# ============================================================

st.divider()

st.subheader("Table 2 – Combined Assessment Configuration")

st.caption(
    "Configure the rescaling method and CO calculation method "
    "in one table."
)

combined_config = (
    data[
        [
            "Assessment",
            "Assessment ID",
            "Date",
            "CO"
        ]
    ]
    .drop_duplicates()
    .sort_values(
        ["Assessment", "Assessment ID", "Date", "CO"]
    )
    .reset_index(drop=True)
)

combined_config["Rescaling Method"] = "Absolute"

combined_config["CO Calculation Method"] = "Average of All"

combined_config["Best N"] = None

combined_config["Weight"] = None


# ============================================================
# 15. COMBINED CONFIGURATION TABLE EDITOR
# ============================================================

if not st.session_state.finalized:

    column_config = {

        "Assessment":
            st.column_config.TextColumn(
                "Assessment",
                disabled=True
            ),

        "Assessment ID":
            st.column_config.TextColumn(
                "Assessment ID",
                disabled=True
            ),

        "Date":
            st.column_config.DateColumn(
                "Date",
                disabled=True
            ),

        "CO":
            st.column_config.TextColumn(
                "CO",
                disabled=True
            ),

        "Rescaling Method":
            st.column_config.SelectboxColumn(
                "Rescaling Method",
                options=RESCALING_OPTIONS,
                required=True
            ),

        "CO Calculation Method":
            st.column_config.SelectboxColumn(
                "CO Calculation Method",
                options=CLUBBING_OPTIONS,
                required=True
            ),

        "Best N":
            st.column_config.NumberColumn(
                "Best N",
                min_value=1,
                step=1,
                format="%d"
            ),

        "Weight":
            st.column_config.NumberColumn(
                "Weight",
                min_value=0.0,
                step=0.5,
                format="%.2f"
            )
    }

    edited_config = st.data_editor(
        combined_config,
        use_container_width=True,
        hide_index=True,

        disabled=[
            "Assessment",
            "Assessment ID",
            "Date",
            "CO"
        ],

        column_config=column_config,

        key="combined_config_editor"
    )

else:

    edited_config = st.session_state["edited_config"]

    st.dataframe(
        edited_config,
        use_container_width=True,
        hide_index=True
    )


st.info(
    "Absolute: Mark Obtained / Max Marks × 100. "
    "Relative: Mark Obtained / Class Average × 100. "
    "Average of All: averages all Assessment IDs. "
    "Average of Best: enter Best N. "
    "Weighted Average: enter a separate weight for each Assessment ID. "
    "Keep Best N and Weight blank when they are not required."
)


# ============================================================
# 16. CONFIGURATION VALIDATION
# ============================================================

def validate_combined_config(config):

    required_columns = [
        "Rescaling Method",
        "CO Calculation Method",
        "Best N",
        "Weight"
    ]

    for col in required_columns:

        if col not in config.columns:

            raise ValueError(
                f"Missing configuration column: {col}"
            )

    if config["Rescaling Method"].isna().any():

        raise ValueError(
            "Select a rescaling method for every row."
        )

    if config["CO Calculation Method"].isna().any():

        raise ValueError(
            "Select a CO calculation method for every row."
        )

    for (assessment, co), group in config.groupby(
        ["Assessment", "CO"]
    ):

        methods = (
            group["CO Calculation Method"]
            .dropna()
            .unique()
        )

        if len(methods) != 1:

            raise ValueError(
                f"Select the same CO Calculation Method for "
                f"all Assessment IDs of {assessment} – {co}."
            )

        method = methods[0]

        if method == "Average of Best":

            best_values = pd.to_numeric(
                group["Best N"],
                errors="coerce"
            ).dropna().unique()

            if len(best_values) == 0:

                raise ValueError(
                    f"Enter Best N for {assessment} – {co}."
                )

            if len(best_values) > 1:

                raise ValueError(
                    f"Use the same Best N for all Assessment IDs "
                    f"of {assessment} – {co}."
                )

        elif method == "Weighted Average":

            weights = pd.to_numeric(
                group["Weight"],
                errors="coerce"
            )

            if weights.isna().any():

                raise ValueError(
                    f"Enter a weight for every Assessment ID "
                    f"of {assessment} – {co}."
                )

            if (weights <= 0).any():

                raise ValueError(
                    f"Weights must be positive for {assessment} – {co}."
                )

    return True


# ============================================================
# 17. RESCALING CALCULATION
# ============================================================

def calculate_rescaling(data, settings):

    result = data.copy()

    rescaling_settings = settings[
        [
            "Assessment",
            "Assessment ID",
            "Date",
            "CO",
            "Rescaling Method"
        ]
    ].copy()

    result = result.merge(
        rescaling_settings,
        on=[
            "Assessment",
            "Assessment ID",
            "Date",
            "CO"
        ],
        how="left"
    )

    if result["Rescaling Method"].isna().any():

        raise ValueError(
            "Rescaling method is missing for some records."
        )

    result["Class Average"] = (
        result.groupby(
            [
                "Assessment",
                "Assessment ID",
                "Date",
                "CO"
            ]
        )["Mark Obtained"].transform("mean")
    )

    result["Rescaled Value"] = 0.0

    for index, row in result.iterrows():

        obtained = float(row["Mark Obtained"])

        maximum = float(row["Max Marks"])

        if maximum <= 0:
            continue

        if row["Rescaling Method"] == "Absolute":

            value = obtained / maximum * 100

        else:

            average = float(row["Class Average"])

            if average > 0:

                value = obtained / average * 100

            else:

                value = 0

        result.loc[index, "Rescaled Value"] = min(
            max(value, 0),
            100
        )

    result["Rescaled Value"] = (
        result["Rescaled Value"].round(2)
    )

    return result


# ============================================================
# 18. CO-WISE CLUBBING
# ============================================================

def calculate_clubbing(rescaled, config):

    student_columns = [
        "Reg No",
        "Admission No",
        "Student Name"
    ]

    results = []

    grouped = rescaled.groupby(
        student_columns + ["Assessment", "CO"],
        dropna=False
    )

    for group_key, group in grouped:

        reg = group_key[0]
        admission = group_key[1]
        name = group_key[2]
        assessment = group_key[3]
        co = group_key[4]

        setting = config[
            (config["Assessment"] == assessment)
            & (config["CO"] == co)
        ]

        if setting.empty:

            raise ValueError(
                f"Configuration missing for {assessment} – {co}."
            )

        methods = (
            setting["CO Calculation Method"]
            .dropna()
            .unique()
        )

        if len(methods) != 1:

            raise ValueError(
                f"Use the same CO Calculation Method for "
                f"all Assessment IDs of {assessment} – {co}."
            )

        method = methods[0]

        # Calculate the average for each Assessment ID first.

        id_values = (
            group.groupby("Assessment ID")["Rescaled Value"]
            .mean()
            .reset_index()
        )

        if method == "Average of All":

            final_value = id_values["Rescaled Value"].mean()

        elif method == "Average of Best":

            best_values = pd.to_numeric(
                setting["Best N"],
                errors="coerce"
            ).dropna().unique()

            if len(best_values) != 1:

                raise ValueError(
                    f"Enter one common Best N for {assessment} – {co}."
                )

            best_n = int(best_values[0])

            if best_n < 1 or best_n > len(id_values):

                raise ValueError(
                    f"Best N for {assessment} – {co} must be between "
                    f"1 and {len(id_values)}."
                )

            final_value = (
                id_values["Rescaled Value"]
                .nlargest(best_n)
                .mean()
            )

        else:

            weighted_total = 0.0
            total_weight = 0.0

            for _, id_row in id_values.iterrows():

                assessment_id = id_row["Assessment ID"]

                matching_setting = setting[
                    setting["Assessment ID"] == assessment_id
                ]

                if matching_setting.empty:

                    raise ValueError(
                        f"Weight setting missing for {assessment} – "
                        f"{co} – {assessment_id}."
                    )

                weights = pd.to_numeric(
                    matching_setting["Weight"],
                    errors="coerce"
                ).dropna().unique()

                if len(weights) != 1 or weights[0] <= 0:

                    raise ValueError(
                        f"Enter one positive weight for "
                        f"{assessment} – {co} – {assessment_id}."
                    )

                weight = float(weights[0])

                weighted_total += (
                    float(id_row["Rescaled Value"]) * weight
                )

                total_weight += weight

            if total_weight <= 0:

                raise ValueError(
                    f"Total weight cannot be zero for {assessment} – {co}."
                )

            final_value = weighted_total / total_weight

        results.append({

            "Reg No": reg,

            "Admission No": admission,

            "Student Name": name,

            "Assessment": assessment,

            "CO": co,

            "CO Calculation Method": method,

            "Number of Assessment IDs": len(id_values),

            "Clubbing Value": round(final_value, 2)

        })

    return pd.DataFrame(results)


# ============================================================
# 19. GENERATE RESULTS
# ============================================================

st.divider()

if not st.session_state.finalized:

    if st.button(
        "Generate CCA Results",
        type="primary",
        use_container_width=True
    ):

        try:

            validate_combined_config(edited_config)

            rescaled = calculate_rescaling(
                data,
                edited_config
            )

            clubbed = calculate_clubbing(
                rescaled,
                edited_config
            )

            generated = rescaled.merge(

                clubbed[
                    [
                        "Reg No",
                        "Admission No",
                        "Student Name",
                        "Assessment",
                        "CO",
                        "Clubbing Value"
                    ]
                ],

                on=[
                    "Reg No",
                    "Admission No",
                    "Student Name",
                    "Assessment",
                    "CO"
                ],

                how="left"
            )

            st.session_state["rescaled_results"] = rescaled

            st.session_state["clubbed_results"] = clubbed

            st.session_state["generation_results"] = generated

            st.session_state["edited_config"] = edited_config.copy()

            st.session_state["generated"] = True

            st.success("CCA results generated successfully!")

            st.rerun()

        except Exception as error:

            st.error(str(error))


# ============================================================
# 20. STEP 5 – GENERATED STUDENT-WISE RESULTS
# ============================================================

if st.session_state.generated:

    st.divider()

    st.header("Step 5 – Generated Results")

    generated_results = st.session_state["generation_results"]

    display_columns = [
        "Student Name",
        "Reg No",
        "Admission No",
        "Assessment",
        "Assessment ID",
        "Date",
        "CO",
        "Mark Obtained",
        "Max Marks",
        "Rescaled Value",
        "Clubbing Value"
    ]

    student_results = generated_results[
        display_columns
    ].copy()

    student_results = student_results.sort_values(
        [
            "Student Name",
            "Assessment",
            "Assessment ID",
            "CO"
        ]
    )

    st.dataframe(
        student_results,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Download Generated Results")

    excel_data = create_excel({

        "Student Wise Results": student_results,

        "Combined Configuration": st.session_state["edited_config"],

        "Rescaled Results": st.session_state["rescaled_results"],

        "Clubbing Results": st.session_state["clubbed_results"]

    })

    st.download_button(
        "Download Generated Results (Excel)",
        data=excel_data,
        file_name="CCA_Generation_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


# ============================================================
# 21. STEP 6 – FINALIZE CCA
# ============================================================

if st.session_state.generated and not st.session_state.finalized:

    st.divider()

    st.header("Step 6 – Finalize CCA")

    st.warning(
        "Finalizing will lock the generation settings."
    )

    if st.button(
        "Finalize CCA",
        type="primary",
        use_container_width=True
    ):

        clubbed = st.session_state["clubbed_results"]

        matrix = clubbed.pivot_table(

            index=[
                "Reg No",
                "Admission No",
                "Student Name"
            ],

            columns="CO",

            values="Clubbing Value",

            aggfunc="mean"
        )

        matrix.columns = [
            f"{co} Final Mark"
            for co in matrix.columns
        ]

        co_columns = list(matrix.columns)

        matrix["Overall CCA"] = (
            matrix[co_columns].mean(axis=1)
        )

        matrix = matrix.round(2).reset_index()

        st.session_state["final_matrix"] = matrix

        st.session_state["finalized"] = True

        st.success("CCA finalized successfully!")

        st.rerun()


# ============================================================
# 22. FINAL CCA MATRIX
# ============================================================

if st.session_state.finalized:

    st.divider()

    st.header("Final CCA Matrix")

    matrix = st.session_state["final_matrix"]

    st.dataframe(
        matrix,
        use_container_width=True,
        hide_index=True
    )

    st.metric(
        "Total Students",
        len(matrix)
    )

    # PDF EXPORT

    pdf_data = create_pdf(
        matrix,
        department,
        course,
        course_code
    )

    st.download_button(
        "Download Final CCA Matrix PDF",
        data=pdf_data,
        file_name="Final_CCA_Matrix.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    # EXCEL EXPORT

    final_excel = create_excel({

        "Final CCA Matrix": matrix,

        "Student Wise Results": st.session_state["generation_results"],

        "Combined Configuration": st.session_state["edited_config"],

        "Rescaled Results": st.session_state["rescaled_results"],

        "Clubbing Results": st.session_state["clubbed_results"]

    })

    st.download_button(
        "Download Final CCA Excel",
        data=final_excel,
        file_name="Final_CCA_Matrix.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


# ============================================================
# 23. FOOTER
# ============================================================

st.divider()

st.caption(
    "CCA Generation | Assessment Rescaling and CO Consolidation"
)
