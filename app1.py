import streamlit as st
import pandas as pd
import sqlite3
from datetime import date, time
import openpyxl
import uuid
from datetime import date

st.set_page_config(page_title="EDC Form", layout="wide")

st.title("Electronic Data Capture (EDC) Form")

# --------------------------------------------------
# Demographic Information
# --------------------------------------------------
st.header("Demographic Information")

col1, col2 = st.columns(2)

with col1:
    if "unique_id" not in st.session_state:
        st.session_state.unique_id = str(uuid.uuid4())[:8].upper()

    unique_id = st.text_input(
        "1. Unique Identifier",
        value=st.session_state.unique_id,
        disabled=True
        )

    dob = st.date_input("2. Date of Birth")
    today = date.today()

    age = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))

    st.number_input(
    "3. Age",
    value=age,
    disabled=False
    )

with col2:
    sex = st.selectbox(
        "4. Sex",
        ["", "Male", "Female", "Other"]
    )

    race = st.multiselect(
        "5. Race (Select all that apply)",
        [
            "American Indian or Alaska Native",
            "Asian",
            "Black or African American",
            "Native Hawaiian or Other Pacific Islander",
            "White",
            "Some Other Race",
            "Unknown",
            "Prefer Not to Answer"
        ]
    )

    ethnicity = st.selectbox(
    "6. Ethnicity",
    [
    "",
    "Hispanic or Latino",
    "Not Hispanic or Latino",
    "Unknown",
    "Prefer Not to Answer"
    ]
    )

# --------------------------------------------------
# Sample Information
# --------------------------------------------------
st.header("Sample Information")

tube_type = st.selectbox(
    "7. Tube Type",
    [
        "",
        "Serum Separator Tube (SST, Gold/Tiger Top)",
        "Serum Tube (Red Top)",
        "K2 EDTA (Lavender Top)",
        "K3 EDTA (Lavender Top)",
        "Lithium Heparin PST (Light Green Top)",
        "Lithium Heparin (Green Top)",
        "Sodium Heparin (Green Top)",
        "Sodium Citrate 3.2% (Light Blue Top)",
        "Fluoride/Oxalate (Gray Top)",
        "PAXgene RNA Tube",
        "PAXgene DNA Tube",
        "BD Barricor Plasma Tube",
        "CPT Cell Preparation Tube",
        "Microtainer EDTA",
        "Microtainer Serum",
        "Other"
    ]
)

draw_date = st.date_input(
    "8. Draw Date",
    key="draw_date"
)

draw_time = st.time_input(
    "9. Draw Time",
    value=None
)

centrifugation_date = st.date_input(
    "10. Centrifugation Date"
)

centrifugation_time = st.time_input(
    "11. Centrifugation Time",
    key="centrifugation_time",
    value=None
)

g_minutes = st.selectbox(
    "12. G-Minutes",
    ["", 500, 1000, 1500, 2000, 2500]
)

centrifugation_speed = st.number_input(
    "13. Centrifugation Speed (RPM/RCF)",
    min_value=0,
    step=1,
    format="%d"
)

centrifugation_duration = st.number_input(
    "14. Centrifugation Duration (minutes)",
    min_value=0,
    format="%d"
)

freeze_date = st.date_input(
    "15. Freeze Date"
)

freeze_time = st.time_input(
    "16. Freeze Time",
    key="freeze_time",
    value=None
)

num_aliquots = st.number_input(
    "17. Number of Aliquots",
    min_value=0,
    step=1
)

freeze_thaw_cycles = st.number_input(
    "18. Number of Freeze/Thaw Cycles",
    min_value=0,
    step=1
)

# --------------------------------------------------
# Instrument Information
# --------------------------------------------------
st.header("Instrument Information")

testing_instrument = st.text_input(
    "19. Testing Instrument"
)

assay_name = st.text_input(
    "20. Assay Name"
)

sid = st.text_input(
    "21. SID"
)

test_completion_date = st.date_input(
    "22. Test Completion Date"
)

test_completion_time = st.time_input(
    "23. Test Completion Time",
    key="completion_time",
    value=None
)

reported_result = st.text_input(
    "24. Reported Result"
)

result_units = st.text_input(
    "25. Result Units"
)

result_flags = st.text_input(
    "26. Result Flags"
)

reagent_lot = st.text_input(
    "27. Reagent Lot"
)

dilution = st.text_input(
    "28. Dilution"
)

# --------------------------------------------------
# Save
# --------------------------------------------------
if st.button("Save Record"):

    data = {
        "UniqueIdentifier": unique_id,
        "DateOfBirth": str(dob),
        "Age": age,
        "Sex": sex,
        "Race": race,
        "Ethnicity": ethnicity,

        "TubeType": tube_type,
        "DrawDate": str(draw_date),
        "DrawTime": str(draw_time),
        "CentrifugationDate": str(centrifugation_date),
        "CentrifugationTime": str(centrifugation_time),
        "GMinutes": g_minutes,
        "CentrifugationSpeed": centrifugation_speed,
        "CentrifugationDuration": centrifugation_duration,
        "FreezeDate": str(freeze_date),
        "FreezeTime": str(freeze_time),
        "NumberAliquots": num_aliquots,
        "FreezeThawCycles": freeze_thaw_cycles,

        "TestingInstrument": testing_instrument,
        "AssayName": assay_name,
        "SID": sid,
        "TestCompletionDate": str(test_completion_date),
        "TestCompletionTime": str(test_completion_time),
        "ReportedResult": reported_result,
        "ResultUnits": result_units,
        "ResultFlags": result_flags,
        "ReagentLot": reagent_lot,
        "Dilution": dilution
    }

    df = pd.DataFrame([data])

    # Save to SQLite
    conn = sqlite3.connect("edc.db")
    df.to_sql(
        "edc_records",
        conn,
        if_exists="append",
        index=False
    )
    conn.close()

    st.success("Record saved successfully.")

# --------------------------------------------------
# Export
# --------------------------------------------------
if st.button("Export to Excel"):

    conn = sqlite3.connect("edc.db")

    export_df = pd.read_sql(
        "SELECT * FROM edc_records",
        conn
    )

    export_df.to_excel(
        "EDC_Export.xlsx",
        index=False
    )

    conn.close()

    st.success("Exported to EDC_Export.xlsx")