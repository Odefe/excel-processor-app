import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Alignment

st.set_page_config(
    page_title="Excel Processor",
    page_icon="📂",
    layout="centered"
)

st.markdown(
    """
    <div style="text-align:center; padding-top:70px;">
        <h1>Upload File Here</h1>
        <p>Upload your spreadsheet and download the processed output.</p>
        <div style="font-size:70px;">📂</div>
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "",
    type=["xlsx"]
)

def format_value(v):

    if pd.isna(v):
        return ""

    if isinstance(v, pd.Timestamp):

        if (
            v.hour == 0 and
            v.minute == 0 and
            v.second == 0
        ):
            return v.strftime("%d/%m/%Y")

        return v.strftime("%H:%M")

    value = str(v).strip()

    if value == "":
        return ""

    try:

        parsed = pd.to_datetime(
            value,
            dayfirst=True,
            errors="raise"
        )

        if (
            parsed.hour == 0 and
            parsed.minute == 0 and
            parsed.second == 0
        ):
            return parsed.strftime("%d/%m/%Y")

        return parsed.strftime("%H:%M")

    except:
        pass

    try:

        parsed = pd.to_datetime(
            value,
            dayfirst=False,
            errors="raise"
        )

        if (
            parsed.hour == 0 and
            parsed.minute == 0 and
            parsed.second == 0
        ):
            return parsed.strftime("%d/%m/%Y")

        return parsed.strftime("%H:%M")

    except:
        pass

    if isinstance(v, float):

        if v.is_integer():
            return str(int(v))

        return str(v)

    return value

def build_single_record(row, fields):

    values = []

    for field in fields:
        values.append(format_value(row.get(field, "")))

    if all(v == "" for v in values):
        return ""

    return ":".join(values)

def build_multi_record(row, groups):

    records = []

    for group in groups:

        values = []

        for field in group:
            values.append(format_value(row.get(field, "")))

        if any(v != "" for v in values):
            records.append(":".join(values))

    return "|".join(records)

def process_file(file):

    df = pd.read_excel(file)

    output = pd.DataFrame()

    final_headers = [
        "Agency",
        "Policy number",
        "Inception Date",
        "Expiry Date",
        "Seq Number",
        "Quote Reference",
        "Client Reference",
        "Transaction Effective Date",
        "Transaction Effective Time",
        "Transaction End Date",
        "Transaction Type",
        "Scheme",
        "Voluntary Excess",
        "Total Policy Premium",
        "Net Policy Premium",
        "Commission %",
        "Commission amount ",
        "ISB %",
        "Policyholder Type",
        "Proposer Title",
        "Proposer First Name",
        "Proposer Initials",
        "Proposer Surname",
        "Proposer Date of Birth",
        "Proposer Gender",
        "Policyholder Company Name",
        "Policyholder Company Business",
        "Date Business Established",
        "Abode Type",
        "Address line 1",
        "Address line 2",
        "Address line 3",
        "Address line 4",
        "Address line 5",
        "Address county",
        "Address postcode",
        "Address country",
        "Driver 1 Details",
        "Driver 1 Claims",
        "Driver 1 Convictions",
        "Driver 1 Non Motoring Convictions",
        "Driver 1 Medical Conditions",
        "Driver 2 Details",
        "Driver 2 Claims",
        "Driver 2 Convictions",
        "Driver 2 Non Motoring convictions",
        "Driver 2 Medical Conditions",
        "Driver 3 Details",
        "Driver 3 Claims",
        "Driver 3 Convictions",
        "Driver 3 Non Motoring convictions",
        "Driver 3 Medical Conditions",
        "Driver 4 Details",
        "Driver 4 Claims",
        "Driver 4 Convictions",
        "Driver 4 Non Motoring convictions",
        "Driver 4 Medical Conditions",
        "Driver 5 Details",
        "Driver 5 Claims",
        "Driver 5 Convictions",
        "Driver 5 Non Motoring convictions",
        "Driver 5 Medical Conditions",
        "Vehicle Reg Number",
        "Vehicle ABI Code",
        "Vehicle  Make",
        "Vehicle  Model",
        "Vehicle  Cover",
        "Vehicle  Excess",
        "Vehicle  Windscreen Excess",
        "Vehicle  Overnight Postcode",
        "Vehicle  Driving Restriction",
        "Vehicle  Main Driver",
        "Vehicle  Body Type",
        "Vehicle  Use",
        "Vehicle  Chassis Number",
        "Vehicle  Cubic Capacity",
        "Vehicle  Gross Weight",
        "Vehicle  Year of Make",
        "Vehicle  Number of Doors",
        "Vehicle  Number of seats",
        "Vehicle  Value",
        "Vehicle  Value Type",
        "Vehicle  Fuel Type",
        "Vehicle  Transmission",
        "Vehicle  Type",
        "Vehicle  NCB years",
        "Vehicle  NCB Type",
        "Vehicle  Previous Insurer Name",
        "Vehicle  Previous Insurer Policy No",
        "Vehicle  Protected NCB",
        "Vehicle  Purchase Date",
        "Vehicle  Overnight storage",
        "Vehicle  Modifications",
        "Vehicle  Owner",
        "Vehicle  Keeper",
        "Vehicle  Annual Mileage",
        "Vehicle  LeftRight Drive",
        "Vehicle  Security Device ",
        "Vehicle  Security Installer",
        "Vehicle  Driving Frequency ",
        "Vehicle  Premium",
        "Endorsements",
        "Sub Broker Name",
        "Sub Broker Address",
        "Sub Broker Post Code"
    ]

    basic_columns = [
        "Agency",
        "Policy number",
        "Inception Date",
        "Expiry Date",
        "Seq Number",
        "Quote Reference",
        "Client Reference",
        "Transaction Effective Date",
        "Transaction Effective Time",
        "Transaction End Date",
        "Transaction Type",
        "Scheme",
        "Voluntary Excess",
        "Total Policy Premium",
        "Net Policy Premium",
        "Commission %",
        "Commission amount ",
        "ISB %",
        "Policyholder Type",
        "Proposer Title",
        "Proposer First Name",
        "Proposer Initials",
        "Proposer Surname",
        "Proposer Date of Birth",
        "Proposer Gender",
        "Policyholder Company Name",
        "Policyholder Company Business",
        "Date Business Established",
        "Abode Type",
        "Address line 1",
        "Address line 2",
        "Address line 3",
        "Address line 4",
        "Address line 5",
        "Address County",
        "Address postcode",
        "Address country"
    ]

    for source_col, target_col in zip(basic_columns, final_headers[:37]):

        if source_col in df.columns:

            output[target_col] = df[source_col].apply(format_value)

    for driver in range(1, 6):

        details_fields = [
            f"Driver {driver} Title ",
            f"Driver {driver}  Forename ",
            f"Driver {driver} MiddleName ",
            f"Driver {driver} Surname ",
            f"Driver {driver} DOB ",
            f"Driver {driver} Age ",
            f"Driver {driver} Licence Type ",
            f"Driver {driver} Date of licence ",
            f"Driver {driver} Employment Status ",
            f"Driver {driver} Business ",
            f"Driver {driver} Occupation ",
            f"Driver {driver} Gender ",
            f"Driver {driver} Marital Status ",
            f"Driver {driver} Relationship to Proposer ",
            f"Driver {driver} Homeowner ",
            f"Driver {driver} Permanent Resident ",
            f"Driver {driver} Proficiency ",
            f"Driver {driver} Number of other vehicles access to ",
            f"Driver {driver} Driver Licence number ",
            f"Driver {driver} Driving Frequency ",
            f"Driver {driver} CountyCourtJudgements ",
            f"Driver {driver} Driving other Vehicles Allowed"
        ]

        claims_groups = [
            [
                f"Driver {driver} Claim 1 Date ",
                f"Driver {driver}  Claim 1 Circumstances (code list 22) ",
                f"Driver {driver} Claim 1 amount ",
                f"Driver {driver} PI claim 1",
                f"Driver {driver} Claim 1 Drivers fault "
            ],
            [
                f"Driver {driver} Claim 2 Date ",
                f"Driver {driver}  Claim 2 Circumstances (code list 22) ",
                f"Driver {driver} Claim 2 amount ",
                f"Driver {driver} PI claim 2",
                f"Driver {driver} Claim 2 Drivers fault "
            ]
        ]

        convictions_groups = [
            [
                f"Driver {driver} Conviction 1 Offence code ",
                f"Driver {driver} Conviction 1  Fine amt ",
                f"Driver {driver} Conviction 1  Points ",
                f"Driver {driver} Conviction 1  Conviction Date ",
                f"Driver {driver} Conviction 1  Disqualified Indt ",
                f"Driver {driver} Conviction 1  SuspensionPeriod (months)"
            ],
            [
                f"Driver {driver} Conviction 2 Offence code ",
                f"Driver {driver} Conviction 2  Fine amt ",
                f"Driver {driver} Conviction 2  Points ",
                f"Driver {driver} Conviction 2  Conviction Date ",
                f"Driver {driver} Conviction 2  Disqualified Indt ",
                f"Driver {driver} Conviction 2  SuspensionPeriod (months)"
            ]
        ]

        non_motoring_groups = [
            [
                f"Offence Date Driver {driver} Non Motoring Conviction 1 ",
                f" Offence Type (Code List 411) Driver {driver} Non Motoring Conviction 1 ",
                f" Sentence Type (Code List 546) Driver {driver} Non Motoring Conviction 1 ",
                f" Length of Sentence (months)Driver {driver} Non Motoring Conviction 1 "
            ]
        ]

        medical_groups = [
            [
                f"Driver {driver} Medical Condition 1 Code (Code List 003)",
                f"Driver {driver} Medical Condition 1 DVLA Notified"
            ]
        ]

        output[f"Driver {driver} Details"] = df.apply(
            lambda row: build_single_record(row, details_fields),
            axis=1
        )

        output[f"Driver {driver} Claims"] = df.apply(
            lambda row: build_multi_record(row, claims_groups),
            axis=1
        )

        output[f"Driver {driver} Convictions"] = df.apply(
            lambda row: build_multi_record(row, convictions_groups),
            axis=1
        )

        output[f"Driver {driver} Non Motoring Convictions"] = df.apply(
            lambda row: build_multi_record(row, non_motoring_groups),
            axis=1
        )

        output[f"Driver {driver} Medical Conditions"] = df.apply(
            lambda row: build_multi_record(row, medical_groups),
            axis=1
        )

    vehicle_columns = [
        "Vehicle Reg Number",
        "Vehicle ABI Code",
        "Vehicle  Make",
        "Vehicle  Model",
        "Vehicle  Cover",
        "Vehicle  Excess",
        "Vehicle  Windscreen Excess",
        "Vehicle  Overnight Postcode",
        "Vehicle  Driving Restriction",
        "Vehicle  Main Driver",
        "Vehicle  Body Type",
        "Vehicle  Use",
        "Vehicle  Chassis Number",
        "Vehicle  Cubic Capacity",
        "Vehicle  Gross Weight",
        "Vehicle  Year of Make",
        "Vehicle  Number of Doors",
        "Vehicle  Number of seats",
        "Vehicle  Value",
        "Vehicle  Value Type",
        "Vehicle  Fuel Type",
        "Vehicle  Transmission",
        "Vehicle  Type",
        "Vehicle  NCB years",
        "Vehicle  NCB Type",
        "Vehicle  Previous Insurer Name",
        "Vehicle  Previous Insurer Policy No",
        "Vehicle  Protected NCB",
        "Vehicle  Purchase Date",
        "Vehicle  Overnight storage",
        "Vehicle  Modifications",
        "Vehicle  Owner",
        "Vehicle  Keeper",
        "Vehicle  Annual Mileage",
        "Vehicle  LeftRight Drive",
        "Vehicle  Security Device ",
        "Vehicle  Security Installer",
        "Vehicle  Driving Frequency ",
        "Vehicle  Premium",
        "Endorsements",
        "Sub Broker Name",
        "Sub Broker Address",
        "Sub Broker Post Code"
    ]

    for col in vehicle_columns:

        if col in df.columns:
            output[col] = df[col].apply(format_value)

    output = output.reindex(columns=final_headers)

    excel_output = BytesIO()

    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:

        output.to_excel(
            writer,
            sheet_name="Final_Export_Output",
            index=False
        )

    excel_output.seek(0)

    workbook = load_workbook(excel_output)

    worksheet = workbook["Final_Export_Output"]

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = column_cells[0].column_letter

        for cell in column_cells:

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass

        worksheet.column_dimensions[column_letter].width = max_length + 4

    final_output = BytesIO()

    workbook.save(final_output)

    final_output.seek(0)

    return final_output

if uploaded_file:

    with st.spinner("Processing file..."):

        result_file = process_file(uploaded_file)

    st.success("Processing Complete")

    st.download_button(
        label="Download Output",
        data=result_file,
        file_name="Final_Export_Output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )