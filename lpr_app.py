import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO


# =========================
# BUSINESS UNIT MAPPING - LPR
# =========================
business_unit_map = {
    "Denmark": {"Sender Name": "Denmark", "Sender Location Id": "C7734"},
    "Denmark - Hanstholm": {"Sender Name": "Denmark - Hanstholm", "Sender Location Id": "C15390"},
    "Driffield": {"Sender Name": "Driffield", "Sender Location Id": "C6964"},
    "France": {"Sender Name": "France", "Sender Location Id": "C4094"},
    "Ireland": {"Sender Name": "Ireland", "Sender Location Id": "C6963"},
    "Larkshall": {"Sender Name": "Larkshall", "Sender Location Id": "C11254"},
    "Netherlands": {"Sender Name": "Netherlands", "Sender Location Id": "C6960"},
    "Spain": {"Sender Name": "Spain", "Sender Location Id": "C14430"},
    "HQ": {"Sender Name": "HQ", "Sender Location Id": "C18571"},
    "Coca-Cola HBC Northern Ireland Ltd": {
        "Sender Name": "Coca-Cola HBC Northern Ireland Ltd",
        "Sender Location Id": "16928"
    }
}


# =========================
# FUNCTIONS
# =========================
def clean_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\xa0", "", regex=False)
    )
    return df


def convert_date_to_ddmmyyyy(value):
    if pd.isna(value):
        return pd.NaT

    try:
        value_str = str(value).strip()

        if value_str.endswith(".0"):
            value_str = value_str[:-2]

        if value_str.isdigit() and len(value_str) == 8:
            return pd.Timestamp(
                year=int(value_str[0:4]),
                month=int(value_str[4:6]),
                day=int(value_str[6:8])
            )

        if value_str.isdigit() and len(value_str) == 6:
            return pd.to_datetime(value_str, format="%d%m%y", errors="coerce")

        if value_str.isdigit() and len(value_str) == 5:
            return pd.to_datetime(
                int(value_str),
                unit="D",
                origin="1899-12-30",
                errors="coerce"
            )

        if len(value_str) >= 10:
            date_part = value_str[:10]

            if "-" in date_part:
                parts = date_part.split("-")
                if len(parts) == 3 and len(parts[0]) == 4:
                    return pd.Timestamp(
                        year=int(parts[0]),
                        month=int(parts[1]),
                        day=int(parts[2])
                    )

            if "/" in date_part:
                parts = date_part.split("/")
                if len(parts) == 3 and len(parts[0]) == 4:
                    return pd.Timestamp(
                        year=int(parts[0]),
                        month=int(parts[1]),
                        day=int(parts[2])
                    )

        return pd.to_datetime(value_str, errors="coerce", dayfirst=True)

    except Exception:
        return pd.NaT


def clean_reference_number(value):
    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    if len(value) > 13:
        value = value[-13:]

    return value


def clean_concat_part(value):
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float, np.integer, np.floating)):
        if float(value).is_integer():
            return str(int(value)).zfill(3)

    value = str(value).strip()

    if value.lower() in ["nan", "none"]:
        return ""

    if value.endswith(".0") and value[:-2].isdigit():
        return value[:-2].zfill(3)

    return value


def clean_code(value):
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        value = value[:-2]

    if value.lower() in ["nan", "none"]:
        return ""

    return value


def working_days(start, end):
    if pd.isna(start):
        return np.nan

    return np.busday_count(
        start.date(),
        end.date()
    )


def is_missing(series):
    return (
        series.isna()
        | (series.astype(str).str.strip() == "")
        | (series.astype(str).str.lower().str.strip() == "nan")
    )


def excel_buffer(df):
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer


def build_tracking_file(final_df, business_unit, batch_number):
    columns = [
        "Movement Date",
        "Business Unit",
        "Pooler",
        "Movement Direction",
        "Pallet Type",
        "Reference 1",
        "Reference 2",
        "Reference 3",
        "Batch Number",
        "Sender Location Id",
        "Sender Name",
        "Sender Town",
        "Sender Postcode",
        "Receiver Location Id",
        "Receiver Name",
        "Receiver Town",
        "Receiver Postcode",
        "Movement Type",
        "Quantity",
        "Savings",
        "Declared Status",
        "Comments"
    ]

    if final_df.empty:
        return pd.DataFrame(columns=columns)

    grouped = final_df.groupby(
        ["Reference", "Declaring Code"],
        as_index=False
    ).agg({
        "Quantity": "sum",
        "Movement Date": "first",
        "Receiver": "first",
        "Comments": "first"
    })

    tracking_df = pd.DataFrame({
        "Movement Date": grouped["Movement Date"],
        "Business Unit": business_unit_map[business_unit]["Sender Name"],
        "Pooler": "LPR",
        "Movement Direction": "Out",
        "Pallet Type": "LPR UK100 - UK",
        "Reference 1": grouped["Reference"],
        "Reference 2": "",
        "Reference 3": "",
        "Batch Number": batch_number,
        "Sender Location Id": business_unit_map[business_unit]["Sender Location Id"],
        "Sender Name": business_unit_map[business_unit]["Sender Name"],
        "Sender Town": "",
        "Sender Postcode": "",
        "Receiver Location Id": grouped["Declaring Code"],
        "Receiver Name": grouped["Receiver"],
        "Receiver Town": "",
        "Receiver Postcode": "",
        "Movement Type": "Out - LPR Drop Point",
        "Quantity": grouped["Quantity"],
        "Savings": "",
        "Declared Status": "Declared",
        "Comments": grouped["Comments"]
    })

    return tracking_df[columns]


# =========================
# APP
# =========================
st.title("LPR Tracking Sheet")

business_unit = st.selectbox(
    "Select Business Unit",
    list(business_unit_map.keys())
)

batch_number = st.text_input("Enter Batch Number")

main_file = st.file_uploader(
    "Upload Main Excel File",
    type=["xlsx"]
)

mapping_file = st.file_uploader(
    "Upload LPR Matching Table - only needed if LPR code is not already in main file",
    type=["xlsx"]
)

hq_mapping_file = None

if business_unit == "HQ":
    hq_mapping_file = st.file_uploader(
        "Upload HQ Location Mapping Table - only needed if using LPR Matching Table",
        type=["xlsx"]
    )


if main_file:

    df = pd.read_excel(main_file)
    df = clean_columns(df)

    mapping_df = None
    if mapping_file:
        mapping_df = pd.read_excel(mapping_file)
        mapping_df = clean_columns(mapping_df)

    st.subheader("Map Main File Columns")

    main_columns = df.columns.tolist()

    date_col = st.selectbox("Select column for Movement Date", main_columns)
    receiver_col = st.selectbox("Select column for Receiver / Customer Name", main_columns)
    reference_col = st.selectbox("Select column for Reference", main_columns)
    quantity_col = st.selectbox("Select column for Quantity", main_columns)

    st.subheader("End Location / LPR Code Mapping")

    code_source = st.radio(
        "Where is the LPR Declaring Code / End Location Code?",
        [
            "Already in Main File",
            "Use LPR Matching Table"
        ]
    )

    selected_concat_cols = []
    existing_mapping_col = None
    mapping_key_col = None
    declaring_code_col = None
    lpr_code_col = None

    main_address_col = None
    hq_address_col = None
    hq_mapping_return_col = None
    hq_mapping_df = None

    if code_source == "Already in Main File":
        lpr_code_col = st.selectbox(
            "Which is your LPR code column?",
            main_columns
        )

    else:
        if mapping_df is None:
            st.error("Please upload LPR Matching Table.")
            st.stop()

        if business_unit == "HQ":
            if hq_mapping_file is None:
                st.error("Please upload HQ Location Mapping Table.")
                st.stop()

            hq_mapping_df = pd.read_excel(hq_mapping_file)
            hq_mapping_df = clean_columns(hq_mapping_df)

            st.subheader("HQ Address Mapping")

            main_address_col = st.selectbox(
                "Select Address Line 1 column from Main File",
                main_columns
            )

            hq_address_col = st.selectbox(
                "Select Address1 column from HQ Mapping Table",
                hq_mapping_df.columns.tolist()
            )

            hq_mapping_return_col = st.selectbox(
                "Select column from HQ Mapping Table to use for LPR Matching",
                hq_mapping_df.columns.tolist()
            )

        else:
            st.subheader("Create Location ID / Mapping Key")

            concatenate_required = st.radio(
                "Do you want to concatenate columns to create Location ID / Mapping Key?",
                ["Yes", "No"]
            )

            if concatenate_required == "Yes":
                selected_concat_cols = st.multiselect(
                    "Select columns to concatenate in order",
                    main_columns
                )
            else:
                existing_mapping_col = st.selectbox(
                    "Select existing Location ID / Mapping Key column",
                    main_columns
                )

        st.subheader("Map LPR Matching Table Columns")

        mapping_key_col = st.selectbox(
            "Select Mapping Key column in LPR Matching Table",
            mapping_df.columns.tolist()
        )

        declaring_code_col = st.selectbox(
            "Select LPR Declaring Code column in LPR Matching Table",
            mapping_df.columns.tolist()
        )

    if st.button("Prepare Data"):

        if not batch_number:
            st.error("Please enter Batch Number.")
            st.stop()

        work_df = df.copy()

        work_df["Movement Date Parsed"] = work_df[date_col].apply(convert_date_to_ddmmyyyy)
        work_df["Movement Date"] = work_df["Movement Date Parsed"].dt.strftime("%d/%m/%Y")

        work_df["Receiver"] = work_df[receiver_col].astype(str).str.strip()

        work_df["Reference"] = work_df[reference_col].apply(clean_reference_number)

        work_df["Quantity"] = pd.to_numeric(
            work_df[quantity_col],
            errors="coerce"
        ).fillna(0)

        work_df = work_df[
            work_df["Quantity"] != 0
        ].copy()

        if code_source == "Already in Main File":

            work_df["Mapping Key"] = work_df[lpr_code_col].apply(clean_code)
            work_df["Declaring Code"] = work_df["Mapping Key"]

        else:

            if business_unit == "HQ":

                work_df["HQ Address Match Key"] = (
                    work_df[main_address_col]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                hq_mapping_df["HQ Address Match Key"] = (
                    hq_mapping_df[hq_address_col]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                hq_lookup = hq_mapping_df[
                    ["HQ Address Match Key", hq_mapping_return_col]
                ].copy()

                hq_lookup.rename(
                    columns={hq_mapping_return_col: "LPR Reporting Code"},
                    inplace=True
                )

                hq_lookup = hq_lookup.drop_duplicates(
                    subset=["HQ Address Match Key"],
                    keep="first"
                )

                work_df = work_df.merge(
                    hq_lookup,
                    on="HQ Address Match Key",
                    how="left"
                )

                work_df["Mapping Key"] = (
                    work_df["LPR Reporting Code"]
                    .apply(clean_code)
                    .str.upper()
                )

            else:

                if concatenate_required == "Yes":
                    if not selected_concat_cols:
                        st.error("Please select columns to concatenate.")
                        st.stop()

                    work_df["Mapping Key"] = ""

                    for col in selected_concat_cols:
                        work_df["Mapping Key"] += work_df[col].apply(clean_concat_part)

                else:
                    work_df["Mapping Key"] = (
                        work_df[existing_mapping_col]
                        .apply(clean_code)
                        .str.upper()
                    )

            mapping_lookup = mapping_df[
                [mapping_key_col, declaring_code_col]
            ].copy()

            mapping_lookup.rename(columns={
                mapping_key_col: "Mapping Key",
                declaring_code_col: "Declaring Code"
            }, inplace=True)

            mapping_lookup["Mapping Key"] = (
                mapping_lookup["Mapping Key"]
                .apply(clean_code)
                .str.upper()
            )

            mapping_lookup["Declaring Code"] = (
                mapping_lookup["Declaring Code"]
                .apply(clean_code)
            )

            mapping_lookup = mapping_lookup.drop_duplicates(
                subset=["Mapping Key"],
                keep="first"
            )

            before_rows = len(work_df)

            work_df = work_df.merge(
                mapping_lookup,
                on="Mapping Key",
                how="left"
            )

            st.info(f"Rows before LPR lookup: {before_rows}")

        today = pd.Timestamp.today().normalize()

        work_df["Working Days Old"] = work_df["Movement Date Parsed"].apply(
            lambda x: working_days(x, today)
        )

        work_df["Comments"] = ""

        work_df.loc[
            work_df["Working Days Old"] > 29,
            "Comments"
        ] = "OVERDUE: Movement exceeds 29 working days"

        st.session_state["lpr_work_df"] = work_df


# =========================
# REVIEW AND OUTPUT
# =========================
if "lpr_work_df" in st.session_state:

    work_df = st.session_state["lpr_work_df"].copy()

    old_date_df = work_df[
        work_df["Working Days Old"] > 29
    ].copy()

    if not old_date_df.empty:
        st.warning("Some movement dates are more than 29 working days old.")
        st.subheader("Review Dates More Than 29 Working Days Old")

        for idx, row in old_date_df.iterrows():
            reference = row["Reference"]
            old_date = row["Movement Date"]
            days_old = row["Working Days Old"]

            st.write(
                f"Reference Number **{reference}** has date **{old_date}**. "
                f"It is **{days_old} working days old**."
            )

            amend_date = st.checkbox(
                f"Amend date for reference {reference}?",
                key=f"lpr_amend_date_{idx}"
            )

            if amend_date:
                new_date = st.date_input(
                    f"Enter new date for reference {reference}",
                    key=f"lpr_new_date_{idx}"
                )

                work_df.loc[idx, "Movement Date Parsed"] = pd.to_datetime(new_date)
                work_df.loc[idx, "Movement Date"] = pd.to_datetime(new_date).strftime("%d/%m/%Y")
                work_df.loc[idx, "Comments"] = ""

    missing_code_df = work_df[
        is_missing(work_df["Declaring Code"])
    ].copy()

    declaring_code_inputs = {}
    remove_rows = []

    if not missing_code_df.empty:
        st.warning("Some rows are missing LPR Declaring Code.")
        st.subheader("Missing LPR Declaring Code Input")

        for idx, row in missing_code_df.iterrows():
            receiver = row["Receiver"]
            reference = row["Reference"]
            mapping_key = row["Mapping Key"]
            quantity = row["Quantity"]
            movement_date = row["Movement Date"]

            with st.container():
                st.markdown("---")
                st.write(f"**Receiver:** {receiver}")
                st.write(f"**Reference Number:** {reference}")
                st.write(f"**Mapping Key:** {mapping_key}")
                st.write(f"**Movement Date:** {movement_date}")
                st.write(f"**Quantity:** {quantity}")

                declaring_code_inputs[idx] = st.text_input(
                    "Please enter LPR Declaring Code for this row",
                    key=f"lpr_declaring_code_{idx}"
                )

                remove = st.checkbox(
                    "Declaring Code not found - remove this row",
                    key=f"lpr_remove_row_{idx}"
                )

                if remove:
                    remove_rows.append(idx)
    else:
        st.success("No missing LPR Declaring Codes found.")

    if st.button("Generate LPR Final Files"):

        for idx, code_value in declaring_code_inputs.items():
            if code_value.strip():
                work_df.loc[idx, "Declaring Code"] = code_value.strip()

        removed_rows_df = work_df.loc[remove_rows].copy()

        declaring_code_required_df = work_df[
            is_missing(work_df["Declaring Code"])
            & ~work_df.index.isin(remove_rows)
        ].copy()

        final_df = work_df[
            ~is_missing(work_df["Declaring Code"])
            & ~work_df.index.isin(remove_rows)
        ].copy()

        tracking_df = build_tracking_file(
            final_df,
            business_unit,
            batch_number
        )

        required_tracking_df = build_tracking_file(
            declaring_code_required_df,
            business_unit,
            batch_number
        )

        removed_tracking_df = build_tracking_file(
            removed_rows_df,
            business_unit,
            batch_number
        )

        st.success("LPR files generated successfully.")

        st.download_button(
            label="Download LPR Tracking Sheet",
            data=excel_buffer(tracking_df),
            file_name=f"{batch_number}_LPR_tracking_sheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download LPR Declaring Code Required File",
            data=excel_buffer(required_tracking_df),
            file_name=f"{batch_number}_LPR_declaring_code_required.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download LPR Removed Rows",
            data=excel_buffer(removed_tracking_df),
            file_name="LPR Removed rows.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
