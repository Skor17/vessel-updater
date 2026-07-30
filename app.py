import io
import pandas as pd
import streamlit as st

# Column indexes (0-based)
EXCEL_COL_NAME = 0
EXCEL_COL_IMO = 2
EXCEL_COL_HANDOVER = 4
EXCEL_COL_PRIMARY_ID = 5
EXCEL_COL_SECONDARY_ID = 6
EXCEL_COL_SHOPTEST = 7

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢")

st.title("🚢 Vessel Data Updater")

uploaded_excel = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
uploaded_csv = st.file_uploader("Upload CSV Database (.csv)", type=["csv"])


def clean_key(value):
    value = str(value).strip()
    if value.lower() == "nan":
        return ""
    if value.endswith(".0"):
        value = value[:-2]
    return value


def valid_csv_value(value):
    if pd.isna(value):
        return False
    value = str(value).strip()
    return value != "" and value.lower() != "nan"


if uploaded_excel and uploaded_csv:

    if st.button("🚀 Update Data", use_container_width=True):

        excel_df = pd.read_excel(uploaded_excel, dtype=object)
        csv_df = pd.read_csv(uploaded_csv, dtype=str).fillna("")

        id_column = csv_df.columns[0]

        csv_lookup = {}
        for _, row in csv_df.iterrows():
            csv_lookup[clean_key(row[id_column])] = row

        rows_checked = len(excel_df)
        rows_modified = 0

        for row_idx in range(len(excel_df)):

            primary = clean_key(excel_df.iat[row_idx, EXCEL_COL_PRIMARY_ID])
            secondary = clean_key(excel_df.iat[row_idx, EXCEL_COL_SECONDARY_ID])

            csv_row = None

            if primary in csv_lookup:
                csv_row = csv_lookup[primary]
            elif secondary in csv_lookup:
                csv_row = csv_lookup[secondary]

            if csv_row is None:
                continue

            row_changed = False

            # -------------------------
            # Name (Column A)
            # -------------------------
            if "Vessel_Name" in csv_row:
                value = csv_row["Vessel_Name"]
                if valid_csv_value(value):
                    if str(excel_df.iat[row_idx, EXCEL_COL_NAME]) != str(value):
                        excel_df.iat[row_idx, EXCEL_COL_NAME] = value
                        row_changed = True

            # -------------------------
            # IMO (Column C)
            # -------------------------
            if "IMO_Number" in csv_row:
                value = csv_row["IMO_Number"]
                if valid_csv_value(value):
                    if str(excel_df.iat[row_idx, EXCEL_COL_IMO]) != str(value):
                        excel_df.iat[row_idx, EXCEL_COL_IMO] = value
                        row_changed = True

            # -------------------------
            # Handover Date (Column E)
            # Update only if CSV date >= 2025-01-01
            # -------------------------
            if "Handover_Date" in csv_row:
                value = csv_row["Handover_Date"]

                if valid_csv_value(value):
                    csv_date = pd.to_datetime(value, errors="coerce")

                    if pd.notna(csv_date) and csv_date >= pd.Timestamp("2025-01-01"):

                        current = excel_df.iat[row_idx, EXCEL_COL_HANDOVER]

                        current_date = pd.to_datetime(current, errors="coerce")

                        if (
                            pd.isna(current_date)
                            or current_date != csv_date
                        ):
                            excel_df.iat[row_idx, EXCEL_COL_HANDOVER] = value
                            row_changed = True

            # -------------------------
            # Shop Test Date (Column H)
            # Skip if Excel already contains "shoptested"
            # -------------------------
            if "Shop_Test_Date" in csv_row:

                current = str(excel_df.iat[row_idx, EXCEL_COL_SHOPTEST])

                if "shoptested" not in current.lower():

                    value = csv_row["Shop_Test_Date"]

                    if valid_csv_value(value):
                        if current != str(value):
                            excel_df.iat[row_idx, EXCEL_COL_SHOPTEST] = value
                            row_changed = True

            if row_changed:
                rows_modified += 1

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            excel_df.to_excel(writer, index=False)

        output.seek(0)

        st.success("Update complete.")

        st.write(f"**Total Rows Checked:** {rows_checked}")
        st.write(f"**Total Rows Actually Modified:** {rows_modified}")

        st.download_button(
            "📥 Download Updated Excel",
            data=output,
            file_name="updated_vessels.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
