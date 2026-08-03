import io
import pandas as pd
import streamlit as st
# Column indices (0-indexed) — same layout in Excel and CSV
COL_VESSEL_NAME = 0  # A — update allowed
COL_HULL = 1  # B — NO UPDATE
COL_IMO = 2  # C — update allowed
COL_OWNER = 3  # D — NO UPDATE
COL_HANDOVER = 4  # E — update allowed if CSV date >= 2025-01-01
COL_PRIMARY = 5  # F — match key
COL_SECONDARY = 6  # G — match key
COL_SHOPTEST = 7  # H — update allowed only if 'shoptested' not in Excel cell
CSV_MATCH_COL = 0  # First column of CSV used for row lookup
HANDOVER_CUTOFF = pd.Timestamp("2025-01-01")
def clean_key(value) -> str:
    return str(value).strip().lower().replace(" 00:00:00", "").replace(".0", "")
def csv_value_present(value) -> bool:
    text = str(value).strip()
    return bool(text) and text.lower() != "nan"
def get_cell(row, col_idx: int) -> str:
    if col_idx >= len(row):
        return ""
    return str(row.iloc[col_idx]).strip()
def set_cell(df: pd.DataFrame, row_idx: int, col_idx: int, value: str) -> None:
    df.iat[row_idx, col_idx] = value
st.set_page_config(page_title="Vessel Data Updater", layout="centered")
st.title("Vessel Data Updater")
uploaded_sheet = st.file_uploader("Upload Excel File", type=["xlsx"])
uploaded_db = st.file_uploader("Upload CSV Database", type=["csv"])
if uploaded_sheet and uploaded_db:
    if st.button("Process & Update File", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet, header=0, dtype=str).fillna("")
            db_df = pd.read_csv(uploaded_db, header=0, dtype=str).fillna("")
            db_lookup = {}
            for _, db_row in db_df.iterrows():
                key = clean_key(get_cell(db_row, CSV_MATCH_COL))
                if key and key not in db_lookup:
                    db_lookup[key] = db_row
            total_rows = len(sheet_df)
            modified_rows = 0
            for row_idx in range(total_rows):
                primary_key = clean_key(sheet_df.iat[row_idx, COL_PRIMARY])
                secondary_key = clean_key(sheet_df.iat[row_idx, COL_SECONDARY])
                if primary_key in db_lookup:
                    db_row = db_lookup[primary_key]
                elif secondary_key in db_lookup:
                    db_row = db_lookup[secondary_key]
                else:
                    continue
                row_modified = False
                # A — Vessel Name
                csv_vessel = get_cell(db_row, COL_VESSEL_NAME)
                if csv_value_present(csv_vessel):
                    excel_vessel = str(sheet_df.iat[row_idx, COL_VESSEL_NAME]).strip()
                    if csv_vessel != excel_vessel:
                        set_cell(sheet_df, row_idx, COL_VESSEL_NAME, csv_vessel)
                        row_modified = True
                # C — IMO Number
                csv_imo = get_cell(db_row, COL_IMO)
                if csv_value_present(csv_imo):
                    excel_imo = str(sheet_df.iat[row_idx, COL_IMO]).strip()
                    if csv_imo != excel_imo:
                        set_cell(sheet_df, row_idx, COL_IMO, csv_imo)
                        row_modified = True
                # E — Handover Date (only if CSV date >= 2025-01-01)
                csv_handover = get_cell(db_row, COL_HANDOVER)
                if csv_value_present(csv_handover):
                    try:
                        handover_date = pd.to_datetime(csv_handover.split(" ")[0])
                        if handover_date >= HANDOVER_CUTOFF:
                            excel_handover = str(sheet_df.iat[row_idx, COL_HANDOVER]).strip()
                            if csv_handover != excel_handover:
                                set_cell(sheet_df, row_idx, COL_HANDOVER, csv_handover)
                                row_modified = True
                    except (ValueError, TypeError):
                        pass
                # H — Shop Test Date (skip if Excel cell contains 'shoptested')
                csv_shoptest = get_cell(db_row, COL_SHOPTEST)
                if csv_value_present(csv_shoptest):
                    excel_shoptest = str(sheet_df.iat[row_idx, COL_SHOPTEST])
                    if "shoptested" not in excel_shoptest.lower():
                        excel_shoptest_stripped = excel_shoptest.strip()
                        if csv_shoptest != excel_shoptest_stripped:
                            set_cell(sheet_df, row_idx, COL_SHOPTEST, csv_shoptest)
                            row_modified = True
                if row_modified:
                    modified_rows += 1
            st.write(f"Total amount of rows: {total_rows}")
            st.write(f"Number of rows that have been modified: {modified_rows}")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet_df.to_excel(writer, index=False)
            output.seek(0)
            download_name = uploaded_sheet.name or "updated.xlsx"
            if not download_name.lower().endswith(".xlsx"):
                download_name = f"{download_name}.xlsx"
            st.download_button(
                label="Download Updated Excel",
                data=output,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.error(f"Error: {exc}")
