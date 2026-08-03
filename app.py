import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"

# Column Indices (0-indexed)
EXCEL_COL_NAME = 0      # A
EXCEL_COL_IMO = 2       # C
EXCEL_COL_HANDOVER = 4  # E
EXCEL_COL_PRIMARY = 5   # F
EXCEL_COL_SECONDARY = 6 # G
EXCEL_COL_SHOPTEST = 7  # H

HANDOVER_CUTOFF = pd.Timestamp("2025-01-01")

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")
st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload CSV Database (.csv)", type=["csv"])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Update File", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet, header=0, dtype=str).fillna("")
            db_df = pd.read_csv(uploaded_db, header=0, dtype=str).fillna("")

            def clean_key(v: str) -> str:
                return str(v).strip().lower().replace(" 00:00:00", "").replace(".0", "")

            # Build Lookup Map (Primary ID is column 0 in CSV)
            db_id_col = db_df.columns[0]
            db_lookup = {}
            for _, db_row in db_df.iterrows():
                key = clean_key(db_row[db_id_col])
                if key and key not in db_lookup:
                    db_lookup[key] = db_row.to_dict()

            rows_checked = len(sheet_df)
            updated_rows = 0

            for i in range(rows_checked):
                primary_key = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY])
                secondary_key = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY])

                if primary_key in db_lookup:
                    db_data = db_lookup[primary_key]
                elif secondary_key in db_lookup:
                    db_data = db_lookup[secondary_key]
                else:
                    continue

                row_modified = False

                # 1. Update Vessel Name (A) and IMO Number (C)
                update_map = {EXCEL_COL_NAME: "Vessel_Name", EXCEL_COL_IMO: "IMO_Number"}
                for col_idx, db_col in update_map.items():
                    val = str(db_data.get(db_col, "")).strip()
                    if val and val.lower() != "nan":
                        if val != str(sheet_df.iat[i, col_idx]).strip():
                            sheet_df.iat[i, col_idx] = val
                            row_modified = True

                # 2. Update Handover Date (E) - Only if >= 2025-01-01
                val_handover = str(db_data.get("Handover_Date", "")).strip()
                if val_handover and val_handover.lower() != "nan":
                    try:
                        dt = pd.to_datetime(val_handover.split(" ")[0])
                        if dt >= HANDOVER_CUTOFF:
                            if val_handover != str(sheet_df.iat[i, EXCEL_COL_HANDOVER]).strip():
                                sheet_df.iat[i, EXCEL_COL_HANDOVER] = val_handover
                                row_modified = True
                    except: pass

                # 3. Update Shop Test Date (H) - Only if "shoptested" NOT in cell
                val_shoptest = str(db_data.get("Shop_Test_Date", "")).strip()
                if val_shoptest and val_shoptest.lower() != "nan":
                    current_cell = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).lower()
                    if "shoptested" not in current_cell:
                        if val_shoptest != str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).strip():
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = val_shoptest
                            row_modified = True

                if row_modified:
                    updated_rows += 1

            st.success("✅ Processing complete.")
            st.write(f"• Total amount of rows: **{rows_checked}**")
            st.write(f"• Number of rows that have been modified: **{updated_rows}**")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet_df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button("📥 Download Updated Excel", data=output, file_name="updated_vessels.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        except Exception as e:
            st.error(f"Error: {e}")
