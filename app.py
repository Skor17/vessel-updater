import streamlit as st
import pandas as pd
from datetime import datetime
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"

# Column Setup (0-indexed)
EXCEL_COL_NAME = 0         # A
EXCEL_COL_IMO = 2          # C
EXCEL_COL_HANDOVER = 4     # E
EXCEL_COL_PRIMARY_DB = 5   # F
EXCEL_COL_SECONDARY_DB = 6 # G
EXCEL_COL_SHOPTEST = 7     # H

HANDOVER_CUTOFF = datetime(2025, 1, 1)

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")
st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Update Data", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet).astype(str)
            db_df = pd.read_csv(uploaded_db, dtype=str).fillna('')

            # Helpers
            def clean_key(v):
                return str(v).strip().lower().replace('.0', '')

            # Create ID map from first CSV column
            db_id_col = db_df.columns[0]
            db_df['CLEAN_KEY'] = db_df[db_id_col].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            total_rows = len(sheet_df)
            updated_rows = 0

            # Loop
            for i in range(total_rows):
                # Get Keys
                p = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB])
                s = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB])

                # Match
                match_id = p if p in db_indexed.index else (s if s in db_indexed.index else None)

                if match_id is None:
                    continue

                row_db = db_indexed.loc[match_id]
                row_modified = False

                # Name
                val_db = str(row_db.get('Vessel_Name', '')).strip()
                if val_db and val_db.lower() != 'nan':
                    if val_db != str(sheet_df.iat[i, EXCEL_COL_NAME]).strip():
                        sheet_df.iat[i, EXCEL_COL_NAME] = val_db
                        row_modified = True

                # IMO
                val_db = str(row_db.get('IMO_Number', '')).strip()
                if val_db and val_db.lower() != 'nan':
                    if val_db != str(sheet_df.iat[i, EXCEL_COL_IMO]).strip():
                        sheet_df.iat[i, EXCEL_COL_IMO] = val_db
                        row_modified = True

                # Handover Date - only update if CSV date >= 2025-01-01
                val_db = str(row_db.get('Handover_Date', '')).strip()
                if val_db and val_db.lower() != 'nan':
                    parsed_date = pd.to_datetime(val_db, errors='coerce')
                    if pd.notna(parsed_date) and parsed_date >= HANDOVER_CUTOFF:
                        if val_db != str(sheet_df.iat[i, EXCEL_COL_HANDOVER]).strip():
                            sheet_df.iat[i, EXCEL_COL_HANDOVER] = val_db
                            row_modified = True

                # Shop Test Date - skip entirely if current cell already contains 'shoptested'
                current_shoptest = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).strip()
                if 'shoptested' not in current_shoptest.lower():
                    val_db = str(row_db.get('Shop_Test_Date', '')).strip()
                    if val_db and val_db.lower() != 'nan':
                        if val_db != current_shoptest:
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = val_db
                            row_modified = True

                if row_modified:
                    updated_rows += 1

            st.success("✅ Processing complete.")
            st.write(f"• Total Rows Checked: **{total_rows}**")
            st.write(f"• Rows Actually Updated: **{updated_rows}**")

            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button(
                "📥 Download Updated Excel",
                data=output.getvalue(),
                file_name="updated_vessels.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Error: {e}")
