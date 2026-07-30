import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"

# Column Setup
EXCEL_COL_NAME = 0         # A
EXCEL_COL_IMO = 2          # C
EXCEL_COL_HANDOVER = 4     # E
EXCEL_COL_PRIMARY_DB = 5   # F
EXCEL_COL_SECONDARY_DB = 6 # G
EXCEL_COL_SHOPTEST = 7     # H

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")
st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Update Data", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet).astype(str)
            db_df = pd.read_csv(uploaded_db, dtype=str).fillna('')

            # Helpers
            def clean_key(v):
                return str(v).strip().lower().replace('.0', '')

            # Create ID map
            db_id_col = db_df.columns[0] # Assuming first col is DB Number
            db_df['CLEAN_KEY'] = db_df[db_id_col].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_rows = 0

            # Loop
            for i in range(len(sheet_df)):
                # Get Keys
                p = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB])
                s = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB])
                
                # Match
                match_id = p if p in db_indexed.index else (s if s in db_indexed.index else None)
                
                if match_id:
                    row_db = db_indexed.loc[match_id]
                    row_modified = False

                    # Mapping: (Excel_Col_Index, DB_Col_Name)
                    cols_to_update = [
                        (EXCEL_COL_NAME, 'Vessel_Name'),
                        (EXCEL_COL_IMO, 'IMO_Number'),
                        (EXCEL_COL_HANDOVER, 'Handover_Date'),
                        (EXCEL_COL_SHOPTEST, 'Shop_Test_Date')
                    ]

                    for col_idx, db_col_name in cols_to_update:
                        if db_col_name in row_db:
                            val_db = str(row_db[db_col_name]).strip()
                            # Only update if DB is not empty
                            if val_db and val_db.lower() != 'nan':
                                if val_db != str(sheet_df.iat[i, col_idx]).strip():
                                    sheet_df.iat[i, col_idx] = val_db
                                    row_modified = True
                    
                    if row_modified:
                        updated_rows += 1

            st.success(f"✅ Processing complete.")
            st.write(f"• Total Rows Checked: **{len(sheet_df)}**")
            st.write(f"• Rows Actually Updated: **{updated_rows}**")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="updated_vessels.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        except Exception as e:
            st.error(f"Error: {e}")
