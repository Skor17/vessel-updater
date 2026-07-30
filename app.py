import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"
DB_ID_COL = 'DB Number'
DB_NAME_COL = 'Vessel_Name'
DB_IMO_COL = 'IMO_Number'
DB_HANDOVER_COL = 'Handover_Date'
DB_SHOPTEST_COL = 'Shop_Test_Date' 

# Excel Column Mapping
EXCEL_COL_NAME = 0         
EXCEL_COL_IMO = 2          
EXCEL_COL_HANDOVER = 4     
EXCEL_COL_PRIMARY_DB = 5   
EXCEL_COL_SECONDARY_DB = 6 
EXCEL_COL_SHOPTEST = 7     

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")
st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

# Debug Input
debug_id = st.text_input("DEBUG: Enter a DB Number to trace updates for that vessel:")

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Create Updated File", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db, dtype=str, sep=None, engine='python')
            db_df = db_df.fillna('')

            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            def get_best_val(row_dict, target_col):
                # Fuzzy match headers
                clean_target = target_col.replace('_', '').replace(' ', '').lower()
                for col_name, val in row_dict.items():
                    if col_name.replace('_', '').replace(' ', '').lower() == clean_target:
                        return val
                return ''

            db_id_actual = next((c for c in db_df.columns if c.replace('_', '').replace(' ', '').lower() == DB_ID_COL.replace('_', '').replace(' ', '').lower()), db_df.columns[0])
            db_df['CLEAN_KEY'] = db_df[db_id_actual].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_rows_count = 0
            total_rows = len(sheet_df)
            
            for i in range(total_rows):
                primary = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB] if EXCEL_COL_PRIMARY_DB < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB] if EXCEL_COL_SECONDARY_DB < sheet_df.shape[1] else '')
                match_id = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match_id:
                    row = db_indexed.loc[match_id]
                    row_dict = row.to_dict() if not isinstance(row, pd.DataFrame) else {col: row[col].tolist()[-1] for col in row.columns}
                    db_h_val = str(get_best_val(row_dict, DB_SHOPTEST_COL)).strip()

                    # DEBUGGING: If this matches the Debug ID, show us what's happening
                    if debug_id and (debug_id == primary or debug_id == secondary):
                        st.write(f"--- DEBUGGING VESSEL {debug_id} ---")
                        st.write(f"Excel Current Value in Col H: '{sheet_df.iat[i, EXCEL_COL_SHOPTEST]}'")
                        st.write(f"Database Value in Col H: '{db_h_val}'")
                        st.write(f"Should update? {'shoptested' not in str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).lower()}")

                    row_was_modified = False

                    # Name/IMO/Handover
                    for col_idx, db_col in [(EXCEL_COL_NAME, DB_NAME_COL), (EXCEL_COL_IMO, DB_IMO_COL), (EXCEL_COL_HANDOVER, DB_HANDOVER_COL)]:
                        val = str(get_best_val(row_dict, db_col)).strip()
                        if val and val != str(sheet_df.iat[i, col_idx]).strip():
                            sheet_df.iat[i, col_idx] = val
                            row_was_modified = True

                    # Shop Test Date - Hard Overwrite logic
                    current_h_val = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).lower()
                    if "shoptested" not in current_h_val:
                        if db_h_val and db_h_val != str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).strip():
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = db_h_val
                            row_was_modified = True
                    
                    if row_was_modified:
                        updated_rows_count += 1
            
            st.success(f"✅ Processing complete.")
            st.write(f"• Total Rows Checked: **{total_rows}**")
            st.write(f"• Actually Updated Rows: **{updated_rows_count}**")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e:
            st.error(f"Error: {e}")
