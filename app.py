import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"

# Database CSV Column Names
DB_ID_COL = 'DB Number'
DB_NAME_COL = 'Vessel_Name'
DB_IMO_COL = 'IMO_Number'
DB_HANDOVER_COL = 'Handover_Date'
DB_SHOPTEST_COL = 'Shop_Test_Date' 

# Excel Column Mapping
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
    if st.button("🚀 Process & Update File", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db, dtype=str, sep=None, engine='python')
            db_df = db_df.fillna('')

            # Helpers
            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            def get_db_val(row_dict, target_col):
                clean_target = target_col.replace('_', '').replace(' ', '').lower()
                for col_name, val in row_dict.items():
                    if col_name.replace('_', '').replace(' ', '').lower() == clean_target:
                        return str(val).strip()
                return ''

            # Indexing
            db_id_actual = next((c for c in db_df.columns if c.replace('_', '').replace(' ', '').lower() == DB_ID_COL.replace('_', '').replace(' ', '').lower()), db_df.columns[0])
            db_df['CLEAN_KEY'] = db_df[db_id_actual].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_rows_count = 0
            
            # --- LOOP ---
            for i in range(len(sheet_df)):
                primary = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB] if EXCEL_COL_PRIMARY_DB < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB] if EXCEL_COL_SECONDARY_DB < sheet_df.shape[1] else '')
                match_id = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match_id:
                    row = db_indexed.loc[match_id]
                    row_dict = row.to_dict() if not isinstance(row, pd.DataFrame) else {col: row[col].tolist()[-1] for col in row.columns}
                    
                    row_modified = False

                    # 1. Update Name, IMO, Handover
                    for col_idx, db_col in [(EXCEL_COL_NAME, DB_NAME_COL), (EXCEL_COL_IMO, DB_IMO_COL), (EXCEL_COL_HANDOVER, DB_HANDOVER_COL)]:
                        val_db = get_db_val(row_dict, db_col)
                        val_ex = str(sheet_df.iat[i, col_idx]).strip()
                        if val_db and val_db != val_ex:
                            sheet_df.iat[i, col_idx] = val_db
                            row_modified = True

                    # 2. Update Shop Test (H) - Direct copy, no logic, ignore if already "shoptested"
                    current_h = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).lower()
                    if "shoptested" not in current_h:
                        val_db_h = get_db_val(row_dict, DB_SHOPTEST_COL)
                        if val_db_h and val_db_h != str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).strip():
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = val_db_h
                            row_modified = True
                    
                    if row_modified:
                        updated_rows_count += 1
            
            st.success(f"✅ Processing complete.")
            st.write(f"• Total Rows Checked: **{len(sheet_df)}**")
            st.write(f"• Vessels Actually Updated: **{updated_rows_count}**")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e:
            st.error(f"Error: {e}")
