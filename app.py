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
EXCEL_COL_NAME = 0         # Column A
EXCEL_COL_IMO = 2          # Column C
EXCEL_COL_HANDOVER = 4     # Column E
EXCEL_COL_PRIMARY_DB = 5   # Column F (First DB Number)
EXCEL_COL_SECONDARY_DB = 6 # Column G (Second DB Number)
EXCEL_COL_SHOPTEST = 7     # Column H (Shop Test Date)

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")

st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    
    if st.button("🚀 Process & Create Updated File", use_container_width=True):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db, dtype=str, sep=None, engine='python')
            db_df = db_df.fillna('')

            # --- HELPER FUNCTIONS ---
            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            def get_best_val(row_dict, target_col):
                clean_target = target_col.replace('_', '').replace(' ', '').lower()
                for col_name, val in row_dict.items():
                    if col_name.replace('_', '').replace(' ', '').lower() == clean_target:
                        return val
                return ''

            # Indexing
            db_id_actual = next((c for c in db_df.columns if c.replace('_', '').replace(' ', '').lower() == DB_ID_COL.replace('_', '').replace(' ', '').lower()), db_df.columns[0])
            db_df['CLEAN_KEY'] = db_df[db_id_actual].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            # --- UPDATE LOOP ---
            updated_count = 0
            total_rows = len(sheet_df)
            
            for i in range(total_rows):
                primary = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB] if EXCEL_COL_PRIMARY_DB < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB] if EXCEL_COL_SECONDARY_DB < sheet_df.shape[1] else '')
                match_id = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match_id:
                    # WE FOUND A MATCH
                    updated_count += 1
                    row = db_indexed.loc[match_id]
                    row_dict = row.to_dict() if not isinstance(row, pd.DataFrame) else {col: row[col].tolist()[-1] for col in row.columns}
                    
                    # Update Vessel Name
                    if val := get_best_val(row_dict, DB_NAME_COL): 
                        sheet_df.iat[i, EXCEL_COL_NAME] = str(val).strip()
                    
                    # Update IMO
                    if val := get_best_val(row_dict, DB_IMO_COL): 
                        sheet_df.iat[i, EXCEL_COL_IMO] = str(val).strip()
                    
                    # Update Handover Date (Keep raw, don't force date formatting)
                    if val := get_best_val(row_dict, DB_HANDOVER_COL):
                        sheet_df.iat[i, EXCEL_COL_HANDOVER] = str(val).strip()
                    
                    # Update Shop Test Date (Column H)
                    current_val_h = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).lower()
                    
                    # ONLY update if "shoptested" is NOT already in the cell
                    if "shoptested" not in current_val_h:
                        val_from_db = get_best_val(row_dict, DB_SHOPTEST_COL)
                        if val_from_db:
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = str(val_from_db).strip()
            
            # --- SUMMARY OUTPUT ---
            st.success(f"✅ Processing complete.")
            st.write(f"**Results Summary:**")
            st.write(f"• Total Rows Checked: **{total_rows}**")
            st.write(f"• Matching Vessels Updated: **{updated_count}**")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            
            st.download_button(
                label="📥 Download Updated Excel File", 
                data=output.getvalue(), 
                file_name="sheet_updated_audited.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error processing files: {e}")
