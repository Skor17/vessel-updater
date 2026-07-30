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

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢")

st.title(f"🚢 Welcome, {BOSS_NAME}!")
st.write("Processing files to update the database while preserving existing data.")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Update Data"):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db, dtype=str)
            db_df = db_df.fillna('')

            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            def force_date(v):
                v_str = str(v).strip()
                # Catch all forms of empty or invalid data that Excel creates
                if not v_str or v_str.lower() in ['nan', 'nat', 'none', 'null', '']: 
                    return None
                try:
                    # errors='coerce' forces bad dates to become blank instead of crashing
                    dt = pd.to_datetime(v_str, errors='coerce')
                    if pd.isna(dt): 
                        return v_str.split(' ')[0]
                    return dt.strftime('%Y-%m-%d')
                except:
                    return v_str.split(' ')[0]

            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_key)
            
            # CRITICAL FIX: If there are duplicates in the DB, always keep the LAST one (the newest)
            db_df = db_df.drop_duplicates(subset=['CLEAN_KEY'], keep='last')
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_count = 0
            for i in range(len(sheet_df)):
                primary = clean_key(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                
                match_id = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match_id:
                    row = db_indexed.loc[match_id]
                    
                    # Column A (Name)
                    val = row.get(DB_NAME_COL, '')
                    if val and str(val).strip() and str(val).strip().lower() != 'nan': 
                        sheet_df.iat[i, 0] = str(val).strip()
                    
                    # Column C (IMO)
                    val = row.get(DB_IMO_COL, '')
                    if val and str(val).strip() and str(val).strip().lower() != 'nan': 
                        sheet_df.iat[i, 2] = str(val).strip()
                    
                    # Column E (Handover)
                    val = row.get(DB_HANDOVER_COL, '')
                    formatted_val = force_date(val)
                    if formatted_val: 
                        sheet_df.iat[i, 4] = formatted_val
                    
                    # Column H (Shop Test)
                    val = row.get(DB_SHOPTEST_COL, '')
                    formatted_val = force_date(val)
                    if formatted_val: 
                        sheet_df.iat[i, 7] = formatted_val
                    
                    updated_count += 1

            st.success(f"✅ Success! Processed {updated_count} matches. Data preserved where DB was empty.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
