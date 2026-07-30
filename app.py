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
st.write("Processing files to update Vessel Database.")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Update Data"):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            
            # CRITICAL: Read CSV as string (dtype=str) to prevent time-tag auto-formatting
            db_df = pd.read_csv(uploaded_db, dtype=str)
            db_df = db_df.fillna('') # Convert empty cells to empty strings

            # ID Cleaning Function
            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            # Date Formatting: Force YYYY-MM-DD
            def force_date(v):
                if not v or v.lower() == 'nan' or v.strip() == '': return ''
                try:
                    # Parse whatever is in there and force YYYY-MM-DD
                    return pd.to_datetime(v).strftime('%Y-%m-%d')
                except:
                    # If it fails, return the string as-is (but stripped)
                    return v.split(' ')[0]

            # Indexing
            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_count = 0
            for i in range(len(sheet_df)):
                primary = clean_key(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                
                match_id = None
                if primary and primary in db_indexed.index:
                    match_id = primary
                elif secondary and secondary in db_indexed.index:
                    match_id = secondary

                if match_id:
                    row = db_indexed.loc[match_id]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    
                    # FORCED OVERWRITE: Regardless of what was there, put the new value
                    # A = 0, C = 2, E = 4, H = 7
                    sheet_df.iat[i, 0] = str(row.get(DB_NAME_COL, ''))
                    sheet_df.iat[i, 2] = str(row.get(DB_IMO_COL, ''))
                    sheet_df.iat[i, 4] = force_date(row.get(DB_HANDOVER_COL, ''))
                    sheet_df.iat[i, 7] = force_date(row.get(DB_SHOPTEST_COL, ''))
                    
                    updated_count += 1

            if updated_count > 0:
                st.success(f"✅ Success! Overwrote {updated_count} rows with database data.")
            else:
                st.warning("No matches found. Please check if your ID columns are correct.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
