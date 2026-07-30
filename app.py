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
            db_df = pd.read_csv(uploaded_db)

            # Robust cleaning: Converts everything to string, strips whitespace, lowers case
            def clean_key(v):
                if pd.isna(v): return ''
                return str(v).strip().lower()

            # Aggressive Date formatting: Force YYYY-MM-DD and remove time
            def format_date_strict(v):
                if pd.isna(v) or v == '': return ''
                # Try to parse as date
                try:
                    return pd.to_datetime(v).strftime('%Y-%m-%d')
                except:
                    # Fallback: if it's already a string like "2024-01-01 00:00:00", split it
                    s = str(v).split(' ')[0]
                    return s

            # Indexing (using the aggressive clean_key)
            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_count = 0
            for i in range(len(sheet_df)):
                # Clean the IDs in the Excel sheet exactly the same way
                primary = clean_key(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                
                # Try to find a match
                match_id = None
                if primary and primary in db_indexed.index:
                    match_id = primary
                elif secondary and secondary in db_indexed.index:
                    match_id = secondary

                if match_id:
                    row = db_indexed.loc[match_id]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    
                    # Apply changes ONLY if data exists in DB
                    new_name = row.get(DB_NAME_COL)
                    if pd.notna(new_name): sheet_df.iat[i, 0] = str(new_name).strip()
                    
                    new_imo = row.get(DB_IMO_COL)
                    if pd.notna(new_imo): sheet_df.iat[i, 2] = str(new_imo).strip()
                    
                    new_handover = row.get(DB_HANDOVER_COL)
                    if pd.notna(new_handover): sheet_df.iat[i, 4] = format_date_strict(new_handover)
                    
                    new_shoptest = row.get(DB_SHOPTEST_COL)
                    if pd.notna(new_shoptest): sheet_df.iat[i, 7] = format_date_strict(new_shoptest)
                    
                    updated_count += 1

            st.success(f"✅ Success! Updated {updated_count} rows.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
