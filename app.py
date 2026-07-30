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

# --- PAGE SETUP ---
st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title(f"🚢 Welcome, {BOSS_NAME}!")
st.subheader("Vessel Database Updater")
st.write("Upload your files below to sync the database.")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Update Data"):
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db)

            # Cleanup helper
            def clean_val(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.replace(' 00:00:00', '').strip()

            # Date formatter
            def format_date(v):
                try:
                    # Convert to datetime object and format as YYYY-MM-DD
                    return pd.to_datetime(v).strftime('%Y-%m-%d')
                except:
                    return clean_val(v) # Fallback if not a date

            # Indexing
            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_val)
            db_indexed = db_df.set_index('CLEAN_KEY')

            updated_count = 0
            for i in range(len(sheet_df)):
                primary = clean_val(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_val(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                match = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match:
                    row = db_indexed.loc[match]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    
                    # Apply changes ONLY to specific columns (A, C, E, H)
                    sheet_df.iat[i, 0] = clean_val(row.get(DB_NAME_COL))      # Col A
                    sheet_df.iat[i, 2] = clean_val(row.get(DB_IMO_COL))       # Col C
                    sheet_df.iat[i, 4] = format_date(row.get(DB_HANDOVER_COL)) # Col E (Handover)
                    sheet_df.iat[i, 7] = clean_val(row.get(DB_SHOPTEST_COL))   # Col H (Shop Test)
                    
                    updated_count += 1

            st.success(f"✅ Success! Updated {updated_count} rows.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error: {e}")
