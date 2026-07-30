import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Boss"  # Change this to your boss's name
# Hardcoded Column Names (Matches the database CSV)
DB_ID_COL = 'DB Number'
DB_NAME_COL = 'Vessel_Name'
DB_IMO_COL = 'IMO_Number'
DB_HANDOVER_COL = 'Handover_Date'
DB_SHOPTEST_COL = 'Shop_Test_Date' # Update this string if the CSV header is different

# --- PAGE SETUP ---
st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢")

# Aesthetics: Custom CSS for a professional look
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title(f"🚢 Welcome, {BOSS_NAME}!")
st.subheader("Vessel Database Updater")
st.write("Please upload the **Excel sheet** and the **CSV database** below to sync your files.")

# --- FILE UPLOADS ---
uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Process & Update Data"):
        try:
            # Load Data
            sheet_df = pd.read_excel(uploaded_sheet)
            db_df = pd.read_csv(uploaded_db)

            # Cleanup helper
            def clean_val(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.replace(' 00:00:00', '').strip()

            # Indexing Database
            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_val)
            db_indexed = db_df.set_index('CLEAN_KEY')

            # Processing
            updated_count = 0
            for i in range(len(sheet_df)):
                # Match against Col G (Index 6) or Col H (Index 7)
                primary = clean_val(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_val(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                
                match = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match:
                    row = db_indexed.loc[match]
                    if isinstance(row, pd.DataFrame): row = row.iloc[0]
                    
                    # Apply changes to specific columns
                    sheet_df.iat[i, 0] = clean_val(row.get(DB_NAME_COL)) # Col A
                    sheet_df.iat[i, 2] = clean_val(row.get(DB_IMO_COL))  # Col C
                    sheet_df.iat[i, 4] = clean_val(row.get(DB_HANDOVER_COL)) # Col E
                    # Use Index 7 (Col H) as requested for Shop Test
                    sheet_df.iat[i, 7] = clean_val(row.get(DB_SHOPTEST_COL)) 
                    
                    updated_count += 1

            # Success & Download
            st.success(f"✅ Success! Updated {updated_count} rows.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button(
                label="📥 Download Updated Excel",
                data=output.getvalue(),
                file_name="sheet_updated_audited.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")
