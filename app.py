import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Vessel Data Processor", layout="wide")
st.title("🚢 Vessel Data Processor")
st.write("Upload your files below. The processor will only update columns A through H.")

# --- FILE UPLOADS ---
col1, col2 = st.columns(2)
with col1:
    sheet_file = st.file_uploader("1. Upload Excel sheet (Target)", type=['xlsx'])
with col2:
    db_file = st.file_uploader("2. Upload CSV Database (Source)", type=['csv'])

if sheet_file and db_file:
    # Load files
    sheet_df = pd.read_excel(sheet_file)
    db_df = pd.read_csv(db_file)
    
    # --- PREVIEW AREA ---
    with st.expander("Click to preview data (To verify column names)"):
        st.write("Database Preview:", db_df.head(2))
    
    # --- DYNAMIC MAPPING ---
    st.subheader("Map your Database Columns")
    db_cols = db_df.columns.tolist()
    sheet_cols = sheet_df.columns.tolist()

    # User picks which columns to use
    id_col = st.selectbox("Which DB column contains the ID (Key)?", db_cols)
    name_col = st.selectbox("Database column for Vessel Name:", db_cols)
    imo_col = st.selectbox("Database column for IMO Number:", db_cols)
    handover_col = st.selectbox("Database column for Handover Date:", db_cols)
    shoptest_col = st.selectbox("Database column for Shop Test Date (Select Column H):", db_cols)
    
    st.write("---")
    sheet_id_col = st.selectbox("Which Excel column contains the ID (to match)?", sheet_cols)

    if st.button("Process Data"):
        # Processing Logic
        def clean_val(v):
            if pd.isna(v) or v is None: return ''
            s = str(v).strip()
            if s.endswith('.0'): s = s[:-2]
            return s.replace(' 00:00:00', '').strip()

        # Index the database
        db_df['CLEAN_KEY'] = db_df[id_col].apply(clean_val)
        db_indexed = db_df.set_index('CLEAN_KEY')

        # Update Loop
        updated_count = 0
        for i in range(len(sheet_df)):
            key = clean_val(sheet_df.at[i, sheet_id_col])
            
            if key in db_indexed.index:
                row = db_indexed.loc[key]
                if isinstance(row, pd.DataFrame): row = row.iloc[0] # Handle duplicates
                
                # Update specific columns
                sheet_df.at[i, 'Vessel_Name'] = row[name_col]
                sheet_df.at[i, 'IMO_Number'] = row[imo_col]
                sheet_df.at[i, 'Handover_Date'] = row[handover_col]
                sheet_df.at[i, 'Shop_Test_Date'] = row[shoptest_col]
                
                updated_count += 1

        # Download
        output = io.BytesIO()
        sheet_df.to_excel(output, index=False)
        st.success(f"Done! Updated {updated_count} rows.")
        st.download_button("Download Updated File", data=output.getvalue(), file_name="sheet_updated_audited.xlsx")
