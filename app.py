import streamlit as st
import pandas as pd
import io

st.title("Vessel Data Processor")
st.write("Upload your Excel sheet and CSV database to merge.")

# File Uploader
sheet_file = st.file_uploader("Upload Excel file (sheet_copy.xlsx)", type=['xlsx'])
db_file = st.file_uploader("Upload CSV database (db_sample.csv)", type=['csv'])

if sheet_file and db_file:
    if st.button("Process Data"):
        # Load data
        sheet_df = pd.read_excel(sheet_file)
        db_df = pd.read_csv(db_file)
        
        # --- YOUR LOGIC HERE ---
        # (I've kept your exact cleaning logic)
        WINDB_KEY_COL = 'DB Number'; WINDB_NAME_COL = 'Vessel_Name'
        WINDB_IMO_COL = 'IMO_Number'; WINDB_HANDOVER_COL = 'Handover_Date'
        WINDB_SHOPTEST_COL = 'Shop_Test_Date'

        def clean_val(v):
            if pd.isna(v) or v is None: return ''
            s = str(v).strip()
            if s.endswith('.0'): s = s[:-2]
            return s.replace(' 00:00:00', '').strip()

        db_df['CLEAN_KEY'] = db_df[WINDB_KEY_COL].apply(clean_val)
        db_indexed = db_df.set_index('CLEAN_KEY')

        for i in range(len(sheet_df)):
            primary_db = clean_val(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
            secondary_db = clean_val(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
            matched_id = primary_db if (primary_db in db_indexed.index) else (secondary_db if (secondary_db in db_indexed.index) else None)

            if matched_id:
                db_row = db_indexed.loc[matched_id]
                if isinstance(db_row, pd.DataFrame): db_row = db_row.iloc[0]
                sheet_df.iat[i, 0] = clean_val(db_row.get(WINDB_NAME_COL))
                sheet_df.iat[i, 2] = clean_val(db_row.get(WINDB_IMO_COL))
                sheet_df.iat[i, 4] = clean_val(db_row.get(WINDB_HANDOVER_COL))
                sheet_df.iat[i, 8] = clean_val(db_row.get(WINDB_SHOPTEST_COL))

        # --- DOWNLOAD ---
        output = io.BytesIO()
        sheet_df.to_excel(output, index=False)
        st.download_button("Download Updated File", data=output.getvalue(), file_name="sheet_updated.xlsx")
        st.success("Done!")
