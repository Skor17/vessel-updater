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
    
    # --- X-RAY DIAGNOSTICS ---
    with st.expander("🔍 Diagnostics (Use this if a vessel isn't updating)"):
        st.write("Type a DB Number below to see EXACTLY what the script reads from the CSV.")
        test_id = st.text_input("Enter DB Number (e.g., 12345):")
        
    if st.button("🚀 Process & Update Data") or test_id:
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
            # CRITICAL FIX 1: Auto-detect separator (handles European ; format)
            db_df = pd.read_csv(uploaded_db, dtype=str, sep=None, engine='python')
            db_df = db_df.fillna('')

            def clean_key(v):
                if pd.isna(v) or v is None: return ''
                s = str(v).strip()
                if s.endswith('.0'): s = s[:-2]
                return s.lower()

            def force_date(v):
                if v is None: return None
                v_str = str(v).strip()
                if not v_str or v_str.lower() in ['nan', 'nat', 'none', 'null', '']: 
                    return None
                try:
                    dt = pd.to_datetime(v_str, errors='coerce')
                    if pd.isna(dt): return v_str.split(' ')[0]
                    return dt.strftime('%Y-%m-%d')
                except:
                    return v_str.split(' ')[0]

            # CRITICAL FIX 2: Fuzzy Column Matching
            def get_best_val(row_dict, target_col):
                clean_target = target_col.replace('_', '').replace(' ', '').lower()
                for col_name, val in row_dict.items():
                    if col_name.replace('_', '').replace(' ', '').lower() == clean_target:
                        if str(val).strip() and str(val).strip().lower() not in ['nan', 'nat', 'none', '']:
                            return val
                return ''

            db_df['CLEAN_KEY'] = db_df[DB_ID_COL].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            # --- DIAGNOSTICS OUTPUT ---
            if test_id:
                clean_test = clean_key(test_id)
                st.subheader(f"Diagnostic Results for: {test_id}")
                st.write(f"Columns found in CSV: {db_df.columns.tolist()}")
                if clean_test in db_indexed.index:
                    raw_data = db_indexed.loc[clean_test]
                    st.success("✅ ID Found in Database!")
                    st.write(raw_data)
                else:
                    st.error("❌ ID NOT FOUND in Database! Check for typos.")
                st.stop() # Stop the rest of the script so you can read the output

            # --- MAIN PROCESSING ---
            updated_count = 0
            for i in range(len(sheet_df)):
                primary = clean_key(sheet_df.iat[i, 6] if 6 < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, 7] if 7 < sheet_df.shape[1] else '')
                match_id = primary if primary in db_indexed.index else (secondary if secondary in db_indexed.index else None)

                if match_id:
                    row = db_indexed.loc[match_id]
                    
                    # CRITICAL FIX 3: Handling Duplicates gracefully
                    if isinstance(row, pd.DataFrame):
                        # If multiple rows exist, create a dict of the LAST valid entry
                        row_dict = {}
                        for col in row.columns:
                            valid_vals = [v for v in row[col].tolist() if str(v).strip() and str(v).strip().lower() != 'nan']
                            row_dict[col] = valid_vals[-1] if valid_vals else ''
                    else:
                        row_dict = row.to_dict()
                    
                    # Updates using Fuzzy Matching
                    val_name = get_best_val(row_dict, DB_NAME_COL)
                    if val_name: sheet_df.iat[i, 0] = str(val_name).strip()
                    
                    val_imo = get_best_val(row_dict, DB_IMO_COL)
                    if val_imo: sheet_df.iat[i, 2] = str(val_imo).strip()
                    
                    val_handover = get_best_val(row_dict, DB_HANDOVER_COL)
                    if force_date(val_handover): sheet_df.iat[i, 4] = force_date(val_handover)
                    
                    val_shop = get_best_val(row_dict, DB_SHOPTEST_COL)
                    if force_date(val_shop): sheet_df.iat[i, 7] = force_date(val_shop)
                    
                    updated_count += 1

            st.success(f"✅ Success! Processed {updated_count} matches. Data preserved where DB was empty.")
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            st.download_button("📥 Download Updated Excel", data=output.getvalue(), file_name="sheet_updated_audited.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        except Exception as e:
            st.error(f"Error: {e}")
