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

# Exact Column Index Mapping (0-indexed: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7)
EXCEL_COL_NAME = 0         # Column A
EXCEL_COL_IMO = 2          # Column C
EXCEL_COL_HANDOVER = 4     # Column E
EXCEL_COL_PRIMARY_DB = 5   # Column F (First DB Number)
EXCEL_COL_SECONDARY_DB = 6 # Column G (Second DB Number)
EXCEL_COL_SHOPTEST = 7     # Column H (Shop Test Date)

# Setting layout="centered" ensures the app stays contained and readable on ultra-wide screens, 
# while naturally squeezing down perfectly for mobile screens.
st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")

st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (sheet_copy.xlsx)", type=['xlsx'])
uploaded_db = st.file_uploader("Upload CSV Database (db_sample.csv)", type=['csv'])

if uploaded_sheet and uploaded_db:
    
    with st.expander("🔍 Diagnostics (Use this if a vessel isn't updating)"):
        st.write("Type a DB Number below to verify what the script reads from the CSV.")
        test_id = st.text_input("Enter DB Number (e.g., 12345):")
        
    # use_container_width=True makes the button stretch nicely to fit its container
    if st.button("🚀 Process & Update Data", use_container_width=True) or test_id:
        try:
            sheet_df = pd.read_excel(uploaded_sheet)
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

            def get_best_val(row_dict, target_col):
                clean_target = target_col.replace('_', '').replace(' ', '').lower()
                for col_name, val in row_dict.items():
                    if col_name.replace('_', '').replace(' ', '').lower() == clean_target:
                        if str(val).strip() and str(val).strip().lower() not in ['nan', 'nat', 'none', '']:
                            return val
                return ''

            # Find key column in CSV
            db_id_actual = None
            for col in db_df.columns:
                if col.replace('_', '').replace(' ', '').lower() == DB_ID_COL.replace('_', '').replace(' ', '').lower():
                    db_id_actual = col
                    break
            
            if not db_id_actual:
                db_id_actual = db_df.columns[0] 

            db_df['CLEAN_KEY'] = db_df[db_id_actual].apply(clean_key)
            db_indexed = db_df.set_index('CLEAN_KEY')

            # --- DIAGNOSTICS MODE ---
            if test_id:
                clean_test = clean_key(test_id)
                st.subheader(f"Diagnostic Results for: {test_id}")
                if clean_test in db_indexed.index:
                    raw_data = db_indexed.loc[clean_test]
                    st.success("✅ ID Found in Database!")
                    st.write(raw_data)
                else:
                    st.error("❌ ID NOT FOUND in Database! Check if DB Number exists in CSV key column.")
                st.stop()

            # --- MAIN UPDATE LOOP ---
            total_rows = len(sheet_df)
            updated_count = 0
            
            for i in range(total_rows):
                
                # Fetch keys strictly from Column F (index 5) and Column G (index 6)
                primary = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB] if EXCEL_COL_PRIMARY_DB < sheet_df.shape[1] else '')
                secondary = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB] if EXCEL_COL_SECONDARY_DB < sheet_df.shape[1] else '')
                
                # Check Primary DB Number first, fallback to Second DB Number
                match_id = None
                if primary and primary in db_indexed.index:
                    match_id = primary
                elif secondary and secondary in db_indexed.index:
                    match_id = secondary

                if match_id:
                    row = db_indexed.loc[match_id]
                    
                    if isinstance(row, pd.DataFrame):
                        row_dict = {}
                        for col in row.columns:
                            valid_vals = [v for v in row[col].tolist() if str(v).strip() and str(v).strip().lower() not in ['nan', 'nat', 'none', '']]
                            row_dict[col] = valid_vals[-1] if valid_vals else ''
                    else:
                        row_dict = row.to_dict()
                    
                    # Update Columns securely
                    val_name = get_best_val(row_dict, DB_NAME_COL)
                    if val_name: sheet_df.iat[i, EXCEL_COL_NAME] = str(val_name).strip()
                    
                    val_imo = get_best_val(row_dict, DB_IMO_COL)
                    if val_imo: sheet_df.iat[i, EXCEL_COL_IMO] = str(val_imo).strip()
                    
                    val_handover = get_best_val(row_dict, DB_HANDOVER_COL)
                    formatted_handover = force_date(val_handover)
                    if formatted_handover: sheet_df.iat[i, EXCEL_COL_HANDOVER] = formatted_handover
                    
                    val_shop = get_best_val(row_dict, DB_SHOPTEST_COL)
                    formatted_shop = force_date(val_shop)
                    if formatted_shop: sheet_df.iat[i, EXCEL_COL_SHOPTEST] = formatted_shop
                    
                    updated_count += 1

            # --- ENHANCED SUCCESS MESSAGE ---
            st.success(f"✅ Data Merge Complete!")
            st.info(
                f"**Update Summary:**\n\n"
                f"• **{total_rows}** total vessels checked in the Excel file.\n"
                f"• **{updated_count}** vessels successfully matched with the database.\n"
                f"• Only outdated or blank cells were updated. All existing valid data was preserved."
            )
            
            output = io.BytesIO()
            sheet_df.to_excel(output, index=False)
            
            # The download button is also responsive now
            st.download_button(
                label="📥 Download Updated Excel", 
                data=output.getvalue(), 
                file_name="sheet_updated_audited.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        except Exception as e:
            st.error(f"Error processing files: {e}")
