import streamlit as st
import pandas as pd
import io

# --- CONFIGURATION ---
BOSS_NAME = "Xu Zhi Jun"

# Updated Column Indices (0-indexed) based on new requirements
EXCEL_COL_NAME = 0          # A - Vessel Name
EXCEL_COL_IMO = 2           # C - IMO Number
EXCEL_COL_HANDOVER = 4      # E - Handover Date
EXCEL_COL_PRIMARY_DB = 5    # F - Primary ID (Column F)
EXCEL_COL_SECONDARY_DB = 6  # G - Secondary ID (Column G)
EXCEL_COL_SHOPTEST = 7      # H - Shop Test Date (Column H)

HANDOVER_CUTOFF = pd.Timestamp("2025-01-01")

st.set_page_config(page_title="Vessel Data Updater", page_icon="🚢", layout="centered")
st.title(f"🚢 Welcome, {BOSS_NAME}!")

uploaded_sheet = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
uploaded_db = st.file_uploader("Upload CSV Database (.csv)", type=["csv"])

if uploaded_sheet and uploaded_db:
    if st.button("🚀 Update Data", use_container_width=True):
        try:
            # Read files as strings to avoid dtype surprises
            sheet_df = pd.read_excel(uploaded_sheet, header=0, dtype=str).fillna("")
            db_df = pd.read_csv(uploaded_db, header=0, dtype=str).fillna("")

            # Normalise a lookup key: strip, lowercase, remove trailing ".0" and " 00:00:00"
            def clean_key(v: str) -> str:
                s = str(v).strip().lower().replace(" 00:00:00", "")
                if s.endswith(".0"):
                    s = s[:-2]
                return s

            # Build lookup: key -> first matching row as a plain dict
            # Handling repeated vessels in DB: first one wins
            db_id_col = db_df.columns[0]
            db_lookup: dict[str, dict] = {}
            for _, db_row in db_df.iterrows():
                key = clean_key(db_row[db_id_col])
                if key and key not in db_lookup:
                    db_lookup[key] = db_row.to_dict()

            csv_cols = list(db_df.columns)

            def get_csv_val(db_row_dict: dict, preferred_name: str, position: int):
                """Return a non-empty, non-nan CSV value or None."""
                if preferred_name in db_row_dict:
                    v = str(db_row_dict[preferred_name]).strip()
                elif position < len(csv_cols):
                    v = str(db_row_dict.get(csv_cols[position], "")).strip()
                else:
                    return None
                return v if v and v.lower() != "nan" else None

            rows_checked = len(sheet_df)
            updated_rows = 0

            # Handling repeated vessels in Excel: iterate through all rows
            for i in range(rows_checked):
                # Resolve match: Primary ID (Col F) first, then Secondary ID (Col G)
                primary_key = clean_key(sheet_df.iat[i, EXCEL_COL_PRIMARY_DB])
                secondary_key = clean_key(sheet_df.iat[i, EXCEL_COL_SECONDARY_DB])

                if primary_key in db_lookup:
                    db_row_dict = db_lookup[primary_key]
                elif secondary_key in db_lookup:
                    db_row_dict = db_lookup[secondary_key]
                else:
                    continue  # No match — leave row untouched

                row_modified = False

                # Name (Col A) — direct overwrite
                val_name = get_csv_val(db_row_dict, "Vessel_Name", 1)
                if val_name is not None:
                    if val_name != sheet_df.iat[i, EXCEL_COL_NAME]:
                        sheet_df.iat[i, EXCEL_COL_NAME] = val_name
                        row_modified = True

                # IMO (Col C) — direct overwrite
                val_imo = get_csv_val(db_row_dict, "IMO_Number", 2)
                if val_imo is not None:
                    if val_imo != sheet_df.iat[i, EXCEL_COL_IMO]:
                        sheet_df.iat[i, EXCEL_COL_IMO] = val_imo
                        row_modified = True

                # Handover Date (Col E) — update only if CSV date >= 2025-01-01
                val_handover = get_csv_val(db_row_dict, "Handover_Date", 3)
                if val_handover is not None:
                    try:
                        # Clean date string before parsing
                        clean_date_str = val_handover.replace(" 00:00:00", "")
                        csv_date = pd.to_datetime(clean_date_str, dayfirst=False, errors="raise")
                        if csv_date >= HANDOVER_CUTOFF:
                            if val_handover != sheet_df.iat[i, EXCEL_COL_HANDOVER]:
                                sheet_df.iat[i, EXCEL_COL_HANDOVER] = val_handover
                                row_modified = True
                    except Exception:
                        pass  # Unparseable date — skip

                # Shop Test Date (Col H) — update only if cell does NOT contain 'shoptested'
                val_shoptest = get_csv_val(db_row_dict, "Shop_Test_Date", 4)
                if val_shoptest is not None:
                    current_shoptest = str(sheet_df.iat[i, EXCEL_COL_SHOPTEST]).strip().lower()
                    if "shoptested" not in current_shoptest:
                        if val_shoptest != sheet_df.iat[i, EXCEL_COL_SHOPTEST]:
                            sheet_df.iat[i, EXCEL_COL_SHOPTEST] = val_shoptest
                            row_modified = True

                if row_modified:
                    updated_rows += 1

            # --- Report ---
            st.success("✅ Processing complete.")
            st.write(f"• Total Rows Checked: **{rows_checked}**")
            st.write(f"• Rows Actually Modified: **{updated_rows}**")

            # --- Export ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                sheet_df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label="📥 Download Updated Excel",
                data=output.getvalue(),
                file_name="updated_vessels.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"Error: {e}")
