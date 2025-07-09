import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image
import os

# ----------------------------
# PAGE CONFIGURATION
# ----------------------------
st.set_page_config(page_title="Denta Quick Merger", page_icon="🦷", layout="centered")

# ----------------------------
# LOGO + HEADER
# ----------------------------
if os.path.exists("DentaQuickEgypt.png"):
    logo = Image.open("DentaQuickEgypt.png")
    st.image(logo)

st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Branch Order Merger</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload old and (optional) new branch order Excel files. Each sheet must start from row 15 and include columns: <strong>الصنف</strong> and <strong>الكمية</strong>.</p>", unsafe_allow_html=True)
st.divider()

# ----------------------------
# UPLOAD SECTION
# ----------------------------
st.subheader("🗂️ Step 1: Upload Excel Files")
old_file = st.file_uploader("Upload OLD Orders File (multi-sheet Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW Orders File (optional)", type=["xlsx"])

# ----------------------------
# HELPER FUNCTIONS
# ----------------------------
def process_multisheet_excel(uploaded_file, header_row=14):
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=header_row)
    cleaned_sheets = {}

    for name, df in all_sheets.items():
        if 'الصنف' in df.columns and 'الكمية' in df.columns:
            df = df[['الصنف', 'الكمية']].dropna()
            df['الصنف'] = df['الصنف'].astype(str).str.strip().str.lower()
            df = df.groupby('الصنف', as_index=False).sum()
            df.columns = ['الصنف', name]
            cleaned_sheets[name] = df
    return cleaned_sheets

def merge_sheets(sheet_dict):
    if not sheet_dict:
        return pd.DataFrame()
    return reduce(lambda left, right: pd.merge(left, right, on='الصنف', how='outer'), sheet_dict.values()).fillna(0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Merged Orders")
    return output.getvalue()

# ----------------------------
# MAIN LOGIC
# ----------------------------
if old_file:
    try:
        old_sheets = process_multisheet_excel(old_file)
        old_merged = merge_sheets(old_sheets)

        if not old_merged.empty:
            if new_file:
                new_sheets = process_multisheet_excel(new_file)
                new_merged = merge_sheets(new_sheets)

                # Merge old + new into combined
                combined = pd.merge(old_merged, new_merged, on='الصنف', how='outer').fillna(0)

                # Recalculate column names from final combined
                old_cols = [col for col in old_merged.columns if col != 'الصنف' and col in combined.columns]
                new_cols = [col for col in new_merged.columns if col != 'الصنف' and col in combined.columns]

                # Summary columns
                combined['Old Quantity'] = combined[old_cols].sum(axis=1)
                combined['New Quantity'] = combined[new_cols].sum(axis=1)
                combined['Total Quantity'] = combined['Old Quantity'] + combined['New Quantity']
                combined['الصنف'] = combined['الصنف'].str.title()

                final_df = combined[['الصنف', 'Old Quantity', 'New Quantity', 'Total Quantity']]
                st.subheader("📋 Step 2: Summary of Old + New Orders")

            else:
                old_cols = [col for col in old_merged.columns if col != 'الصنف']
                old_merged['Old Quantity'] = old_merged[old_cols].sum(axis=1)
                old_merged['الصنف'] = old_merged['الصنف'].str.title()
                final_df = old_merged[['الصنف', 'Old Quantity']]
                st.subheader("📋 Step 2: Summary of OLD Orders Only")

            final_df = final_df.sort_values(by=final_df.columns[-1], ascending=False)
            st.dataframe(final_df, use_container_width=True)

            # ----------------------------
            # DOWNLOAD EXCEL
            # ----------------------------
            excel_data = to_excel(final_df)
            st.subheader("📥 Step 3: Download Excel Report")
            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_data,
                file_name="Merged_Orders_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.warning("⚠️ The uploaded OLD file doesn't contain any sheets with the required columns.")

    except Exception as e:
        st.error("❌ Error while processing the uploaded files.")
        st.exception(e)
