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
    st.image(logo, width=180)

st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Branch Order Merger</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload old and (optionally) new branch order Excel files. Each sheet must start from row 15 and include columns: <strong>الصنف</strong> and <strong>الكمية</strong>.</p>", unsafe_allow_html=True)
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
            # 🔹 Step 2: Show OLD Summary
            old_cols = [col for col in old_merged.columns if col != 'الصنف']
            old_merged['Old Quantity'] = old_merged[old_cols].sum(axis=1)
            old_merged['الصنف'] = old_merged['الصنف'].str.title()
            df_old_summary = old_merged[['الصنف', 'Old Quantity']]
            df_old_summary = df_old_summary.sort_values(by='Old Quantity', ascending=False)

            st.subheader("📋 Step 2: Summary of OLD Orders")
            st.dataframe(df_old_summary, use_container_width=True)

            if new_file:
                new_sheets = process_multisheet_excel(new_file)
                new_merged = merge_sheets(new_sheets)

                if not new_merged.empty:
                    # 🔹 Step 3: Show NEW Summary
                    new_cols = [col for col in new_merged.columns if col != 'الصنف']
                    new_merged['New Quantity'] = new_merged[new_cols].sum(axis=1)
                    new_merged['الصنف'] = new_merged['الصنف'].str.title()
                    df_new_summary = new_merged[['الصنف', 'New Quantity']]
                    df_new_summary = df_new_summary.sort_values(by='New Quantity', ascending=False)

                    st.subheader("📋 Step 3: Summary of NEW Orders")
                    st.dataframe(df_new_summary, use_container_width=True)

                    # 🔹 Step 4: Merge both summaries
                    combined = pd.merge(df_old_summary, df_new_summary, on='الصنف', how='outer').fillna(0)
                    combined['Total Quantity'] = combined['Old Quantity'] + combined['New Quantity']
                    combined = combined.sort_values(by='Total Quantity', ascending=False)

                    st.subheader("📋 Step 4: Summary of OLD + NEW Orders")
                    st.dataframe(combined, use_container_width=True)

                    # 🔹 Step 5: Download
                    excel_data = to_excel(combined)
                    st.subheader("📥 Step 5: Download Merged Excel Report")
                    st.download_button(
                        label="⬇️ Download Excel File",
                        data=excel_data,
                        file_name="Merged_Orders_Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ The NEW file didn’t contain valid data sheets.")
        else:
            st.warning("⚠️ The OLD file didn’t contain valid data sheets.")

    except Exception as e:
        st.error("❌ Error while processing the uploaded files.")
        st.exception(e)
