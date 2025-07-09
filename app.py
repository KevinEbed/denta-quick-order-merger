import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image

st.set_page_config(page_title="Denta Quick", page_icon="🦷", layout="centered")

# Show logo and app title
logo = Image.open("DentaQuickEgypt.png")
st.image(logo)

st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Branch Order Merger</h2>", unsafe_allow_html=True)
st.divider()

st.set_page_config(page_title="📦 Order Merger Tool", layout="centered")

st.markdown("<h1 style='text-align: center;'>📦 Multi-Branch Order Merger</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload your old and new order Excel files. Each sheet must contain quantities per branch starting at row 15.</p>", unsafe_allow_html=True)
st.divider()

# Upload Section
st.subheader("🗂️ Step 1: Upload Your Files")
old_file = st.file_uploader("Upload OLD Orders File (multi-sheet Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW Orders File (multi-sheet Excel)", type=["xlsx"])

def process_multisheet_excel(uploaded_file, header_row=14):
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=header_row)
    cleaned_sheets = {}

    for name, df in all_sheets.items():
        if 'الصنف' in df.columns and 'الكمية' in df.columns:
            df = df[['الصنف', 'الكمية']].dropna()
            df['الصنف'] = df['الصنف'].str.strip().str.lower()
            df = df.groupby('الصنف', as_index=False).sum()
            df.columns = ['الصنف', name]  # Name = sheet name (branch)
            cleaned_sheets[name] = df
    return cleaned_sheets

def merge_sheets(sheet_dict):
    if not sheet_dict:
        return pd.DataFrame()
    merged = reduce(lambda left, right: pd.merge(left, right, on='الصنف', how='outer'), sheet_dict.values())
    return merged.fillna(0)

if old_file and new_file:
    try:
        st.success("✅ Files uploaded successfully!")

        old_sheets = process_multisheet_excel(old_file)
        new_sheets = process_multisheet_excel(new_file)

        old_merged = merge_sheets(old_sheets)
        new_merged = merge_sheets(new_sheets)

        # Merge old and new merged data
        combined = pd.merge(old_merged, new_merged, on='الصنف', how='outer').fillna(0)

        # Find all quantity columns (exclude 'الصنف')
        quantity_cols = [col for col in combined.columns if col != 'الصنف']
        combined['اجمالي الكمية'] = combined[quantity_cols].sum(axis=1)
        combined['الصنف'] = combined['الصنف'].str.title()

        combined = combined[['الصنف'] + quantity_cols + ['اجمالي الكمية']]
        combined = combined.sort_values(by='اجمالي الكمية', ascending=False)

        st.subheader("📋 Step 2: Preview Combined Orders")
        st.dataframe(combined, use_container_width=True)

        # Downloadable Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = to_excel(combined)

        st.subheader("📥 Step 3: Download Final Merged File")
        st.download_button(
            label="Download Combined Orders Excel",
            data=excel_data,
            file_name="Combined_Orders.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("❌ Error while processing files.")
        st.exception(e)
