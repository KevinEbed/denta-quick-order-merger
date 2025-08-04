import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image
import os

st.set_page_config(page_title="Denta Quick Merger", page_icon="🦷", layout="centered")

if os.path.exists("DentaQuickEgypt.png"):
    logo = Image.open("DentaQuickEgypt.png")
    st.image(logo)

st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Branch Order Merger</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload old and (optionally) new branch order Excel files. Handles both old and new formats automatically.</p>", unsafe_allow_html=True)
st.divider()

st.subheader("🗂️ Step 1: Upload Excel Files")
uploaded_file = st.file_uploader("Upload Orders File (multi-sheet Excel)", type=["xlsx"])

# ----------- New Generalized Header Logic -----------

HEADER_KEYWORDS = {
    'equipment_name': ['equipment name', 'اسم الجهاز', 'الصنف', 'item', 'product'],
    'number': ['number', 'العدد', 'qty', 'quantity', 'qyt', 'الكمية'],
    'notes': ['notes', 'ملاحظات'],
    'serial': ['serial', 'رقم', 'الرقم']
}

def detect_column(col):
    col = str(col).strip().lower()
    for key, options in HEADER_KEYWORDS.items():
        if any(opt in col for opt in options):
            return key
    return col

def find_header_row(df):
    for i in range(min(20, len(df))):
        row = df.iloc[i].astype(str).str.lower().tolist()
        if any(any(h in cell for h in HEADER_KEYWORDS['equipment_name']) for cell in row):
            return i
    return None

def process_multisheet_general(file):
    xl = pd.ExcelFile(file)
    all_data = []

    for sheet_name in xl.sheet_names:
        df_raw = xl.parse(sheet_name, header=None)
        header_row = find_header_row(df_raw)
        if header_row is None:
            continue

        df = xl.parse(sheet_name, header=header_row)
        df.columns = [detect_column(c) for c in df.columns]

        for col in ['equipment_name', 'number']:
            if col not in df.columns:
                df[col] = None

        if 'notes' not in df.columns:
            df['notes'] = ''

        df = df[['equipment_name', 'number', 'notes']].dropna(subset=['equipment_name'])
        df['equipment_name'] = df['equipment_name'].astype(str).str.strip().str.title()
        df['notes'] = df['notes'].astype(str).str.strip()
        df['number'] = pd.to_numeric(df['number'], errors='coerce').fillna(0)

        all_data.append(df)

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        summary_df = combined_df.groupby(['equipment_name', 'notes'], dropna=False)['number'].sum().reset_index()
        summary_df = summary_df.sort_values(by='number', ascending=False)
        return summary_df
    else:
        return pd.DataFrame()

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Merged Summary")
    return output.getvalue()

# ----------- Streamlit Execution -----------

if uploaded_file:
    try:
        result_df = process_multisheet_general(uploaded_file)
        if not result_df.empty:
            st.subheader("📋 Combined Equipment Summary")
            st.dataframe(result_df, use_container_width=True)

            excel_data = to_excel(result_df)
            st.subheader("📥 Download Summary Report")
            st.download_button("⬇️ Download Excel File", data=excel_data, file_name="Combined_Equipment_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ No valid sheets or data found in the file.")
    except Exception as e:
        st.error("❌ Error while processing the uploaded file.")
        st.exception(e)
