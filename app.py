import streamlit as st
import pandas as pd
from io import BytesIO

# Page setup
st.set_page_config(page_title="Branch Order Merger", page_icon="📦", layout="centered")

# Header
st.markdown("""
    <h1 style='text-align: center; color: #3B7A57;'>📦 Branch Order Merger</h1>
    <p style='text-align: center;'>Easily merge order files from multiple branches and generate a clean summary report.</p>
""", unsafe_allow_html=True)

st.divider()

# Upload section
st.subheader("Step 1: Upload Your Order Files")

file1 = st.file_uploader("🗂️ Upload First Branch Order (Excel or CSV)", type=["xlsx", "xls", "csv"])
file2 = st.file_uploader("🗂️ Upload Second Branch Order (Excel or CSV)", type=["xlsx", "xls", "csv"])

def read_file(file):
    if file.name.endswith('.csv'):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

if file1 and file2:
    try:
        df1 = read_file(file1)[['الصنف', 'الكمية']]
        df2 = read_file(file2)[['الصنف', 'الكمية']]

        df1.columns = ['Item', 'First Order Quantity']
        df2.columns = ['Item', 'Second Order Quantity']

        df1['Item'] = df1['Item'].str.strip().str.lower()
        df2['Item'] = df2['Item'].str.strip().str.lower()

        merged = pd.merge(df1, df2, on='Item', how='outer').fillna(0)
        merged['First Order Quantity'] = merged['First Order Quantity'].astype(int)
        merged['Second Order Quantity'] = merged['Second Order Quantity'].astype(int)
        merged['Total Quantity'] = merged['First Order Quantity'] + merged['Second Order Quantity']
        merged['Item'] = merged['Item'].str.title()

        final = merged[['Item', 'First Order Quantity', 'Second Order Quantity', 'Total Quantity']]

        st.success("✅ Orders merged successfully!")

        st.subheader("Step 2: Preview Merged Report")
        st.dataframe(final, use_container_width=True)

        # Excel export
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = to_excel(final)

        st.subheader("Step 3: Download Your Merged Report")
        st.download_button(
            label="📥 Download Excel File",
            data=excel_data,
            file_name="Merged_Branch_Orders.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("⚠️ Something went wrong. Please check your files.")
        st.exception(e)
