import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Order Merger Tool", layout="wide")
st.title("📦 Order Merger Tool")

st.markdown("Upload two order files (Excel or CSV) from different branches. This app will combine them by item and show total quantities.")

# File upload
file1 = st.file_uploader("Upload First Order File", type=["xlsx", "xls", "csv"])
file2 = st.file_uploader("Upload Second Order File", type=["xlsx", "xls", "csv"])

def read_file(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)

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
        merged = merged[['Item', 'First Order Quantity', 'Second Order Quantity', 'Total Quantity']]

        st.success("✅ Files merged successfully!")
        st.dataframe(merged)

        # Export to Excel
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_data = to_excel(merged)

        st.download_button(
            label="📥 Download Combined Orders Excel",
            data=excel_data,
            file_name="Combined_Orders.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Something went wrong: {e}")
