import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image
import os

# ------------------ Page Config & Branding ------------------ #
st.set_page_config(page_title="Denta Quick Merger", page_icon="🦷", layout="centered")

# Display Logo
if os.path.exists("DentaQuickEgypt.png"):
    logo = Image.open("DentaQuickEgypt.png")
    st.image(logo)

# Custom Title and Subtitle
st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Order & Equipment Merger</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload Excel files and choose a mode to either merge old/new orders or combine equipment lists.</p>", unsafe_allow_html=True)
st.divider()

# ------------------ Mode Selection ------------------ #
mode = st.radio("Choose a Function:", ["1️⃣ Old + New Order Merger", "2️⃣ Equipment Summary Merger"])

# ------------------ Header Row Input ------------------ #
header_row = st.number_input("Enter the row number where the headers are (0-indexed)", min_value=0, max_value=50, step=1)

st.divider()
st.subheader("📤 Upload Excel Files")

if mode == "1️⃣ Old + New Order Merger":
    old_file = st.file_uploader("Upload OLD Orders File", type=["xlsx"], key="old_file")
    new_file = st.file_uploader("Upload NEW Orders File (optional)", type=["xlsx"], key="new_file")
else:
    files_uploaded = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True, key="equip_files")

# ------------------ Helper Functions ------------------ #
HEADERS = {
    'serial': ['serial', 'رقم', 'الرقم'],
    'equipment_name': ['equipment name', 'اسم الجهاز', 'اسم المعدة', 'item', 'product', 'الصنف'],
    'number': ['number', 'العدد', 'qty', 'quantity', 'الكمية'],
    'notes': ['notes', 'ملاحظات'],
    'price': ['price', 'السعر']
}

def normalize_column_name(col):
    col = str(col).strip().lower()
    for key, values in HEADERS.items():
        if any(val.lower() in col for val in values):
            return key
    return col

def normalize_columns(df):
    df.columns = [normalize_column_name(c) for c in df.columns]
    return df

def process_multisheet_excel(uploaded_file, header_row):
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    cleaned_sheets = {}
    price_df = pd.DataFrame()

    for name, raw_df in all_sheets.items():
        try:
            df = pd.read_excel(uploaded_file, sheet_name=name, header=header_row)
            df = normalize_columns(df)

            if 'equipment_name' in df.columns and 'number' in df.columns:
                temp = df[['equipment_name', 'number']].dropna()
                temp['equipment_name'] = temp['equipment_name'].astype(str).str.strip().str.lower()
                temp = temp.groupby('equipment_name', as_index=False).sum()
                temp.columns = ['equipment_name', name]
                cleaned_sheets[name] = temp

                if 'price' in df.columns:
                    price_temp = df[['equipment_name', 'price']].dropna()
                    price_temp['equipment_name'] = price_temp['equipment_name'].astype(str).str.strip().str.lower()
                    price_df = pd.concat([price_df, price_temp], ignore_index=True)
        except Exception as e:
            st.warning(f"⚠️ Skipped sheet '{name}' due to error: {e}")

    if not price_df.empty:
        price_df = price_df.drop_duplicates(subset='equipment_name')
        price_df = price_df.groupby('equipment_name', as_index=False).first()

    return cleaned_sheets, price_df

def merge_sheets(sheet_dict):
    if not sheet_dict:
        return pd.DataFrame()
    return reduce(lambda left, right: pd.merge(left, right, on='equipment_name', how='outer'), sheet_dict.values()).fillna(0)

def to_excel(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def process_equipment_summary(files_uploaded, header_row):
    all_data = []

    for uploaded in files_uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet, header=header_row)
                df = normalize_columns(df)

                for col in ['equipment_name', 'number', 'notes']:
                    if col not in df.columns:
                        df[col] = None

                df = df[['equipment_name', 'number', 'notes']].copy()
                df['number'] = pd.to_numeric(df['number'], errors='coerce').fillna(0)
                all_data.append(df)
        except Exception as e:
            st.warning(f"⚠️ Skipped file '{uploaded.name}' due to error: {e}")

    if not all_data:
        return pd.DataFrame()

    combined = pd.concat(all_data, ignore_index=True)
    grouped = combined.groupby(['equipment_name', 'notes'], dropna=False)['number'].sum().reset_index()
    return grouped

# ------------------ Execution Logic ------------------ #
if mode == "1️⃣ Old + New Order Merger" and old_file:
    try:
        old_sheets, old_prices = process_multisheet_excel(old_file, header_row)
        old_merged = merge_sheets(old_sheets)

        if not old_merged.empty:
            old_cols = [col for col in old_merged.columns if col != 'equipment_name']
            old_merged['Old Quantity'] = old_merged[old_cols].sum(axis=1)
            old_merged['equipment_name'] = old_merged['equipment_name'].str.title()
            df_old_summary = old_merged[['equipment_name', 'Old Quantity']].copy()

            st.subheader("📋 OLD Orders Summary")
            st.dataframe(df_old_summary, use_container_width=True)

            price_df = old_prices.copy()

            if new_file:
                new_sheets, new_prices = process_multisheet_excel(new_file, header_row)
                new_merged = merge_sheets(new_sheets)

                if not new_merged.empty:
                    new_cols = [col for col in new_merged.columns if col != 'equipment_name']
                    new_merged['New Quantity'] = new_merged[new_cols].sum(axis=1)
                    new_merged['equipment_name'] = new_merged['equipment_name'].str.title()
                    df_new_summary = new_merged[['equipment_name', 'New Quantity']].copy()

                    combined = pd.merge(df_old_summary, df_new_summary, on='equipment_name', how='outer').fillna(0)
                    combined['Total Quantity'] = combined['Old Quantity'] + combined['New Quantity']

                    if price_df.empty and not new_prices.empty:
                        price_df = new_prices.copy()

                    if not price_df.empty:
                        price_df['equipment_name'] = price_df['equipment_name'].str.title()
                        combined = pd.merge(combined, price_df, on='equipment_name', how='left')

                    combined = combined[['equipment_name', 'price', 'Old Quantity', 'New Quantity', 'Total Quantity']]
                    combined = combined.sort_values(by='Total Quantity', ascending=False)

                    st.subheader("📋 Combined Summary with Prices")
                    st.dataframe(combined, use_container_width=True)

                    excel_data = to_excel(combined, sheet_name="Merged Summary")
                    st.download_button("⬇️ Download Merged Excel", data=excel_data, file_name="Merged_Orders_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("⚠️ New file has no valid sheets.")
            else:
                if not old_prices.empty:
                    old_prices['equipment_name'] = old_prices['equipment_name'].str.title()
                    df_old_summary = pd.merge(df_old_summary, old_prices, on='equipment_name', how='left')
                    df_old_summary = df_old_summary[['equipment_name', 'price', 'Old Quantity']]

                st.subheader("📋 OLD Orders with Prices")
                st.dataframe(df_old_summary, use_container_width=True)

                excel_data = to_excel(df_old_summary)
                st.download_button("⬇️ Download Excel", data=excel_data, file_name="Old_Orders_With_Prices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ Old file has no valid sheets.")
    except Exception as e:
        st.error("❌ Error while processing the files.")
        st.exception(e)

elif mode == "2️⃣ Equipment Summary Merger" and files_uploaded:
    try:
        result_df = process_equipment_summary(files_uploaded, header_row)
        if not result_df.empty:
            result_df = result_df.sort_values(by='number', ascending=False)
            st.subheader("📊 Equipment Summary")
            st.dataframe(result_df, use_container_width=True)

            excel_data = to_excel(result_df, sheet_name="Equipment Summary")
            st.download_button("⬇️ Download Equipment Summary", data=excel_data, file_name="Equipment_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ No valid data found.")
    except Exception as e:
        st.error("❌ Error while processing equipment summary.")
        st.exception(e)
