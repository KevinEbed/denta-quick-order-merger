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
st.divider()
st.subheader("📄 Upload Excel Files")

if mode == "1️⃣ Old + New Order Merger":
    old_file = st.file_uploader("Upload OLD Orders File", type=["xlsx"], key="old_file")
    new_file = st.file_uploader("Upload NEW Orders File (optional)", type=["xlsx"], key="new_file")
else:
    files_uploaded = st.file_uploader("Upload Excel Files", type=["xlsx"], accept_multiple_files=True, key="equip_files")

# ------------------ Common Functions ------------------ #
def normalize_columns(df):
    col_map = {
        'product': 'الصنف',
        'qyt': 'الكمية',
        'الكمية': 'الكمية',
        'الصنف': 'الصنف',
        'vendor': 'vendor',
        'السعر': 'السعر',
        'price': 'السعر'
    }
    df.columns = [col_map.get(str(c).strip().lower(), str(c).strip().lower()) for c in df.columns]
    return df

def find_header_row(df):
    for i in range(20):
        row = df.iloc[i].astype(str).str.lower()
        if any(val in row.values for val in ['الصنف', 'product']):
            return i
    return None

def process_multisheet_excel(uploaded_file):
    all_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)
    cleaned_sheets = {}
    price_df = pd.DataFrame()

    for name, raw_df in all_sheets.items():
        header_row = find_header_row(raw_df)
        if header_row is not None:
            df = pd.read_excel(uploaded_file, sheet_name=name, header=header_row)
            df = normalize_columns(df)

            if 'الصنف' in df.columns and 'الكمية' in df.columns:
                temp = df[['الصنف', 'الكمية']].dropna()
                temp['الصنف'] = temp['الصنف'].astype(str).str.strip().str.lower()
                temp = temp.groupby('الصنف', as_index=False).sum()
                temp.columns = ['الصنف', name]
                cleaned_sheets[name] = temp

                if 'السعر' in df.columns:
                    price_temp = df[['الصنف', 'السعر']].dropna()
                    price_temp['الصنف'] = price_temp['الصنف'].astype(str).str.strip().str.lower()
                    price_df = pd.concat([price_df, price_temp], ignore_index=True)

    if not price_df.empty:
        price_df = price_df.drop_duplicates(subset='الصنف')
        price_df = price_df.groupby('الصنف', as_index=False).first()

    return cleaned_sheets, price_df

def merge_sheets(sheet_dict):
    if not sheet_dict:
        return pd.DataFrame()
    return reduce(lambda left, right: pd.merge(left, right, on='الصنف', how='outer'), sheet_dict.values()).fillna(0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Merged Orders")
    return output.getvalue()

# ------------------ Updated Equipment Merger Functions ------------------ #
HEADERS = {
    'serial': ['serial', 'رقم', 'الرقم'],
    'equipment_name': ['equipment name', 'اسم الجهاز', 'اسم المعدة', 'item', 'product', 'الصنف'],
    'number': ['number', 'العدد', 'qty', 'qyt', 'quantity', 'الكمية'],
    'notes': ['notes', 'ملاحظات']
}

def normalize_column_name(col):
    col = str(col).strip().lower()
    for key, values in HEADERS.items():
        if any(val in col for val in values):
            return key
    return col

def find_equipment_header_row(df):
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        normalized = [normalize_column_name(c) for c in row]
        if {'equipment_name', 'number'}.issubset(set(normalized)):
            return i
    return None

def process_equipment_summary(files_uploaded):
    all_data = []
    for uploaded in files_uploaded:
        try:
            xl = pd.ExcelFile(uploaded)
            for sheet in xl.sheet_names:
                raw_df = xl.parse(sheet, header=None)
                header_row = find_equipment_header_row(raw_df)

                if header_row is not None:
                    df = xl.parse(sheet, header=header_row)
                    df.columns = [normalize_column_name(c) for c in df.columns]

                    for col in ['equipment_name', 'number', 'notes']:
                        if col not in df.columns:
                            df[col] = None

                    df = df[['equipment_name', 'number', 'notes']]
                    df['number'] = pd.to_numeric(df['number'], errors='coerce').fillna(0)
                    all_data.append(df)
                else:
                    st.warning(f"❌ Skipped sheet '{sheet}' in file '{uploaded.name}': No valid header found.")
        except Exception as e:
            st.warning(f"⚠️ Skipped file '{uploaded.name}' due to error: {e}")

    if not all_data:
        return pd.DataFrame()

    combined = pd.concat(all_data, ignore_index=True)
    grouped = combined.groupby(['equipment_name', 'notes'], dropna=False)['number'].sum().reset_index()
    grouped['notes'] = grouped['notes'].fillna('').astype(str).str.strip()
    return grouped

# ------------------------- Main Logic ------------------------- #
if mode == "1️⃣ Old + New Order Merger" and old_file:
    try:
        old_sheets, old_prices = process_multisheet_excel(old_file)
        old_merged = merge_sheets(old_sheets)

        if not old_merged.empty:
            old_cols = [col for col in old_merged.columns if col != 'الصنف']
            old_merged['Old Quantity'] = old_merged[old_cols].sum(axis=1)
            old_merged['الصنف'] = old_merged['الصنف'].str.title()
            df_old_summary = old_merged[['الصنف', 'Old Quantity']].copy()

            st.subheader("📋 OLD Orders Summary")
            st.dataframe(df_old_summary, use_container_width=True)

            price_df = old_prices.copy()

            if new_file:
                new_sheets, new_prices = process_multisheet_excel(new_file)
                new_merged = merge_sheets(new_sheets)

                if not new_merged.empty:
                    new_cols = [col for col in new_merged.columns if col != 'الصنف']
                    new_merged['New Quantity'] = new_merged[new_cols].sum(axis=1)
                    new_merged['الصنف'] = new_merged['الصنف'].str.title()
                    df_new_summary = new_merged[['الصنف', 'New Quantity']].copy()

                    combined = pd.merge(df_old_summary, df_new_summary, on='الصنف', how='outer').fillna(0)
                    combined['Total Quantity'] = combined['Old Quantity'] + combined['New Quantity']

                    if price_df.empty and not new_prices.empty:
                        price_df = new_prices.copy()

                    if not price_df.empty:
                        price_df['الصنف'] = price_df['الصنف'].str.title()
                        combined = pd.merge(combined, price_df, on='الصنف', how='left')

                    combined = combined[['الصنف', 'السعر', 'Old Quantity', 'New Quantity', 'Total Quantity']]
                    combined = combined.sort_values(by='Total Quantity', ascending=False)

                    st.subheader("📋 Combined Summary with Prices")
                    st.dataframe(combined, use_container_width=True)

                    excel_data = to_excel(combined)
                    st.download_button("⬇️ Download Merged Excel", data=excel_data, file_name="Merged_Orders_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("⚠️ The NEW file didn’t contain valid data sheets.")
            else:
                if not old_prices.empty:
                    old_prices['الصنف'] = old_prices['الصنف'].str.title()
                    df_old_summary = pd.merge(df_old_summary, old_prices, on='الصنف', how='left')
                    df_old_summary = df_old_summary[['الصنف', 'السعر', 'Old Quantity']]
                st.subheader("📋 OLD Orders with Prices")
                st.dataframe(df_old_summary, use_container_width=True)

                excel_data = to_excel(df_old_summary)
                st.download_button("⬇️ Download Excel File", data=excel_data, file_name="Old_Orders_With_Prices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ The OLD file didn’t contain valid data sheets.")
    except Exception as e:
        st.error("❌ Error while processing the uploaded files.")
        st.exception(e)

elif mode == "2️⃣ Equipment Summary Merger" and files_uploaded:
    try:
        result_df = process_equipment_summary(files_uploaded)
        if not result_df.empty:
            result_df = result_df.sort_values(by='number', ascending=False)
            st.subheader("📊 Equipment Summary")
            st.dataframe(result_df, use_container_width=True)

            excel_data = to_excel(result_df)
            st.download_button("⬇️ Download Equipment Summary", data=excel_data, file_name="Equipment_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ No valid data found.")
    except Exception as e:
        st.error("❌ Error while processing equipment summary.")
        st.exception(e)
