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
old_file = st.file_uploader("Upload OLD Orders File (multi-sheet Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW Orders File (optional)", type=["xlsx"])

# ------------------- Processing Functions ------------------- #

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

# ------------------------- Main App Logic ------------------------- #

if old_file:
    try:
        old_sheets, old_prices = process_multisheet_excel(old_file)
        old_merged = merge_sheets(old_sheets)

        if not old_merged.empty:
            old_cols = [col for col in old_merged.columns if col != 'الصنف']
            old_merged['Old Quantity'] = old_merged[old_cols].sum(axis=1)
            old_merged['الصنف'] = old_merged['الصنف'].str.title()
            df_old_summary = old_merged[['الصنف', 'Old Quantity']].copy()

            st.subheader("📋 Step 2: Summary of OLD Orders")
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

                    st.subheader("📋 Step 3: Summary of NEW Orders")
                    st.dataframe(df_new_summary, use_container_width=True)

                    combined = pd.merge(df_old_summary, df_new_summary, on='الصنف', how='outer').fillna(0)
                    combined['Total Quantity'] = combined['Old Quantity'] + combined['New Quantity']

                    if price_df.empty and not new_prices.empty:
                        price_df = new_prices.copy()

                    if not price_df.empty:
                        price_df['الصنف'] = price_df['الصنف'].str.title()
                        combined = pd.merge(combined, price_df, on='الصنف', how='left')

                    combined = combined[['الصنف', 'السعر', 'Old Quantity', 'New Quantity', 'Total Quantity']]
                    combined = combined.sort_values(by='Total Quantity', ascending=False)

                    st.subheader("📋 Step 4: Summary of OLD + NEW Orders with Prices")
                    st.dataframe(combined, use_container_width=True)

                    excel_data = to_excel(combined)
                    st.subheader("📥 Step 5: Download Merged Excel Report")
                    st.download_button("⬇️ Download Excel File", data=excel_data, file_name="Merged_Orders_Summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.warning("⚠️ The NEW file didn’t contain valid data sheets.")
            else:
                if not old_prices.empty:
                    old_prices['الصنف'] = old_prices['الصنف'].str.title()
                    df_old_summary = pd.merge(df_old_summary, old_prices, on='الصنف', how='left')
                    df_old_summary = df_old_summary[['الصنف', 'السعر', 'Old Quantity']]
                st.subheader("📋 Step 3: OLD Orders + Prices")
                st.dataframe(df_old_summary, use_container_width=True)

                excel_data = to_excel(df_old_summary)
                st.subheader("📥 Step 4: Download Excel File")
                st.download_button("⬇️ Download Excel File", data=excel_data, file_name="Old_Orders_With_Prices.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ The OLD file didn’t contain valid data sheets.")
    except Exception as e:
        st.error("❌ Error while processing the uploaded files.")
        st.exception(e)
