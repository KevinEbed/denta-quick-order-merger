import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image
import os

# ------------------------------ Page Setup ------------------------------ #
st.set_page_config(page_title="Denta Quick Merger", page_icon="🦷", layout="centered")

if os.path.exists("DentaQuickEgypt.png"):
    logo = Image.open("DentaQuickEgypt.png")
    st.image(logo)

st.markdown("<h2 style='text-align: center; color: #3B7A57;'>🦷 Denta Quick – Branch Order Merger</h2>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center;'>Upload OLD and (optional) NEW order files. Headers can start on <strong>row 1</strong> or <strong>row 15</strong>. All valid sheets will be combined and summarized.</p>", unsafe_allow_html=True)
st.divider()

# ------------------------------ Upload Section ------------------------------ #
st.subheader("🗂️ Step 1: Upload Excel Files")
old_file = st.file_uploader("Upload OLD Orders File (multi-sheet Excel)", type=["xlsx"])
new_file = st.file_uploader("Upload NEW Orders File (optional)", type=["xlsx"])

# ------------------------------ Utility Functions ------------------------------ #
def detect_and_clean_sheet(df):
    for header_row in [0, 14]:  # Try both header locations: row 1 (index 0), row 15 (index 14)
        temp_df = pd.read_excel(df, header=header_row)
        if 'الصنف' in temp_df.columns and 'الكمية' in temp_df.columns:
            return pd.read_excel(df, sheet_name=None, header=header_row)
    return {}

def normalize_columns(df):
    rename_map = {
        'الصنف': 'product',
        'product': 'product',
        'الكمية': 'quantity',
        'QYT': 'quantity',
        'vendor': 'vendor',
        'السعر': 'price'
    }
    df = df.rename(columns={col: rename_map.get(col, col) for col in df.columns})
    if 'product' in df.columns and 'quantity' in df.columns:
        df['product'] = df['product'].astype(str).str.strip().str.lower()
        return df[['product', 'quantity'] + ([col for col in ['price'] if col in df.columns])]
    return pd.DataFrame()

def process_file(uploaded_file):
    all_sheets_combined = pd.DataFrame()
    price_df = pd.DataFrame()

    excel_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None)

    for sheet_name, raw_df in excel_sheets.items():
        for header_row in [0, 14]:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row)
                df = normalize_columns(df)
                if not df.empty:
                    grouped = df.groupby('product', as_index=False).sum(numeric_only=True)
                    all_sheets_combined = pd.concat([all_sheets_combined, grouped], ignore_index=True)
                    if 'price' in df.columns:
                        price_df = pd.concat([price_df, df[['product', 'price']].dropna()], ignore_index=True)
                    break  # Header found, stop searching
            except Exception:
                continue

    # Group again after full concatenation
    final_df = all_sheets_combined.groupby('product', as_index=False).sum()
    if not price_df.empty:
        price_df = price_df.drop_duplicates(subset='product').groupby('product', as_index=False).first()
    return final_df, price_df

def merge_quantities(df_old, df_new):
    combined = pd.merge(df_old, df_new, on='product', how='outer', suffixes=('_old', '_new')).fillna(0)
    combined['Total Quantity'] = combined['quantity_old'] + combined['quantity_new']
    return combined

def to_excel(dfs: dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=name)
    return output.getvalue()

# ------------------------------ Main Processing ------------------------------ #
if old_file:
    try:
        df_old, prices_old = process_file(old_file)
        df_old['product'] = df_old['product'].str.title()
        df_old.rename(columns={'quantity': 'Old Quantity'}, inplace=True)

        st.subheader("📋 Summary of OLD Orders")
        st.dataframe(df_old, use_container_width=True)

        if new_file:
            df_new, prices_new = process_file(new_file)
            df_new['product'] = df_new['product'].str.title()
            df_new.rename(columns={'quantity': 'New Quantity'}, inplace=True)

            st.subheader("📋 Summary of NEW Orders")
            st.dataframe(df_new, use_container_width=True)

            combined = merge_quantities(df_old, df_new)

            # Add price (prefer old price, then new)
            prices_old['product'] = prices_old['product'].str.title()
            prices_new['product'] = prices_new['product'].str.title()
            price_df = pd.concat([prices_old, prices_new], ignore_index=True)
            price_df = price_df.drop_duplicates(subset='product')

            combined = pd.merge(combined, price_df, on='product', how='left')
            combined = combined[['product', 'price', 'Old Quantity', 'New Quantity', 'Total Quantity']]
            combined = combined.sort_values(by='Total Quantity', ascending=False)

            st.subheader("📋 Summary of OLD + NEW Orders Combined")
            st.dataframe(combined, use_container_width=True)

            # Download
            excel_bytes = to_excel({
                "Old Orders": df_old,
                "New Orders": df_new,
                "Combined Summary": combined
            })
            st.download_button(
                label="⬇️ Download Excel Report",
                data=excel_bytes,
                file_name="DentaQuick_Merged_Orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Only Old
            if not prices_old.empty:
                prices_old['product'] = prices_old['product'].str.title()
                df_old = pd.merge(df_old, prices_old, on='product', how='left')
                df_old = df_old[['product', 'price', 'Old Quantity']]

            st.subheader("📋 OLD Orders + Prices")
            st.dataframe(df_old, use_container_width=True)

            excel_bytes = to_excel({"Old Orders": df_old})
            st.download_button(
                label="⬇️ Download Excel Report",
                data=excel_bytes,
                file_name="DentaQuick_Old_Orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error("❌ Error processing the uploaded files.")
        st.exception(e)
