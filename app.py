import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from PIL import Image
import os

# --- Streamlit Config ---
st.set_page_config(page_title="Denta Quick Merger", page_icon="🦷", layout="centered")

# --- Load Logo ---
if os.path.exists("DentaQuickEgypt.png"):
    logo = Image.open("DentaQuickEgypt.png")
    st.image(logo)

st.markdown("<h2 style='text-align: center;'>📊 Denta Quick File Merger</h2>", unsafe_allow_html=True)
st.markdown("---")

# --- Headers Map for Equipment Summary Merger ---
HEADERS = {
    'serial': ['serial', 'رقم', 'الرقم'],
    'equipment_name': ['equipment name', 'اسم الجهاز', 'اسم المعدة', 'item', 'product'],
    'number': ['number', 'العدد', 'qty', 'quantity'],
    'notes': ['notes', 'ملاحظات']
}

def normalize_column_name(col):
    col = str(col).strip().lower()
    for key, values in HEADERS.items():
        if any(val.lower() in col for val in values):
            return key
    return col

# --- Equipment Summary Merger Logic ---
def process_equipment_file(file, header_row):
    xl = pd.ExcelFile(file)
    all_data = []

    for sheet in xl.sheet_names:
        try:
            df = xl.parse(sheet, header=header_row)
            df.columns = [normalize_column_name(c) for c in df.columns]
            needed = ['equipment_name', 'number', 'notes']
            for col in needed:
                if col not in df.columns:
                    df[col] = None
            df = df[needed]
            df['number'] = pd.to_numeric(df['number'], errors='coerce').fillna(0)
            all_data.append(df)
        except Exception as e:
            st.warning(f"⚠️ Skipped sheet '{sheet}': {e}")

    if not all_data:
        return pd.DataFrame()

    combined = pd.concat(all_data, ignore_index=True)
    grouped = combined.groupby(['equipment_name', 'notes'], dropna=False)['number'].sum().reset_index()
    return grouped

# --- Old + New Order Merger Logic ---
def merge_old_and_new(old_file, new_file):
    try:
        old_df = pd.read_excel(old_file, sheet_name=None)
        new_df = pd.read_excel(new_file, sheet_name=None)
    except Exception as e:
        st.error(f"❌ Failed to read Excel files: {e}")
        return None, None

    old_orders_with_prices = pd.DataFrame()
    old_orders_summary = pd.DataFrame()

    # Look for sheets containing specific keywords
    for sheet_name, df in old_df.items():
        lower_name = sheet_name.lower()
        if "summary" in lower_name:
            old_orders_summary = df.copy()
        elif "price" in lower_name or "with" in lower_name:
            old_orders_with_prices = df.copy()

    # If failed to detect specific sheets, fallback
    if old_orders_summary.empty and len(old_df) >= 1:
        old_orders_summary = list(old_df.values())[0]
    if old_orders_with_prices.empty and len(old_df) >= 2:
        old_orders_with_prices = list(old_df.values())[1]

    # Combine new file sheets
    new_orders = pd.concat([df for df in new_df.values()], ignore_index=True)

    return old_orders_summary, old_orders_with_prices, new_orders

# --- User Interface ---
mode = st.radio("Choose Merger Mode", ["Equipment Summary Merger", "Old + New Order Merger"], horizontal=True)

if mode == "Equipment Summary Merger":
    st.subheader("🔧 Equipment Summary Merger")
    header_row_input = st.number_input("Enter the row number where the headers are (0-indexed):", min_value=0, step=1)
    uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file and st.button("▶️ Process"):
        st.success("✅ Processing file...")
        result_df = process_equipment_file(uploaded_file, header_row_input)

        if not result_df.empty:
            st.dataframe(result_df)
            output_file = "equipment_summary_merged.xlsx"
            result_df.to_excel(output_file, index=False)
            with open(output_file, "rb") as f:
                st.download_button("📥 Download Merged Summary", f, file_name=output_file)
        else:
            st.warning("⚠️ No valid data found.")

elif mode == "Old + New Order Merger":
    st.subheader("🧾 Old + New Order Merger")
    old_file = st.file_uploader("📤 Upload Old Orders Excel File", type=["xlsx", "xls"], key="old")
    new_file = st.file_uploader("📤 Upload New Orders Excel File", type=["xlsx", "xls"], key="new")

    if old_file and new_file and st.button("▶️ Merge"):
        st.success("✅ Merging files...")
        old_summary, old_with_prices, new_orders = merge_old_and_new(old_file, new_file)

        if old_summary is not None:
            st.markdown("### 📋 OLD Orders Summary")
            st.dataframe(old_summary)

            st.markdown("### 💰 OLD Orders with Prices")
            st.dataframe(old_with_prices)

            st.markdown("### 🆕 NEW Orders")
            st.dataframe(new_orders)

            merged_df = pd.concat([old_summary, new_orders], ignore_index=True)
            output_file = "old_new_orders_merged.xlsx"
            merged_df.to_excel(output_file, index=False)
            with open(output_file, "rb") as f:
                st.download_button("📥 Download Merged Orders", f, file_name=output_file)
        else:
            st.warning("⚠️ Could not process the provided files.")
