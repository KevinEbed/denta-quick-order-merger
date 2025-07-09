# 🦷 Denta Quick – Branch Order Merger

A simple, user-friendly Streamlit app for merging and analyzing Excel orders from multiple dental branches. Designed for **internal use by Denta Quick** and compatible with Excel files exported from your branch order system.

---

## 🚀 Features

- ✅ Upload two Excel files: **Old Orders** and **(optional) New Orders**
- 📊 Each file may contain **multiple sheets** (one per branch)
- 🔁 Automatically **merges items by name**
- 🧾 Displays:
  - Old Order Quantities
  - New Order Quantities
  - **Total Quantity**
  - **Item Price (السعر)** (if available in any sheet)
- 📥 Download final report as Excel
- 🖼️ Custom branded with Denta Quick logo

---

## 📁 File Requirements

- Format: `.xlsx` Excel file
- Each sheet represents a branch order
- Must include these columns (starting from **row 15**):

| Arabic Column | Meaning       | Required |
|---------------|----------------|----------|
| `الصنف`        | Item name      | ✅ Yes    |
| `الكمية`       | Quantity        | ✅ Yes    |
| `السعر`        | Unit price      | Optional |

---

## 🧑‍💻 How to Run Locally

1. Clone the repo:
   ```bash
   git clone https://github.com/KevinEbed/denta-quick-order-merger.git
   cd denta-quick-order-merger
