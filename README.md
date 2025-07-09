# 📦 Order Merger App

This is a simple Streamlit web app that lets users upload two Excel or CSV order files and merges them based on item names to calculate total quantities.

---

## 🔧 Features

- Upload two order files from different branches
- Merges by item name (Arabic column: `الصنف`)
- Sums quantities (Arabic column: `الكمية`)
- Exports result as downloadable Excel file
- Clean and user-friendly interface

---

## 📂 File Requirements

Each Excel or CSV file must have:
- A column called **`الصنف`** (Item)
- A column called **`الكمية`** (Quantity)

---

## ▶️ How to Run Locally

1. Clone the repo:

```bash
git clone https://github.com/yourusername/order-merger-app.git
cd order-merger-app
