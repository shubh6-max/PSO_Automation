# 📊 Attendance Compliance Dashboard

A dynamic Streamlit dashboard to analyze employee attendance compliance across service lines using uploaded Excel files. It calculates working days (Monday to Thursday) based on selected month and year, evaluates attendance, and provides pivot summaries and compliance metrics with downloadable Excel reports.

---

## 🔧 Features

- Upload Excel files with attendance data.
- Select month and year to compute working days.
- Automatically identify first and last working days.
- Filter attendance columns dynamically.
- Create service-line wise summaries with compliance %, charts, and raw data.
- Export final report with multiple formatted sheets.

---

## 🖼️ UI Preview

![App Preview](https://media.licdn.com/dms/image/v2/D4D0BAQHV_-WdH8NGCw/company-logo_200_200/B4DZdoK0PjHkAM-/0/1749799355979/themathcompany_logo?e=1756944000&v=beta&t=3N3rldQGIH1FsqUhgbyI2qnELA8Txh4ZJvHFtbNeRhQ)

---

## 🚀 How to Run

1. **Install requirements:**

   ```bash
   pip install streamlit pandas openpyxl xlsxwriter
