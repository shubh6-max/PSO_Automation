import pandas as pd
import streamlit as st
from io import BytesIO
import calendar
from datetime import datetime

st.set_page_config(layout="wide")

st.markdown("""
    <style>
    /* Background and layout */
    .stApp {
        background-color: #ebe3d6;
    }

    /* Push content to the right */
    .block-container {
        padding-left: 180px;  /* Adjust based on logo width */
    }

    /* Fixed logo in top-left corner */
    .fixed-logo {
        position: fixed;
        top: 20px;
        left: 20px;
        width: 140px;
        z-index: 1000;
        background-color: white;
        padding: 5px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.2);
    }
    </style>

    <img src="https://upload.wikimedia.org/wikipedia/commons/8/88/MathCo_Logo.png" class="fixed-logo">
""", unsafe_allow_html=True)




# # Add image in top-left corner using HTML
# st.markdown("""
#     <img src="https://upload.wikimedia.org/wikipedia/commons/8/88/MathCo_Logo.png" class="fixed-logo">
# """, unsafe_allow_html=True)

st.title("📊 Attendance Compliance Dashboard")

uploaded_file = st.file_uploader("**Upload the Excel File**", type=["xlsx"])

# Function to count working days (Monday to Thursday)
def count_working_days(month_name: str, year: int) -> int:
    try:
        month = list(calendar.month_name).index(month_name)
    except ValueError:
        return 0

    if month == 0:
        return 0

    num_days = calendar.monthrange(year, month)[1]
    working_days = sum(
        1 for day in range(1, num_days + 1)
        if datetime(year, month, day).weekday() in [0, 1, 2, 3]  # Mon–Thu
    )

    return working_days

def get_working_day_range(month_name: str, year: int) -> tuple:
    month_num = list(calendar.month_name).index(month_name.capitalize())
    num_days = calendar.monthrange(year, month_num)[1]
    working_days = [
        datetime(year, month_num, d).strftime("%d-%b-%Y").upper()
        for d in range(1, num_days + 1)
        if datetime(year, month_num, d).weekday() in [0, 1, 2, 3]
    ]
    return (working_days[0], working_days[-1]) if working_days else (None, None)

# Handle unnamed or duplicate columns
def sanitize_columns(columns):
    seen = {}
    new_cols = []
    for col in columns:
        if pd.isna(col):
            col = "Unnamed"
        col = str(col).strip()
        count = seen.get(col, 0)
        new_col = col if count == 0 else f"{col}_{count}"
        new_cols.append(new_col)
        seen[col] = count + 1
    return new_cols

# Detect attendance date columns
def is_parseable_date(col):
    try:
        pd.to_datetime(str(col), errors='raise')
        return True
    except:
        return False

# Two-column layout
col1, col2 = st.columns(2)

with col1:
    months = ["Select Month"] + list(calendar.month_name)[1:]
    selected_month = st.selectbox("**Select Month**", months)

with col2:
    years = ["Select Year"] + list(range(2025, datetime.now().year + 6))
    selected_year = st.selectbox("**Select Year**", years, index=0)

# Only show result and file name input when valid selections are made
start_col, end_col = None, None
if selected_month != "Select Month" and selected_year != "Select Year":
    selected_year = int(selected_year)
    result = count_working_days(selected_month, selected_year)
    st.success(f"✅ Working days (Mon–Thu) in {selected_month} {selected_year}: **{result}** for Compliance.")
    
    start_col, end_col = get_working_day_range(selected_month, selected_year)

    output_name = st.text_input("**Enter output Excel file name (without extension)**", 
                                value=f"{selected_month}-{selected_year} Attendance Compliance Report")
else:
    st.info("👈 Please select both Month and Year to see the result.")


# start_col,end_col=get_working_day_range(selected_month, selected_year)

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    selected_sheet = st.selectbox("**Select sheet to process**", sheet_names, index=1 if len(sheet_names) > 1 else 0)

    df = pd.read_excel(xls, sheet_name=selected_sheet, engine="openpyxl")

    # Check if start_col and end_col exist in the selected sheet
    if start_col in df.columns and end_col in df.columns:
        start_idx = df.columns.get_loc(start_col)
        end_idx = df.columns.get_loc(end_col)

        selected_columns = df.iloc[:, start_idx:end_idx+1].columns # type: ignore

        final_column_list = list(df.columns[:12]) + list(selected_columns)
        df = df[final_column_list]
    else:
        st.error(f"❌ The date columns `{start_col}` and/or `{end_col}` were not found in the selected sheet: `{selected_sheet}`.")
        st.stop()


    # Slice the columns (inclusive)
    selected_columns = df.iloc[:, start_idx:end_idx+1].columns # type: ignore

    final_column_list=[]
    for i in df.columns[:12]:
        final_column_list.append(i)

    for i in selected_columns:
        final_column_list.append(i)

    df=df[final_column_list]

    df=df[df["Status"]=="Active"]

    try:
        df["Accounts"].replace("Mathapps","MathApps",inplace=True)
    except:
        print("Select sheet")


    df.columns = sanitize_columns(df.columns)

    required_cols = ['EmpID', 'Accounts', 'Service Line']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Required columns missing: {', '.join(missing_cols)}")
        st.stop()


    attendance_day_cols = [col for col in df.columns if is_parseable_date(col)]

    df['Adjusted Attendance'] = df[attendance_day_cols].apply(
        lambda row: sum(
            str(val).strip().upper() in ['1', 'L', 'WFH','Miss Punch Out'] or pd.to_numeric(val, errors='coerce') == 1
            for val in row
        ),
        axis=1
    )

    # Step 2: Count "L" or "WFH" per row in those columns
    df['L_WFH_Count'] = df[attendance_day_cols].apply(
        lambda row: sum(str(val).strip().upper() in ['L', 'WFH'] for val in row),
        axis=1
    )

    df["Total Attendance (Without L/WFH)"]=df["Adjusted Attendance"]-df["L_WFH_Count"]
    # === Excel Writer with formatting and chart ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)

        workbook = writer.book

        # ---------------------------------------
        # Sheet 1: Summary (Service Line-wise)
        # ---------------------------------------
        summary_ws = workbook.add_worksheet("Summary") # type: ignore
        writer.sheets["Summary"] = summary_ws

        header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9E1F2', 'border': 1}) # type: ignore
        cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}) # type: ignore
        title_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#4F81BD', 'align': 'left', 'border': 1}) # type: ignore

        start_row, left_col, right_col = 0, 0, 6

        for service_line in df['Service Line'].dropna().unique():
            filtered_df = df[df['Service Line'] == service_line].copy()
            pivot = pd.pivot_table(filtered_df, index='Accounts', values='EmpID', aggfunc='count', margins=True, margins_name='Grand Total').reset_index()

            filtered_df['Compliant'] = filtered_df['Adjusted Attendance'] >= result # type: ignore
            filtered_df['Non_Compliant'] = filtered_df['Adjusted Attendance'] < result # type: ignore

            summary = filtered_df.groupby('Accounts').agg(
                Count_of_Emp=('EmpID', 'count'),
                Compliant_Count=('Compliant', 'sum'),
                Non_Compliant_Count=('Non_Compliant', 'sum')
            ).reset_index()

            summary['Compliance %'] = (summary['Non_Compliant_Count'] / summary['Count_of_Emp']).apply(lambda x: f"{x:.0%}")
            summary['% Non-compliance'] = (summary['Compliant_Count'] / summary['Count_of_Emp']).apply(lambda x: f"{x:.0%}")

            grand_total = pd.DataFrame({
                'Accounts': ['Grand Total'],
                'Count_of_Emp': [summary['Count_of_Emp'].sum()],
                'Compliant_Count': [summary['Compliant_Count'].sum()],
                'Non_Compliant_Count': [summary['Non_Compliant_Count'].sum()],
                '% Non-compliance': [f"{summary['Non_Compliant_Count'].sum() / summary['Count_of_Emp'].sum():.0%}"],
                'Compliance %': [f"{summary['Compliant_Count'].sum() / summary['Count_of_Emp'].sum():.0%}"]
            })

            final_summary = pd.concat([summary, grand_total], ignore_index=True)

            # Title
            summary_ws.write(start_row, left_col, f"Service Line: {service_line} - Pivot Table", title_format)
            summary_ws.write(start_row, right_col, f"Service Line: {service_line} - Compliance Summary", title_format)

            # Pivot
            for col_num, value in enumerate(pivot.columns):
                summary_ws.write(start_row + 1, left_col + col_num, value, header_format)
            for row_num in range(len(pivot)):
                for col_num in range(len(pivot.columns)):
                    summary_ws.write(start_row + 2 + row_num, left_col + col_num, pivot.iloc[row_num, col_num], cell_format)

            # Compliance Table
            for col_num, value in enumerate(final_summary.columns):
                summary_ws.write(start_row + 1, right_col + col_num, value, header_format)
            for row_num in range(len(final_summary)):
                for col_num in range(len(final_summary.columns)):
                    summary_ws.write(start_row + 2 + row_num, right_col + col_num, final_summary.iloc[row_num, col_num], cell_format)

            # Chart
            buckets = {
                "<=50%": final_summary[final_summary['Compliance %'].str.rstrip('%').astype(float) <= 50].shape[0],
                "50% - 75%": final_summary[
                    (final_summary['Compliance %'].str.rstrip('%').astype(float) > 50) &
                    (final_summary['Compliance %'].str.rstrip('%').astype(float) <= 75)
                ].shape[0],
                "Above 75%": final_summary[final_summary['Compliance %'].str.rstrip('%').astype(float) > 75].shape[0]
            }

            chart_data_row = start_row + 2 + max(len(pivot), len(final_summary)) + 2
            summary_ws.write_row(chart_data_row, right_col, ['Compliance Band', 'Count'])
            for i, (band, count) in enumerate(buckets.items(), start=1):
                summary_ws.write_row(chart_data_row + i, right_col, [band, count])

            chart = workbook.add_chart({'type': 'column'})  # type: ignore # Vertical bars
            chart.add_series({
                'name': 'Compliance Status',
                'categories': ["Summary", chart_data_row + 1, right_col, chart_data_row + 3, right_col],
                'values': ["Summary", chart_data_row + 1, right_col + 1, chart_data_row + 3, right_col + 1],
                'fill': {'color': '#2F75B5'}
            })
            chart.set_title({'name': f'{service_line} Compliance Status'})
            chart.set_x_axis({'name': 'Compliance Band'})
            chart.set_y_axis({'name': 'Count'})  # Normal Y-axis (not reversed)
            chart.set_style(10)
            summary_ws.insert_chart(chart_data_row + 1, right_col + 3, chart)


            start_row += max(len(pivot), len(final_summary)) + 6 + 4

        # ---------------------------------------
        # Sheet 2: ORG Summary
        # ---------------------------------------
        org_ws = workbook.add_worksheet("ORG") # type: ignore
        writer.sheets["ORG"] = org_ws

        # ORG-level Pivot
        org_pivot = pd.pivot_table(df, index='Service Line', values='EmpID', aggfunc='count', margins=True, margins_name='Grand Total').reset_index()
        for col_num, col in enumerate(org_pivot.columns):
            org_ws.write(0, col_num, col, header_format)
        for row_num in range(len(org_pivot)):
            for col_num in range(len(org_pivot.columns)):
                org_ws.write(1 + row_num, col_num, org_pivot.iloc[row_num, col_num], cell_format)

        # ORG Compliance
        df['Compliant'] = df['Adjusted Attendance'] >= result # type: ignore
        df['Non_Compliant'] = df['Adjusted Attendance'] < result # type: ignore
        org_summary = df.groupby('Service Line').agg(
            Count_of_Emp=('EmpID', 'count'),
            Compliant_Count=('Compliant', 'sum'),
            Non_Compliant_Count=('Non_Compliant', 'sum')
        ).reset_index()
        org_summary['% Non-compliance'] = (org_summary['Non_Compliant_Count'] / org_summary['Count_of_Emp']).apply(lambda x: f"{x:.0%}")
        org_summary['Compliance %'] = (org_summary['Compliant_Count'] / org_summary['Count_of_Emp']).apply(lambda x: f"{x:.0%}")

        grand_total = pd.DataFrame({
            'Service Line': ['Grand Total'],
            'Count_of_Emp': [org_summary['Count_of_Emp'].sum()],
            'Compliant_Count': [org_summary['Compliant_Count'].sum()],
            'Non_Compliant_Count': [org_summary['Non_Compliant_Count'].sum()],
            '% Non-compliance': [f"{org_summary['Non_Compliant_Count'].sum() / org_summary['Count_of_Emp'].sum():.0%}"],
            'Compliance %': [f"{org_summary['Compliant_Count'].sum() / org_summary['Count_of_Emp'].sum():.0%}"]
        })
        org_final = pd.concat([org_summary, grand_total], ignore_index=True)

        org_start = len(org_pivot) + 4
        for col_num, col in enumerate(org_final.columns):
            org_ws.write(org_start, col_num, col, header_format)
        for row_num in range(len(org_final)):
            for col_num in range(len(org_final.columns)):
                org_ws.write(org_start + 1 + row_num, col_num, org_final.iloc[row_num, col_num], cell_format)

    # Download button
    st.success("🎉 File generated successfully!")
    st.download_button(
        label="📥 Download Final Excel",
        data=output.getvalue(),
        file_name=f"{output_name.strip()}.xlsx", # type: ignore
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# st.stop()
