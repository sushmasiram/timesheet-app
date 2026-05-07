import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
from io import BytesIO

# PDF
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Timesheet Generator",
    layout="wide"
)

# =========================================================
# CUSTOM CSS (PASTEL UI)
# =========================================================

st.markdown("""
<style>

.main {
    background-color: #f7f9fc;
}

h1 {
    color: #4b6584;
    font-weight: 700;
}

h2, h3 {
    color: #5c6b73;
}

.stButton > button {
    background-color: #a8d8ea;
    color: #1f2937;
    border-radius: 12px;
    border: none;
    padding: 0.6rem 1.4rem;
    font-weight: 600;
}

.stButton > button:hover {
    background-color: #8ecae6;
    color: black;
}

.stDownloadButton > button {
    background-color: #b8f2e6;
    color: #1f2937;
    border-radius: 12px;
    border: none;
    padding: 0.6rem 1.4rem;
    font-weight: 600;
}

.stDownloadButton > button:hover {
    background-color: #9de2d0;
    color: black;
}

div[data-baseweb="select"] > div {
    background-color: #ffffff;
    border-radius: 10px;
}

input {
    border-radius: 10px !important;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}

.metric-card {
    background-color: #ffffff;
    padding: 20px;
    border-radius: 10px;
    border: 1px solid #e6ecf2;
    text-align: center;
    box-shadow: 0px 2px 8px rgba(0,0,0,0.04);
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# TITLE
# =========================================================

st.title("📅 Timesheet Generator")


# =========================================================
# COMPREHENSIVE HOLIDAY DATA
# =========================================================

holiday_calendar = [
    # 2024 Holidays
    {"date": "2024-01-26", "name": "Republic Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-03-25", "name": "Holi", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-04-11", "name": "Eid ul-Fitr", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-04-17", "name": "Ram Navami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-04-21", "name": "Mahavir Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-05-01", "name": "May/Maharashtra Day", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-08-15", "name": "Independence Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-08-26", "name": "Janmashtami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-10-02", "name": "Gandhi Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-10-12", "name": "Dussehra", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-11-01", "name": "Diwali", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2024-12-25", "name": "Christmas Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},

    # 2025 Holidays
    {"date": "2025-01-01", "name": "New Year's Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-01-26", "name": "Republic Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-03-14", "name": "Holi", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-03-30", "name": "Eid ul-Fitr", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-04-06", "name": "Ram Navami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-04-10", "name": "Mahavir Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-05-01", "name": "May/Maharashtra Day", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-08-15", "name": "Independence Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-08-16", "name": "Janmashtami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-10-02", "name": "Gandhi Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-10-01", "name": "Dussehra", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-10-20", "name": "Diwali", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2025-12-25", "name": "Christmas Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},

    # 2026 Holidays
    {"date": "2026-01-01", "name": "New Year's Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-01-26", "name": "Republic Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-03-06", "name": "Holi", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-03-19", "name": "Eid ul-Fitr", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-03-27", "name": "Ram Navami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-03-29", "name": "Mahavir Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-05-01", "name": "May/Maharashtra Day", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-08-15", "name": "Independence Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-09-07", "name": "Janmashtami", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-10-02", "name": "Gandhi Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-09-25", "name": "Dussehra", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-10-13", "name": "Diwali", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-12-25", "name": "Christmas Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
]


# =========================================================
# INPUTS SECTION
# =========================================================

st.markdown("###📝 Timesheet Details")

col1, col2 = st.columns(2)

with col1:
    name = st.text_input("Employee Name", placeholder="e.g., Sushma Motamarri")
    email = st.text_input("Email ID", placeholder="e.g., sushma@hankeranalytics.com")
    
with col2:
    department = st.selectbox(
        "Department",
        ["Engineering", "Sales", "HR", "Finance", "Operations", "Marketing", "Other"]
    )
    manager_name = st.text_input("Manager/Approver Name", placeholder="e.g., John Doe")

col3, col4, col5 = st.columns(3)

with col3:
    month_name = st.selectbox(
        "Select Month",
        list(calendar.month_name)[1:]
    )
    month = list(calendar.month_name).index(month_name)

with col4:
    year = st.number_input(
        "Enter Year",
        value=2026,
        min_value=2024,
        max_value=2030
    )

with col5:
    working_hours = st.number_input(
        "Daily Working Hours",
        value=8.0,
        min_value=1.0,
        max_value=12.0,
        step=0.5
    )

location = st.selectbox(
    "Select Location",
    ["Gurugram", "Bengaluru", "Chennai", "Pune", "Hyderabad"]
)

leave_dates = st.multiselect(
    "Select Leave Dates",
    list(range(1, 32))
)

remarks = st.text_area(
    "Add Remarks (Optional)",
    placeholder="e.g., Attended training, Conference, Remote work period, etc.",
    height=80
)

# =========================================================
# INPUT VALIDATION
# =========================================================

def validate_inputs(name, email, year, month):
    """Validate user inputs"""
    errors = []
    
    if not name.strip():
        errors.append("⚠ Employee Name is required")
    
    if not email.strip():
        errors.append("⚠ Email ID is required")
    elif "@" not in email:
        errors.append("⚠ Please enter a valid email")
    
    if year < 2024 or year > 2030:
        errors.append("⚠ Year should be between 2024-2030")
    
    if month < 1 or month > 12:
        errors.append("⚠ Invalid month selected")
    
    return errors


# =========================================================
# BUILD WEEKS STRUCTURE FOR WEEKLY WORK INPUT
# =========================================================

def build_holiday_map(location, year, month):
    """Build holiday map for the month"""
    return {
        h["date"]: h["name"]
        for h in holiday_calendar
        if h.get(location) == "Yes" and h["date"].startswith(f"{year}-{month:02d}")
    }


def build_weeks(year, month, holiday_map):
    """Build week structure for the month"""
    num_days = calendar.monthrange(year, month)[1]
    working_dates = []

    for day in range(1, num_days + 1):
        date_obj = datetime(year, month, day)
        date_str = date_obj.strftime("%Y-%m-%d")
        weekday = date_obj.weekday()

        if weekday >= 5:  # Skip weekends
            continue
        if date_str in holiday_map:  # Skip holidays
            continue

        working_dates.append(day)

    weeks_in_month = []
    current_week = []

    for day in working_dates:
        if not current_week:
            current_week.append(day)
        else:
            current_date = datetime(year, month, day)
            if current_date.weekday() == 0:  # New week if Monday
                weeks_in_month.append(current_week)
                current_week = [day]
            else:
                current_week.append(day)

    if current_week:
        weeks_in_month.append(current_week)

    return weeks_in_month


# =========================================================
# WEEKLY WORK DETAILS - APPEAR BEFORE GENERATE BUTTON
# =========================================================

st.markdown("### 🗓 Weekly Work Details")

# Build holiday map and weeks
holiday_map = build_holiday_map(location, year, month)
weeks_in_month = build_weeks(year, month, holiday_map)

# Create input fields for each week
weekly_work = {}

if len(weeks_in_month) == 0:
    st.warning("⚠️ No working weeks found for the selected month and location.")
else:
    for idx, week_days in enumerate(weeks_in_month, start=1):
        start_day = week_days[0]
        end_day = week_days[-1]
        
        week_label = f"{start_day}" if start_day == end_day else f"{start_day}-{end_day}"
        
        weekly_work[idx] = st.text_input(
            f"Week {idx} ({month_name} {week_label})",
            value=f"Client Work {idx}",
            key=f"week_{idx}"
        )


# =========================================================
# GENERATE BUTTON
# =========================================================

st.markdown("---")

if st.button("✅ Generate Timesheet", use_container_width=True):
    
    # Validate inputs
    validation_errors = validate_inputs(name, email, year, month)
    
    if validation_errors:
        for error in validation_errors:
            st.error(error)
        st.stop()

    # Warn if too many leaves
    num_days = calendar.monthrange(year, month)[1]
    if len(leave_dates) > num_days * 0.5:
        st.warning(f"⚠️ You've selected {len(leave_dates)} leave days. That's more than 50% of the month!")

    # =========================================================
    # BUILD TIMESHEET DATA - ALL DAYS INCLUDING WEEKENDS
    # =========================================================

    data = []
    working_days = 0
    holiday_count = 0
    leave_count = 0
    weekly_off_count = 0

    for day in range(1, num_days + 1):
        date_obj = datetime(year, month, day)
        date_str = date_obj.strftime("%Y-%m-%d")
        weekday = date_obj.weekday()

        # Find week number - only for working days
        week_number = None
        for idx, week_days in enumerate(weeks_in_month, start=1):
            if day in week_days:
                week_number = idx
                break

        project_work = ""

        # Determine status for THIS day
        if date_str in holiday_map:
            value = f"Holiday - {holiday_map[date_str]}"
            holiday_count += 1
        elif weekday >= 5:
            value = "Weekly Off"
            weekly_off_count += 1
        elif day in leave_dates:
            value = "Leave"
            leave_count += 1
        else:
            value = working_hours
            working_days += 1
            project_work = weekly_work.get(week_number, "Client Work")

        data.append({
            "Date": date_obj.strftime("%d-%m-%Y"),
            "Week": week_number,
            "Project": project_work,
            "Working hrs": value
        })

    df = pd.DataFrame(data)
    display_df = df.drop(columns=["Week"])

    # =========================================================
    # DISPLAY PREVIEW & SUMMARY
    # =========================================================

    st.subheader("📊 Preview & Summary")

    # Show metrics
    metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)

    with metric_col1:
        st.metric("Working Days", working_days, delta=f"{working_days * working_hours:.1f} hrs")
    with metric_col2:
        st.metric("Holidays", holiday_count)
    with metric_col3:
        st.metric("Leaves", leave_count)
    with metric_col4:
        st.metric("Weekly Offs", weekly_off_count)

    st.dataframe(display_df, use_container_width=True)

    # =========================================================
    # EXCEL EXPORT - SIMPLIFIED LOGIC
    # =========================================================

    def generate_excel():
        output = BytesIO()
        sheet_name = f"{month_name[:3]} Timesheet"

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet(sheet_name)

            # Formats
            header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#CDE7F0',
                'font_size': 11
            })

            title_format = workbook.add_format({
                'bold': True,
                'font_size': 12,
                'border': 0
            })

            normal_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })

            working_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#E8F6EF'
            })

            weekend_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#F5F5F5'
            })

            holiday_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFE6E6'
            })

            leave_format = workbook.add_format({
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#FFF4E6'
            })

            # Header section
            worksheet.write(0, 0, f"{month_name} {year} Timesheet", title_format)
            worksheet.write(1, 0, f"Name: {name}", normal_format)
            worksheet.write(1, 1, f"Email: {email}", normal_format)
            worksheet.write(2, 0, f"Department: {department}", normal_format)
            worksheet.write(2, 1, f"Manager: {manager_name}", normal_format)
            worksheet.write(3, 0, f"Generated: {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}", normal_format)
            worksheet.write(3, 1, f"Location: {location}", normal_format)

            # Column widths
            worksheet.set_column("A:A", 18)
            worksheet.set_column("B:B", 35)
            worksheet.set_column("C:C", 18)

            # Headers (at row 5)
            worksheet.write(5, 0, "Date", header_format)
            worksheet.write(5, 1, "Project", header_format)
            worksheet.write(5, 2, "Working hrs", header_format)

            # =========================================================
            # SIMPLIFIED DATA WRITING - WRITE ALL ROWS DIRECTLY
            # =========================================================
            # =========================================================
            # DATA WRITING WITH WEEKLY PROJECT MERGING
            # =========================================================

            excel_row = 6

            # ---------------------------------------------------------
            # FIRST WRITE ALL ROWS
            # ---------------------------------------------------------

            for idx, (_, row_data) in enumerate(df.iterrows()):

                current_row = excel_row + idx
                working_hrs_value = row_data['Working hrs']

                # Determine format
                if "Weekly Off" in str(working_hrs_value):

                    cell_format = weekend_format

                elif "Holiday" in str(working_hrs_value):

                    cell_format = holiday_format

                elif "Leave" in str(working_hrs_value):

                    cell_format = leave_format

                elif isinstance(working_hrs_value, (int, float)):

                    cell_format = working_format

                else:

                    cell_format = normal_format

                # -----------------------------------------------------
                # DATE COLUMN
                # -----------------------------------------------------

                worksheet.write(
                    current_row,
                    0,
                    row_data['Date'],
                    cell_format
                )

                # -----------------------------------------------------
                # PROJECT COLUMN
                # Keep blank for now
                # (will merge later)
                # -----------------------------------------------------

                worksheet.write(
                    current_row,
                    1,
                    "",
                    cell_format
                )

                # -----------------------------------------------------
                # WORKING HOURS COLUMN
                # -----------------------------------------------------

                worksheet.write(
                    current_row,
                    2,
                    working_hrs_value,
                    cell_format
                )


            # =========================================================
            # MERGE PROJECT COLUMN WEEK-WISE
            # =========================================================

            for week_num in sorted(df['Week'].dropna().unique()):

                # Get rows for this week
                week_df = df[df['Week'] == week_num]

                # Keep only ACTUAL WORKING DAYS
                week_df = week_df[
                    week_df['Working hrs'].apply(
                        lambda x: isinstance(x, (int, float))
                    )
                ]

                # Skip if no working days
                if len(week_df) == 0:
                    continue

                # Excel row positions
                start_idx = week_df.index[0]
                end_idx = week_df.index[-1]

                start_row = excel_row + start_idx
                end_row = excel_row + end_idx

                # Project name
                project_name = week_df.iloc[0]['Project']

                # -----------------------------------------------------
                # SINGLE ROW WEEK
                # -----------------------------------------------------

                if start_row == end_row:

                    worksheet.write(
                        start_row,
                        1,
                        project_name,
                        working_format
                    )

                # -----------------------------------------------------
                # MULTIPLE ROWS -> MERGE
                # -----------------------------------------------------

                else:

                    worksheet.merge_range(
                        start_row,   # first row
                        1,           # first col
                        end_row,     # last row
                        1,           # last col
                        project_name,
                        working_format
                    )


            # =========================================================
            # WRITE NON-WORKING DAY TEXTS IN PROJECT COLUMN
            # =========================================================

            for idx, (_, row_data) in enumerate(df.iterrows()):

                current_row = excel_row + idx
                working_hrs_value = row_data['Working hrs']

                # Only for non-working rows
                if not isinstance(working_hrs_value, (int, float)):

                    if "Weekly Off" in str(working_hrs_value):

                        cell_format = weekend_format

                    elif "Holiday" in str(working_hrs_value):

                        cell_format = holiday_format

                    elif "Leave" in str(working_hrs_value):

                        cell_format = leave_format

                    else:

                        cell_format = normal_format

                    worksheet.write(
                        current_row,
                        1,
                        working_hrs_value,
                        cell_format
                    )

            # =========================================================
            # SUMMARY SECTION
            # =========================================================

            start_row = len(df) + 8

            worksheet.write(start_row, 0, "Summary", header_format)
            worksheet.write(start_row + 1, 0, "Working Days", normal_format)
            worksheet.write(start_row + 1, 1, working_days, normal_format)
            worksheet.write(start_row + 2, 0, "Total Hours", normal_format)
            worksheet.write(start_row + 2, 1, working_days * working_hours, normal_format)
            worksheet.write(start_row + 3, 0, "Holidays", normal_format)
            worksheet.write(start_row + 3, 1, holiday_count, normal_format)
            worksheet.write(start_row + 4, 0, "Leaves", normal_format)
            worksheet.write(start_row + 4, 1, leave_count, normal_format)
            worksheet.write(start_row + 5, 0, "Weekly Offs", normal_format)
            worksheet.write(start_row + 5, 1, weekly_off_count, normal_format)

            if remarks.strip():
                worksheet.write(start_row + 7, 0, "Remarks", normal_format)
                worksheet.write(start_row + 7, 1, remarks, normal_format)

        output.seek(0)
        return output

    # =========================================================
    # PDF EXPORT
    # =========================================================

    def generate_pdf():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph(f"{month_name} {year} Timesheet", styles['Title']))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"<b>Name:</b> {name} | <b>Email:</b> {email} | <b>Department:</b> {department}", styles['Normal']))
        elements.append(Paragraph(f"<b>Manager:</b> {manager_name} | <b>Location:</b> {location}", styles['Normal']))
        elements.append(Spacer(1, 10))

        table_data = [display_df.columns.tolist()] + display_df.values.tolist()
        table = Table(table_data)
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER')
        ]))

        elements.append(table)
        elements.append(Spacer(1, 20))

        summary_data = [
            ["Working Days", working_days],
            ["Total Hours", f"{working_days * working_hours:.1f}"],
            ["Holidays", holiday_count],
            ["Leaves", leave_count],
            ["Weekly Offs", weekly_off_count]
        ]

        summary_table = Table(summary_data)
        summary_table.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black)]))
        elements.append(summary_table)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    # =========================================================
    # DOWNLOAD BUTTONS
    # =========================================================

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "📊 Download Excel",
            data=generate_excel(),
            file_name=f"{name}_{month_name}_{year}_Timesheet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        st.download_button(
            "📄 Download PDF",
            data=generate_pdf(),
            file_name=f"{name}_{month_name}_{year}_Timesheet.pdf",
            mime="application/pdf"
        )

    with col3:
        csv_data = display_df.to_csv(index=False).encode()
        st.download_button(
            "📋 Download CSV",
            data=csv_data,
            file_name=f"{name}_{month_name}_{year}_Timesheet.csv",
            mime="text/csv"
        )

    st.success("✅ Timesheet generated successfully!")
