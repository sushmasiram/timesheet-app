import streamlit as st
import pandas as pd
from datetime import datetime
import calendar
from io import BytesIO

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet


st.title("📊 Timesheet Generator")

# ---------------- INPUTS ----------------
name = st.text_input("Enter Name")
email = st.text_input("Enter Email ID")
month_name = st.selectbox("Select Month", list(calendar.month_name)[1:])
month = list(calendar.month_name).index(month_name)

year = st.number_input("Enter Year", value=2026)

location = st.selectbox("Select Location", ["Gurugram", "Bengaluru", "Chennai", "Pune", "Hyderabad"])

leave_dates = st.multiselect("Select Leave Dates", list(range(1, 32)))

# ---------------- HOLIDAY DATA ----------------
holiday_calendar = [
    {"date": "2026-01-01", "name": "New Year’s Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-01-15", "name": "Pongal", "Gurugram": "No", "Bengaluru": "No", "Chennai": "Yes", "Pune": "No", "Hyderabad": "Yes"},
    {"date": "2026-01-26", "name": "Republic Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-03-04", "name": "Holi", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "No", "Pune": "Yes", "Hyderabad": "No"},
    {"date": "2026-03-19", "name": "Ugadi/Gudi Padwa", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "No", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-04-03", "name": "Good Friday", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "No", "Pune": "No", "Hyderabad": "Yes"},
    {"date": "2026-04-14", "name": "Tamil New Year", "Gurugram": "No", "Bengaluru": "No", "Chennai": "Yes", "Pune": "No", "Hyderabad": "No"},
    {"date": "2026-05-01", "name": "May/Maharashtra Day", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-06-02", "name": "Telangana Foundation Day", "Gurugram": "No", "Bengaluru": "No", "Chennai": "No", "Pune": "No", "Hyderabad": "Yes"},
    {"date": "2026-08-28", "name": "Raksha Bandhan", "Gurugram": "Yes", "Bengaluru": "No", "Chennai": "No", "Pune": "No", "Hyderabad": "No"},
    {"date": "2026-09-04", "name": "Janmashtami", "Gurugram": "Yes", "Bengaluru": "No", "Chennai": "No", "Pune": "No", "Hyderabad": "No"},
    {"date": "2026-09-14", "name": "Ganesh Chaturthi", "Gurugram": "No", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-10-02", "name": "Gandhi Jayanti", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
    {"date": "2026-10-20", "name": "Ayudha Pooja/Dussehra", "Gurugram": "Yes", "Bengaluru": "No", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "No"},
    {"date": "2026-11-09", "name": "Govardhan Pooja", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "No"},
    {"date": "2026-12-25", "name": "Christmas Day", "Gurugram": "Yes", "Bengaluru": "Yes", "Chennai": "Yes", "Pune": "Yes", "Hyderabad": "Yes"},
]

# ---------------- GENERATE BUTTON ----------------
if st.button("Generate Timesheet"):

    # Build holiday map for selected location
    holiday_map = {
        h["date"]: h["name"]
        for h in holiday_calendar
        if h.get(location) == "Yes"
    }

    num_days = calendar.monthrange(year, month)[1]

    data = []
    working_days = 0
    holiday_count = 0
    leave_count = 0
    weekly_off_count = 0

    # --------- MAIN LOOP ---------
    for day in range(1, num_days + 1):
        date_obj = datetime(year, month, day)
        date_str = date_obj.strftime("%Y-%m-%d")
        weekday = date_obj.weekday()

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
            value = 8
            working_days += 1

        data.append({
            "Date": date_obj.strftime("%d-%m-%Y"),
            "Project": "Client Work",
            "Working hrs": value
        })

    df = pd.DataFrame(data)

    st.subheader("Preview")
    st.dataframe(df)

    # ---------------- EXCEL ----------------
    def generate_excel():
        output = BytesIO()

        file_month = f"{month_name}"
        sheet_name = f"{month_name[:3]} Timesheet"

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=3)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]

            # ---- FORMATS ----
            header_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'align': 'center'
            })

            normal_format = workbook.add_format({
                'border': 1
            })

            # ---- NAME + EMAIL ----
            worksheet.write(0, 0, f"Name: {name}")
            worksheet.write(1, 0, f"Email: {email}")

            # ---- COLUMN WIDTH ----
            worksheet.set_column("A:A", 15)
            worksheet.set_column("B:B", 20)
            worksheet.set_column("C:C", 25)

            # ---- APPLY HEADER FORMAT ----
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(3, col_num, col_name, header_format)

            # ---- APPLY CELL BORDER ----
            for row in range(len(df)):
                for col in range(len(df.columns)):
                    worksheet.write(row + 4, col, df.iloc[row, col], normal_format)

            # ---- SUMMARY ----
            start_row = len(df) + 6

            worksheet.write(start_row, 0, "Summary", header_format)

            worksheet.write(start_row + 1, 0, "Working Days")
            worksheet.write(start_row + 1, 1, working_days)

            worksheet.write(start_row + 2, 0, "Holidays")
            worksheet.write(start_row + 2, 1, holiday_count)

            worksheet.write(start_row + 3, 0, "Leaves")
            worksheet.write(start_row + 3, 1, leave_count)


        output.seek(0)
        return output

    # ---------------- PDF ----------------
    def generate_pdf():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)

        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph(f"{month_name} {year} Timesheet", styles['Title']))
        elements.append(Spacer(1, 10))

        table_data = [df.columns.tolist()] + df.values.tolist()

        table = Table(table_data)
        table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
            ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 20))

        summary_data = [
            ["Working Days", working_days],
            ["Holidays", holiday_count],
            ["Leaves", leave_count]
        ]

        summary_table = Table(summary_data)
        summary_table.setStyle(TableStyle([
            ('GRID', (0,0), (-1,-1), 1, colors.black),
        ]))

        elements.append(summary_table)

        doc.build(elements)
        buffer.seek(0)
        return buffer

    # ---------------- DOWNLOAD BUTTONS ----------------
    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Download Excel",
            data=generate_excel(),
            file_name = f"{name}'s Timesheet_{month_name}01-{month_name}{len(df)}, {year}.xlsx"
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        st.download_button(
            "📄 Download PDF",
            data=generate_pdf(),
            file_name = f"{name}'s Timesheet_{month_name}01-{month_name}{len(df)}, {year}.pdf"
            mime="application/pdf"
        )
