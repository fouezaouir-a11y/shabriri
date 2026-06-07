import os
import re
import calendar
import pandas as pd
import streamlit as st
from datetime import datetime

# PDF generation libraries
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
# PDF reading library
from pypdf import PdfReader

# Set up browser layout frame config
st.set_page_config(page_title="Advanced Payroll Matrix", layout="wide")

st.title("📊 Advanced Attendance Parsing & Payroll Engine")
st.markdown("---")

# ----------------------------------------------------
# SIDEBAR CONFIGURATIONS (Global System Standards)
# ----------------------------------------------------
st.sidebar.header("🏢 Company Settings")
company_name = st.sidebar.text_input("Company Name", "Tech Den & DJ Freeshop")
uploaded_logo = st.sidebar.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
st.sidebar.header("📅 Pay Period Selection")

months_list = [
    ("January", "01"), ("February", "02"), ("March", "03"), ("April", "04"),
    ("May", "05"), ("June", "06"), ("July", "07"), ("August", "08"),
    ("September", "09"), ("October", "10"), ("November", "11"), ("December", "12")
]
selected_month_name, selected_month_code = st.sidebar.selectbox(
    "Select Month", months_list, index=4, format_func=lambda x: x[0]
)
selected_year = st.sidebar.selectbox("Select Year", ["2026", "2027", "2025"], index=0)
target_period = f"{selected_month_code}/{selected_year}"

# Dynamic Days in Month tracking variables
year_int = int(selected_year)
month_int = int(selected_month_code)
num_days_in_month = calendar.monthrange(year_int, month_int)[1]

st.sidebar.markdown("---")
st.sidebar.header("⏱️ Shift & Rules Matrix")
shift_start_time = st.sidebar.time_input(
    "Shift Start Time", 
    value=datetime.strptime("09:00", "%H:%M").time()
)
shift_hours = st.sidebar.radio("Paid Shift Length (Including 1H Paid Break)", [8, 9], index=0)
bonus_rule = st.sidebar.radio("7-Day Streak Bonus", [1.0, 0.5], format_func=lambda x: f"+{x} Day Pay")

shift_start_str = shift_start_time.strftime("%H:%M")

logo_path = None
if uploaded_logo:
    with open(uploaded_logo.name, "wb") as f:
        f.write(uploaded_logo.getbuffer())
    logo_path = uploaded_logo.name

# Calendar formatting helper
def get_calendar_label(day_num, month_num, year_num):
    dt = datetime(year_num, month_num, day_num)
    day_name = dt.strftime("%A") 
    return f"{day_num:02d}/{month_num:02d} ({day_name})"

# ----------------------------------------------------
# ADVANCED MULTI-EMPLOYEE & MULTI-PUNCH PARSING ENGINE
# ----------------------------------------------------
def analyze_master_biometric_pdf(pdf_file, target_month_str):
    master_database = {}
    if pdf_file is None:
        return master_database

    try:
        reader = PdfReader(pdf_file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() or ""
            
        lines = full_text.split("\n")
        for line in lines:
            ts_match = re.search(r"(\d{2})/(\d{2})/(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", line)
            if ts_match:
                day = int(ts_match.group(1))
                m_y = f"{ts_match.group(2)}/{ts_match.group(3)}"
                time_str = f"{ts_match.group(4)}:{ts_match.group(5)}"
                
                if m_y == target_month_str:
                    clean_line = line.replace("OUR COMPANY", "").replace("COMPANY", "").replace(ts_match.group(0), "")
                    clean_line = re.sub(r'[\"\',]', '', clean_line).strip()
                    emp_name_found = re.sub(r'\b\d\b', '', clean_line).strip() 
                    
                    if not emp_name_found or len(emp_name_found) < 2:
                        backup_digits = re.findall(r'\b\d+\b', clean_line)
                        emp_name_found = backup_digits[0] if backup_digits else "Unknown Staff"
                    
                    if emp_name_found not in master_database:
                        master_database[emp_name_found] = {}
                    if day not in master_database[emp_name_found]:
                        master_database[emp_name_found][day] = []
                    
                    if time_str not in master_database[emp_name_found][day]:
                        master_database[emp_name_found][day].append(time_str)
                        
    except Exception as e:
        st.error(f"Engine parsing alert: {e}")
        
    for emp in master_database:
        for day in master_database[emp]:
            master_database[emp][day] = sorted(list(set(master_database[emp][day])))
            
    return master_database

# ----------------------------------------------------
# MAIN DASHBOARD INTERFACE
# ----------------------------------------------------
st.subheader("1. File Ingestion Setup")
uploaded_pdf = st.file_uploader("Drop machine log PDF report sheet here", type=["pdf"])

parsed_master_db = analyze_master_biometric_pdf(uploaded_pdf, target_period)

st.markdown("---")
st.subheader("2. Target Employee Management")

default_name_value = ""
if parsed_master_db:
    default_name_value = sorted(list(parsed_master_db.keys()))[0]

selected_employee_key = st.text_input("👤 Employee Name (Extracted automatically or type manually):", value=default_name_value)

if selected_employee_key:
    st.success(f"Processing sheet records layout for: **{selected_employee_key}**")
else:
    st.warning("⚠️ Please provide an Employee Name to label the calculation profiles.")

# ----------------------------------------------------
# PROFILE CONFIGURATIONS PANEL
# ----------------------------------------------------
col_e1, col_e2 = st.columns(2)
with col_e1:
    base_salary = st.number_input("Base Monthly Salary (DA)", min_value=0.0, value=40000.0, step=500.0)
with col_e2:
    advance_pay = st.number_input("Advance Deductions (DA)", min_value=0.0, value=0.0, step=500.0)

selected_employee_logs = parsed_master_db.get(selected_employee_key, {})
if not selected_employee_logs and parsed_master_db:
    first_file_key = sorted(list(parsed_master_db.keys()))[0]
    selected_employee_logs = parsed_master_db.get(first_file_key, {})

structured_data = {}
for d in range(1, num_days_in_month + 1):
    punches = sorted(selected_employee_logs.get(d, []))
    if punches:
        fmt = "%H:%M"
        act = datetime.strptime(punches[0], fmt)
        exp = datetime.strptime(shift_start_str, fmt)
        late_min = int((act - exp).total_seconds() / 60) if act > exp else 0
        punch_display_text = " -> ".join(punches)
        
        structured_data[d] = {
            "Status": "Present",
            "Clock Pointers": punch_display_text,
            "Late Minutes": late_min,
            "Lateness Status": "Non-Justified" if late_min > 0 else "N/A"
        }
    else:
        structured_data[d] = {
            "Status": "Unauthorized Absence",
            "Clock Pointers": "--",
            "Late Minutes": 0,
            "Lateness Status": "N/A"
        }

# ----------------------------------------------------
# INTERACTIVE DATA MATRIX (Fully Editable Punches)
# ----------------------------------------------------
st.markdown("---")
st.subheader("3. 📋 Review Logs & Modify Row Profiles")
st.caption("💡 CRITICAL: The 'Biometric Punches' column is fully editable! You can fix punch mistakes by typing timestamps separated by arrows, e.g., '09:02 -> 13:00 -> 14:15 -> 18:00'.")

initial_rows = []
for d in range(1, num_days_in_month + 1):
    day_profile = structured_data[d]
    date_label = get_calendar_label(d, month_int, year_int)
    initial_rows.append({
        "Date": date_label,
        "Attendance Status": day_profile["Status"],
        "Biometric Punches": day_profile["Clock Pointers"],
        "Lateness Status": day_profile["Lateness Status"]
    })

df_initial = pd.DataFrame(initial_rows)

edited_df = st.data_editor(
    df_initial,
    use_container_width=True,
    height=400,
    disabled=["Date"], 
    column_config={
        "Attendance Status": st.column_config.SelectboxColumn(
            "Attendance Status",
            options=["Present", "Day Off", "Paid Leave", "Authorized Absence", "Unauthorized Absence"],
            required=True
        ),
        "Biometric Punches": st.column_config.TextColumn(
            "Biometric Punches",
            help="Type directly into this field to manually correct or add timestamps.",
            required=True
        ),
        "Lateness Status": st.column_config.SelectboxColumn(
            "Lateness Status",
            options=["Non-Justified", "Justified", "N/A"],
            required=True
        )
    },
    hide_index=True,
    key="attendance_editor_v10"
)

# ----------------------------------------------------
# LIVE DETAILED CALCULATION PREVIEW ENGINE
# ----------------------------------------------------
st.markdown("---")
st.subheader("4. 🧮 Live Day-by-Day Financial Breakdown (Verify Before Printing)")

daily_rate = base_salary / float(num_days_in_month)
hourly_rate = daily_rate / shift_hours

calculated_preview_rows = []
final_attendance_profile = {}
total_pay = 0.0
consecutive_work_days = 0

for idx, row in edited_df.iterrows():
    day_num = idx + 1
    date_string = row["Date"]
    status = row["Attendance Status"]
    late_status = row["Lateness Status"]
    
    raw_punch_str = str(row["Biometric Punches"]).strip()
    if raw_punch_str in ["--", "None", "nan", ""]:
        punches = []
    else:
        punches = sorted([p.strip() for p in raw_punch_str.split("->") if re.match(r"^\d{2}:\d{2}$", p.strip())])
    
    # Isolate exact clock slots for the printed report data columns
    c_in = punches[0] if len(punches) >= 1 else "--"
    b_out = punches[1] if len(punches) >= 4 else "--"
    b_in = punches[2] if len(punches) >= 4 else "--"
    c_out = punches[-1] if len(punches) >= 2 else "--"
    
    late_min = 0
    if status == "Present" and punches:
        try:
            fmt = "%H:%M"
            act = datetime.strptime(punches[0], fmt)
            exp = datetime.strptime(shift_start_str, fmt)
            if act > exp:
                late_min = int((act - exp).total_seconds() / 60)
        except:
            pass

    hours_worked = 0.0
    overtime_hours = 0.0
    day_earnings = 0.0
    penalty = 0.0
    late_penalty = 0.0
    break_overstay_minutes = 0
    
    if status == "Present":
        if len(punches) >= 2:
            fmt = "%H:%M"
            total_seconds_worked = 0
            
            for i in range(0, len(punches) - 1, 2):
                try:
                    t_in = datetime.strptime(punches[i], fmt)
                    t_out = datetime.strptime(punches[i+1], fmt)
                    total_seconds_worked += (t_out - t_in).total_seconds()
                except:
                    continue
            
            if len(punches) >= 4:
                try:
                    break_out_dt = datetime.strptime(punches[1], fmt)
                    break_in_dt = datetime.strptime(punches[2], fmt)
                    total_break_seconds = (break_in_dt - break_out_dt).total_seconds()
                    total_break_minutes = total_break_seconds / 60.0
                    
                    if total_break_minutes > 60.0:
                        break_overstay_minutes = int(total_break_minutes - 60)
                except:
                    pass
            
            try:
                first_punch = datetime.strptime(punches[0], fmt)
                last_punch = datetime.strptime(punches[-1], fmt)
                total_span_hours = (last_punch - first_punch).total_seconds() / 3600.0
                hours_worked = total_span_hours
            except:
                hours_worked = float(shift_hours)
            
            if hours_worked == 0:
                hours_worked = float(shift_hours)

            if hours_worked > shift_hours:
                overtime_hours = hours_worked - shift_hours
                regular_hours = shift_hours
            else:
                regular_hours = hours_worked
                
            day_earnings = (regular_hours * hourly_rate) + (overtime_hours * hourly_rate * 1.51)
        else:
            hours_worked = float(shift_hours)
            day_earnings = daily_rate
        
        if late_min > 15 and late_status == "Non-Justified":
            late_penalty += (1.5 * hourly_rate)
            extra_time = late_min - 15
            intervals = (extra_time + 14) // 15
            late_penalty += (intervals * 1.5 * hourly_rate)
            
        if break_overstay_minutes > 0 and late_status == "Non-Justified":
            late_penalty += (1.5 * hourly_rate)
            extra_break_time = break_overstay_minutes - 15
            if extra_break_time > 0:
                break_intervals = (extra_break_time + 14) // 15
                late_penalty += (break_intervals * 1.5 * hourly_rate)
            
        day_earnings -= late_penalty
        consecutive_work_days += 1
    else:
        consecutive_work_days = 0
        if status in ["Day Off", "Paid Leave"]:
            day_earnings = daily_rate
        elif status == "Authorized Absence":
            day_earnings = 0.0
            if late_min > 0: late_penalty = (0.25 * hourly_rate)
            day_earnings -= late_penalty
        elif status == "Unauthorized Absence":
            penalty = daily_rate * 0.5
            day_earnings = -penalty

    if consecutive_work_days == 7:
        day_earnings += (daily_rate * bonus_rule)
        consecutive_work_days = 0

    total_pay += day_earnings
    total_deductions = penalty + late_penalty
    
    info_label = status
    if status == "Present":
        notes = []
        if late_min > 0: notes.append(f"Late {late_min}m")
        if break_overstay_minutes > 0: notes.append(f"Break Over {break_overstay_minutes}m")
        if notes:
            info_label += f" ({', '.join(notes)}"
            if late_status == "Justified": info_label += " - Justified)"
            else: info_label += ")"

    final_attendance_profile[day_num] = {
        "DateText": date_string,
        "ClockIn": c_in,
        "BreakOut": b_out,
        "BreakIn": b_in,
        "ClockOut": c_out,
        "LateStatusText": info_label,
        "Deductions": total_deductions,
        "NetEarnings": day_earnings
    }
    
    calculated_preview_rows.append({
        "Date": date_string,
        "Status": status,
        "Biometric Flow": raw_punch_str if status == "Present" else "--",
        "Hours Worked": f"{hours_worked:.1f} hrs",
        "Penalties Deducted": f"-{total_deductions:.2f} DA" if total_deductions > 0 else "0.00 DA",
        "Net Earnings": f"{day_earnings:.2f} DA"
    })

df_financial_verification = pd.DataFrame(calculated_preview_rows)
st.dataframe(df_financial_verification, use_container_width=True, height=400, hide_index=True)

# Summary Panels
net_salary = total_pay - advance_pay
col_m1, col_m2, col_m3 = st.columns(3)
with col_m1:
    st.metric("Gross Parsed Earnings", f"{total_pay:.2f} DA")
with col_m2:
    st.metric("Salary Advances (Deductions)", f"-{advance_pay:.2f} DA")
with col_m3:
    st.metric("Final Employee Net Pay-Out", f"{net_salary:.2f} DA")

# ----------------------------------------------------
# REPORT COMPILER EXPORT
# ----------------------------------------------------
st.markdown("---")
st.subheader("5. Export Final Report Documents")

if st.button("Compile & Print Final PDF Statement", type="primary"):
    pdf_rows = []
    for d in sorted(final_attendance_profile.keys()):
        p = final_attendance_profile[d]
        pdf_rows.append([
            p["DateText"], 
            p["ClockIn"], p["BreakOut"], p["BreakIn"], p["ClockOut"],
            p["LateStatusText"], 
            f"-{p['Deductions']:.2f} DA" if p['Deductions'] > 0 else "0.00 DA",
            f"{p['NetEarnings']:.2f} DA"
        ])
        
    filename = f"Payroll_Report_{selected_employee_key.replace(' ', '_')}_{target_period.replace('/', '_')}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"), spaceAfter=10)
    normal_style = styles['Normal']
    bold_style = ParagraphStyle('BoldText', parent=normal_style, fontName='Helvetica-Bold')

    header_data = []
    if logo_path and os.path.exists(logo_path):
        img = Image(logo_path, width=50, height=50)
        header_data = [[img, Paragraph(f"<b>{company_name}</b><br/>Payroll Statement - {target_period}", title_style)]]
    else:
        header_data = [[Paragraph(f"<b>{company_name}</b><br/>Payroll Statement - {target_period}", title_style)]]
        
    header_table = Table(header_data, colWidths=[100, 470] if logo_path else [570])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(header_table)
    story.append(Spacer(1, 15))
    
    emp_info = [
        [Paragraph("<b>Employee Name:</b>", normal_style), Paragraph(selected_employee_key, normal_style), Paragraph("<b>Base Monthly Salary:</b>", normal_style), Paragraph(f"{base_salary:.2f} DA", normal_style)],
        [Paragraph("<b>Pay Period:</b>", normal_style), Paragraph(target_period, normal_style), Paragraph("<b>Advance Reductions:</b>", normal_style), Paragraph(f"{advance_pay:.2f} DA", normal_style)],
        ["", "", Paragraph("<b>Net Pay Out:</b>", bold_style), Paragraph(f"<b>{net_salary:.2f} DA</b>", bold_style)]
    ]
    info_table = Table(emp_info, colWidths=[130, 150, 150, 140])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (2,2), (3,2), colors.HexColor("#E2E8F0")),
        ('PADDING', (0,0), (-1,-1), 5),
        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.lightgrey)
    ]))
    story.append(info_table)
    story.append(Spacer(1, 15))
    
    # Clean, detailed header with columns for every shift state
    table_header = ["Date", "In", "Break Out", "Break In", "Out", "Status / Info", "Penalties", "Net Earnings"]
    table_data = [table_header] + pdf_rows
    
    # Document dimensions adjusted beautifully to fit perfectly onto standard print layouts
    report_table = Table(table_data, colWidths=[105, 40, 55, 50, 40, 120, 75, 85])
    report_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,0), 4),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")])
    ]))
    
    story.append(report_table)
    doc.build(story)
    
    st.success(f"✅ Calculation successfully processed! Report compiled as: '{filename}'")
    
    with open(filename, "rb") as pdf_file:
        st.download_button(
            label="📥 Download Detailed PDF Payroll Report",
            data=pdf_file,
            file_name=filename,
            mime="application/pdf"
        )