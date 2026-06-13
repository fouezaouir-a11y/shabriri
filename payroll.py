import os
import re
import calendar
import pandas as pd
import streamlit as st
from datetime import datetime

# PDF generation libraries
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
# PDF reading library
from pypdf import PdfReader

# Set up browser layout frame config
st.set_page_config(page_title="Advanced Payroll Matrix", layout="wide")

st.title("?? Universal Attendance Parsing & Payroll Engine")
st.markdown("---")

# ----------------------------------------------------
# SIDEBAR CONFIGURATIONS (Global System Standards)
# ----------------------------------------------------
st.sidebar.header("?? Company Settings")
company_name = st.sidebar.text_input("Company Name", "Tech Den & DJ Freeshop")
uploaded_logo = st.sidebar.file_uploader("Upload Company Logo", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
st.sidebar.header("?? Pay Period Selection")

months_list = [
    ("January", "01"), ("February", "02"), ("March", "03"), ("April", "04"),
    ("May", "05"), ("June", "06"), ("July", "07"), ("August", "08"),
    ("September", "09"), ("October", "10"), ("November", "11"), ("December", "12")
]
# Defaults to November to align smoothly with your real data period logs
selected_month_name, selected_month_code = st.sidebar.selectbox(
    "Select Month", months_list, index=10, format_func=lambda x: x[0]
)
selected_year = st.sidebar.selectbox("Select Year", ["2025", "2026", "2027"], index=0)
target_period = f"{selected_month_code}/{selected_year}"

# Dynamic Days in Month tracking variables
year_int = int(selected_year)
month_int = int(selected_month_code)
num_days_in_month = calendar.monthrange(year_int, month_int)[1]

st.sidebar.markdown("---")
st.sidebar.header("?? Shift & Rules Matrix")
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
# UNIVERSAL PARSING ENGINE (PDF, EXCEL, & PLAIN TEXT)
# ----------------------------------------------------
def analyze_master_biometric_log(uploaded_file, target_month_str):
    master_database = {}
    if uploaded_file is None:
        return master_database

    filename = uploaded_file.name.lower()
    
    parts = target_month_str.split('/')
    if len(parts) == 2:
        target_m, target_y = parts[0], parts[1]
    else:
        target_m, target_y = "11", "2025"

    # --- METHOD A: EXCEL PARSING CORE (.XLSX / .XLS) ---
    if filename.endswith(('.xlsx', '.xls')):
        try:
            if filename.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            else:
                df = pd.read_excel(uploaded_file, engine='xlrd')
                
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            name_col = next((c for c in df.columns if any(k in c for k in ['name', 'employee', 'nom', 'user', 'id', 'person'])), None)
            date_col = next((c for c in df.columns if any(k in c for k in ['date', 'time', 'punch', 'horaire', 'check', 'mouvement'])), None)
            
            if not name_col or not date_col:
                name_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]
                date_col = df.columns[3] if len(df.columns) > 3 else df.columns[1]

            if name_col and date_col:
                for _, row in df.iterrows():
                    emp_raw = str(row[name_col]).strip()
                    date_raw = str(row[date_col]).strip()
                    
                    if emp_raw.lower() in ['nan', 'null', '', 'none'] or date_raw.lower() in ['nan', 'null', '', 'none']:
                        continue
                    
                    day, m_str, y_str, time_str = None, None, None, None

                    # Pattern 1: Machine Standard (YYYY/MM/DD HH:MM:SS AM/PM)
                    machine_match = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2}):?(\d{2})?\s*(AM|PM)?", date_raw, re.IGNORECASE)
                    # Pattern 2: Manual Test Standard (DD/MM/YYYY HH:MM)
                    test_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})", date_raw)

                    if machine_match:
                        y_str = machine_match.group(1)
                        m_str = f"{int(machine_match.group(2)):02d}"
                        day = int(machine_match.group(3))
                        hr = int(machine_match.group(4))
                        minute = machine_match.group(5)
                        ampm = machine_match.group(7)
                        
                        if ampm:
                            ampm = ampm.upper()
                            if ampm == "PM" and hr < 12: hr += 12
                            elif ampm == "AM" and hr == 12: hr = 0
                        time_str = f"{hr:02d}:{minute}"

                    elif test_match:
                        day = int(test_match.group(1))
                        m_str = f"{int(test_match.group(2)):02d}"
                        y_str = test_match.group(3)
                        time_str = f"{int(test_match.group(4)):02d}:{test_match.group(5)}"

                    if day and m_str == target_m and y_str == target_y:
                        emp_name_found = re.sub(r'[\"\',]', '', emp_raw).strip()
                        if emp_name_found not in master_database:
                            master_database[emp_name_found] = {}
                        if day not in master_database[emp_name_found]:
                            master_database[emp_name_found][day] = []
                        if time_str not in master_database[emp_name_found][day]:
                            master_database[emp_name_found][day].append(time_str)
            return master_database
        except Exception as e:
            st.error(f"Excel Sheet Parsing Alert: {e}")
            return master_database

    # --- METHOD B: PLAIN TEXT LOG PARSING CORE (.TXT) ---
    elif filename.endswith('.txt'):
        try:
            raw_data = uploaded_file.read()
            full_text = raw_data.decode('utf-8', errors='ignore')
            lines = full_text.splitlines()
            
            for line in lines:
                if not line.strip():
                    continue
                
                day, m_str, y_str, time_str = None, None, None, None
                
                machine_match = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})\s+(\d{1,2}):(\d{2}):?(\d{2})?\s*(AM|PM)?", line, re.IGNORECASE)
                test_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})", line)

                if machine_match:
                    y_str = machine_match.group(1)
                    m_str = f"{int(machine_match.group(2)):02d}"
                    day = int(machine_match.group(3))
                    hr = int(machine_match.group(4))
                    minute = machine_match.group(5)
                    ampm = machine_match.group(7)
                    
                    if ampm:
                        ampm = ampm.upper()
                        if ampm == "PM" and hr < 12: hr += 12
                        elif ampm == "AM" and hr == 12: hr = 0
                    time_str = f"{hr:02d}:{minute}"
                    matched_str = machine_match.group(0)
                elif test_match:
                    day = int(test_match.group(1))
                    m_str = f"{int(test_match.group(2)):02d}"
                    y_str = test_match.group(3)
                    time_str = f"{int(test_match.group(4)):02d}:{test_match.group(5)}"
                    matched_str = test_match.group(0)

                if day and m_str == target_m and y_str == target_y:
                    clean_line = line.replace(matched_str, "")
                    clean_line = re.sub(r'[\"\',\t\s\:]+', ' ', clean_line).strip()
                    emp_name_found = re.sub(r'\b\d\b', '', clean_line).strip()
                    
                    if not emp_name_found or len(emp_name_found) < 2:
                        backup_digits = re.findall(r'\b\d+\b', clean_line)
                        emp_name_found = f"Staff ID: {backup_digits[0]}" if backup_digits else "Unknown Identity"
                    
                    if emp_name_found not in master_database:
                        master_database[emp_name_found] = {}
                    if day not in master_database[emp_name_found]:
                        master_database[emp_name_found][day] = []
                    if time_str not in master_database[emp_name_found][day]:
                        master_database[emp_name_found][day].append(time_str)
            return master_database
        except Exception as e:
            st.error(f"Text File Parsing Alert: {e}")
            return master_database

    # --- METHOD C: ROBUST COMPILER PDF PARSING CORE (.PDF) ---
    else:
        try:
            reader = PdfReader(uploaded_file)
            full_text = ""
            for page in reader.pages:
                page_text = page.extract_text() or ""
                if page_text and "\n" not in page_text:
                    page_text = re.sub(r'(\d{4}/\d{1,2}/\d{1,2})', r'\n\1', page_text)
                    page_text = re.sub(r'(\d{1,2}/\d{1,2}/\d{4})', r'\n\1', page_text)
                full_text += page_text + "\n"
                
            full_text = full_text.encode('utf-8', errors='ignore').decode('utf-8')
            lines = full_text.split("\n")
            
            for line in lines:
                day, m_str, y_str, time_str = None, None, None, None
                
                machine_match = re.search(r"(\d{4})/(\d{1,2})/(\d{1,2})\s*(\d{1,2}):(\d{2}):?(\d{2})?\s*(AM|PM)?", line, re.IGNORECASE)
                test_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*(\d{1,2}):(\d{2})", line)

                if machine_match:
                    y_str = machine_match.group(1)
                    m_str = f"{int(machine_match.group(2)):02d}"
                    day = int(machine_match.group(3))
                    hr = int(machine_match.group(4))
                    minute = machine_match.group(5)
                    ampm = machine_match.group(7)
                    
                    if ampm:
                        ampm = ampm.upper()
                        if ampm == "PM" and hr < 12: hr += 12
                        elif ampm == "AM" and hr == 12: hr = 0
                    time_str = f"{hr:02d}:{minute}"
                    matched_str = machine_match.group(0)
                elif test_match:
                    day = int(test_match.group(1))
                    m_str = f"{int(test_match.group(2)):02d}"
                    y_str = test_match.group(3)
                    time_str = f"{int(test_match.group(4)):02d}:{test_match.group(5)}"
                    matched_str = test_match.group(0)

                if day and m_str == target_m and y_str == target_y:
                    clean_line = line.replace("OUR COMPANY", "").replace("COMPANY", "").replace(matched_str, "")
                    clean_line = re.sub(r'[\"\',]', '', clean_line).strip()
                    emp_name_found = re.sub(r'\b\d\b', '', clean_line).strip() 
                    
                    # Target layout fix for specific clipped PDF formats
                    if ":" in emp_name_found:
                        emp_name_found = emp_name_found.split(":")[0].strip()
                    
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
            st.error(f"PDF Engine parsing alert: {e}")
            
    # Universal Sort Cleanup Pass
    for emp in list(master_database.keys()):
        for day in list(master_database[emp].keys()):
            master_database[emp][day] = sorted(list(set(master_database[emp][day])))
            
    return master_database

# ----------------------------------------------------
# MAIN DASHBOARD INTERFACE
# ----------------------------------------------------
st.subheader("1. File Ingestion Setup")
uploaded_file = st.file_uploader("Drop machine log PDF report, Excel logs, or plain Text dumps here", type=["pdf", "xlsx", "xls", "txt"])

parsed_master_db = analyze_master_biometric_log(uploaded_file, target_period)

st.markdown("---")
st.subheader("2. Target Employee Management")

default_name_value = ""
if parsed_master_db:
    default_name_value = sorted(list(parsed_master_db.keys()))[0]

selected_employee_key = st.text_input("?? Employee Name:", value=default_name_value)

# Profile configurations
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
# INTERACTIVE DATA MATRIX
# ----------------------------------------------------
st.markdown("---")
st.subheader("3. ?? Review Logs & Modify Row Profiles")

initial_rows = []
for d in range(1, num_days_in_month + 1):
    day_profile = structured_data[d]
    try:
        date_label = get_calendar_label(d, month_int, year_int)
    except:
        date_label = f"{d:02d}/{month_int:02d}"
        
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
    height=380,
    disabled=["Date"], 
    column_config={
        "Attendance Status": st.column_config.SelectboxColumn(
            "Attendance Status",
            options=["Present", "Day Off", "Paid Leave", "Authorized Absence", "Unauthorized Absence"],
            required=True
        ),
        "Biometric Punches": st.column_config.TextColumn("Biometric Punches (Format: IN -> OUT)", required=True),
        "Lateness Status": st.column_config.SelectboxColumn("Lateness Status", options=["Non-Justified", "Justified", "N/A"], required=True)
    },
    hide_index=True,
    key="universal_attendance_engine_matrix"
)

# ----------------------------------------------------
# LIVE FINANCIAL CALCULATION PREVIEW ENGINE
# ----------------------------------------------------
st.markdown("---")
st.subheader("4. ?? Live Calculation Summary Table")

daily_rate = base_salary / float(num_days_in_month)
hourly_rate = daily_rate / shift_hours

calculated_preview_rows = []
final_attendance_profile = {}
total_pay = 0.0
total_all_bonus = 0.0
total_all_penalty = 0.0
total_hours_worked = 0.0
total_overtime_hours = 0.0
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
    
    c_in = punches[0] if len(punches) >= 1 else "--"
    l_out = punches[1] if len(punches) >= 3 else "--" 
    l_in = punches[2] if len(punches) >= 4 else "--"
    c_out = punches[-1] if len(punches) >= 2 else "--"

    late_min = 0
    if status == "Present" and punches:
        try:
            fmt = "%H:%M"
            act = datetime.strptime(punches[0], fmt)
            exp = datetime.strptime(shift_start_str, fmt)
            if act > exp:
                late_min = int((act - exp).total_seconds() / 60)
        except: pass

    hours_worked = 0.0
    overtime_hours = 0.0
    day_earnings = 0.0
    penalty = 0.0
    late_penalty = 0.0
    day_bonus = 0.0
    
    if status == "Present":
        if len(punches) >= 2:
            fmt = "%H:%M"
            try:
                first_punch = datetime.strptime(punches[0], fmt)
                last_punch = datetime.strptime(punches[-1], fmt)
                hours_worked = (last_punch - first_punch).total_seconds() / 3600.0
            except:
                hours_worked = float(shift_hours)
            
            if hours_worked == 0: hours_worked = float(shift_hours)
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
            
        penalty = late_penalty
        day_earnings -= penalty
        consecutive_work_days += 1
    else:
        consecutive_work_days = 0
        if status in ["Day Off", "Paid Leave"]:
            day_earnings = daily_rate
        elif status == "Authorized Absence":
            day_earnings = 0.0
        elif status == "Unauthorized Absence":
            penalty = daily_rate * 0.5
            day_earnings = -penalty

    if consecutive_work_days == 7:
        day_bonus = (daily_rate * bonus_rule)
        day_earnings += day_bonus
        consecutive_work_days = 0

    total_pay += day_earnings
    total_all_bonus += day_bonus
    total_all_penalty += penalty
    total_hours_worked += hours_worked
    total_overtime_hours += overtime_hours
    
    info_label = status

    final_attendance_profile[day_num] = {
        "DateText": date_string, "ClockIn": c_in, "LunchOut": l_out, "LunchIn": l_in, "ClockOut": c_out,
        "HoursWorked": f"{hours_worked:.1f}", "Overtime": f"{overtime_hours:.1f}", "Bonus": f"{day_bonus:.2f}",
        "Penalty": f"{penalty:.2f}", "StatusText": info_label, "NetEarnings": f"{day_earnings:.2f}"
    }
    
    calculated_preview_rows.append({
        "Date": date_string, "In": c_in, "Out": c_out, "Hours": f"{hours_worked:.1f} hrs",
        "O.T.": f"{overtime_hours:.1f} hrs", "Day Net Total": f"{day_earnings:.2f} DA"
    })

df_financial_verification = pd.DataFrame(calculated_preview_rows)
st.dataframe(df_financial_verification, use_container_width=True, height=380, hide_index=True)

net_salary = total_pay - advance_pay
col_m1, col_m2, col_m3 = st.columns(3)
with col_m1: st.metric("Gross Earnings (Inc. Bonuses)", f"{total_pay:.2f} DA")
with col_m2: st.metric("Advances / Deductions", f"-{advance_pay:.2f} DA")
with col_m3: st.metric("Net Salary Pay-Out", f"{net_salary:.2f} DA")

# ----------------------------------------------------
# EXPORT EXECUTIVE LEDGER WITH EDGE-PINNED STYLING
# ----------------------------------------------------
st.markdown("---")
st.subheader("5. Export Final Corporate Ledger Document")

if st.button("Compile & Print Final PDF Statement", type="primary"):
    pdf_rows = []
    for d in sorted(final_attendance_profile.keys()):
        p = final_attendance_profile[d]
        pdf_rows.append([
            p["DateText"], p["ClockIn"], p["LunchOut"], p["LunchIn"], p["ClockOut"],
            p["HoursWorked"], p["Overtime"], p["Bonus"], p["Penalty"], p["NetEarnings"]
        ])
        
    filename = f"Corporate_Payroll_{selected_employee_key.replace(' ', '_')}_{target_period.replace('/', '_')}.pdf"
    
    doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"), spaceAfter=5)
    section_style = ParagraphStyle('SectionTitle', parent=styles['Heading2'], fontSize=13, textColor=colors.HexColor("#1A365D"), spaceBefore=20, spaceAfter=12)
    normal_style = styles['Normal']
    
    label_style = ParagraphStyle('ProfileLabel', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#4A5568"))
    val_style = ParagraphStyle('ProfileVal', fontName='Helvetica', fontSize=10, textColor=colors.black)
    
    matrix_header_style = ParagraphStyle('MatrixHeader', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor("#1A365D"))
    matrix_label_style = ParagraphStyle('MatrixLabel', fontName='Helvetica', fontSize=10.5, textColor=colors.black)
    matrix_bold_style = ParagraphStyle('MatrixBold', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#16A34A"))

    def build_header_block():
        if logo_path and os.path.exists(logo_path):
            img = Image(logo_path, width=50, height=50)
            h_table = Table([[img, Paragraph(f"<b>{company_name}</b><br/>Official Salary & Attendance Ledger", title_style)]], colWidths=[65, 485])
        else:
            h_table = Table([[Paragraph(f"<b>{company_name}</b><br/>Official Salary & Attendance Ledger", title_style)]], colWidths=[550])
        h_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        return h_table

    # ====================================================
    # PAGE 1: DOMINANT ACCOUNTING MATRIX & EDGE SIGNATURES
    # ====================================================
    story.append(build_header_block())
    story.append(Spacer(1, 15))
    story.append(Paragraph("<b>Employee Profile Details</b>", section_style))
    
    emp_info_data = [
        [Paragraph("Employee Target Identity:", label_style), Paragraph(selected_employee_key, val_style)],
        [Paragraph("Statement Accounting Period:", label_style), Paragraph(f"{selected_month_name} {selected_year} ({target_period})", val_style)],
        [Paragraph("Contractual Base Salary Standard:", label_style), Paragraph(f"{base_salary:.2f} DA", val_style)],
        [Paragraph("Metrics Profile Standard:", label_style), Paragraph(f"{shift_hours} Hours Per Shift / 9:00 AM Entry", val_style)]
    ]
    info_table = Table(emp_info_data, colWidths=[180, 370])
    info_table.setStyle(TableStyle([
        ('PADDING', (0,0), (-1,-1), 6), ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story.append(info_table)
    story.append(Spacer(1, 20))
    
    story.append(Paragraph("<b>Financial Accounting Summary Matrix</b>", section_style))
    summary_data = [
        [Paragraph("Payroll Component Matrix", matrix_header_style), Paragraph("Statement Metric Value", matrix_header_style), Paragraph("Payroll Component Matrix", matrix_header_style), Paragraph("Statement Metric Value", matrix_header_style)],
        [Paragraph("Calculated Hourly Rate", matrix_label_style), Paragraph(f"{hourly_rate:.2f} DA", matrix_label_style), Paragraph("Total Attendance Penalty", matrix_label_style), Paragraph(f"{total_all_penalty:.2f} DA", matrix_label_style)],
        [Paragraph("Total Work Hours Logged", matrix_label_style), Paragraph(f"{total_hours_worked:.1f} Hrs", matrix_label_style), Paragraph("Base Advance Reductions", matrix_label_style), Paragraph(f"{advance_pay:.2f} DA", matrix_label_style)],
        [Paragraph("Total Overtime Logged", matrix_label_style), Paragraph(f"{total_overtime_hours:.1f} Hrs", matrix_label_style), Paragraph("Total Statement Deductions", matrix_label_style), Paragraph(f"{(total_all_penalty+advance_pay):.2f} DA", matrix_label_style)],
        [Paragraph("Accumulated Monthly Bonus", matrix_label_style), Paragraph(f"{total_all_bonus:.2f} DA", matrix_label_style), Paragraph("Final Net Payout (DZD)", matrix_header_style), Paragraph(f"{net_salary:.2f} DA", matrix_bold_style)]
    ]
    
    summary_table = Table(summary_data, colWidths=[165, 110, 165, 110])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EDF2F7")),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 16),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (2,4), (3,4), colors.HexColor("#F0FDF4")), 
        ('BOX', (2,4), (3,4), 2, colors.HexColor("#16A34A"))
    ]))
    story.append(summary_table)
    
    # Precise Spacer pushes the verification cards right onto the bottom baseline grid 
    story.append(Spacer(1, 180))
    
    sig_label_style = ParagraphStyle('SigLabel', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#2D3748"))
    sub_text_style = ParagraphStyle('SubText', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#718096"))
    
    sig_data = [
        [Paragraph("Prepared By:", sig_label_style), Paragraph("Verified & Approved By:", sig_label_style)],
        [Spacer(1, 45), Spacer(1, 45)],
        [Paragraph("....................................................<br/>HR Operations Representative", sub_text_style), 
         Paragraph("....................................................<br/>Responsible Director Signature", sub_text_style)]
    ]
    sig_table = Table(sig_data, colWidths=[275, 275])
    sig_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('VALIGN', (0,0), (-1,-1), 'BOTTOM'), ('PADDING', (0,0), (-1,-1), 4)
    ]))
    story.append(sig_table)
    
    story.append(PageBreak())
    
    # ====================================================
    # PAGE 2: ITEMIZED DAILY LOG DATASTREAM
    # ====================================================
    story.append(Paragraph(f"<b>Itemized Daily Biometric Log — Staff: {selected_employee_key}</b>", section_style))
    story.append(Spacer(1, 4))
    
    table_header = ["Date", "In", "Lunch Out", "Lunch In", "Out", "Hrs Worked", "O.T.", "Bonus", "Penalty", "Net Total"]
    table_data = [table_header] + pdf_rows
    
    report_table = Table(table_data, colWidths=[72, 40, 48, 48, 40, 55, 40, 50, 52, 105])
    report_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 8.5),
        ('FONTSIZE', (0,1), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 3.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3.5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
    ]))
    story.append(report_table)
    story.append(Spacer(1, 12))
    
    bottom_summary_data = [
        [Paragraph("<b>Calculated Hourly Salary:</b>", normal_style), Paragraph(f"{hourly_rate:.2f} DA / Hr", normal_style), Paragraph("<b>Final Net Salary Payout:</b>", ParagraphStyle('B', fontName='Helvetica-Bold', fontSize=9.5)), Paragraph(f"<b>{net_salary:.2f} DA</b>", ParagraphStyle('G', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor("#16A34A")))]
    ]
    bottom_table = Table(bottom_summary_data, colWidths=[150, 120, 140, 140])
    bottom_table.setStyle(TableStyle([
        ('BACKGROUND', (2,0), (3,0), colors.HexColor("#F0FDF4")), ('PADDING', (0,0), (-1,-1), 8), ('BOX', (2,0), (3,0), 1.2, colors.HexColor("#16A34A")) 
    ]))
    story.append(bottom_table)
    
    doc.build(story)
    
    st.success(f"? Premium 2-Page Ledger Compiled Successfully: '{filename}'")
    
    with open(filename, "rb") as pdf_file:
        st.download_button(
            label="?? Download Upgraded Executive PDF Statement",
            data=pdf_file,
            file_name=filename,
            mime="application/pdf"
        )
