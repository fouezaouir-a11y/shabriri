import os
import re
import calendar
import pandas as pd
import streamlit as st
from datetime import datetime

# PDF Document Generation Engine Libraries
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
# PDF Ingestion & Scraping Engine Libraries
from pypdf import PdfReader

# ----------------------------------------------------
# GLOBAL BROWSER CONFIGURATION & INTERFACE THEME
# ----------------------------------------------------
st.set_page_config(
    page_title="Executive Payroll Matrix & Biometric Parser",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Injector to polish the Streamlit data views
st.markdown("""
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stMetric { background-color: #f8fafc; padding: 1rem; border-radius: 0.5rem; border: 1px solid #e2e8f0; }
    div.stDataFrame div[data-testid="stTable"] { font-family: monospace; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 Universal Attendance Parsing & Payroll Engine")
st.markdown("##### *Production Grade Operational Ledger Framework for Tech Den & DJ Freeshop*")
st.markdown("---")

# ----------------------------------------------------
# SIDEBAR CONTROL PANEL & BUSINESS LOGIC SETTINGS
# ----------------------------------------------------
st.sidebar.image("https://img.icons8.com/fluent/96/000000/payroll.png", width=80)
st.sidebar.header("🏢 Corporate Workspace Profile")
company_name = st.sidebar.text_input("Active Business Entity", "Tech Den & DJ Freeshop")
uploaded_logo = st.sidebar.file_uploader("Corporate Logo Identity (Brand Header)", type=["png", "jpg", "jpeg"])

st.sidebar.markdown("---")
st.sidebar.header("📅 Timeline Accounting Frame")

months_list = [
    ("January", "01"), ("February", "02"), ("March", "03"), ("April", "04"),
    ("May", "05"), ("June", "06"), ("July", "07"), ("August", "08"),
    ("September", "09"), ("October", "10"), ("November", "11"), ("December", "12")
]

# Synchronize selector to current month automatically
current_system_date = datetime.now()
selected_month_name, selected_month_code = st.sidebar.selectbox(
    "Target Payroll Month", 
    months_list, 
    index=current_system_date.month - 1, 
    format_func=lambda x: x[0]
)

# FOREVER FUTURE-PROOF TIMELINE CORE
# Tracks current calendar year dynamically and provides an offset balance of ±5 years
active_current_year = current_system_date.year
dynamic_years_range = [str(year_item) for year_item in range(active_current_year - 5, active_current_year + 6)]
selected_year = st.sidebar.selectbox(
    "Target Accounting Year", 
    dynamic_years_range, 
    index=dynamic_years_range.index(str(active_current_year))
)

target_period = f"{selected_month_code}/{selected_year}"

# Resolve calendar attributes for the targeted period
year_int = int(selected_year)
month_int = int(selected_month_code)
num_days_in_month = calendar.monthrange(year_int, month_int)[1]

st.sidebar.markdown("---")
st.sidebar.header("⏱️ Operational Metrics & Shift Constraints")

shift_start_time = st.sidebar.time_input(
    "Contractual Shift Target (IN)", 
    value=datetime.strptime("09:00", "%H:%M").time()
)
shift_start_str = shift_start_time.strftime("%H:%M")

shift_hours = st.sidebar.radio(
    "Paid Shift Window Definition", 
    [8, 9], 
    index=0,
    help="Total expected daily hours allocation, including standard built-in breaks."
)

st.sidebar.markdown("**Streak Reward Configurations**")
enable_streak_bonus = st.sidebar.checkbox("Enable 7/7 Continuous Labor Streak Reward", value=True)

if enable_streak_bonus:
    bonus_rule = st.sidebar.radio(
        "Streak Reward Scale", 
        [1.0, 0.5], 
        index=0,
        format_func=lambda multiplier_val: f"+{multiplier_val} Days Base Wage Allocation"
    )
else:
    st.sidebar.caption("⚠️ 7/7 Day Streak Bonus is currently disabled (Evaluated at 0 DA)")
    bonus_rule = 0.0

# Handle physical workspace logo buffering
logo_path = None
if uploaded_logo:
    try:
        with open(uploaded_logo.name, "wb") as logo_buffer:
            logo_buffer.write(uploaded_logo.getbuffer())
        logo_path = uploaded_logo.name
    except Exception as logo_err:
        st.sidebar.error(f"Logo Buffer Exception: {logo_err}")

# Calendar date labeling framework engine
def get_calendar_label(day_index, month_index, year_index):
    try:
        built_date = datetime(year_index, month_index, day_index)
        return f"{day_index:02d}/{month_index:02d} ({built_date.strftime('%A')})"
    except Exception:
        return f"{day_index:02d}/{month_index:02d} (Unknown)"

# ----------------------------------------------------
# UNIVERSAL PARSING CORE (4-WAY FLEX EXTRACTION LOGIC)
# ----------------------------------------------------
def analyze_master_biometric_log(uploaded_file, target_month_str):
    """
    Highly elastic data-scraping matrix core. Ingests raw file formats, parses lines,
    applies flex regex loops to match names, single/double dates, and normalizes log blocks.
    """
    master_database = {}
    if uploaded_file is None:
        return master_database

    filename = uploaded_file.name.lower()
    target_m, target_y = target_month_str.split('/')
    
    # --- METHOD A: EXCEL MACHINE INTERFACES (.XLSX / .XLS) ---
    if filename.endswith(('.xlsx', '.xls')):
        try:
            if filename.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file, engine='openpyxl')
            else:
                df = pd.read_excel(uploaded_file, engine='xlrd')
                
            # Clean structural row headers to lowercase to strip variant machine spacing anomalies
            df.columns = [str(header_item).strip().lower() for header_item in df.columns]
            
            # Search for headers using loose phrase matching
            name_col = next((c for c in df.columns if any(k in c for k in ['name', 'employee', 'nom', 'user', 'id', 'person', 'employe', 'id count'])), None)
            date_col = next((c for c in df.columns if any(k in c for k in ['date', 'time', 'punch', 'horaire', 'check', 'mouvement', 'temps', 'p_date'])), None)
            
            # Fallback index maps if names don't align perfectly
            if not name_col or not date_col:
                name_col = df.columns[0] if len(df.columns) > 0 else None
                date_col = df.columns[1] if len(df.columns) > 1 else None

            if name_col and date_col:
                for _, row_item in df.iterrows():
                    emp_raw = str(row_item[name_col]).strip()
                    date_raw = str(row_item[date_col]).strip()
                    
                    if emp_raw.lower() in ['nan', 'null', '', 'none'] or date_raw.lower() in ['nan', 'null', '', 'none']:
                        continue
                    
                    # Flex Regex: Supports 5/6/2026, 05/06/2026, spaces, and timestamp lines perfectly
                    ts_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})", date_raw)
                    if ts_match:
                        extracted_day = int(ts_match.group(1))
                        m_str = f"{int(ts_match.group(2)):02d}"
                        y_str = ts_match.group(3)
                        time_str = f"{int(ts_match.group(4)):02d}:{ts_match.group(5)}"
                        
                        if m_str == target_m and y_str == target_y:
                            emp_name_clean = re.sub(r'[\"\',]', '', emp_raw).strip()
                            if emp_name_clean not in master_database:
                                master_database[emp_name_clean] = {}
                            if extracted_day not in master_database[emp_name_clean]:
                                master_database[emp_name_clean][extracted_day] = []
                            if time_str not in master_database[emp_name_clean][extracted_day]:
                                master_database[emp_name_clean][extracted_day].append(time_str)
            return master_database
        except Exception as excel_err:
            st.error(f"Excel Ingestion Architecture Alert: {excel_err}")
            return master_database

    # --- METHOD B: RAW LOG TEXT MATRIX STREAMS (.TXT) ---
    elif filename.endswith('.txt'):
        try:
            raw_binary_stream = uploaded_file.read()
            decoded_text = raw_binary_stream.decode('utf-8', errors='ignore')
            text_lines = decoded_text.splitlines()
            
            for standard_line in text_lines:
                if not standard_line.strip():
                    continue
                
                ts_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2})", standard_line)
                if ts_match:
                    extracted_day = int(ts_match.group(1))
                    m_str = f"{int(ts_match.group(2)):02d}"
                    y_str = ts_match.group(3)
                    time_str = f"{int(ts_match.group(4)):02d}:{ts_match.group(5)}"
                    
                    if m_str == target_m and y_str == target_y:
                        # Strip timestamp markers to isolate name payloads safely
                        clean_payload = standard_line.replace(ts_match.group(0), "")
                        clean_payload = re.sub(r'[\"\',\t\s\:]+', ' ', clean_payload).strip()
                        emp_name_clean = re.sub(r'\b\d\b', '', clean_payload).strip()
                        
                        if not emp_name_clean or len(emp_name_clean) < 2:
                            backup_identifiers = re.findall(r'\b\d+\b', clean_payload)
                            emp_name_clean = f"Staff ID: {backup_identifiers[0]}" if backup_identifiers else "Unknown Log Identity"
                        
                        if emp_name_clean not in master_database:
                            master_database[emp_name_clean] = {}
                        if extracted_day not in master_database[emp_name_clean]:
                            master_database[emp_name_clean][extracted_day] = []
                        if time_str not in master_database[emp_name_clean][extracted_day]:
                            master_database[emp_name_clean][extracted_day].append(time_str)
            return master_database
        except Exception as text_err:
            st.error(f"Text Database Stream Ingestion Alert: {text_err}")
            return master_database

    # --- METHOD C: ROBUST COMPILER PDF REPORTS (.PDF) ---
    else:
        try:
            pdf_scraping_reader = PdfReader(uploaded_file)
            aggregated_pdf_text = ""
            for page_index in range(len(pdf_scraping_reader.pages)):
                extracted_page_content = pdf_scraping_reader.pages[page_index].extract_text() or ""
                
                # Hidden Break Fix: If the text stream is flattened into one line, force newlines
                if extracted_page_content and "\n" not in extracted_page_content:
                    extracted_page_content = re.sub(r'(\d{1,2}/\d{1,2}/\d{4})', r'\n\1', extracted_page_content)
                aggregated_pdf_text += extracted_page_content + "\n"
                
            normalized_pdf_text = aggregated_pdf_text.encode('utf-8', errors='ignore').decode('utf-8')
            compiled_lines = normalized_pdf_text.split("\n")
            
            for line_item in compiled_lines:
                ts_match = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})\s*(\d{1,2}):(\d{2})", line_item)
                if ts_match:
                    extracted_day = int(ts_match.group(1))
                    m_str = f"{int(ts_match.group(2)):02d}"
                    y_str = ts_match.group(3)
                    time_str = f"{int(ts_match.group(4)):02d}:{ts_match.group(5)}"
                    
                    if m_str == target_m and y_str == target_y:
                        # Clean standard system headers out of data lines dynamically
                        clean_payload = line_item.replace("OUR COMPANY", "").replace("COMPANY", "").replace(ts_match.group(0), "")
                        clean_payload = re.sub(r'[\"\',]', '', clean_payload).strip()
                        emp_name_clean = re.sub(r'\b\d\b', '', clean_payload).strip()
                        
                        if not emp_name_clean or len(emp_name_clean) < 2:
                            backup_identifiers = re.findall(r'\b\d+\b', clean_payload)
                            emp_name_clean = backup_identifiers[0] if backup_identifiers else "Unknown Staff"
                        
                        if emp_name_clean not in master_database:
                            master_database[emp_name_clean] = {}
                        if extracted_day not in master_database[emp_name_clean]:
                            master_database[emp_name_clean][extracted_day] = []
                        if time_str not in master_database[emp_name_clean][extracted_day]:
                            master_database[emp_name_clean][extracted_day].append(time_str)
                            
        except Exception as pdf_err:
            st.error(f"PDF Processing Engine Extraction Alert: {pdf_err}")
            
    # Universal Chronological Array Pass
    for active_emp in list(master_database.keys()):
        for active_day in list(master_database[active_emp].keys()):
            master_database[active_emp][active_day] = sorted(list(set(master_database[active_emp][active_day])))
            
    return master_database

# ----------------------------------------------------
# INTERFACE CONTROL INTERACTION LAYERS
# ----------------------------------------------------
st.subheader("1. Ingestion Interface Node")
uploaded_file = st.file_uploader(
    "Drop active biometric system ledger exports here directly", 
    type=["pdf", "xlsx", "xls", "txt"],
    help="Supports standard cross-platform terminal drops."
)

# Parse file data instantly into reactive memory
parsed_master_db = analyze_master_biometric_log(uploaded_file, target_period)

st.markdown("---")
st.subheader("2. Target Employee Workspace Identity")

resolved_default_selection = ""
if parsed_master_db:
    resolved_default_selection = sorted(list(parsed_master_db.keys()))[0]

selected_employee_key = st.text_input("👤 Active Employee Focus Key", value=resolved_default_selection)

# Individual payment allocation parameters
col_profile_w1, col_profile_w2 = st.columns(2)
with col_profile_w1:
    base_salary = st.number_input("Contract Base Monthly Compensation (DA)", min_value=0.0, value=40000.0, step=500.0)
with col_profile_w2:
    advance_pay = st.number_input("Account Advances / Upfront Reductions (DA)", min_value=0.0, value=0.0, step=500.0)

# Extract logs tied to selection
selected_employee_logs = parsed_master_db.get(selected_employee_key, {})
if not selected_employee_logs and parsed_master_db:
    fallback_key_item = sorted(list(parsed_master_db.keys()))[0]
    selected_employee_logs = parsed_master_db.get(fallback_key_item, {})

# Build baseline profile map for calculations
structured_data = {}
for active_day_idx in range(1, num_days_in_month + 1):
    extracted_punches = sorted(selected_employee_logs.get(active_day_idx, []))
    if extracted_punches:
        time_conversion_format = "%H:%M"
        actual_arrival_dt = datetime.strptime(extracted_punches[0], time_conversion_format)
        expected_arrival_dt = datetime.strptime(shift_start_str, time_conversion_format)
        
        calculated_lateness = int((actual_arrival_dt - expected_arrival_dt).total_seconds() / 60) if actual_arrival_dt > expected_arrival_dt else 0
        punch_path_string = " -> ".join(extracted_punches)
        
        structured_data[active_day_idx] = {
            "Status": "Present",
            "Clock Pointers": punch_path_string,
            "Late Minutes": calculated_lateness,
            "Lateness Status": "Non-Justified" if calculated_lateness > 0 else "N/A"
        }
    else:
        structured_data[active_day_idx] = {
            "Status": "Unauthorized Absence",
            "Clock Pointers": "--",
            "Late Minutes": 0,
            "Lateness Status": "N/A"
        }

# ----------------------------------------------------
# INTERACTIVE DATA MATRIX (SECTION 3)
# ----------------------------------------------------
st.markdown("---")
st.subheader("3. 📋 Live Auditing Grid Matrix")
st.info("💡 You can manually adjust attendance statuses, update biometric punches, or toggle lateness justifications directly in the grid below.")

audited_rows_collector = []
for active_day_idx in range(1, num_days_in_month + 1):
    day_profile_ref = structured_data[active_day_idx]
    date_label_string = get_calendar_label(active_day_idx, month_int, year_int)
    audited_rows_collector.append({
        "Date": date_label_string,
        "Attendance Status": day_profile_ref["Status"],
        "Biometric Punches": day_profile_ref["Clock Pointers"],
        "Lateness Status": day_profile_ref["Lateness Status"]
    })

df_audited_initial = pd.DataFrame(audited_rows_collector)

edited_grid_dataframe = st.data_editor(
    df_audited_initial,
    use_container_width=True,
    height=400,
    disabled=["Date"], 
    column_config={
        "Attendance Status": st.column_config.SelectboxColumn(
            "Attendance Status",
            options=["Present", "Day Off", "Paid Leave", "Authorized Absence", "Unauthorized Absence"],
            required=True
        ),
        "Biometric Punches": st.column_config.TextColumn("Biometric Punch Log Strings (IN -> OUT)", required=True),
        "Lateness Status": st.column_config.SelectboxColumn("Lateness Status", options=["Non-Justified", "Justified", "N/A"], required=True)
    },
    hide_index=True,
    key="production_audit_grid_instance"
)

# ----------------------------------------------------
# BUSINESS COMPLIANCE CALCULATION KERNEL (SECTION 4)
# ----------------------------------------------------
st.markdown("---")
st.subheader("4. 🧮 Granular Financial Statement Computations")

daily_compensation_baseline = base_salary / float(num_days_in_month)
hourly_compensation_baseline = daily_compensation_baseline / shift_hours

calculated_statement_rows = []
final_attendance_profile = {}

accumulated_gross_earnings = 0.0
accumulated_bonuses = 0.0
accumulated_penalties = 0.0
accumulated_hours_worked = 0.0
accumulated_overtime_hours = 0.0
rolling_consecutive_days_worked = 0

for structural_idx, row_item in edited_grid_dataframe.iterrows():
    day_numerical_key = structural_idx + 1
    date_label_string = row_item["Date"]
    attendance_status = row_item["Attendance Status"]
    lateness_status = row_item["Lateness Status"]
    
    raw_punch_input = str(row_item["Biometric Punches"]).strip()
    if raw_punch_input in ["--", "None", "nan", ""]:
        sanitized_punches = []
    else:
        sanitized_punches = sorted([punch_block.strip() for punch_block in raw_punch_input.split("->") if re.match(r"^\d{1,2}:\d{2}$", punch_block.strip())])
    
    # Map layout pointers based on 4-punch biometric definitions
    pointer_in = sanitized_punches[0] if len(sanitized_punches) >= 1 else "--"
    pointer_lunch_out = sanitized_punches[1] if len(sanitized_punches) >= 3 else "--" 
    pointer_lunch_in = sanitized_punches[2] if len(sanitized_punches) >= 4 else "--"
    pointer_out = sanitized_punches[-1] if len(sanitized_punches) >= 2 else "--"

    calculated_late_minutes = 0
    if attendance_status == "Present" and sanitized_punches:
        try:
            time_conversion_format = "%H:%M"
            actual_arrival_dt = datetime.strptime(sanitized_punches[0], time_conversion_format)
            expected_arrival_dt = datetime.strptime(shift_start_str, time_conversion_format)
            if actual_arrival_dt > expected_arrival_dt:
                calculated_late_minutes = int((actual_arrival_dt - expected_arrival_dt).total_seconds() / 60)
        except:
            pass

    day_hours_total = 0.0
    day_overtime_total = 0.0
    day_net_earnings = 0.0
    day_penalty_total = 0.0
    day_late_penalty_allocation = 0.0
    day_bonus_earned = 0.0
    
    if attendance_status == "Present":
        if len(sanitized_punches) >= 2:
            time_conversion_format = "%H:%M"
            try:
                first_punch_dt = datetime.strptime(sanitized_punches[0], time_conversion_format)
                last_punch_dt = datetime.strptime(sanitized_punches[-1], time_conversion_format)
                day_hours_total = (last_punch_dt - first_punch_dt).total_seconds() / 3600.0
            except:
                day_hours_total = float(shift_hours)
            
            if day_hours_total == 0: 
                day_hours_total = float(shift_hours)
                
            if day_hours_total > shift_hours:
                day_overtime_total = day_hours_total - shift_hours
                regular_contract_hours = shift_hours
            else:
                day_overtime_total = 0.0
                regular_contract_hours = day_hours_total
                
            day_net_earnings = (regular_contract_hours * hourly_compensation_baseline) + (day_overtime_total * hourly_compensation_baseline * 1.51)
        else:
            day_hours_total = float(shift_hours)
            day_net_earnings = daily_compensation_baseline
        
        # Staggered Lateness Penalty Engine Implementation
        if calculated_late_minutes > 15 and lateness_status == "Non-Justified":
            day_late_penalty_allocation += (1.5 * hourly_compensation_baseline)
            remaining_excess_minutes = calculated_late_minutes - 15
            staggered_intervals = (remaining_excess_minutes + 14) // 15
            day_late_penalty_allocation += (staggered_intervals * 1.5 * hourly_compensation_baseline)
            
        day_penalty_total = day_late_penalty_allocation
        day_net_earnings -= day_penalty_total
        rolling_consecutive_days_worked += 1
    else:
        rolling_consecutive_days_worked = 0
        if attendance_status in ["Day Off", "Paid Leave"]:
            day_net_earnings = daily_compensation_baseline
        elif attendance_status == "Authorized Absence":
            day_net_earnings = 0.0
        elif attendance_status == "Unauthorized Absence":
            day_penalty_total = daily_compensation_baseline * 0.5
            day_net_earnings = -day_penalty_total

    # Track 7-day labor streak reward matrix
    if rolling_consecutive_days_worked == 7:
        day_bonus_earned = (daily_compensation_baseline * bonus_rule)
        day_net_earnings += day_bonus_earned
        rolling_consecutive_days_worked = 0

    accumulated_gross_earnings += day_net_earnings
    accumulated_bonuses += day_bonus_earned
    accumulated_penalties += day_penalty_total
    accumulated_hours_worked += day_hours_total
    accumulated_overtime_hours += day_overtime_total

    final_attendance_profile[day_numerical_key] = {
        "DateText": date_label_string, "ClockIn": pointer_in, "LunchOut": pointer_lunch_out, "LunchIn": pointer_lunch_in, "ClockOut": pointer_out,
        "HoursWorked": f"{day_hours_total:.1f}", "Overtime": f"{day_overtime_total:.1f}", "Bonus": f"{day_bonus_earned:.2f}",
        "Penalty": f"{day_penalty_total:.2f}", "StatusText": attendance_status, "NetEarnings": f"{day_net_earnings:.2f}"
    }
    
    calculated_statement_rows.append({
        "Target Date Frame": date_label_string, 
        "Shift IN": pointer_in, 
        "Shift OUT": pointer_out, 
        "Total Hours Logged": f"{day_hours_total:.1f} hrs",
        "Overtime Metrics": f"{day_overtime_total:.1f} hrs", 
        "Computed Row Balance": f"{day_net_earnings:.2f} DA"
    })

df_statement_preview = pd.DataFrame(calculated_statement_rows)
st.dataframe(df_statement_preview, use_container_width=True, height=380, hide_index=True)

final_net_salary_payout = accumulated_gross_earnings - advance_pay

col_m1, col_m2, col_m3 = st.columns(3)
with col_m1: 
    st.metric("Gross Account Balance (Inc. Streaks)", f"{accumulated_gross_earnings:.2f} DA")
with col_m2: 
    st.metric("Total Outstanding Cash Advances", f"-{advance_pay:.2f} DA")
with col_m3: 
    st.metric("Final Statement Net Payout (DZD)", f"{final_net_salary_payout:.2f} DA")

# ----------------------------------------------------
# REPORTLAB EXECUTIVE EXPORT COMPILED ENGINE (SECTION 5)
# ----------------------------------------------------
st.markdown("---")
st.subheader("5. Document Compilation Core")

if st.button("Compile & Sign Executive PDF Document", type="primary"):
    pdf_rows_payload = []
    for day_key_item in sorted(final_attendance_profile.keys()):
        profile_pointer = final_attendance_profile[day_key_item]
        pdf_rows_payload.append([
            profile_pointer["DateText"], profile_pointer["ClockIn"], profile_pointer["LunchOut"], profile_pointer["LunchIn"], profile_pointer["ClockOut"],
            profile_pointer["HoursWorked"], profile_pointer["Overtime"], profile_pointer["Bonus"], profile_pointer["Penalty"], profile_pointer["NetEarnings"]
        ])
        
    compiled_output_filename = f"Corporate_Payroll_{selected_employee_key.replace(' ', '_')}_{target_period.replace('/', '_')}.pdf"
    
    # Initialize Document Matrix Frame Setup
    doc_canvas_template = SimpleDocTemplate(
        compiled_output_filename, 
        pagesize=letter, 
        rightMargin=30, 
        leftMargin=30, 
        topMargin=30, 
        bottomMargin=30
    )
    story_stream_collector = []
    
    # Custom Document Style Matrices
    document_style_sheet = getSampleStyleSheet()
    title_text_style = ParagraphStyle('DocTitle', parent=document_style_sheet['Heading1'], fontSize=18, textColor=colors.HexColor("#1A365D"), spaceAfter=5)
    section_text_style = ParagraphStyle('SectionTitle', parent=document_style_sheet['Heading2'], fontSize=13, textColor=colors.HexColor("#1A365D"), spaceBefore=20, spaceAfter=12)
    standard_text_style = document_style_sheet['Normal']
    
    label_text_style = ParagraphStyle('ProfileLabel', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#4A5568"))
    value_text_style = ParagraphStyle('ProfileVal', fontName='Helvetica', fontSize=10, textColor=colors.black)
    
    matrix_header_cell_style = ParagraphStyle('MatrixHeader', fontName='Helvetica-Bold', fontSize=11, textColor=colors.HexColor("#1A365D"))
    matrix_label_cell_style = ParagraphStyle('MatrixLabel', fontName='Helvetica', fontSize=10.5, textColor=colors.black)
    matrix_bold_cell_style = ParagraphStyle('MatrixBold', fontName='Helvetica-Bold', fontSize=12, textColor=colors.HexColor("#16A34A"))

    def execute_header_generation():
        if logo_path and os.path.exists(logo_path):
            brand_image_element = Image(logo_path, width=50, height=50)
            header_layout_table = Table([[brand_image_element, Paragraph(f"<b>{company_name}</b><br/>Official Salary & Attendance Ledger", title_text_style)]], colWidths=[65, 485])
        else:
            header_layout_table = Table([[Paragraph(f"<b>{company_name}</b><br/>Official Salary & Attendance Ledger", title_text_style)]], colWidths=[550])
        header_layout_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        return header_layout_table

    # ====================================================
    # PAGE 1: STRATEGIC BALANCE FRAME & ACCRUAL ACCOUNTING
    # ====================================================
    story_stream_collector.append(execute_header_generation())
    story_stream_collector.append(Spacer(1, 15))
    story_stream_collector.append(Paragraph("<b>Employee Profile Matrix Details</b>", section_text_style))
    
    employee_metadata_block = [
        [Paragraph("Target Individual Identity Key:", label_text_style), Paragraph(selected_employee_key, value_text_style)],
        [Paragraph("Statement Accounting Window:", label_text_style), Paragraph(f"{selected_month_name} {selected_year} ({target_period})", value_text_style)],
        [Paragraph("Contractual Base Remuneration Standard:", label_text_style), Paragraph(f"{base_salary:.2f} DA", value_text_style)],
        [Paragraph("Metrics Profile Standard Definitions:", label_text_style), Paragraph(f"{shift_hours} Hours Per Assigned Shift / 9:00 AM Entry Boundary", value_text_style)]
    ]
    employee_metadata_table = Table(employee_metadata_block, colWidths=[180, 370])
    employee_metadata_table.setStyle(TableStyle([
        ('PADDING', (0,0), (-1,-1), 6), 
        ('LINEBELOW', (0,0), (-1,-1), 0.5, colors.HexColor("#E2E8F0")), 
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE')
    ]))
    story_stream_collector.append(employee_metadata_table)
    story_stream_collector.append(Spacer(1, 20))
    
    story_stream_collector.append(Paragraph("<b>Financial Summary Accumulation Balance Matrix</b>", section_text_style))
    summary_matrix_block = [
        [Paragraph("Payroll Component Element Matrix", matrix_header_cell_style), Paragraph("Calculated Account Metric", matrix_header_cell_style), Paragraph("Payroll Component Element Matrix", matrix_header_cell_style), Paragraph("Calculated Account Metric", matrix_header_cell_style)],
        [Paragraph("Calculated Labor Hourly Rate", matrix_label_cell_style), Paragraph(f"{hourly_compensation_baseline:.2f} DA", matrix_label_cell_style), Paragraph("Total Accrued Penalty Deductions", matrix_label_cell_style), Paragraph(f"{accumulated_penalties:.2f} DA", matrix_label_cell_style)],
        [Paragraph("Total Cumulative Hours Worked", matrix_label_cell_style), Paragraph(f"{accumulated_hours_worked:.1f} Hrs", matrix_label_cell_style), Paragraph("Outstanding Advance Reductions", matrix_label_cell_style), Paragraph(f"{advance_pay:.2f} DA", matrix_label_cell_style)],
        [Paragraph("Total Overtime Window Accrued", matrix_label_cell_style), Paragraph(f"{accumulated_overtime_hours:.1f} Hrs", matrix_label_cell_style), Paragraph("Total Aggregated Deductions", matrix_label_cell_style), Paragraph(f"{(accumulated_penalties+advance_pay):.2f} DA", matrix_label_cell_style)],
        [Paragraph("Accumulated Continuity Bonuses", matrix_label_cell_style), Paragraph(f"{accumulated_bonuses:.2f} DA", matrix_label_cell_style), Paragraph("Final Net Payout Balance (DZD)", matrix_header_cell_style), Paragraph(f"{final_net_salary_payout:.2f} DA", matrix_bold_cell_style)]
    ]
    
    summary_matrix_table = Table(summary_matrix_block, colWidths=[165, 110, 165, 110])
    summary_matrix_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#EDF2F7")),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E0")),
        ('PADDING', (0,0), (-1,-1), 16),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (2,4), (3,4), colors.HexColor("#F0FDF4")), 
        ('BOX', (2,4), (3,4), 2, colors.HexColor("#16A34A"))
    ]))
    story_stream_collector.append(summary_matrix_table)
    
    # Precise baseline spacer to keep verification cards pinned neatly to the margin floor
    story_stream_collector.append(Spacer(1, 180))
    
    signature_label_style = ParagraphStyle('SigLabel', fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#2D3748"))
    signature_sub_text_style = ParagraphStyle('SubText', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#718096"))
    
    signature_card_layout = [
        [Paragraph("Prepared Internally By:", signature_label_style), Paragraph("Verified & Authorized Globally By:", signature_label_style)],
        [Spacer(1, 45), Spacer(1, 45)],
        [Paragraph("....................................................<br/>Human Resources Operations Controller", signature_sub_text_style), 
         Paragraph("....................................................<br/>Managing Director Executive Signature", signature_sub_text_style)]
    ]
    signature_verification_table = Table(signature_card_layout, colWidths=[275, 275])
    signature_verification_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'LEFT'), 
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'), 
        ('PADDING', (0,0), (-1,-1), 4)
    ]))
    story_stream_collector.append(signature_verification_table)
    
    story_stream_collector.append(PageBreak())
    
    # ====================================================
    # PAGE 2: GRANULAR BIOMETRIC DAILY TIMELINE STREAMS
    # ====================================================
    story_stream_collector.append(Paragraph(f"<b>Itemized Daily Log Ledger Datastream — Target Profile: {selected_employee_key}</b>", section_text_style))
    story_stream_collector.append(Spacer(1, 4))
    
    itemized_table_headers = ["Date", "IN", "Lunch OUT", "Lunch IN", "OUT", "Worked", "O.T.", "Bonus", "Penalty", "Net Row Balance"]
    itemized_matrix_payload = [itemized_table_headers] + pdf_rows_payload
    
    itemized_ledger_table = Table(itemized_matrix_payload, colWidths=[72, 40, 48, 48, 40, 55, 40, 50, 52, 105])
    itemized_ledger_table.setStyle(TableStyle([
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
    story_stream_collector.append(itemized_ledger_table)
    story_stream_collector.append(Spacer(1, 12))
    
    document_footer_summary_block = [
        [Paragraph("<b>Contract Hourly Base Standard:</b>", standard_text_style), Paragraph(f"{hourly_compensation_baseline:.2f} DA / Hr", standard_text_style), Paragraph("<b>Final Net Statement Payout Balance:</b>", ParagraphStyle('B', fontName='Helvetica-Bold', fontSize=9.5)), Paragraph(f"<b>{final_net_salary_payout:.2f} DA</b>", ParagraphStyle('G', fontName='Helvetica-Bold', fontSize=9.5, textColor=colors.HexColor("#16A34A")))]
    ]
    document_footer_summary_table = Table(document_footer_summary_block, colWidths=[150, 120, 140, 140])
    document_footer_summary_table.setStyle(TableStyle([
        ('BACKGROUND', (2,0), (3,0), colors.HexColor("#F0FDF4")), 
        ('PADDING', (0,0), (-1,-1), 8), 
        ('BOX', (2,0), (3,0), 1.2, colors.HexColor("#16A34A")) 
    ]))
    story_stream_collector.append(document_footer_summary_table)
    
    try:
        doc_canvas_template.build(story_stream_collector)
        st.success(f"✅ Executive Payroll Balance Statement compiled successfully: '{compiled_output_filename}'")
        
        with open(compiled_output_filename, "rb") as final_pdf_asset:
            st.download_button(
                label="📥 Download Executive Statement Asset (PDF)",
                data=final_pdf_asset,
                file_name=compiled_output_filename,
                mime="application/pdf"
            )
    except Exception as build_err:
        st.error(f"ReportLab Document Stream Construction Fault: {build_err}")
