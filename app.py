import streamlit as st
import xlsxwriter
import io
import os
import re

# --- PAGE CONFIG ---
st.set_page_config(page_title="CO Mark Generator | St. Dominic's College", layout="wide", page_icon="🏫")

# --- CUSTOM HEADER WITH LOGO ---
col_logo, col_title = st.columns([1, 6])

with col_logo:
    # Checks if logo file exists to prevent app crash if missing
    if os.path.exists("logo.png"):
        st.image("logo.png", width=120) 
    else:
        st.markdown("<div style='height: 120px; width: 120px; border: 2px dashed #ccc; display: flex; align-items: center; justify-content: center;'><small>Add logo.png</small></div>", unsafe_allow_html=True)

with col_title:
    st.markdown("<h1 style='color: #2E4053; margin-top: -15px;'>ST. DOMINIC'S COLLEGE, KANJIRAPALLY</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='color: #5D6D7E; margin-top: -10px;'>Advanced Assessment & CO Mark Entry Generator</h3>", unsafe_allow_html=True)

st.write("<p>Configure sections, 'best of' logic, and CO weightages to generate a multi-sheet Excel tracker.</p>", unsafe_allow_html=True)
st.markdown("---")

# --- UI FOR INPUTS ---
course_name = st.text_input("Name of the Course", placeholder="e.g., CHE101 - Inorganic Chemistry I")

col1, col2 = st.columns(2)
with col1:
    num_students = st.number_input("Number of Students", min_value=1, max_value=200, value=20)
    num_sections = st.number_input("Number of Sections (e.g. Parts A, B, C)", min_value=1, max_value=10, value=3)
with col2:
    num_cos = st.number_input("Number of Course Outcomes (COs)", min_value=1, max_value=12, value=4)

st.markdown("---")

# 1. SECTION CONFIGURATION
st.subheader("1. Section (Parts) Configuration")
sections_data = []
prev_q_end = 0

for i in range(num_sections):
    st.markdown(f"**Section {i+1}**")
    sc1, sc2, sc3, sc4, sc5 = st.columns(5)
    
    # Set default max marks based on section index
    if i == 0:
        default_max = 1.0
    elif i == 1:
        default_max = 5.0
    elif i == 2:
        default_max = 10.0
    else:
        default_max = 5.0 

    default_start = int(prev_q_end + 1)
    
    with sc1:
        sec_name = st.text_input("Section Name", value=chr(65+i), key=f"sec_name_{i}")
    with sc2:
        q_start = st.number_input("Start Q No.", min_value=1, value=default_start, key=f"q_start_{i}")
    with sc3:
        q_end = st.number_input("End Q No.", min_value=int(q_start), value=int(q_start)+4, key=f"q_end_{i}")
    with sc4:
        sec_max = st.number_input("Max Marks PER Question", min_value=1.0, value=default_max, key=f"sec_max_{i}")
    with sc5:
        ans_any = st.number_input("Answer Any (Count)", min_value=1, value=4, key=f"ans_any_{i}")
    
    prev_q_end = q_end
    q_list = [str(q) for q in range(int(q_start), int(q_end) + 1)]
    
    sections_data.append({
        "name": sec_name,
        "questions": q_list,
        "mark_per_q": sec_max,
        "answer_any": ans_any
    })

st.markdown("---")

# 2. CO CONFIGURATION
st.subheader("2. Course Outcome (CO) Mapping & Weightage")
cos_data = {}
for i in range(num_cos):
    cc1, cc2, cc3 = st.columns([1, 1, 3])
    with cc1:
        co_name = st.text_input(f"CO Name", value=f"CO{i+1}", key=f"co_name_{i}")
    with cc2:
        co_weight = st.number_input(f"Weightage for {co_name}", min_value=1.0, value=5.0, key=f"co_w_{i}")
    with cc3:
        co_qs = st.text_input(f"Questions mapped to {co_name} (comma separated)", key=f"co_qs_{i}")
    
    q_list = [q.strip() for q in co_qs.split(",") if q.strip()]
    cos_data[co_name] = {
        "weightage": co_weight,
        "questions": q_list
    }

# --- EXCEL GENERATION ---
if st.button("Generate Excel File", type="primary"):
    display_course = course_name.strip() if course_name.strip() else "Course Assessment"
    
    # Calculate Max Marks for each CO
    q_to_max_mark = {}
    for sec in sections_data:
        for q in sec["questions"]:
            q_to_max_mark[q] = sec["mark_per_q"]
            
    co_max_marks = {}
    for co, data in cos_data.items():
        max_m = sum([q_to_max_mark.get(q, 0) for q in data["questions"]])
        co_max_marks[co] = max_m if max_m > 0 else 1 
        
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    
    # Title & Header Formats
    fmt_title = wb.add_format({'bold': True, 'font_size': 13, 'bg_color': '#2E4053', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    fmt_hdr = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_sec_tot_hdr = wb.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    fmt_grand_hdr = wb.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    
    # Protection Cell Formats (Locked/Unlocked)
    fmt_input = wb.add_format({'border': 1, 'align': 'center', 'locked': False}) 
    fmt_locked = wb.add_format({'border': 1, 'align': 'center', 'locked': True}) 
    fmt_sec_tot_cell = wb.add_format({'bold': True, 'bg_color': '#FCE4D6', 'border': 1, 'align': 'center', 'locked': True})
    fmt_grand_cell = wb.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'locked': True})
    
    # Error & Highlight Formats
    fmt_excess = wb.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
    fmt_error = wb.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'bold': True})

    # Row offsets (Row 0 = Course Banner, Row 1 = Headers, Row 2+ = Student Data)
    hdr_row = 1
    data_start_row = 2

    # --- SHEET 1: MARK ENTRY ---
    ws1 = wb.add_worksheet("1. Mark Entry")
    ws1.protect()
    ws1.set_row(0, 25)
    
    ws1.write(hdr_row, 0, "Roll No", fmt_hdr)
    ws1.write(hdr_row, 1, "Student Name", fmt_hdr)
    ws1.set_column(0, 0, 8)
    ws1.set_column(1, 1, 20)
    
    col_idx = 2
    q_col_mapping = {}
    sec_total_cols = []
    
    for sec in sections_data:
        if not sec["questions"]:
            continue
            
        start_col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
        
        for q in sec["questions"]:
            ws1.write(hdr_row, col_idx, f"Q{q}\n(Max:{sec['mark_per_q']})", fmt_hdr)
            ws1.set_column(col_idx, col_idx, 8)
            q_col_mapping[q] = xlsxwriter.utility.xl_col_to_name(col_idx)
            col_idx += 1
            
        end_col_letter = xlsxwriter.utility.xl_col_to_name(col_idx - 1)
        first_excel_row = data_start_row + 1 # Excel row 3
        last_excel_row = data_start_row + num_students
        range_str = f"{start_col_letter}{first_excel_row}:{end_col_letter}{last_excel_row}"
        
        # Max Mark Error Formatting & Data Validation
        ws1.conditional_format(range_str, {'type': 'cell', 'criteria': '>', 'value': sec['mark_per_q'], 'format': fmt_error})
        ws1.data_validation(range_str, {'validate': 'decimal', 'criteria': '<=', 'value': sec['mark_per_q'], 'error_title': 'Invalid Mark', 'error_message': f'Marks for this section cannot exceed {sec["mark_per_q"]}.'})
        
        # Section Total Header (Wrapped)
        ws1.write(hdr_row, col_idx, f"Part {sec['name']} Total\n(Best {sec['answer_any']})", fmt_sec_tot_hdr)
        ws1.set_column(col_idx, col_idx, 12)
        sec_total_cols.append(xlsxwriter.utility.xl_col_to_name(col_idx))
        
        # Highlight Excess Answers
        if len(sec["questions"]) > sec["answer_any"]:
            ans_any = sec["answer_any"]
            formula = f"=AND({start_col_letter}{first_excel_row}<>\"\", COUNT(${start_col_letter}{first_excel_row}:${end_col_letter}{first_excel_row})>{ans_any}, RANK.EQ({start_col_letter}{first_excel_row},${start_col_letter}{first_excel_row}:${end_col_letter}{first_excel_row}, 1) + COUNTIF(${start_col_letter}{first_excel_row}:{start_col_letter}{first_excel_row}, {start_col_letter}{first_excel_row}) - 1 <= COUNT(${start_col_letter}{first_excel_row}:${end_col_letter}{first_excel_row})-{ans_any})"
            ws1.conditional_format(f"${start_col_letter}{first_excel_row}:${end_col_letter}{last_excel_row}", {'type': 'formula', 'criteria': formula, 'format': fmt_excess})
            
        # Data Rows
        for s in range(num_students):
            r = data_start_row + s
            excel_r = r + 1
            ws1.write(r, 0, "", fmt_input) # Blank Roll No (Unlocked)
            ws1.write(r, 1, "", fmt_input) # Blank Name (Unlocked)
            
            for c in range(col_idx - len(sec["questions"]), col_idx):
                ws1.write(r, c, "", fmt_input) # Mark Entry Cells (Unlocked)
                
            # Best of N Formula (Blank if no data)
            large_parts = [f"IFERROR(LARGE({start_col_letter}{excel_r}:{end_col_letter}{excel_r}, {n}), 0)" for n in range(1, sec["answer_any"] + 1)]
            best_n_formula = f'=IF(COUNT({start_col_letter}{excel_r}:{end_col_letter}{excel_r})=0, "", {"+".join(large_parts)})'
            ws1.write_formula(r, col_idx, best_n_formula, fmt_sec_tot_cell)
            
        col_idx += 1

    # Grand Total Column Sheet 1
    ws1.write(hdr_row, col_idx, "Grand Total", fmt_grand_hdr)
    ws1.set_column(col_idx, col_idx, 12)
    for s in range(num_students):
        r = data_start_row + s
        excel_r = r + 1
        sum_cells = [f"{c}{excel_r}" for c in sec_total_cols]
        if sum_cells:
            sum_args = ",".join(sum_cells)
            ws1.write_formula(r, col_idx, f'=IF(COUNT({sum_args})=0, "", SUM({sum_args}))', fmt_grand_cell)
        else:
            ws1.write(r, col_idx, "", fmt_grand_cell)

    # Merge Course Banner across Sheet 1
    ws1.merge_range(0, 0, 0, col_idx, f"Course: {display_course} — Mark Entry", fmt_title)
    grand_total_col_sheet1 = xlsxwriter.utility.xl_col_to_name(col_idx)

    # --- SHEET 2: CONSOLIDATION ---
    ws2 = wb.add_worksheet("2. CO Consolidation")
    ws2.protect()
    ws2.set_row(0, 25)
    
    ws2.write(hdr_row, 0, "Roll No", fmt_hdr)
    ws2.write(hdr_row, 1, "Student Name", fmt_hdr)
    ws2.set_column(0, 0, 8)
    ws2.set_column(1, 1, 20)
    
    col2_idx = 2
    for co, data in cos_data.items():
        ws2.write(hdr_row, col2_idx, f"{co} Total\n(Max:{co_max_marks[co]})", fmt_hdr)
        ws2.set_column(col2_idx, col2_idx, 10)
        
        ws2.write(hdr_row, col2_idx + 1, f"{co} Wgt.\n(Out of {data['weightage']})", fmt_sec_tot_hdr)
        ws2.set_column(col2_idx + 1, col2_idx + 1, 12)
        
        for s in range(num_students):
            r = data_start_row + s
            excel_r = r + 1
            # Dynamic linking to Sheet 1 Name & Roll No (Blank if empty)
            ws2.write_formula(r, 0, f"=IF('1. Mark Entry'!A{excel_r}=\"\", \"\", '1. Mark Entry'!A{excel_r})", fmt_locked)
            ws2.write_formula(r, 1, f"=IF('1. Mark Entry'!B{excel_r}=\"\", \"\", '1. Mark Entry'!B{excel_r})", fmt_locked)
            
            # CO Raw Total
            cells_to_sum = [f"'1. Mark Entry'!{q_col_mapping[q]}{excel_r}" for q in data["questions"] if q in q_col_mapping]
            if cells_to_sum:
                sum_args = ",".join(cells_to_sum)
                ws2.write_formula(r, col2_idx, f'=IF(COUNT({sum_args})=0, "", SUM({sum_args}))', fmt_locked)
            else:
                ws2.write(r, col2_idx, "", fmt_locked)
                
            # CO Weightage Conversion (Rounded to 2 decimal points)
            raw_cell = f"{xlsxwriter.utility.xl_col_to_name(col2_idx)}{excel_r}"
            weight_formula = f'=IF({raw_cell}="", "", ROUND(({raw_cell}/{co_max_marks[co]})*{data["weightage"]}, 2))'
            ws2.write_formula(r, col2_idx + 1, weight_formula, fmt_sec_tot_cell)
            
        col2_idx += 2
        
    # Final Exam Grand Total pulled from Sheet 1
    ws2.write(hdr_row, col2_idx, "Exam Grand Total", fmt_grand_hdr)
    ws2.set_column(col2_idx, col2_idx, 15)
    for s in range(num_students):
        r = data_start_row + s
        excel_r = r + 1
        ws2.write_formula(r, col2_idx, f"=IF('1. Mark Entry'!{grand_total_col_sheet1}{excel_r}=\"\", \"\", '1. Mark Entry'!{grand_total_col_sheet1}{excel_r})", fmt_grand_cell)

    # Merge Course Banner across Sheet 2
    ws2.merge_range(0, 0, 0, col2_idx, f"Course: {display_course} — CO Consolidation", fmt_title)

    wb.close()
    
    # Sanitize course name for the Excel filename
    safe_course_filename = re.sub(r'[^A-Za-z0-9_\-]', '_', course_name.strip()) if course_name.strip() else "SDC"
    
    st.success("Advanced multi-sheet template generated successfully!")
    st.download_button("📥 Download Excel File", data=output.getvalue(), file_name=f"{safe_course_filename}_CO_Marks.xlsx")

# --- CUSTOM FOOTER ---
st.markdown("<br><br><hr>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #808B96; font-size: 14px;'><strong>Developed by:</strong> Mr. Francis Thomas, Assistant Professor, Department of Chemistry</p>", unsafe_allow_html=True)
