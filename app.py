import streamlit as st
import xlsxwriter
import io

st.set_page_config(page_title="CO Mark Entry Generator", layout="centered")

st.title("Excel Mark Entry Template Generator")
st.write("Define your assessment structure to generate an Excel file with automated CO and Total calculations.")

# Setup assessment parameters
with st.form("setup_form"):
    num_students = st.number_input("Number of Students", min_value=1, max_value=200, value=60)
    num_questions = st.number_input("Number of Questions", min_value=1, max_value=20, value=5)
    
    st.write("Assign Course Outcomes (COs) to each question:")
    cols = st.columns(5)
    co_mapping = {}
    for i in range(1, num_questions + 1):
        with cols[(i-1) % 5]:
            co_mapping[f"Q{i}"] = st.selectbox(f"Q{i} CO", ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"], key=f"q{i}")
            
    submit = st.form_submit_button("Generate Excel Template")

if submit:
    # Identify unique COs used in this assessment
    unique_cos = sorted(list(set(co_mapping.values())))
    
    # Create an in-memory Excel file
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Mark Entry")
    
    # Formatting
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1, 'align': 'center'})
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})
    formula_format = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center'})

    # Write Headers
    headers = ["Roll No", "Student Name"] + [f"{q}\n({co})" for q, co in co_mapping.items()] + [f"Total {co}" for co in unique_cos] + ["Grand Total"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header, header_format)
        worksheet.set_column(col_num, col_num, 15 if col_num == 1 else 10)

    # Write Data Rows with Formulas
    for row in range(1, num_students + 1):
        worksheet.write(row, 0, row, cell_format) # Roll No
        worksheet.write(row, 1, "", cell_format)  # Name placeholder
        
        # Placeholder for question marks
        q_start_col = 2
        q_end_col = q_start_col + num_questions - 1
        for col in range(q_start_col, q_end_col + 1):
            worksheet.write(row, col, "", cell_format)
            
        # Write CO Total Formulas
        co_start_col = q_end_col + 1
        for i, co in enumerate(unique_cos):
            # Find columns that belong to this CO
            co_cols = [q_start_col + j for j, (q, mapped_co) in enumerate(co_mapping.items()) if mapped_co == co]
            
            if co_cols:
                # Build formula like: =SUM(C2, E2, F2)
                cells = [f"{xlsxwriter.utility.xl_col_to_name(c)}{row + 1}" for c in co_cols]
                formula = f"=SUM({','.join(cells)})"
                worksheet.write_formula(row, co_start_col + i, formula, formula_format)
            else:
                worksheet.write(row, co_start_col + i, 0, formula_format)
                
        # Write Grand Total Formula
        total_col = co_start_col + len(unique_cos)
        first_co_cell = f"{xlsxwriter.utility.xl_col_to_name(co_start_col)}{row + 1}"
        last_co_cell = f"{xlsxwriter.utility.xl_col_to_name(total_col - 1)}{row + 1}"
        total_formula = f"=SUM({first_co_cell}:{last_co_cell})"
        
        worksheet.write_formula(row, total_col, total_formula, formula_format)

    workbook.close()
    
    st.success("Template generated successfully!")
    st.download_button(
        label="📥 Download Excel File",
        data=output.getvalue(),
        file_name="CO_Mark_Entry_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
