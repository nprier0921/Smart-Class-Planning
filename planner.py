import openpyxl
import networkx as nx
import os
import platform
import subprocess

def generatePlan(graph, max_courses_per_semester=4, start_semester="Fall", start_year=2026):
    """
    Generates a semester-by-semester plan using Kahn's algorithm 
    with a course limit per semester.
    """
    g = graph.copy()
    
    unmet_prereqs = {node: g.in_degree(node) for node in g.nodes()}
    
    toSchedule = []
    for node in g.nodes():
        if unmet_prereqs[node] == 0:
            toSchedule.append(node)
            
    toSchedule.sort()
    
    academicTerms = ["Fall", "Spring"]
    semIndex = 0
    currentSem = start_year
    
    semester_plan = {}
    
    while toSchedule:
        semester_key = f"{academicTerms[semIndex]} {currentSem}"
        
        currentTerm = []
        tempQueue = []
        
        for course in toSchedule:
            if len(currentTerm) < max_courses_per_semester:
                currentTerm.append(course)
            else:
                tempQueue.append(course)
                
        semester_plan[semester_key] = currentTerm
        
        for course in currentTerm:
            for successor in g.successors(course):
                unmet_prereqs[successor] -= 1
                if unmet_prereqs[successor] == 0:
                    if successor not in tempQueue:
                        tempQueue.append(successor)
                        
        toSchedule = sorted(tempQueue)
        
        semIndex = (semIndex + 1) % len(academicTerms)
        if semIndex == 0:
            currentSem += 1
            
    return semester_plan

def get_course_credits(course_code):
    """
    Assigns standard credit hours based on course numbering conventions.
    """
    if "1301" in course_code or "1302" in course_code or "2105" in course_code:
        return 4
    return 3

def export_plan_to_excel(semester_plan, filename="student_graduation_plan.xlsx", auto_open=True):
    """
    Exports the semester plan into an organized spreadsheet where 
    each semester is grouped with its own headers, courses, and credit totals.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Graduation Plan"
    
    header_fill = openpyxl.styles.PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = openpyxl.styles.Font(color="FFFFFF", bold=True)
    sem_font = openpyxl.styles.Font(size=12, bold=True, color="1F4E78")
    bold_font = openpyxl.styles.Font(bold=True)
    thin_border = openpyxl.styles.Border(
        left=openpyxl.styles.Side(style='thin', color='D9D9D9'),
        right=openpyxl.styles.Side(style='thin', color='D9D9D9'),
        top=openpyxl.styles.Side(style='thin', color='D9D9D9'),
        bottom=openpyxl.styles.Side(style='thin', color='D9D9D9')
    )
    
    current_row = 1
    
    for semester, courses in semester_plan.items():
        ws.cell(row=current_row, column=1, value=semester).font = sem_font
        current_row += 1
        
        headers = ["Course Code", "Credits"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=current_row, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        current_row += 1
        
        totalCreditHours = 0
        for course in courses:
            credits = get_course_credits(course)
            totalCreditHours += credits
            
            c_cell = ws.cell(row=current_row, column=1, value=course)
            cr_cell = ws.cell(row=current_row, column=2, value=credits)
            
            c_cell.border = thin_border
            cr_cell.border = thin_border
            c_cell.alignment = openpyxl.styles.Alignment(horizontal="left")
            cr_cell.alignment = openpyxl.styles.Alignment(horizontal="center")
            
            current_row += 1
            
        total_label_cell = ws.cell(row=current_row, column=1, value="Total Semester Credits")
        total_label_cell.font = bold_font
        total_label_cell.alignment = openpyxl.styles.Alignment(horizontal="right")
        
        total_val_cell = ws.cell(row=current_row, column=2, value=totalCreditHours)
        total_val_cell.font = bold_font
        total_val_cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        
        current_row += 2
        
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 5, 18)
        
    wb.save(filename)
    print(f"\nDegree plan successfully exported and formatted in {filename}")
    
    if auto_open:
        try:
            if platform.system() == "Windows":
                os.startfile(filename)
            elif platform.system() == "Darwin":
                subprocess.run(["open", filename])
            else:
                subprocess.run(["xdg-open", filename])
            print("Opening the formatted Excel file automatically...")
        except Exception as error:
            print(f"Could not auto-open file: {error}.")