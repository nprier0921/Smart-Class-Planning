"""Three-input class planning and Excel prerequisite checking for a desktop application."""

import csv
import json
import re
from collections import defaultdict
from io import BytesIO
from pathlib import Path

import networkx as nx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from class_schedule import load_class_schedule, normalize_course_code
from pdf_reader import read_pdf

COURSE_PATTERN = re.compile(r"\b[A-Z]{2,5}\s*\d{4}[A-Z]?\b", re.IGNORECASE)
SEMESTER_PATTERN = re.compile(r"^(Spring|Summer|Fall)\s+(20\d{2})$", re.IGNORECASE)
TERM_NUMBER = {"Spring": 0, "Summer": 1, "Fall": 2}


def course_code(value):
    # Standardizes course codes for matching across all input files.
    code = normalize_course_code(str(value).strip())
    return code if COURSE_PATTERN.fullmatch(code) else ""


def load_catalog(folder):
    # Loads course titles and credit values for the selected degree.
    file_path = Path(folder) / "degree_requirements.json"
    with file_path.open(encoding="utf-8") as handle:
        return json.load(handle).get("courses", {})


def extract_remaining_courses(pdf_file, catalog):
    # Extracts single-course requirements from DegreeWorks Still Needed sections.
    text = read_pdf(str(pdf_file))
    if not text.strip():
        raise ValueError("No readable text was found in the DegreeWorks PDF.")

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    selected = []
    seen = set()
    ignored_electives = 0

    for index, line in enumerate(lines):
        if "Still Needed:" not in line:
            continue

        title = line.split("Still Needed:", 1)[0].strip()
        next_line = lines[index + 1] if index + 1 < len(lines) else ""

        # Skips elective-credit groups because automatic elective suggestions are outside project scope.
        if ("elective" in title.lower() or
                re.match(r"^\d+\s+credits?\s+in\b", next_line, flags=re.IGNORECASE)):
            ignored_electives += 1
            continue

        # Processes requirements requesting one specific course or one alternative.
        if not re.match(r"^1\s+class\s+in\b", next_line, flags=re.IGNORECASE):
            continue

        choices = [course_code(match.group()) for match in COURSE_PATTERN.finditer(next_line)]
        choices = [choice for choice in choices if choice]
        title_codes = [course_code(match.group()) for match in COURSE_PATTERN.finditer(title)]
        title_codes = [value for value in title_codes if value]

        # Prefers a named catalog requirement when the audit heading contains its course code.
        preferred = next((value for value in title_codes if value in catalog), None)
        selected_code = preferred or (choices[0] if choices else "")
        if selected_code and selected_code not in seen:
            selected.append(selected_code)
            seen.add(selected_code)

    if not selected:
        raise ValueError("No single-course 'Still Needed' requirements were found in the PDF.")
    return selected, ignored_electives


def read_prerequisite_graph(path):
    # Converts prerequisite arrows into a directed acyclic course graph.
    graph = nx.DiGraph()
    with Path(path).open(encoding="utf-8-sig") as source:
        for number, raw_line in enumerate(source, 1):
            if "->" not in raw_line:
                continue
            parts = raw_line.split("->")
            if len(parts) != 2:
                raise ValueError(f"Invalid prerequisite relationship on line {number}.")
            first, second = [course_code(part) for part in parts]
            if not first or not second:
                raise ValueError(f"Invalid prerequisite course code on line {number}.")
            graph.add_edge(first, second)
    if not nx.is_directed_acyclic_graph(graph):
        raise ValueError("Prerequisite data contains a cycle and cannot be scheduled.")
    return graph


def term_index(semester, year):
    # Converts a semester into a chronological number for comparison.
    return int(year) * 3 + TERM_NUMBER[semester.title()]


def generate_plan(folder, pdf_file, start_year=2026, max_courses=4):
    # Combines DegreeWorks, prerequisites, and semester offerings into a class plan.
    folder = Path(folder).resolve()
    catalog = load_catalog(folder)
    remaining, ignored_electives = extract_remaining_courses(pdf_file, catalog)
    graph = read_prerequisite_graph(folder / "prereq courses.txt")
    schedule = load_class_schedule(folder / "data" / "course_schedule.csv")
    max_courses = int(max_courses)
    start_year = int(start_year)
    if not 1 <= max_courses <= 8:
        raise ValueError("Maximum courses per semester must be between 1 and 8.")

    offerings = defaultdict(set)
    descriptions = {}
    for entry in schedule:
        code = course_code(entry.course_code)
        offerings[code].add((entry.semester, entry.year))
        descriptions.setdefault(code, entry.course_title)

    terms = sorted(
        {(entry.semester, entry.year) for entry in schedule
         if term_index(entry.semester, entry.year) >= term_index("Fall", start_year)},
        key=lambda item: term_index(*item),
    )

    # Treats prerequisites absent from the PDF's remaining-course list as previously satisfied.
    pending = set(remaining)
    required_set = set(remaining)
    term_for_course = {}
    plan = {}
    for semester, year in terms:
        current_slot = term_index(semester, year)
        eligible = []
        for code in pending:
            if (semester, year) not in offerings[code]:
                continue
            parents = list(graph.predecessors(code)) if code in graph else []
            if all(parent not in required_set or
                   (parent in term_for_course and term_for_course[parent] < current_slot)
                   for parent in parents):
                eligible.append(code)

        # Prioritizes courses that unlock the greatest number of remaining requirements.
        eligible.sort(key=lambda code: (
            -len(set(nx.descendants(graph, code)).intersection(pending)) if code in graph else 0,
            len(offerings[code]), code,
        ))
        chosen = eligible[:max_courses]
        if not chosen:
            continue
        label = f"{semester} {year}"
        plan[label] = chosen
        for code in chosen:
            pending.remove(code)
            term_for_course[code] = current_slot
        if not pending:
            break

    unscheduled = []
    for code in remaining:
        if code not in pending:
            continue
        missing = [parent for parent in graph.predecessors(code)
                   if parent in required_set and parent in pending] if code in graph else []
        if not offerings[code]:
            reason = "Course has no entry in the class schedule CSV."
        elif missing:
            reason = "Prerequisite not scheduled: " + ", ".join(sorted(missing))
        else:
            reason = "No eligible listed semester remained after prerequisite checking."
        unscheduled.append((code, reason))

    def description(code):
        # Selects a course title and credit value from the supplied degree data.
        item = catalog.get(code, {})
        return item.get("title", descriptions.get(code, "Course")), int(item.get("credits", 3))

    rows = []
    for semester, courses in plan.items():
        for code in courses:
            title, credits = description(code)
            rows.append((semester, code, title, credits))

    return {
        "plan": plan,
        "rows": rows,
        "remaining": remaining,
        "unscheduled": unscheduled,
        "electives_skipped": ignored_electives,
        "graph": graph,
        "schedule": schedule,
        "total_credits": sum(row[3] for row in rows),
    }


def export_plan(result, folder):
    # Creates a formatted Excel workbook automatically in the output folder.
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Class Plan"
    sheet.merge_cells("A1:D1")
    sheet["A1"] = "Smart Class Planning Tool - Recommended Schedule"
    sheet["A1"].font = Font(size=16, bold=True, color="1C3153")
    sheet["A2"] = "Program"
    sheet["B2"] = "Computer Science (B.S.) - Software Systems"
    sheet["A3"] = "Total scheduled credits"
    sheet["B3"] = result["total_credits"]
    sheet.append(["Semester", "Course Code", "Course Title", "Credits"])
    for row in result["rows"]:
        sheet.append(row)

    # Groups and totals the generated schedule by semester.
    totals = workbook.create_sheet("Semester Totals")
    totals.append(["Semester", "Courses", "Total Credits"])
    for semester, courses in result["plan"].items():
        totals.append([semester, len(courses), sum(
            row[3] for row in result["rows"] if row[0] == semester)])

    if result["unscheduled"]:
        unplaced = workbook.create_sheet("Unscheduled")
        unplaced.append(["Course", "Reason"])
        for code, reason in result["unscheduled"]:
            unplaced.append([code, reason])
        unplaced.column_dimensions["A"].width = 21
        unplaced.column_dimensions["B"].width = 80

    for tab in workbook.worksheets:
        header_row = 4 if tab == sheet else 1
        for cell in tab[header_row]:
            cell.fill = PatternFill(fill_type="solid", fgColor="233F6B")
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(vertical="center")
        tab.row_dimensions[header_row].height = 26
        tab.freeze_panes = "A5" if tab == sheet else "A2"
        tab.sheet_view.showGridLines = False
    for col, width in {"A": 20, "B": 20, "C": 52, "D": 14}.items():
        sheet.column_dimensions[col].width = width
    for col, width in {"A": 22, "B": 16, "C": 20}.items():
        totals.column_dimensions[col].width = width

    output_folder = Path(folder) / "output"
    output_folder.mkdir(parents=True, exist_ok=True)
    file_path = output_folder / "student_graduation_plan.xlsx"
    number = 2
    while file_path.exists():
        file_path = output_folder / f"student_graduation_plan_{number}.xlsx"
        number += 1
    workbook.save(file_path)
    return file_path


def read_excel_courses(excel_path):
    # Reads either a table-style Excel plan or semester-heading course lists.
    workbook = load_workbook(excel_path, read_only=True, data_only=True)
    sheet = workbook["Class Plan"] if "Class Plan" in workbook.sheetnames else workbook.active
    course_terms = []
    current_term = None
    for row in sheet.iter_rows(values_only=True):
        a = str(row[0] or "").strip() if row else ""
        b = str(row[1] or "").strip() if len(row) > 1 else ""
        semester = SEMESTER_PATTERN.fullmatch(a)
        if semester:
            current_term = f"{semester.group(1).title()} {semester.group(2)}"
            if course_code(b):
                course_terms.append((current_term, course_code(b)))
        elif current_term and course_code(a):
            course_terms.append((current_term, course_code(a)))
    workbook.close()
    if not course_terms:
        raise ValueError("No courses found. Use columns 'Semester' and 'Course Code', or semester headings followed by course codes.")
    return course_terms


def check_excel_plan(excel_file, folder):
    # Finds courses scheduled before or during the same term as their prerequisites.
    graph = read_prerequisite_graph(Path(folder) / "prereq courses.txt")
    entries = read_excel_courses(excel_file)
    placements = defaultdict(list)
    for term, code in entries:
        match = SEMESTER_PATTERN.fullmatch(term)
        placements[code].append((term, term_index(match.group(1), match.group(2))))

    errors = []
    for code, values in placements.items():
        if len(values) > 1:
            errors.append(f"{code}: listed in more than one semester.")
        if code not in graph:
            continue
        for term, term_number in values:
            for prior in graph.predecessors(code):
                # Prerequisites absent from this workbook are assumed to be completed earlier.
                if prior in placements and not any(
                    earlier_number < term_number for _, earlier_number in placements[prior]
                ):
                    errors.append(f"{code} in {term}: prerequisite {prior} must be in an earlier semester.")
    return entries, errors
