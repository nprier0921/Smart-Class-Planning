"""DegreeWorks review, availability-aware scheduling, and Excel plan validation."""

import json
import re
from collections import defaultdict
from io import BytesIO
from pathlib import Path

import networkx as nx
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from class_schedule import load_class_schedule, normalize_course_code

COURSE_RE = re.compile(r"\b[A-Z]{2,5}\s*\d{4}[A-Z]?\b", re.IGNORECASE)
TERM_RE = re.compile(r"\b(Spring|Summer|Fall)\s+(20\d{2})\b", re.IGNORECASE)
TERM_ORDER = {"Spring": 0, "Summer": 1, "Fall": 2}
NOT_SELECTED = "Skip / already satisfied"


def course_codes(text):
    # Extracts and standardizes course identifiers while preserving their order.
    found = []
    previous_subject = None
    for match in COURSE_RE.finditer(str(text or "")):
        code = normalize_course_code(match.group())
        subject, number = code.split(" ", 1)
        if subject == "OR" and previous_subject:
            code = f"{previous_subject} {number}"
        elif subject in {"OR", "AND"}:
            continue
        else:
            previous_subject = subject
        if code not in found:
            found.append(code)
    return found


def term_value(semester, year):
    # Converts a semester and year into a sortable chronological value.
    return int(year) * 3 + TERM_ORDER[semester.title()]


def term_label(semester, year):
    # Returns a consistent semester label for the interface and spreadsheet.
    return f"{semester.title()} {int(year)}"


def read_catalog(project_dir):
    # Loads verified credit values and titles from the bundled catalog excerpt.
    path = Path(project_dir) / "degree_requirements.json"
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def read_prerequisites(file_path):
    # Loads directed prerequisite relationships from the team's existing file format.
    graph = nx.DiGraph()
    with Path(file_path).open("r", encoding="utf-8-sig") as handle:
        for line_number, line in enumerate(handle, 1):
            line = line.strip()
            if not line or "->" not in line:
                continue
            parts = line.split("->")
            if len(parts) != 2:
                raise ValueError(f"Invalid prerequisite rule at line {line_number}.")
            before, after = (normalize_course_code(part.strip()) for part in parts)
            if not COURSE_RE.fullmatch(before) or not COURSE_RE.fullmatch(after):
                raise ValueError(f"Invalid course code in prerequisite line {line_number}.")
            graph.add_edge(before, after)
    if not nx.is_directed_acyclic_graph(graph):
        raise ValueError("The prerequisite file contains a cycle. Correct the file before scheduling.")
    return graph


def catalog_data_warnings(prerequisite_graph, catalog):
    # Identifies prerequisite identifiers that resemble different catalog codes.
    catalog_codes = set(catalog.get("courses", {}))
    findings = []
    for code in prerequisite_graph.nodes:
        if code in catalog_codes:
            continue
        base = re.sub(r"[A-Z]$", "", code)
        matches = sorted(candidate for candidate in catalog_codes
                         if re.sub(r"[A-Z]$", "", candidate) == base)
        if matches:
            findings.append(
                f"Prerequisite file lists {code}; supplied degree catalog lists {', '.join(matches)}. "
                "Verify the correct course identifier before relying on this dependency."
            )
    return findings


def parse_degreeworks(text, catalog=None):
    # Groups explicit DegreeWorks requirements and their alternative course options.
    known = (catalog or {}).get("courses", {})
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]
    groups = []
    warnings = []
    for index, line in enumerate(lines):
        if not re.search(r"Still\s+Needed\s*:", line, flags=re.IGNORECASE):
            continue
        title = re.split(r"Still\s+Needed\s*:", line, flags=re.IGNORECASE)[0].strip()
        body = []
        for next_line in lines[index + 1:]:
            if re.search(r"Still\s+Needed\s*:", next_line, flags=re.IGNORECASE):
                break
            if re.match(r"(?:AREA\b|Remark:)", next_line, flags=re.IGNORECASE):
                break
            body.append(next_line)
        body_text = " ".join(body)
        choices = course_codes(" ".join([title, body_text]))
        elective = bool(re.search(r"elective|\d+\s+credits\s+in\b", title + " " + body_text, re.IGNORECASE))
        if elective:
            groups.append({"title": title, "options": [], "kind": "elective", "raw": body_text})
            continue
        if not choices:
            groups.append({"title": title, "options": [], "kind": "unresolved", "raw": body_text})
            continue
        choices.sort(key=lambda code: (code not in known, code not in course_codes(title), code))
        groups.append({"title": title, "options": choices, "kind": "required", "raw": body_text})
    if not groups:
        fallbacks = course_codes(text)
        groups = [
            {"title": "Unclassified course: " + code, "options": [code],
             "kind": "unverified", "raw": ""}
            for code in fallbacks
        ]
        warnings.append(
            "No 'Still Needed' sections were detected. Each extracted course must be checked against DegreeWorks before scheduling."
        )
    if not groups:
        raise ValueError("No recognizable course requirements were found in the PDF.")
    completed_suggestions = []
    in_progress_suggestions = []
    for line in lines:
        if re.search(r"\b(?:in\s+progress|currently\s+enrolled)\b", line, re.IGNORECASE):
            in_progress_suggestions.extend(course_codes(line))
        elif re.search(r"\b(?:completed|course\s+satisfied)\b", line, re.IGNORECASE):
            completed_suggestions.extend(course_codes(line))
    return {
        "groups": groups, "warnings": warnings,
        "detected_courses": course_codes(text),
        "suggested_completed": list(dict.fromkeys(completed_suggestions)),
        "suggested_in_progress": list(dict.fromkeys(in_progress_suggestions)),
    }


def validate_selection(review, selections, catalog):
    # Resolves one course per required group and preserves unresolved requirements.
    selected = []
    unresolved = []
    warnings = list(review["warnings"])
    catalog_courses = catalog.get("courses", {})
    for index, group in enumerate(review["groups"]):
        if group["kind"] in ("elective", "unresolved"):
            unresolved.append({"course": "", "title": group["title"],
                               "reason": "Requirement needs an elective or manual selection."})
            continue
        selected_course = selections.get(index, NOT_SELECTED)
        if selected_course == NOT_SELECTED:
            if len(group.get("options", [])) > 1:
                unresolved.append({
                    "course": "",
                    "title": group["title"],
                    "reason": "Alternative required course options need confirmation; use Review DegreeWorks courses."
                })
            continue
        if selected_course not in group["options"]:
            raise ValueError(f"Invalid course selection for '{group['title']}'.")
        if selected_course not in selected:
            selected.append(selected_course)
        if selected_course not in catalog_courses:
            warnings.append(
                f"{selected_course} is not identified as a named requirement in the supplied Software Systems catalog. Verify the student's catalog year."
            )
    return selected, unresolved, list(dict.fromkeys(warnings))


def generate_availability_plan(selected_courses, completed_courses, prereq_graph,
                               offerings, catalog, start_year=2026, max_courses=4):
    # Schedules only courses with exact CSV offerings and satisfied prior-term prerequisites.
    if not 1 <= int(max_courses) <= 8:
        raise ValueError("Maximum courses per term must be between 1 and 8.")
    known = catalog.get("courses", {})
    completed = set(course_codes(" ".join(completed_courses)))
    selected = list(dict.fromkeys(course_codes(" ".join(selected_courses))))
    pending = [course for course in selected if course not in completed]
    offerings_by_course = defaultdict(set)
    for entry in offerings:
        offerings_by_course[normalize_course_code(entry.course_code)].add(
            (entry.semester, int(entry.year))
        )
    start = term_value("Fall", start_year)
    candidate_terms = sorted(
        {(entry.semester, int(entry.year)) for entry in offerings
         if term_value(entry.semester, entry.year) >= start},
        key=lambda pair: term_value(*pair)
    )
    future_priority = {}
    for course in selected:
        future_priority[course] = (
            len(nx.descendants(prereq_graph, course)) if course in prereq_graph else 0
        )
    plan = {}
    planned_term = {}
    for term, year in candidate_terms:
        # Restricts selection to courses available in the current term.
        eligible = []
        for course in pending:
            if course not in known or (term, year) not in offerings_by_course[course]:
                continue
            prerequisites = list(prereq_graph.predecessors(course)) if course in prereq_graph else []
            if all(prereq in completed or (
                prereq in planned_term and planned_term[prereq] < term_value(term, year)
            ) for prereq in prerequisites):
                eligible.append(course)
        eligible.sort(key=lambda code: (-future_priority[code], len(offerings_by_course[code]), code))
        chosen = eligible[:int(max_courses)]
        if chosen:
            plan[term_label(term, year)] = chosen
        for course in chosen:
            pending.remove(course)
            planned_term[course] = term_value(term, year)
    unscheduled = []
    for course in pending:
        reasons = []
        if course not in known:
            reasons.append("Outside the named courses in the selected Software Systems catalog")
        if not any(term_value(term, year) >= start for term, year in offerings_by_course[course]):
            reasons.append("No offering listed at or after the chosen start term")
        prerequisites = list(prereq_graph.predecessors(course)) if course in prereq_graph else []
        external = [p for p in prerequisites if p not in completed and p not in selected]
        if external:
            reasons.append("Unconfirmed earlier prerequisites: " + ", ".join(sorted(external)))
        waiting = [p for p in prerequisites if p in selected and p not in completed and p not in planned_term]
        if waiting:
            reasons.append("Prerequisites still unscheduled: " + ", ".join(sorted(waiting)))
        if prerequisites and not waiting and not external and not reasons:
            reasons.append("No offered term after prerequisites become satisfied")
        if not reasons:
            reasons.append("No available place within the terms listed in the CSV")
        unscheduled.append({"course": course, "title": known.get(course, {}).get("title", "Unverified course"),
                            "reason": "; ".join(reasons)})
    return {"plan": plan, "unscheduled": unscheduled, "planned_term": planned_term,
            "completed": sorted(completed), "candidate_terms": candidate_terms}


def export_plan_bytes(plan_result, catalog, warnings=(), review=None, offering_statuses=None):
    # Creates a formatted workbook with scheduled, unscheduled, and review details.
    wb = Workbook()
    ws = wb.active
    ws.title = "Proposed Plan"
    ws.append(["PRELIMINARY CLASS PLAN - VERIFY WITH ACADEMIC ADVISOR"])
    ws.append(["Semester", "Course Code", "Course Title", "Credits", "Availability"])
    courses = catalog["courses"]
    offering_statuses = offering_statuses or {}
    for semester, group in plan_result["plan"].items():
        for course in group:
            details = courses.get(course, {})
            status = offering_statuses.get((course, semester), "UNVERIFIED")
            availability = ("Verified source record" if status == "VERIFIED" else
                            "DEMO / unverified course offering")
            ws.append([semester, course, details.get("title", "Unverified"),
                       details.get("credits"), availability])
    ws.append(["KNOWN CREDIT TOTAL", None, None,
               sum(courses[code]["credits"] for group in plan_result["plan"].values()
                   for code in group if code in courses), "Excludes unspecified electives/unknown credits"])
    other = wb.create_sheet("Unscheduled")
    other.append(["Course", "Course Title", "Reason"])
    for record in plan_result["unscheduled"]:
        other.append([record["course"], record["title"], record["reason"]])
    notes = wb.create_sheet("Review Notes")
    notes.append(["Notes requiring review"])
    for message in warnings:
        notes.append([message])
    for record in (review or {}).get("unresolved", []):
        notes.append([f"{record['title']}: {record['reason']}"])
    notes.append(["CSV term records are not verified against live university enrollment data."])
    notes.append(["The selected PDF and the supplied 2026 Software Systems catalog may reflect different catalog years."])
    notes.append(["Course completion, grades, alternative/concurrent prerequisites, and graduation eligibility require advisor verification."])
    for sheet in wb.worksheets:
        sheet.freeze_panes = "A3" if sheet == ws else "A2"
        header_row = 2 if sheet == ws else 1
        for cell in sheet[header_row]:
            cell.fill = PatternFill("solid", fgColor="203B66")
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(wrap_text=True)
        sheet.row_dimensions[header_row].height = 30
        widths = ([22, 20, 43, 14, 42] if sheet == ws else [27, 42, 90] if sheet == other else [110])
        for index, width in enumerate(widths, 1):
            sheet.column_dimensions[chr(64 + index)].width = width
        if sheet == notes:
            for row in sheet:
                row[0].alignment = Alignment(wrap_text=True, vertical="top")
                sheet.row_dimensions[row[0].row].height = 34
    ws["A1"].font = Font(size=15, bold=True, color="203B66")
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def parse_existing_excel(excel_bytes):
    # Reads semester/course pairs from either the team's workbook or a simple table.
    if not excel_bytes:
        raise ValueError("Select a nonempty Excel workbook.")
    try:
        wb = load_workbook(BytesIO(excel_bytes), data_only=True, read_only=True)
    except Exception as error:
        raise ValueError(f"Unable to read Excel workbook: {error}") from error
    selected = wb["Proposed Plan"] if "Proposed Plan" in wb.sheetnames else wb.active
    pairs = []
    current_term = None
    for row in selected.iter_rows(values_only=True):
        values = list(row)
        if not values:
            continue
        first = str(values[0] or "").strip()
        second = str(values[1] or "").strip() if len(values) > 1 else ""
        term_match = TERM_RE.fullmatch(first)
        if term_match:
            current_term = term_label(term_match.group(1), term_match.group(2))
        if term_match:
            if COURSE_RE.fullmatch(second) and not TERM_RE.fullmatch(second):
                pairs.append((current_term, normalize_course_code(second)))
        elif current_term and COURSE_RE.fullmatch(first) and not TERM_RE.fullmatch(first):
            pairs.append((current_term, normalize_course_code(first)))
    wb.close()
    if not pairs:
        raise ValueError("No semester/course entries found. Use 'Fall 2026' and 'CPSC 4175' labels in the first two columns.")
    return pairs


def check_existing_plan(excel_bytes, prereq_graph, offerings, catalog, completed_courses=()):
    # Reports prerequisite order violations, missing evidence, and offering discrepancies.
    pairs = parse_existing_excel(excel_bytes)
    completed = set(course_codes(" ".join(completed_courses)))
    positions = defaultdict(list)
    offerings_by_course = defaultdict(set)
    known = catalog.get("courses", {})
    for course_term, course in pairs:
        match = TERM_RE.fullmatch(course_term)
        positions[course].append((course_term, term_value(match.group(1), match.group(2))))
    for item in offerings:
        offerings_by_course[item.course_code].add((item.semester, item.year))
    findings = []
    for course, entries in positions.items():
        if len(entries) > 1:
            findings.append({"level":"Error", "course": course, "message":"Course appears multiple times in the plan."})
        if course not in known:
            findings.append({"level":"Review", "course": course,
                             "message":"Course is not a named requirement in the supplied 2026 Software Systems catalog."})
        prereqs = list(prereq_graph.predecessors(course)) if course in prereq_graph else []
        for term, position in entries:
            match = TERM_RE.fullmatch(term)
            slot = (match.group(1), int(match.group(2)))
            if slot not in offerings_by_course[course]:
                status = "No CSV offering records" if not offerings_by_course[course] else "Not listed for " + term
                findings.append({"level":"Review", "course":course, "message":status + "; verify schedule data."})
            for prerequisite in prereqs:
                if prerequisite in completed:
                    continue
                if prerequisite not in positions:
                    findings.append({"level":"Review", "course":course,
                                     "message":f"Cannot verify completed prerequisite {prerequisite}; not in this plan."})
                elif not any(prior < position for _, prior in positions[prerequisite]):
                    findings.append({"level":"Error", "course":course,
                                     "message":f"Prerequisite {prerequisite} must be completed before {term}."})
    if not findings:
        findings.append({"level":"Pass", "course":"All listed courses",
                         "message":"No violations found in the recorded prerequisite edges and term offerings."})
    return {"pairs": pairs, "findings": findings,
            "errors": sum(f["level"] == "Error" for f in findings),
            "reviews": sum(f["level"] == "Review" for f in findings)}
