"""Desktop application adapter for the existing PDF and schedule input modules."""

from io import BytesIO
import csv
from pathlib import Path

import networkx as nx

from class_schedule import get_course_offerings, load_class_schedule
from pdf_reader import read_pdf
from planning_upgrade import (
    check_existing_plan,
    catalog_data_warnings,
    course_codes,
    export_plan_bytes,
    generate_availability_plan,
    parse_degreeworks,
    read_catalog,
    read_prerequisites,
    validate_selection,
)


def review_pdf(pdf_bytes, project_dir):
    # Extracts DegreeWorks requirement groups and suggested alternative selections.
    if not pdf_bytes:
        raise ValueError("Select a nonempty DegreeWorks PDF.")
    project_dir = Path(project_dir).resolve()
    catalog = read_catalog(project_dir)
    text = read_pdf(BytesIO(pdf_bytes))
    review = parse_degreeworks(text, catalog)
    review["catalog"] = catalog
    return review


def build_plan(pdf_bytes, project_dir, max_courses=4, start_year=2026,
               selections=None, completed_courses=(), in_progress_courses=()):
    # Combines reviewed PDF requirements with verified local scheduling inputs.
    project_dir = Path(project_dir).resolve()
    prerequisite_file = project_dir / "prereq courses.txt"
    schedule_file = project_dir / "data" / "course_schedule.csv"
    for path, name in ((prerequisite_file, "prerequisite file"),
                       (schedule_file, "class schedule CSV")):
        if not path.is_file():
            raise FileNotFoundError(f"Missing {name}: {path}")
    review = review_pdf(pdf_bytes, project_dir)
    if selections is None:
        raise ValueError("Review and confirm the DegreeWorks course selections before generating a plan.")
    catalog = review["catalog"]
    selected, unresolved, warnings = validate_selection(review, selections, catalog)
    if not selected and not unresolved:
        raise ValueError("No remaining requirements were selected.")
    in_progress = set(course_codes(" ".join(in_progress_courses)))
    completed_confirmed = set(course_codes(" ".join(completed_courses)))
    overlap = in_progress.intersection(completed_confirmed)
    if overlap:
        raise ValueError("Courses cannot be both completed and in progress: " + ", ".join(sorted(overlap)))
    prerequisites = read_prerequisites(prerequisite_file)
    warnings.extend(catalog_data_warnings(prerequisites, catalog))
    schedule = load_class_schedule(schedule_file)
    # Reads the optional verification marker attached to each schedule record.
    offering_statuses = {}
    with schedule_file.open("r", newline="", encoding="utf-8-sig") as handle:
        for entry in csv.DictReader(handle):
            course = entry["course_code"].strip().upper()
            course = " ".join(course.split())
            semester = f"{entry['semester'].strip().title()} {int(entry['year'])}"
            offering_statuses[(course, semester)] = entry.get("verification_status", "UNVERIFIED").upper()
    demo_count = sum(value != "VERIFIED" for value in offering_statuses.values())
    active_selected = [course for course in selected if course not in in_progress]
    result = generate_availability_plan(
        active_selected, completed_courses, prerequisites, schedule, catalog,
        start_year=start_year, max_courses=max_courses
    )
    result["unscheduled"].extend(unresolved)
    for course in selected:
        if course in in_progress:
            result["unscheduled"].append({
                "course": course,
                "title": catalog["courses"].get(course, {}).get("title", "Unverified course"),
                "reason": "Currently in progress; requires a confirmed passing result before dependent courses can be planned."
            })
    if in_progress:
        warnings.append("In-progress courses are not treated as completed prerequisites.")
    if result["unscheduled"]:
        warnings.insert(0, "Some requirements could not be placed using the supplied data. Check the Unscheduled tab.")
    if not result["plan"]:
        warnings.insert(0, "No requested courses could be placed with the currently confirmed prerequisites and CSV term records.")
    if any(course not in prerequisites for course in selected):
        warnings.append("Some selected courses have no entries in the prerequisite file; absence of a recorded rule is not proof of eligibility.")
    if demo_count:
        warnings.append(
            f"The schedule CSV includes {demo_count} DEMO / unverified course-term entries. "
            "They are illustrative planning assumptions, NOT confirmed CSU class offerings. "
            "Verify term dates using the university registration search."
        )
    else:
        warnings.append("Verify the CSV offerings against CSU registration before relying on the plan.")
    warnings.append("The student's DegreeWorks catalog year may differ from the supplied Software Systems track PDF.")
    warnings.append("Minimum grades, OR prerequisites, concurrent enrollment, and complete graduation eligibility are not automatically verified.")
    notes = {"unresolved": unresolved}
    excel_bytes = export_plan_bytes(result, catalog, warnings, notes, offering_statuses)
    # Limits graph display to selected courses and their recorded prerequisite ancestors.
    relevant = set(selected)
    for course in selected:
        if course in prerequisites:
            relevant.update(nx.ancestors(prerequisites, course))
    graph = prerequisites.subgraph(relevant).copy()
    graph.add_nodes_from(selected)
    titles = {item.course_code: item.course_title for item in schedule}
    rows = []
    for semester, courses in result["plan"].items():
        for course in courses:
            item = catalog["courses"].get(course, {})
            rows.append({"semester":semester, "course":course,
                         "title":item.get("title", titles.get(course, "Unknown title")),
                         "credits":item.get("credits"),
                         "availability":("Verified record" if offering_statuses.get((course, semester)) == "VERIFIED"
                                         else "DEMO / not CSU verified"),
                         "listed_terms":", ".join(f"{o.semester} {o.year}" for o in get_course_offerings(schedule, course))})
    return {
        "extracted_courses": review["detected_courses"],
        "selected_courses": selected,
        "graph": graph,
        "schedule": schedule,
        "plan": result["plan"],
        "rows": rows,
        "unscheduled": result["unscheduled"],
        "warnings": list(dict.fromkeys(warnings)),
        "xlsx_bytes": excel_bytes,
        "unresolved_requirements": unresolved,
        "completed_courses": result["completed"],
        "in_progress_courses": sorted(in_progress),
        "catalog": catalog,
    }


def check_excel(excel_bytes, project_dir, completed_courses=()):
    # Validates an uploaded semester plan against recorded prerequisites and offerings.
    project_dir = Path(project_dir).resolve()
    catalog = read_catalog(project_dir)
    prerequisites = read_prerequisites(project_dir / "prereq courses.txt")
    offerings = load_class_schedule(project_dir / "data" / "course_schedule.csv")
    return check_existing_plan(excel_bytes, prerequisites, offerings, catalog, completed_courses)
