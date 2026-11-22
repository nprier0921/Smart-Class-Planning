import os
from pathlib import Path
from pdf_reader import read_pdf
from course_extractor import extract_courses
from dag_creator import DAGCreator
from class_schedule import load_class_schedule, get_course_offerings
from planner import generatePlan, export_plan_to_excel

# Automatically locate and set the project directory as the working environment
PROJECT_DIR = Path(__file__).resolve().parent
os.chdir(PROJECT_DIR)

def main():
    # Step 1: Input and extract remaining required courses from the student's DegreeWorks PDF
    input1 = input("Enter your DegreeWorks PDF file: ").strip()
    print("\nReading DegreeWorks PDF...")
    text = read_pdf(input1)
    courses = extract_courses(text)
    
    print(f"\nFound {len(courses)} courses in DegreeWorks:")
    for course in courses:
        print(f"  - {course}")

    # Step 2: Build the Directed Acyclic Graph (DAG) for course prerequisites
    input2 = input("\nEnter your prerequisite graph file: ").strip()
    dag_creator = DAGCreator(input2)
    dag_creator.graph_creation(courses)
    graph = dag_creator.get_graph()

    print("\nPrerequisite graph created.")
    print("Number of courses:", graph.number_of_nodes())
    print("Number of prerequisite relationships:", graph.number_of_edges())

    # Step 3: Load the official class schedule CSV to verify term offerings
    input3 = input("\nEnter your class schedule CSV file: ").strip()
    schedule = load_class_schedule(input3)
    print("\nClass schedule loaded successfully.")

    # Cross-reference degree requirements with actual term offerings
    print("\nCourse Availability Check:")
    for course in courses:
        offerings = get_course_offerings(schedule, course)
        if offerings:
            print(f"{course} is offered in: " + ", ".join([f"{o.semester} {o.year}" for o in offerings]))
        else:
            print(f"{course}: No scheduled offerings found.")

    # Step 4: Generate the semester-by-semester graduation plan using Kahn's algorithm
    print("\nGenerating semester-by-semester graduation plan...")
    student_plan = generatePlan(graph, max_courses_per_semester=4)

    print("\nGenerated Graduation Plan:")
    for semester, sem_courses in student_plan.items():
        print(f"\n{semester}:")
        for course in sem_courses:
            print(f"  - {course}")

    # Step 5: Export the finalized plan directly to an Excel document (.xlsx)
    export_plan_to_excel(student_plan, "student_graduation_plan.xlsx")

    # Display the visual representation of the prerequisite graph
    print("\nDisplaying prerequisite graph...")
    dag_creator.show_graph()

if __name__ == "__main__":
    main()