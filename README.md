# Smart-Class-Planning

The Smart Class Planning System is a desktop application that helps users create a recommended class schedule based on their DegreeWorks requirements, completed courses, prerequisites, and available course offerings.

## Description
The Smart Class Planning Tool is a Python desktop application designed to assist students with academic course planning. The program uses a student's DegreeWorks PDF, prerequisite information, and course availability data to create a recommended semester-by-semester class plan. The application checks prerequisites and course offerings to help prevent students from being scheduled for courses they are not yet eligible to take. After generating a plan, the application displays the recommended schedule and can automatically save the plan as an Excel file. The project is designed to run locally through Visual Studio 2022 on Windows.

## Getting-Started
### Dependencies
- Windows 10 or later
- Visual Studio 2022
- Python 3.14
- Python packages
  - pypdf
  - networkx
  - matplotlib
  - openpyxl

Install the required packages with:
py -3.14 -m pip install pypdf networkx matplotlib openpyxl

### Installing
1. Download or clone the Smart Class Planning Tool project.
2. Open the project folder in Visual Studio 2022.
3. Make sure the following files are in the same project folder:
   - desktop_app.py
   - ui_backend.py
   - planning_upgrade.py
   - degree_requirements.json
   - main.py
   - planner.py
   - pdf_reader.py
   - course_extractor.py
   - dag_creator.py
   - class_schedule.py
   - prereq courses.txt
   - Sample Input1.pdf
4. Make sure course_schedule.csv is located in the data folder.
5. Set desktop_app.py as the Startup File in Visual Studio 2022.

## Executing-Program
1. Open the project in Visual Studio 2022.
2. Set desktop_app.py as the Startup File.
3. Press Ctrl + F5.
4. When the application opens, select Computer Science Software Systems.
5. Select Sample Input1.pdf as the DegreeWorks file.
6. Click Generate Class Plan.
7. Review the recommended schedule shown in the application.
8. The application will automatically create an Excel file containing the class plan.

Ctrl + F5

The upgraded planning logic is activated through desktop_app.py. The original main.py continues to use the team's original planner logic.

## Help
If the program does not run:
- Make sure desktop_app.py is set as the Startup File.
- Make sure all required project files are in the correct folder.
- Make sure prereq courses.txt and Sample Input1.pdf are present.
- Make sure data/course_schedule.csv is in the data folder.
- Check that the required Python packages are installed.

If dependencies need to be installed again, run:
py -3.14 -m pip install pypdf networkx matplotlib openpyxl

The included course_schedule.csv is a small test dataset and does not represent live course availability. Therefore, some courses may appear under 'Unscheduled' when using Sample Input1.pdf.

The application does not verify live enrollment, earned grades, transfer credits, course substitutions, or graduation eligibility. The generated plan should be reviewed with official degree requirements before being used as an actual graduation plan.

## Authors
- Sharmelle Bourne
- Kenye Brooks
- Nia Prier

