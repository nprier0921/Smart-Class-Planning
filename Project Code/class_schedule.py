import csv
from pathlib import Path


# Stores one scheduled course offering.
class CourseOffering:

    def __init__(self, course_code, course_title, semester, year):
        # Stores the standardized course code.
        self.course_code = course_code

        # Stores the official course title.
        self.course_title = course_title

        # Stores the semester when the course is offered.
        self.semester = semester

        # Stores the year when the course is offered.
        self.year = year

    def __str__(self):
        # Returns a readable representation of the course offering.
        return (
            f"{self.course_code} - {self.course_title} "
            f"({self.semester} {self.year})"
        )


# Removes unnecessary spaces and standardizes course codes.
def normalize_course_code(course_code):

    # Removes spaces from the beginning and end.
    course_code = course_code.strip()

    # Converts letters to uppercase.
    course_code = course_code.upper()

    # Separates the department code from the course number.
    parts = course_code.split()

    # Returns an already separated course code.
    if len(parts) == 2:
        return f"{parts[0]} {parts[1]}"

    # Handles course codes entered without a space.
    letters = ""
    numbers = ""

    for character in course_code:

        # Collects alphabetic characters for the department code.
        if character.isalpha():
            letters += character

        # Collects numeric and remaining course-number characters.
        else:
            numbers += character

    # Creates a consistent department-and-number format.
    return f"{letters} {numbers}".strip()


# Validates supported semester names.
def validate_semester(semester):

    # Defines semester names accepted by the planning system.
    valid_semesters = [
        "Spring",
        "Summer",
        "Fall"
    ]

    # Standardizes capitalization.
    semester = semester.strip().title()

    # Rejects unsupported semester values.
    if semester not in valid_semesters:
        raise ValueError(
            f"Invalid semester '{semester}'. "
            "Expected Spring, Summer, or Fall."
        )

    # Returns the standardized semester value.
    return semester


# Validates the year stored in the schedule.
def validate_year(year):

    try:
        # Converts the year from text into an integer.
        year = int(year)

    except ValueError:
        # Reports non-numeric year values.
        raise ValueError(
            f"Invalid year '{year}'. Year must contain only numbers."
        )

    # Rejects unrealistic schedule years.
    if year < 2000 or year > 2100:
        raise ValueError(
            f"Invalid year '{year}'."
        )

    # Returns the validated year.
    return year


# Reads class schedule information from a CSV file.
def load_class_schedule(file_path):

    # Converts the supplied file location into a Path object.
    file_path = Path(file_path)

    # Stops execution when the input file cannot be found.
    if not file_path.exists():
        raise FileNotFoundError(
            f"Class schedule file not found: {file_path}"
        )

    # Stores all successfully parsed course offerings.
    schedule = []

    # Opens the CSV file for reading.
    with open(
        file_path,
        mode="r",
        newline="",
        encoding="utf-8-sig"
    ) as csv_file:

        # Reads each CSV row using the header names.
        reader = csv.DictReader(csv_file)

        # Defines the columns required by the schedule parser.
        required_columns = {
            "course_code",
            "course_title",
            "semester",
            "year"
        }

        # Collects the column names found in the CSV file.
        existing_columns = set(reader.fieldnames or [])

        # Finds any required columns missing from the file.
        missing_columns = required_columns - existing_columns

        # Rejects CSV files with missing required columns.
        if missing_columns:
            raise ValueError(
                "Missing required CSV columns: "
                + ", ".join(sorted(missing_columns))
            )

        # Processes every scheduled course row.
        for row_number, row in enumerate(reader, start=2):

            try:
                # Standardizes the course code.
                course_code = normalize_course_code(
                    row["course_code"]
                )

                # Removes unnecessary spaces from the course title.
                course_title = row["course_title"].strip()

                # Validates and standardizes the semester.
                semester = validate_semester(
                    row["semester"]
                )

                # Validates and converts the year.
                year = validate_year(
                    row["year"]
                )

                # Rejects blank course codes.
                if not course_code:
                    raise ValueError(
                        "Course code cannot be blank."
                    )

                # Rejects blank course titles.
                if not course_title:
                    raise ValueError(
                        "Course title cannot be blank."
                    )

                # Creates a course offering from the validated row.
                offering = CourseOffering(
                    course_code,
                    course_title,
                    semester,
                    year
                )

                # Adds the course offering to the schedule.
                schedule.append(offering)

            except ValueError as error:

                # Reports the CSV row containing invalid data.
                raise ValueError(
                    f"Error in row {row_number}: {error}"
                )

    # Returns every valid course offering.
    return schedule


# Returns every course offered during a specific semester.
def get_courses_by_semester(schedule, semester, year):

    # Standardizes and validates the requested semester.
    semester = validate_semester(semester)

    # Standardizes and validates the requested year.
    year = validate_year(year)

    # Stores courses matching the requested semester.
    matching_courses = []

    # Checks every course in the schedule.
    for offering in schedule:

        # Selects courses matching both semester and year.
        if (
            offering.semester == semester
            and offering.year == year
        ):
            matching_courses.append(offering)

    # Returns all matching course offerings.
    return matching_courses


# Determines whether a specific course is offered during a semester.
def is_course_offered(
    schedule,
    course_code,
    semester,
    year
):

    # Standardizes the requested course code.
    course_code = normalize_course_code(
        course_code
    )

    # Standardizes and validates the requested semester.
    semester = validate_semester(
        semester
    )

    # Standardizes and validates the requested year.
    year = validate_year(
        year
    )

    # Checks every scheduled course offering.
    for offering in schedule:

        # Checks for a matching course, semester, and year.
        if (
            offering.course_code == course_code
            and offering.semester == semester
            and offering.year == year
        ):
            return True

    # Reports that no matching course offering exists.
    return False


# Returns every known semester when a course is offered.
def get_course_offerings(
    schedule,
    course_code
):

    # Standardizes the requested course code.
    course_code = normalize_course_code(
        course_code
    )

    # Stores matching offerings.
    offerings = []

    # Searches the complete schedule.
    for offering in schedule:

        # Selects entries matching the requested course.
        if offering.course_code == course_code:
            offerings.append(offering)

    # Returns all semester offerings for the course.
    return offerings


# Converts the schedule into a dictionary grouped by semester.
def group_schedule_by_semester(schedule):

    # Stores semester groups and their courses.
    grouped_schedule = {}

    # Processes every scheduled course.
    for offering in schedule:

        # Creates a semester label.
        semester_key = (
            f"{offering.semester} {offering.year}"
        )

        # Creates an empty semester list when needed.
        if semester_key not in grouped_schedule:
            grouped_schedule[semester_key] = []

        # Adds the course to the correct semester.
        grouped_schedule[semester_key].append(
            offering
        )

    # Returns the grouped semester schedule.
    return grouped_schedule


# Displays every course in the loaded schedule.
def print_full_schedule(schedule):

    # Groups courses by semester.
    grouped_schedule = group_schedule_by_semester(
        schedule
    )

    # Displays every semester and its courses.
    for semester, courses in grouped_schedule.items():

        print(f"\n{semester}")
        print("-" * len(semester))

        # Displays every course offered during the semester.
        for course in courses:
            print(
                f"{course.course_code} - "
                f"{course.course_title}"
            )


# Displays every known offering for one course.
def print_course_offerings(
    schedule,
    course_code
):

    # Finds all offerings matching the course.
    offerings = get_course_offerings(
        schedule,
        course_code
    )

    # Displays a message when no offering is found.
    if not offerings:
        print(
            f"No scheduled offerings found for "
            f"{normalize_course_code(course_code)}."
        )
        return

    # Displays the matching course offerings.
    for offering in offerings:
        print(offering)


# Runs a basic test when the file executes directly.
if __name__ == "__main__":

    # Defines the location of the third input file.
    schedule_file = (
        "data/course_schedule.csv"
    )

    try:
        # Loads and validates the class schedule.
        class_schedule = load_class_schedule(
            schedule_file
        )

        # Displays confirmation after successful loading.
        print("Class schedule loaded successfully.")

        # Displays the complete schedule.
        print_full_schedule(
            class_schedule
        )

        # Tests whether a course is offered in a semester.
        offered = is_course_offered(
            class_schedule,
            "CPSC 4175",
            "Fall",
            2026
        )

        # Displays the result of the availability test.
        print(
            "\nCPSC 4175 offered Fall 2026:",
            offered
        )

    except (
        FileNotFoundError,
        ValueError
    ) as error:

        # Displays input or parsing errors.
        print(
            f"Schedule error: {error}"
        )
