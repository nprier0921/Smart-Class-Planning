from pdf_reader import read_pdf
from course_extractor import extract_courses

def main():
    pdf_file = "Sample Input1.pdf"
    print("DegreeWorks PDF reading.....")
    text = read_pdf(pdf_file)
    print("PDF was read successfully")
    courses = extract_courses(text)
    print("\nThe following courses were found in your DegreeWorks Plan:")

    for course in courses:
        print(course)

if __name__ == "__main__":
    main()
