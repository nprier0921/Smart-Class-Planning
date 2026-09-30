import re

def extract_courses(text):

    pattern = r'\b[A-Z]{2,5}\s?\d{4}\b'

    courses = re.findall(pattern, text)

    courses = list(dict.fromkeys(courses))

    return courses