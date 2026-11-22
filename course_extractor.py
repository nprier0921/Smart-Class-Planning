import re

def extract_courses(text):

    pattern =  r'\b[A-Z]{2,5}\s+\d{4}[A-Z]?\b'

    courses = re.findall(pattern, text)

    # Replaces multiple spaces with a single space.
    courses = [
        re.sub(r'\s+', ' ', course).strip()
        for course in courses
    ]

    courses = list(dict.fromkeys(courses))

    return courses