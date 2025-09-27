import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import get_column_letter
from collections import Counter

input_file = "../data/Check_In_Cleaned.xlsx"
output_file = "../results/course_retention.xlsx"
sheetName = "course retention"

def addCourseRetention(original):
    #course specific retention
    uniqueCourses = original["Course Name"].unique()
    rows = []

    for course in uniqueCourses:
        single_course_students = original[original["Course Name"] == course]["NetID"]
        student_counts = single_course_students.value_counts()
        student_count_dict = student_counts.to_dict()
        course_occurance_histogram = Counter(student_count_dict.values())
        course_occurance_histogram = dict(sorted(course_occurance_histogram.items()))

        occurance_chunks = {
            "1": course_occurance_histogram.get(1, 0),
            "2+": sum(v for k, v in course_occurance_histogram.items() if k >= 2),
            "5+": sum(v for k, v in course_occurance_histogram.items() if k >= 5),
            "10+": sum(v for k, v in course_occurance_histogram.items() if k >= 10),
        }
        total_people = single_course_students.nunique()

        retention={
            "Course Name": course,
            "Total People": total_people, 
            "1 Occurance Count": occurance_chunks["1"],
            "2+ Count": occurance_chunks["2+"],
            "2+ percent retention": occurance_chunks["2+"]/total_people,
            "5+ Count": occurance_chunks["5+"],
            "5+ percent retention": occurance_chunks["5+"]/total_people, 
            "10+ Count": occurance_chunks["10+"], 
            "10+ percent retention": occurance_chunks["10+"]/total_people,
        }
        rows.append(retention)

    #all student retention
    all_students = original["NetID"]
    student_counts = all_students.value_counts()
    student_count_dict = student_counts.to_dict()
    occurance_histogram = Counter(student_count_dict.values())
    occurance_histogram = dict(sorted(occurance_histogram.items()))

    occurance_chunks = {
        "1": occurance_histogram.get(1, 0),
        "2+": sum(v for k, v in occurance_histogram.items() if k >= 2),
        "5+": sum(v for k, v in occurance_histogram.items() if k >= 5),
        "10+": sum(v for k, v in occurance_histogram.items() if k >= 10),
    }
    total_people = all_students.nunique()
    retention={
        "Course Name": "All Courses",
        "Total People": total_people, 
        "1 Occurance Count": occurance_chunks["1"],
        "2+ Count": occurance_chunks["2+"],
        "2+ percent retention": occurance_chunks["2+"]/total_people,
        "5+ Count": occurance_chunks["5+"],
        "5+ percent retention": occurance_chunks["5+"]/total_people, 
        "10+ Count": occurance_chunks["10+"], 
        "10+ percent retention": occurance_chunks["10+"]/total_people,
    }
    rows.append(retention)
    return pd.DataFrame(rows)

#creates and columns data frame to add the columns to the excel file
df = pd.DataFrame(columns = ["Course Name", "Total People", "1 Occurance Count", 
                             "2+ Count", "2+ percent retention", "5+ Count", "5+ percent retention", 
                             "10+ Count", "10+ percent retention"])


original = pd.read_excel(input_file)
df = addCourseRetention(original)

df.to_excel(output_file, sheet_name = sheetName, index=False)

# format the excel file
wb = load_workbook(output_file)
ws = wb[sheetName]

for col in ws.columns:
    max_len = 0
    col_letter = get_column_letter(col[0].column)
    for cell in col:
        if cell.value:
            max_len = max(max_len, len(str(cell.value)))
    ws.column_dimensions[col_letter].width = max_len + 2

wb.active = wb.sheetnames.index(sheetName)
wb.save(output_file)
