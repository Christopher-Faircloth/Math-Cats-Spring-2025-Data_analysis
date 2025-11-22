import add_new_columns
import add_student_grades
import anonymize_names



# import pandas as pd
# original = pd.read_excel("../../data/Check_In_Cleaned.xlsx")
# print("Total rows:", len(original))
# print("Unique NetIDs:", original["NetID"].nunique())
# print("NaN NetIDs:", original["NetID"].isna().sum())
# print(original["NetID"].dtype)


# important note:
# as of 10/25/2025, after merging the checkin file with the student grade report file
# i lose 3 students that are on the check in file and i lose 23 studetns that are 
# on the grade report file
# I go from 876 unique net ids to 850 unique ids

print("Done Cleaning")