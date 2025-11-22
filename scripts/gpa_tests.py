import pandas as pd

file_input = ("../data/Check_In_Cleaned.xlsx")

df = pd.read_excel(file_input)
df = df[df["Final Grade"] == 4]
print("Average GPA: " , df["Final Grade"].mean())