#import excel
#Clean excel data
#Write back to excel

import pandas as pd

excel_workbook = 'students_addresses_2.xlsx'
Students = pd.read_excel(excel_workbook, sheet_name='Students')

print('\t\t\t============== Summary: ==================')

first_name_list=[]
last_name_list=[]

excelnames=Students['Student Name']
print(excelnames)
Students[["First Name", "Last Name"]] = Students["Student Name"].str.split(
    " ", n=1, expand=True
)
all_other_cols = [
    col for col in Students.columns if col not in ["First Name", "Last Name"]
]
Students = Students[["First Name", "Last Name"] + all_other_cols]
del Students['Student Name']
print(Students.head(10))
for name in excelnames:
    first_name, last_name = name.split(' ', 1)
    first_name_list.append(first_name)
    last_name_list.append(last_name)

    print(f'First Name: {first_name}, Last Name: {last_name}')

    Students.to_excel('students_addresses_cleaned.xlsx', index=False)
