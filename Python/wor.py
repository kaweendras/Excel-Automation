import openpyxl

files = [] #include paths to your files here
values = []

# Section 2
for file in files:
    wb = openpyxl.load_workbook(file)
    sheet = wb['Sheet1']
    value = sheet['F5'].value
    values.append(value)