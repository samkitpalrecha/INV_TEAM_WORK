import xlwings as xw

final_output = {}

# Load the workbook
workbook = xw.Book('Fin_Statements_Template.xlsx')

# Iterate through the sheets of the workbook
for sheet in workbook.sheets:  # Skip the first sheet
    value = sheet.range('P5').value
    final_output[sheet.name] = value

print(final_output)