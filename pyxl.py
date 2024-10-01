import openpyxl

workbook_path = "C:\\Users\\T552140\\Downloads\\Vlan Spreadsheet.xlsx"

workbook = openpyxl.load_workbook(workbook_path)
common_text = {}  # Create an empty dictionary to store common text

for sheet in workbook.worksheets:
    print(sheet)
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell_value = str(cell.value)  # Convert to string

                if cell_value in common_text:
                    common_text[cell_value] += 1
                else:
                    common_text[cell_value] = 1

# Identify and display common text (appears in all sheets)
common_text = {key: value for key, value in common_text.items() if value == len(workbook.worksheets)}

print("Common Text Found:")
for text, count in common_text.items():
    print(text)

