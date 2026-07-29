from openpyxl import Workbook
wb = Workbook()
sheet = wb.active
sheet.title = "MyData"
sheet["A1"] = "Fish"
sheet["B1"] = "Number"

sheet["A2"] = "Clown"
sheet["B2"] = 3

sheet["A3"] = "Puffer"
sheet["B3"] = 1

wb.save("example_sheet_fish.xlsx")