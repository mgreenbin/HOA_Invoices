# methods for accessing cells
# 1: read by sheet reference
#ws['G10'].value = "Due Date: 03/31/2020"
#wb.save("Edmondson 2020 - 6935 Woodvale UPDATED.xlsx")
#wb.close()


#2: created a reference to specific cell, use that reference to access row, column, value, etc.
#e = ws['A2']

#print(e.value)

#3: Use row and column numbers via cell
#c = ws.cell(row=2, column=1)