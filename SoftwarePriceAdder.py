import math
import xlrd
import xlwt

#open software wb
software = "software.xlsx"
workbook = xlrd.open_workbook(software)
sheet = workbook.sheet_by_index(0)
endOfXL = 100

# write to new workbook 
newWorkbook = xlwt.Workbook()
newSheet = newWorkbook.add_sheet("Software to add")
newWorkbookName = "softwareToAdd.xls"

# globals
thisMapp = "Software"
modelNumber = sheet.cell_value(0,0)
modelName = sheet.cell_value(0,1)

# Write model
newSheet.write(0, 0, "Model")
newSheet.write(0, 3, thisMapp)
newSheet.write(0, 4, modelName)
newSheet.write(0, 5, "Equipment")
newSheet.write(0, 6,  "IT SW")
newSheet.write(0, 11, modelNumber)
newSheet.write(0, 13, modelName)
newSheet.write(0, 14, 0)
newSheet.write(0, 15, 0)
newSheet.write(0, 16, 0)
newSheet.write(0, 17, 0)
newSheet.write(0, 18, 0)
newSheet.write(0, 19, 0)
newSheet.write(0, 20, 0)
newSheet.write(0, 28, "Document Direction Ltd")
newSheet.write(0, 29, 0)
newSheet.write(0, 31, 0)
newSheet.write(0, 32, 0)

for x in range(1, endOfXL):
  try:
    accNumber = sheet.cell_value(x,0)
    accDesc = sheet.cell_value(x,1)
    accName = accDesc
    newSheet.write(x, 0, "Access")
    newSheet.write(x, 3, thisMapp)
    newSheet.write(x, 8, "ACCESSORY")
    newSheet.write(x, 9, "N")
    newSheet.write(x, 11, accNumber)
    newSheet.write(x, 13, accName)
    newSheet.write(x, 14, 0)
    newSheet.write(x, 15, 0)
    newSheet.write(x, 16, 0)
    newSheet.write(x, 17, 0)
    newSheet.write(x, 18, 0)
    newSheet.write(x, 19, 0)
    newSheet.write(x, 20, 0)
    newSheet.write(x, 25, 0)
    newSheet.write(x, 29, accDesc)
    newSheet.write(x, 31, 0)
    newSheet.write(x, 32, 0)
  except:
    break

newWorkbook.save(newWorkbookName)
print("saved: " + str(newWorkbookName))


