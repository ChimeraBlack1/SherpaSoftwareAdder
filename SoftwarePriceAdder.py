import math
import xlrd
import xlwt

#open software wb
software = "software.xlsx"
softwareWB = xlrd.open_workbook(software)
softwareSHEET = softwareWB.sheet_by_index(0)

