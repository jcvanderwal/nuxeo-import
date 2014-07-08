"""This script will be used to check whether the list of files in a excel sheet does 
actually exist, if so the line will be marked green, else it will be marked red"""


import os.path, os ,glob, xlrd, xlwt, csv, fnmatch
from xlrd import open_workbook
from xlwt import Workbook
from xlutils.copy import copy



def cell2string(cell):
     if cell.ctype==xlrd.XL_CELL_EMPTY:
          return ""
     elif cell.ctype==xlrd.XL_CELL_TEXT:
          return str(cell.value.encode('utf-8'))
     elif cell.ctype==xlrd.XL_CELL_NUMBER:
          return str(int(cell.value))
     elif cell.ctype==xlrd.XL_CELL_DATE:
          return str(cell.value)
     elif cell.ctype==xlrd.XL_CELL_BOOLEAN:
          return str(cell.value)
     elif cell.ctype==xlrd.XL_CELL_ERROR:
          return ""
     elif cell.ctype== xlrd.XL_CELL_BLANK:
          return ""

def getmatches(string1, string2, string3, string4, string5, string6):

     if not (string5=="0"):
          string23 = string2+string3
          string2punt3=string2 +"."+string3
          string56 = "."+string5+"."+string6
          lijst = []

          for root, dirs, filenames in os.walk("./"):
               for filename in filenames:
                    pad = os.path.abspath(filename)
                    lijst.append(pad)
          matching = [s for s in lijst if string4 in s]
          for filenames in matching:
               if string1 in filenames:
                    if string23 or string2punt3 in filenames:
                              if string4 in filenames:
                                   if string56 in filenames:
                                        if filenames:
                                             return "gevonden"

all_book_list = glob.glob(os.getcwd()+"/*.xls*")
for excel_file_idx in range(len(all_book_list)):
     path=all_book_list[excel_file_idx]
     book=open_workbook(path)
     writingbook=copy(book)
     found=xlwt.easyxf('pattern: pattern solid, fore_color green;')
     notfound=xlwt.easyxf('pattern: pattern solid, fore_color red;')
     sheet = book.sheet_by_index(0)
     for row_idx in range(2,sheet.nrows):
          row=sheet.row(row_idx)
          string=row[0].value
          sheet2 = writingbook.get_sheet(0)
          if(getmatches(cell2string(row[0]), cell2string(row[1]),cell2string(row[2]), cell2string(row[3]), cell2string(row[4]), cell2string(row[5])))=="gevonden":
            sheet2.write(row_idx, 0, string, found)
          else:
                sheet2.write(row_idx, 0, string, notfound)
     writingbook.save(path+"Checked"+".xls")






