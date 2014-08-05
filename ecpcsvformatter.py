
""" This module provides functions to format Excel files to CSV so that Nuxeo Importer can use them.
If run directly in a directory, it will try to generate csv files for all available Excel files.
A function is also available to importe one file only (formatExcelFile)
This module relies on xlrd that can be found at http://www.python-excel.org/
or just install with pip install xlrd
"""
import os,glob,xlrd,csv

# Function that format every cell value to a string as ints are actually floats in excel
# Based on http://www.lexicon.net/sjmachin/xlrd.html#xlrd.Cell-class
def cell2string(cell):
    if cell.ctype==xlrd.XL_CELL_EMPTY:
        return ""
    elif cell.ctype==xlrd.XL_CELL_TEXT:
        return str(cell.value.encode('utf-8'))
    elif cell.ctype==xlrd.XL_CELL_NUMBER:
        return str(int(cell.value))
    elif cell.ctype==xlrd.XL_CELL_DATE:
        return str(int(cell.value))
    elif cell.ctype==xlrd.XL_CELL_BOOLEAN:
        return str(int(cell.value))
    elif cell.ctype==xlrd.XL_CELL_ERROR:
        return ""
    elif cell.ctype==xlrd.XL_CELL_BLANK:
        return ""

# Main Function to format a single a file and create a CSV file with a specific name
def format_excel_file(path, csv_name):
    DELIMITOR = "."
    ECP_CONTAINER="Subject"
    ECP_FILE="ECP_file"
    AUTO_IMPORT_FILE="0"
    book = xlrd.open_workbook(path)
    sh = book.sheet_by_index(0)

    final_csv = open(csv_name+".csv",'w')
    csv_writer= csv.writer(final_csv,delimiter=',',quotechar='"',quoting=csv.QUOTE_ALL)

    #First Create the header of the csv file
    header_row=[]

    #property
    header_row.append("def:Property")
    #Department
    header_row.append("def:Department")
    #Subject ( only the first 2 characters, otherwise, it defines a sub-subject)
    header_row.append("def:Subject")
    #Sub-Subject (if argomento is greater than 3 characters)
    header_row.append("def:SubSubject")
    #Define date
    header_row.append("def:Date")
    #DepartementSubject (should follow dpt/subject/subsubject)
    header_row.append("def:DepartmentSubject")
    #DocumentKind
    header_row.append("def:DocumentKind")
    #DocumentDescription
    header_row.append("dc:description")
    #document type
    header_row.append("type")
    #document name
    header_row.append("name")
    #auto file import (to know if attached files should be imported or not)
    header_row.append("autofileimport")
    #Document Title
    header_row.append("dc:title")
    #Cadastral Information
    header_row.append("def:Cadastral")
    #Data Room boolean
    header_row.append("def:DataRoom")
    #Document Format
    header_row.append("def:Format")
    #File Location
    header_row.append("def:Location")   
    #Document Note
    header_row.append("def:Note")
    #Document Asset
    header_row.append("def:Asset")
    #Document Number
    header_row.append("def:DocumentNumber")
    #Brand
    header_row.append("def:Brand")

    #Write the header to the csv file
    csv_writer.writerow(header_row)

    for row_idx in range(2,sh.nrows):
        row=sh.row(row_idx)
        csv_row=[]
        #property
        csv_row.append(cell2string(row[0]))

        #Department
        csv_row.append(cell2string(row[1]))
        #Subject ( only the first 2 characters, otherwise, it defines a sub-subject)
        csv_row.append(cell2string(row[2])[:2])
        #Sub-Subject (if argomento is greater than 3 characters)
        csv_row.append(cell2string(row[2])) if len(cell2string(row[2]))>2 else csv_row.append("")
        #Date,  CSV importer needs MM/dd/yyyy, excels files are yyyyMMddxx
        disposable=cell2string(row[3])[:8]
        if len(disposable)<6 and len(disposable) != 0 :
            print "wrong date input on row: ", row_idx
            csv_row.append("")
        elif len(disposable)==6:
            csv_row.append("01/"+disposable[4:6]+"/"+disposable[:4])
            print "disposable"
        else:
            csv_row.append(disposable[4:6]+"/"+disposable[6:8]+"/"+disposable[:4])
        #DepartementSubject (should looks like dpt/subject/subsubject)
        dpt_subj=cell2string(row[1])+"/"+cell2string(row[2])[:2] #dpt_subj is reused for the name
        if len(cell2string(row[2]))>2:
            dpt_subj+="/"+cell2string(row[2])
        csv_row.append(dpt_subj)
        #DocumentKind
        csv_row.append(cell2string(row[7]))
        #DocumentKind
        csv_row.append(cell2string(row[8]))
        #Document type
        if cell2string(row[5])=="1":
            csv_row.append(ECP_FILE)
        else:
            csv_row.append(ECP_CONTAINER)
        #Document Name (dptsubject/Property Departement.IDArgomento.IDDataDocumento.IDDocumento
        disposable=dpt_subj
        if cell2string(row[5])=="1":
            disposable+=("/"+cell2string(row[0])+" "+cell2string(row[1])+"."+cell2string(row[2]))
            disposable+=("."+cell2string(row[3])+"."+cell2string(row[4]))
        csv_row.append(disposable)
        #AutoImport
        csv_row.append(AUTO_IMPORT_FILE)
        #Document title= Property Departement.IDArgomento.IDDataDocumento.IDDocumento for files only empty otherwise
        disposable=""
        if cell2string(row[5])=="1":
            disposable+=(cell2string(row[0])+" "+cell2string(row[1])+"."+cell2string(row[2]))
            disposable+=("."+cell2string(row[3])+"."+cell2string(row[4]))
        csv_row.append(disposable)
        #Cadastral information : a list for each Fgl for each Mappa for each Subalterno
        fgl_list = cell2string(row[12]).replace(" ","").replace("-",",").split(",")
        mappa_list = cell2string(row[13]).replace(" ","").replace("-",",").split(",")
        subalterno_list = cell2string(row[14]).replace(" ","").replace("-",",").split(",")
        disposable=""
        for fgl in fgl_list:
            if mappa_list[0]!="":
                for mappa in mappa_list:
                    if subalterno_list[0]!="":
                        for subalterno in subalterno_list:
                            disposable+=(fgl+DELIMITOR+mappa+DELIMITOR+subalterno+"|")
                    else:
                        disposable+=(fgl+DELIMITOR+mappa+"|")
            else:
                disposable+=(fgl+"|")
        if len(disposable)>0:
            disposable=disposable[:(len(disposable)-1)]
        csv_row.append(disposable)
        #Data room boolean
        csv_row.append("0" if (cell2string(row[25])=="no" or cell2string(row[25])=="")  else "1")
        # Document format
        csv_row.append(cell2string(row[20]))
        #File Location (sede.scaff.faldo)
        disposable=cell2string(row[23])
        disposable+=(DELIMITOR+cell2string(row[22])) if cell2string(row[22])!="" else ""
        disposable+=(DELIMITOR+cell2string(row[21])) if cell2string(row[21])!="" else ""
        csv_row.append(disposable)
        #Document Note
        csv_row.append(cell2string(row[24]))
        #Document Asset
        csv_row.append(cell2string(row[9]))
        #Document Number
        csv_row.append(cell2string(row[18]))
        #Brand
        csv_row.append(cell2string(row[10]))
        #Write the row to the CSV    
        csv_writer.writerow(csv_row)

    #myTestRow=["header1","header3"]
    final_csv.close()

def format_all_excel_files_current_directory():
    all_book_list = glob.glob(os.getcwd()+"/*.xls*")
    for excel_file_idx in range(len(all_book_list)):
        path=all_book_list[excel_file_idx]
        format_excel_file(path,path[path.rfind("/")+1:path.rfind(".")])


if __name__ == "__main__":
     format_all_excel_files_current_directory()

