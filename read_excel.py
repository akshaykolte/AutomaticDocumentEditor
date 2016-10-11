from xlrd import open_workbook, XL_CELL_EMPTY
from utilities import *
import datetime
from xml.dom.minidom import Text

def view_excel_file(directory_path, excel_file_name):
    data_list = []
    wb = open_workbook(directory_path + excel_file_name)
    for sheet in wb.sheets():
        print 'Sheet:',sheet.name
        if sheet.name.strip() == 'Sheet1':

            unit_details_dict = []
            unit_details = {}

            for row in range(2, sheet.nrows):
                for column in range(sheet.ncols):
                    if sheet.cell(1, column).value != '':
                        print sheet.cell(1,column).value
                        print "---------------------"
                        print sheet.cell(row,column).value
                        print "---------------------"
                        tempText = Text()
                        tempText.data = sheet.cell(row,column).value
                        unit_details[sheet.cell(1,column).value] = tempText.toxml()
                    else:
                        break



                unit_details_dict.append(unit_details)
                unit_details = {}
            print "=====================view_excel_file ends here========================="
            print unit_details_dict
            return unit_details_dict
        else:
            pass

    return unit_details_dict
