from xlrd import open_workbook, XL_CELL_EMPTY
from utilities import *
import datetime

NEW_LINE_LIST = ['DescriptionofGoods', 'Qty', 'CTHNo', 'CIFValue']

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
                        unit_details[convert_to_xml(sheet.cell(1,column).value)] = convert_to_xml(sheet.cell(row,column).value)
                        if convert_to_xml(sheet.cell(1,column).value) in NEW_LINE_LIST:
                            print '\n\n\n', NEW_LINE_LIST, '\n\n\n'
                            unit_details[convert_to_xml(sheet.cell(1,column).value)] = new_linify(unit_details[convert_to_xml(sheet.cell(1,column).value)]).split('\n')
                            print sheet.cell(1,column).value
                            print "---------------------"
                            print unit_details[convert_to_xml(sheet.cell(1,column).value)]
                            print "---------------------"
                    else:
                        break


                unit_details['SrNo'] = [i+1 for i in range(len(unit_details[NEW_LINE_LIST[0]]))]
                unit_details_dict.append(unit_details)
                unit_details = {}
            print "=====================view_excel_file ends here========================="
            print unit_details_dict
            return unit_details_dict
        else:
            pass

    return unit_details_dict
