from xlrd import open_workbook, XL_CELL_EMPTY
from utilities import *
import datetime

def view_excel_file(excel_file_name, unit_name_arg, search_text):
    data_list = []
    wb = open_workbook(excel_file_name)
    for s in wb.sheets():
        print 'Sheet:',s.name
        if s.name.strip() == 'B-17 Bond':
            unit_details = {}
            for i in range(1,646):
                # print i
                if s.cell(i, 2).value != '':
                    unit_name = s.cell(i, 2).value.strip()
                    '''if unit_name.lower().find(search_text) < 0:
                        continue
                    unit_name = unit_name_arg'''

                    if not unit_name in unit_details:
                        unit_details[unit_name] = []

                    current_unit = {}
                    current_unit['pc_date'] = s.cell(i, 16).value
                    try:
                        # print current_unit['pc_date']
                        pc_date = datetime.datetime.strptime(current_unit['pc_date'], '%d.%m.%Y')
                    except:
                        # print 'Datetime error at row', i+1
                        continue

                    if get_finiancial_year(pc_date) == get_finiancial_year(datetime.datetime.now()):
                        current_unit['bond_page_number'] = s.cell(i, 4).value
                        current_unit['bond_page_date'] = s.cell(i, 5).value
                        current_unit['rwc_number'] = s.cell(i, 6).value
                        current_unit['ta_number'] = s.cell(i, 7).value
                        current_unit['ta_date'] = s.cell(i, 8).value
                        # current_unit['rwc_date'] = s.cell(i, 9).value
                        current_unit['invoice_number_and_date'] = s.cell(i, 10).value
                        current_unit['boe_number_and_date'] = s.cell(i, 12).value
                        current_unit['certificate_number'] = s.cell(i, 13).value
                        current_unit['ic_date'] = s.cell(i, 14).value
                        current_unit['procurement_certificate_number'] = s.cell(i, 15).value
                        current_unit['name_of_supplier'] = s.cell(i, 17).value
                        current_unit['country_name'] = s.cell(i, 18).value
                        current_unit['description_of_goods'] = s.cell(i, 20).value
                        current_unit['quantity_of_goods'] = normalize_qty(s.cell(i, 21).value)
                        current_unit['cif_value'] = s.cell(i, 22).value
                        current_unit['assessable_value'] = s.cell(i, 24).value
                        current_unit['duty_foregone'] = s.cell(i, 25).value
                        current_unit['debit-b17'] = s.cell(i, 30).value

                        unit_details[unit_name].append(current_unit)

            return unit_details
