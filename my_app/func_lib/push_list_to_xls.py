import xlsxwriter
import datetime
import os
from my_app.settings import app_cfg


def push_list_to_xls(my_list, excel_file, run_dir=app_cfg['UPDATES_DIR'], tbl_name='table1'):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir, excel_file))
    print('CREATING>>>>>>>>>> ', path_to_run_dir)

    # def push_list_to_xls(my_list, xls_file, xls_time=app_cfg['PROD_DATE']):
    #
    # Get settings for file locations and names
    #

    #
    # Write the Excel File
    #
    workbook = xlsxwriter.Workbook(path_to_run_dir)
    worksheet = workbook.add_worksheet()

    # cell_format = workbook.add_format()
    # cell_format.set_bold()
    # cell_format.set_bg_color('#B7FFF9')
    #
    # cell_format.set_bg_color('#B7D9FF')
    # cell_format.set_bg_color('#FFFEB7')
    # # cell_format.set_font_color('red')

    xls_money = workbook.add_format({'num_format': '$#,##0'})
    xls_date = workbook.add_format({'num_format': 'mm / dd/ yyyy'})

    for row_num, my_row in enumerate(my_list):
        for col_num, cell_val in enumerate(my_row):
            if type(cell_val) is float:
                worksheet.write(row_num, col_num, cell_val, xls_money)
            elif isinstance(cell_val, datetime.datetime):
                worksheet.write(row_num, col_num, cell_val, xls_date)
            else:
                # worksheet.write(row_num, col_num, cell_val, cell_format)
                worksheet.write(row_num, col_num, cell_val)

    # Prep the header row for our table
    header_row = my_list[0]
    col_list = []
    for col_name in header_row:
        col_desc = {'header': col_name}
        col_list.append(col_desc)

    # Make a table of our data (handy for PowerBI
    worksheet.add_table(0, 0, row_num, col_num, {'header_row': True,
                                                     'name': tbl_name,
                                                     'columns': col_list})

    workbook.close()
    return
