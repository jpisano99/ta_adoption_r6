import xlsxwriter
import xlrd
import os
from my_app.settings import app_cfg


def push_xlrd_to_xls(my_list, excel_file, run_dir=app_cfg['UPDATES_DIR'], tbl_name='table1'):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir, excel_file))
    print('CREATING>>>>>>>>>> ', path_to_run_dir)
    # path_to_run_dir is a fully qualified path (ie)
    #    C:\Users\Jim\ta_adoption_data\ta_data_updates

    #
    # Write the Excel File
    #
    workbook = xlsxwriter.Workbook(path_to_run_dir)
    worksheet = workbook.add_worksheet()

    # Some ways we could add custom formats - testing
    # cell_format = workbook.add_format()
    # cell_format.set_bold()
    # cell_format.set_bg_color('#B7FFF9')
    #
    # cell_format.set_bg_color('#B7D9FF')
    # cell_format.set_bg_color('#FFFEB7')
    # # cell_format.set_font_color('red')

    # Define these formats for XLSXWriter
    xls_money = workbook.add_format({'num_format': '$#,##0'})
    xls_date = workbook.add_format({'num_format': 'mm / dd/ yyyy'})

    # Loop over each row and each cell
    # Format Dates and Currency as per our specs
    row_num = 0
    col_num = 0
    for row_num, my_row in enumerate(my_list):
        for col_num, my_cell in enumerate(my_row):
            if my_cell.ctype == xlrd.XL_CELL_DATE:
                worksheet.write(row_num, col_num, my_cell.value, xls_date)
            elif my_cell.ctype == xlrd.XL_CELL_NUMBER:
                worksheet.write(row_num, col_num, my_cell.value, xls_money)
            else:
                worksheet.write(row_num, col_num, my_cell.value)

    # Prep the header row for our table
    header_row = my_list[0]
    col_list = []
    for col_name in header_row:
        col_desc = {'header': col_name.value}
        col_list.append(col_desc)

    # Make a table of our data (handy for PowerBI
    worksheet.add_table(0, 0, row_num, col_num, {'header_row': True,
                                                    'name': tbl_name,
                                                   'columns': col_list})
    workbook.close()

    return
