import xlrd
import os
from my_app.settings import app_cfg


def open_wb(excel_file, run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir))
    path_to_file = os.path.join(path_to_run_dir, excel_file)
    print('OPENING>>>>>>>>>> ', path_to_file)

    #
    # Open up excel workbook
    #
    my_wb = xlrd.open_workbook(path_to_file)
    my_ws = my_wb.sheet_by_index(0)

    return my_wb, my_ws


if __name__ == "__main__":
    my_excel = open_wb(app_cfg['BOOKINGS'])
    print('We have: ', my_excel[0], my_excel[1])
