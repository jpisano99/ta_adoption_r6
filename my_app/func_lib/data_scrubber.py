from my_app.settings import app_cfg
from my_app.func_lib.sheet_desc import sheet_map
from my_app.func_lib.build_sheet_map import build_sheet_map
from my_app.func_lib.open_wb import open_wb
import time, datetime, ntpath, xlrd
from datetime import datetime


def data_scrubber(my_ws, path_to_file):
    #
    # Data Scrubber
    #

    # Initialize and Get Header Row
    list_of_rows = []
    list_of_row = []
    list_headers = []

    list_of_rows.append(my_ws.row(0))
    list_headers.append(my_ws.row_values(0, 0))
    print()
    print(list_of_rows)
    print("Headers\t", list_headers)
    print()

    # Open and Map the XLS file we are scrubbing
    run_dir, my_file = ntpath.split(path_to_file)
    my_map = build_sheet_map(my_file, sheet_map, 'XLS_SUBSCRIPTIONS', run_dir)

    # List comprehension replacement for above
    # Strip out the columns from the sheet map that we don't need
    my_col_map = [x for x in my_map if x[1] == 'XLS_SUBSCRIPTIONS']

    for my_row in range(1, my_ws.nrows):
        print()
        print('Row #:', my_row)
        for my_col in my_col_map:
            my_col_name = list_headers[0][my_col[2]]
            dest_cell_type = my_col[4]
            dest_cell_val = None

            src_cell = my_ws.cell(my_row, my_col[2])
            src_cell_type = my_ws.cell_type(my_row, my_col[2])
            src_cell_val = my_ws.cell_value(my_row, my_col[2])

            print()
            print('\t\t' + my_col_name + ' source type/value is', src_cell)
            print('\t\tDest type needs to be', dest_cell_type)

            if src_cell_type == xlrd.XL_CELL_DATE:
                print("\t\tDate", my_ws.cell_value(my_row, my_col[2]), ' needs to be a ' + dest_cell_type)
                # datetime_object = datetime.strptime('Jun 1 2005  1:33PM', '%b %d %Y %I:%M%p')

            elif src_cell_type == xlrd.XL_CELL_NUMBER:
                # All numbers from xlrd are python floats so we need to check if we need an int
                if dest_cell_type == 'int':
                    dest_cell_val = int(src_cell_val)
                    print('\t\tConverted', src_cell_val, ' to ', dest_cell_val)

            elif src_cell_type == xlrd.XL_CELL_TEXT:
                if dest_cell_type == 'currency' or dest_cell_type == 'int':
                    try:
                        int(src_cell_val)
                        dest_cell_val = int(src_cell_val)
                        print('\t\tConverted', src_cell_val, ' to ', dest_cell_val)
                    except ValueError:
                        print("\t\tERROR: Not an currency, int or float")
                        dest_cell_val = 0
                        print('\t\tConverted', src_cell_val, ' to ', dest_cell_val)
                elif dest_cell_type == 'date':
                    try:
                        datetime.strptime(src_cell_val, '%d %b %Y')
                        dest_cell_val = datetime.strptime(src_cell_val, '%d %b %Y')
                        print('\t\tConverted', src_cell_val, ' to ', dest_cell_val)
                    except ValueError:
                        print('\t\t', type(src_cell_val), src_cell_val)
                        print('\t\tNot a date')
                        dest_cell_val = datetime.strptime('01 Jan 2000', '%d %b %Y')
                        print('\t\tConverted', src_cell_val, ' to ', dest_cell_val)

            elif src_cell_type == xlrd.XL_CELL_BLANK:
                print("\t\tBlank", my_ws.cell_value(my_row, my_col[2]), ' needs to be a ' + dest_cell_type)

            elif src_cell_type == xlrd.XL_CELL_BOOLEAN:
                print("\t\tBoolean", my_ws.cell_value(my_row, my_col[2]), ' needs to be a ' + dest_cell_type)

            elif src_cell_type == xlrd.XL_CELL_EMPTY:
                print("\t\tEmpty", my_ws.cell_value(my_row, my_col[2]), ' needs to be a ' + dest_cell_type)

            elif src_cell_type == xlrd.XL_CELL_ERROR:
                print("\t\tError", my_ws.cell_value(my_row, my_col[2]), ' needs to be' + dest_cell_type)
            time.sleep(0.5)
            print (type(dest_cell_val),dest_cell_val)
            list_of_row.append(dest_cell_val)
        print(list_of_row)

        list_of_rows.append(list_of_row)
        print(list_of_rows)
        time.sleep(2)
    return list_of_rows


if __name__ == "__main__":
    path_to_scrub = 'C:/Users\jpisano/ta_adoption_data/ta_data_updates/TA Master Subscriptions as of 06-12-19.xlsx'
    my_path, my_file = ntpath.split(path_to_scrub)
    wb, ws = open_wb(my_file)
    data_scrubber(ws, path_to_scrub)