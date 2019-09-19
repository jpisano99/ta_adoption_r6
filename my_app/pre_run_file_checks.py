import os
import json
import time
from my_app.settings import app_cfg
from my_app.func_lib.push_xlrd_to_xls import push_xlrd_to_xls
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.build_sku_dict import build_sku_dict


def pre_run_file_checks(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    update_dir = app_cfg['UPDATES_DIR']
    archive_dir = app_cfg['ARCHIVES_DIR']

    # Check that all key directories exist
    path_to_main_dir = (os.path.join(home, working_dir))
    if not os.path.exists(path_to_main_dir):
        print(path_to_main_dir, " does NOT Exist !")
        exit()

    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    if not os.path.exists(path_to_run_dir):
        print(path_to_run_dir, " does NOT Exist !")
        exit()

    path_to_updates = (os.path.join(home, working_dir, update_dir))
    if not os.path.exists(path_to_updates):
        print(path_to_updates, " does NOT Exist !")
        exit()

    path_to_archives = (os.path.join(home, working_dir, archive_dir))
    if not os.path.exists(path_to_archives):
        print(path_to_archives, " does NOT Exist !")
        exit()

    # OK directories are there any files ?
    if not os.listdir(path_to_run_dir):
        print('Directory', path_to_run_dir, 'contains NO files')
        exit()

    #  Get the required Files to begin processing from app_cfg (settings.py)
    files_needed = {}
    # Do we have RAW files to process ?
    for var in app_cfg:
        if var.find('RAW') != -1:
            # Look for any config var containing the word 'RAW' and assume they are "Missing'
            files_needed[app_cfg[var]] = 'Missing'

    # See if we have the files_needed are there and they have consistent dates (date_list)
    run_files = os.listdir(path_to_run_dir)
    date_list = []
    for file_needed, status in files_needed.items():
        for run_file in run_files:
            date_tag = run_file[-13:-13 + 8]  # Grab the date if any
            run_file = run_file[:len(run_file)-14]  # Grab the name without the date
            if run_file == file_needed:
                date_list.append(date_tag)  # Grab the date
                files_needed[file_needed] = 'Found'
                break

    # All time stamps the same ?
    base_date = date_list[0]
    for date_stamp in date_list:
        if date_stamp != base_date:
            print('ERROR: Inconsistent date stamp(s) found')
            exit()

    # Do we have all the files we need ?
    for file_name, status in files_needed.items():
        if status != 'Found':
            print("ERROR: Filename ", "'"+file_name, "MM-DD-YY'  is missing from directory", "'"+run_dir+"'")
            exit()

    # Read the config_dict.json file
    try:
        with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
            config_dict = json.load(json_input)
        print(config_dict)
        print(type(config_dict))
        print(config_dict['last_run_dir'])
    except:
        print('No config_dict file found.')

    # Since we have a consistent date then Create the json file for config_data.json.
    # Put the time_stamp in it
    config_dict = {'data_time_stamp': base_date,
                   'last_run_dir': path_to_run_dir,
                   'files_scrubbed': 'never'}
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE']), 'w') as json_output:
        json.dump(config_dict, json_output)

    # Delete all previous tmp_ files
    for file_name in run_files:
        if file_name[0:4] == 'tmp_':
            os.remove(os.path.join(path_to_run_dir, file_name))

    # Here is what we have - All things should be in place
    print('Our directories:')
    print('\tPath to Main Dir:', path_to_main_dir)
    print('\tPath to Updates Dir:', path_to_updates)
    print('\tPath to Archives Dir:', path_to_archives)
    print('\tPath to Run Dir:', path_to_run_dir)

    # Process the RAW data (Renewals and Bookings)
    # Clean up rows, combine multiple Bookings files, add custom table names
    processing_date = date_list[0]
    file_paths = []
    bookings = []
    subscriptions = []
    as_status = []
    print()
    print('We are processing files:')

    # We need to make a sku_filter_dict here
    tmp_dict = build_sku_dict()
    sku_filter_dict = {}
    for key, val in tmp_dict.items():
        if val[0] == 'Service':
            sku_filter_dict[key] = val

    # Main loop to process files
    for file_name in files_needed:
        file_path = file_name + ' ' + processing_date + '.xlsx'
        file_path = os.path.join(path_to_run_dir, file_path)

        file_paths.append(file_path)

        my_wb, my_ws = open_wb(file_name + ' ' + processing_date + '.xlsx', run_dir)
        print('\t\t', file_name + '', processing_date + '.xlsx', ' has ', my_ws.nrows,
              ' rows and ', my_ws.ncols, 'columns')

        if file_name.find('Bookings') != -1:
            # For the Bookings start_row is here
            start_row = 3
            start_col = 1
            for row in range(start_row, my_ws.nrows):
                bookings.append(my_ws.row_slice(row, start_col))

        elif file_name.find('Subscriptions') != -1:
            # This raw sheet starts on row num 0
            for row in range(0, my_ws.nrows):
                subscriptions.append(my_ws.row_slice(row))

        elif file_name.find('AS Delivery Status') != -1:
            # This AS-F raw sheet starts on row num 0
            # Grab the header row
            as_status.append(my_ws.row_slice(0))
            for row in range(1, my_ws.nrows):
                # Check to see if this is a TA SKU
                if my_ws.cell_value(row, 14) in sku_filter_dict:
                    as_status.append(my_ws.row_slice(row))

    #
    # Push the lists out to an Excel File
    #
    push_xlrd_to_xls(bookings, app_cfg['XLS_BOOKINGS'], run_dir, 'ta_bookings')
    push_xlrd_to_xls(subscriptions, app_cfg['XLS_SUBSCRIPTIONS'], run_dir, 'ta_subscriptions')
    push_xlrd_to_xls(as_status, app_cfg['XLS_AS_DELIVERY_STATUS'], run_dir, 'ta_delivery')

    print('We have ', len(bookings), 'bookings line items')
    print('We have ', len(as_status), 'AS-Fixed SKU line items')
    print('We have ', len(subscriptions), 'subscription line items')

    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
        config_dict = json.load(json_input)
    config_dict['files_scrubbed'] = 'phase_1'
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE']), 'w') as json_output:
        json.dump(config_dict, json_output)

    return


if __name__ == "__main__" and __package__ is None:
    pre_run_file_checks()
    exit()
