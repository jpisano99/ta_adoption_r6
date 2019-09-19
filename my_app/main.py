import datetime
# from datetime import datetime
import os
import json
import xlrd
import time
from my_app.settings import app_cfg
from my_app.func_lib.push_xlrd_to_xls import push_xlrd_to_xls
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.build_coverage_dict import build_coverage_dict
from my_app.func_lib.build_sku_dict import build_sku_dict
from my_app.func_lib.find_team import find_team
from my_app.func_lib.process_renewals import process_renewals
from my_app.func_lib.process_subscriptions import process_subs
from my_app.func_lib.process_delivery_updates import process_delivery
from my_app.func_lib.build_customer_list import build_customer_list
from my_app.func_lib.cleanup_orders import cleanup_orders
from my_app.func_lib.sheet_desc import sheet_map as sm
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.create_customer_order_dict import create_customer_order_dict
from my_app.func_lib.get_linked_sheet_update import get_linked_sheet_update
from my_app.func_lib.build_sheet_map import build_sheet_map
from my_app.func_lib.sheet_desc import sheet_map, sheet_keys
from my_app.func_lib.data_scrubber import data_scrubber


def phase_1(run_dir=app_cfg['UPDATES_DIR']):
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
    # with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
    #     config_dict = json.load(json_input)
    # print(config_dict)

    # Since we have a consistent date then Create the json file for config_data.json.
    # Put the time_stamp in it
    config_dict = {'data_time_stamp': base_date,
                   'last_run_dir': path_to_run_dir}
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
    renewals = []
    subscriptions = []
    as_status = []
    start_row = 0
    print()
    print('We are processing files:')

    for file_name in files_needed:
        file_path = file_name + ' ' + processing_date + '.xlsx'
        file_path = os.path.join(path_to_run_dir, file_path)

        file_paths.append(file_path)

        my_wb, my_ws = open_wb(file_name + ' ' + processing_date + '.xlsx', run_dir)
        # my_wb = xlrd.open_workbook(file_path)
        # my_ws = my_wb.sheet_by_index(0)
        print('\t\t', file_name + '', processing_date + '.xlsx', ' has ', my_ws.nrows,
              ' rows and ', my_ws.ncols, 'columns')

        if file_name.find('Bookings') != -1:
            if start_row == 0:
                # For the first workbook include the header row
                start_row = 2
            elif start_row == 2:
                # For subsequent workbooks skip the header
                start_row = 3
            for row in range(start_row, my_ws.nrows):
                bookings.append(my_ws.row_slice(row))

        elif file_name.find('Subscriptions') != -1:
            # This raw sheet starts on row num 0

            for row in range(0, my_ws.nrows):
                subscriptions.append(my_ws.row_slice(row))

        elif file_name.find('Renewals') != -1:
            # This raw sheet starts on row num 2
            for row in range(2, my_ws.nrows):
                renewals.append(my_ws.row_slice(row))

        elif file_name.find('AS Delivery Status') != -1:
            # This raw sheet starts on row num 2
            for row in range(0, my_ws.nrows):
                as_status.append(my_ws.row_slice(row))

    # Push the lists out to an Excel File
    push_xlrd_to_xls(bookings, app_cfg['XLS_BOOKINGS'], run_dir, 'ta_bookings')

    as_bookings = get_as_skus(bookings)
    push_xlrd_to_xls(as_bookings, app_cfg['XLS_AS_SKUS'], run_dir, 'as_bookings')

    push_xlrd_to_xls(renewals, app_cfg['XLS_RENEWALS'], run_dir, 'ta_renewals')

    push_xlrd_to_xls(subscriptions, app_cfg['XLS_SUBSCRIPTIONS'], run_dir, 'ta_subscriptions')

    push_xlrd_to_xls(as_status, app_cfg['XLS_AS_DELIVERY_STATUS'], run_dir, 'ta_delivery')

    print('We have ', len(bookings), 'bookings line items')
    print('We have ', len(as_bookings), 'Services line items')
    print('We have ', len(renewals), 'renewal line items')
    print('We have ', len(subscriptions), 'subscription line items')
    return

##################
# End of Phase 1
##################

##################
# Start of Phase 2
##################


def phase_2(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    bookings_path = os.path.join(path_to_run_dir, app_cfg['XLS_BOOKINGS'])

    # Read the config_dict.json file
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
        config_dict = json.load(json_input)
    data_time_stamp = datetime.datetime.strptime(config_dict['data_time_stamp'], '%m-%d-%y')
    last_run_dir = config_dict['last_run_dir']

    print("Run Date: ", data_time_stamp, type(data_time_stamp))
    print('Run Directory:', last_run_dir)
    print(bookings_path)

    # Go to Smartsheets and build these two dicts to use reference lookups
    # team_dict: {'sales_levels 1-6':[('PSS','TSA')]}
    # sku_dict: {sku : [sku_type, sku_description]}
    team_dict = build_coverage_dict()
    sku_dict = build_sku_dict()

    #
    # Open up the bookings excel workbooks
    #
    wb_bookings, sheet_bookings = open_wb(app_cfg['XLS_BOOKINGS'], run_dir)

    # From the current up to date bookings file build a simple list
    # that describes the format of the output file we are creating
    # and the columns we need to add (ie PSS, TSA, Renewal Dates)

    my_sheet_map = build_sheet_map(app_cfg['XLS_BOOKINGS'], sheet_map,
                                   'XLS_BOOKINGS', run_dir)
    #
    # init a bunch a variables we need for the main loop
    #
    order_header_row = []
    order_rows = []
    order_row = []
    trash_rows = []

    dest_col_nums = {}
    src_col_nums = {}

    # Build a dict of source sheet {'col_name' : src_col_num}
    # Build a dict of destination sheet {'col_name' : dest_col_num}
    # Build the header row for the output file
    for idx, val in enumerate(my_sheet_map):
        # Add to the col_num dict of col_names
        dest_col_nums[val[0]] = idx
        src_col_nums[val[0]] = val[2]
        order_header_row.append(val[0])

    # Initialize the order_row and trash_row lists
    order_rows.append(order_header_row)
    trash_rows.append(sheet_bookings.row_values(0))

    print('There are ', sheet_bookings.nrows, ' rows in Raw Bookings')

    #
    # Main loop of over raw bookings excel data
    #
    # This loop will build two lists:
    # 1. Interesting orders based on SKUs (order_rows)
    # 2. Trash orders SKUs we don't care about (trash_rows)
    # As determined by the sku_dict
    # We have also will assign team coverage to both rows
    #
    for i in range(1, sheet_bookings.nrows):

        # Is this SKU of interest ?
        sku = sheet_bookings.cell_value(i, src_col_nums['Bundle Product ID'])

        if sku in sku_dict:
            # Let's make a row for this order
            # Since it has an "interesting" sku
            customer = sheet_bookings.cell_value(i, src_col_nums['ERP End Customer Name'])
            order_row = []
            sales_level = ''
            sales_level_cntr = 0

            # Grab SKU data from the SKU dict
            sku_type = sku_dict[sku][0]
            sku_desc = sku_dict[sku][1]
            sku_sensor_cnt = sku_dict[sku][2]

            # Walk across the sheet_map columns
            # to build this output row cell by cell
            for val in my_sheet_map:
                col_name = val[0]  # Source Sheet Column Name
                col_idx = val[2]  # Source Sheet Column Number

                # If this is a 'Sales Level X' column then
                # Capture it's value until we get to level 6
                # then do a team lookup
                if col_name[:-2] == 'Sales Level':
                    sales_level = sales_level + sheet_bookings.cell_value(i, col_idx) + ','
                    sales_level_cntr += 1

                    if sales_level_cntr == 6:
                        # We have collected all 6 sales levels
                        # Now go to find_team to do the lookup
                        sales_level = sales_level[:-1]
                        sales_team = find_team(team_dict, sales_level)
                        pss = sales_team[0]
                        tsa = sales_team[1]
                        order_row[dest_col_nums['pss']] = pss
                        order_row[dest_col_nums['tsa']] = tsa

                if col_idx != -1:
                    # OK we have a cell that we need from the raw bookings
                    # sheet we need so grab it
                    order_row.append(sheet_bookings.cell_value(i, col_idx))
                elif col_name == 'Product Description':
                    # Add in the Product Description
                    order_row.append(sku_desc)
                elif col_name == 'Product Type':
                    # Add in the Product Type
                    order_row.append(sku_type)
                elif col_name == 'Sensor Count':
                    # Add in the Sensor Count
                    order_row.append(sku_sensor_cnt)
                else:
                    # this cell is assigned a -1 in the sheet_map
                    # so assign a blank as a placeholder for now
                    order_row.append('')

            # Done with all the columns in this row
            # Log this row for BOTH customer names and orders
            # Go to next row of the raw bookings data
            order_rows.append(order_row)

        else:
            # The SKU was not interesting so let's trash it
            trash_rows.append(sheet_bookings.row_values(i))

    print('Extracted ', len(order_rows), " rows of interesting SKU's' from Raw Bookings")
    print('Trashed ', len(trash_rows), " rows of trash SKU's' from Raw Bookings")
    #
    # End of main loop
    #
    push_list_to_xls(order_rows,'jim.xlsx')

    #
    # Subscription Analysis
    #
    no_match = [['No Match Found in Subscription update']]
    no_match_cntr = 0
    match_cntr = 0
    subs_sorted_dict, subs__summary_dict = process_subs(run_dir)
    for order_row in order_rows[1:]:
        customer = order_row[dest_col_nums['ERP End Customer Name']]
        if customer in subs_sorted_dict:
            match_cntr += 1
            sub_start_date = datetime.datetime.strptime(subs_sorted_dict[customer][0][0], '%m-%d-%Y')
            sub_initial_term = subs_sorted_dict[customer][0][1]
            sub_renewal_date = datetime.datetime.strptime(subs_sorted_dict[customer][0][2], '%m-%d-%Y')
            sub_days_to_renew = subs_sorted_dict[customer][0][3]
            sub_monthly_charge = subs_sorted_dict[customer][0][4]
            sub_id = subs_sorted_dict[customer][0][5]
            sub_status = subs_sorted_dict[customer][0][6]

            order_row[dest_col_nums['Start Date']] = sub_start_date
            order_row[dest_col_nums['Initial Term']] = sub_initial_term
            order_row[dest_col_nums['Renewal Date']] = sub_renewal_date
            order_row[dest_col_nums['Days Until Renewal']] = sub_days_to_renew
            order_row[dest_col_nums['Monthly Charge']] = sub_monthly_charge
            order_row[dest_col_nums['Subscription ID']] = sub_id
            order_row[dest_col_nums['Status']] = sub_status

            if len(subs_sorted_dict[customer]) > 1:
                renewal_comments = '+' + str(len(subs_sorted_dict[customer])-1) + ' more subscriptions(s)'
                order_row[dest_col_nums['Subscription Comments']] = renewal_comments
        else:
            got_one = False
            for x in no_match:
                if x[0].lower() in customer.lower():
                    got_one = True
                    break
            if got_one is False:
                no_match_cntr += 1
                no_match.append([customer])

    push_list_to_xls(order_rows, 'jim1.xlsx')
    push_list_to_xls(no_match, 'subcription misses.xlsx')

    #
    # AS Delivery Analysis
    #
    as_dict = process_delivery(run_dir)
    print(as_dict)
    print(len(as_dict))
    for order_row in order_rows[1:]:
        customer = order_row[dest_col_nums['ERP End Customer Name']]
        if customer in as_dict:
            match_cntr += 1
            # as_customer = as_dict[customer][0][0]
            # as_pid = as_dict[customer][0][1]
            # as_dm = as_dict[customer][0][2]
            # as_start_date = datetime.datetime.strptime(as_dict[customer][0][3], '%m-%d-%Y')
            #
            # order_row[dest_col_nums['End Customer']] = as_customer
            # order_row[dest_col_nums['PID']] = as_pid
            # order_row[dest_col_nums['Delivery Manager']] = as_dm
            # order_row[dest_col_nums['Project Scheduled Start Date']] = as_start_date

            order_row[dest_col_nums['PID']] = as_dict[customer][0][0]
            order_row[dest_col_nums['Delivery Manager']] = as_dict[customer][0][1]
            order_row[dest_col_nums['Delivery PM']] = as_dict[customer][0][2]
            order_row[dest_col_nums['Tracking status']] = as_dict[customer][0][3]
            order_row[dest_col_nums['Tracking Sub-status']] = as_dict[customer][0][4]
            order_row[dest_col_nums['Comments']] = as_dict[customer][0][5]
            order_row[dest_col_nums['Project Scheduled Start Date']] = datetime.datetime.strptime(as_dict[customer][0][6], '%m-%d-%Y')
            order_row[dest_col_nums['Scheduled End Date']] = datetime.datetime.strptime(as_dict[customer][0][7], '%m-%d-%Y')
            order_row[dest_col_nums['Project Creation Date']] = datetime.datetime.strptime(as_dict[customer][0][8], '%m-%d-%Y')
            # order_row[dest_col_nums['Project Closed Date']] = datetime.datetime.strptime(as_dict[customer][0][9], '%m-%d-%Y')
            order_row[dest_col_nums['Traffic lights (account team)']] = as_dict[customer][0][10]
            order_row[dest_col_nums['Tracking Responsible']] = as_dict[customer][0][11]
            order_row[dest_col_nums['ExecutiveSummary']] = as_dict[customer][0][12]
            order_row[dest_col_nums['Critalpath']] = as_dict[customer][0][13]
            order_row[dest_col_nums['IsssuesRisks']] = as_dict[customer][0][14]
            order_row[dest_col_nums['ActivitiesCurrent']] = as_dict[customer][0][15]
            order_row[dest_col_nums['ActivitiesNext']] = as_dict[customer][0][16]
            order_row[dest_col_nums['LastUpdate']] = as_dict[customer][0][17]
            order_row[dest_col_nums['SO']] = as_dict[customer][0][18]

        else:
            got_one = False
            for x in no_match:
                if x[0].lower() in customer.lower():
                    got_one = True
                    break
            if got_one is False:
                no_match_cntr += 1
                no_match.append([customer])

    #
    # End of  Construction Zone
    #

    # Now we build a an order dict
    # Let's organize as this
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    order_dict = {}
    orders = []
    order = []

    for idx, order_row in enumerate(order_rows):
        if idx == 0:
            continue
        customer = order_row[0]
        orders = []

        # Is this customer in the order dict ?
        if customer in order_dict:
            orders = order_dict[customer]
            orders.append(order_row)
            order_dict[customer] = orders
        else:
            orders.append(order_row)
            order_dict[customer] = orders

    # Create a simple customer_list
    # Contains a full set of unique sorted customer names
    # Example: customer_list = [[erp_customer_name,end_customer_ultimate], [CustA,CustA]]
    customer_list = build_customer_list(run_dir)
    print('There are ', len(customer_list), ' unique Customer Names')

    # Clean up order_dict to remove:
    # 1.  +/- zero sum orders
    # 2. zero revenue orders
    order_dict, customer_platforms = cleanup_orders(customer_list, order_dict, my_sheet_map)

    #
    # Create a summary order file out of the order_dict
    #
    summary_order_rows = [order_header_row]
    for key, val in order_dict.items():
        for my_row in val:
            summary_order_rows.append(my_row)
    print(len(summary_order_rows), ' of scrubbed rows after removing "noise"')

    #
    # Push our lists to an excel file
    #
    # push_list_to_xls(customer_platforms, 'jim ')
    print('order summary name ', app_cfg['XLS_ORDER_SUMMARY'])

    push_list_to_xls(summary_order_rows, app_cfg['XLS_ORDER_SUMMARY'],
                     run_dir, 'ta_summary_orders')
    push_list_to_xls(order_rows, app_cfg['XLS_ORDER_DETAIL'], run_dir, 'ta_order_detail')
    push_list_to_xls(customer_list, app_cfg['XLS_CUSTOMER'], run_dir, 'ta_customers')
    push_list_to_xls(trash_rows, app_cfg['XLS_BOOKINGS_TRASH'], run_dir, 'ta_trash_rows')

    # exit()
    #
    # Push our lists to a smart sheet
    #
    # push_xls_to_ss(wb_file, app_cfg['XLS_ORDER_SUMMARY'])
    # push_xls_to_ss(wb_file, app_cfg['XLS_ORDER_DETAIL'])
    # push_xls_to_ss(wb_file, app_cfg['XLS_CUSTOMER'])
    # exit()
    return

##################
# Start of Phase 3
##################

def phase_3(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    # from my_app.func_lib.sheet_desc import sheet_map
    #
    # Open the order summary
    #
    wb_orders, sheet_orders = open_wb(app_cfg['XLS_ORDER_SUMMARY'], run_dir)

    # wb_orders, sheet_orders = open_wb('tmp_TA Scrubbed Orders_as_of ' + app_cfg['PROD_DATE'])

    # Loop over the orders XLS worksheet
    # Create a simple list of orders with NO headers
    order_list = []
    for row_num in range(1, sheet_orders.nrows):  # Skip the header row start at 1
        tmp_record = []
        for col_num in range(sheet_orders.ncols):
            my_cell = sheet_orders.cell_value(row_num, col_num)

            # If we just read a date save it as a datetime
            if sheet_orders.cell_type(row_num, col_num) == 3:
                my_cell = datetime.datetime(*xlrd.xldate_as_tuple(my_cell, wb_orders.datemode))
            tmp_record.append(my_cell)
        order_list.append(tmp_record)

    # Create a dict of customer orders
    customer_order_dict = create_customer_order_dict(order_list)
    print()
    print('We have summarized ', len(order_list), ' of interesting line items into')
    print(len(customer_order_dict), ' unique customers')
    print()

    # Build Sheet Maps
    sheet_map = build_sheet_map(app_cfg['SS_CX'], sm, 'SS_CX', run_dir)
    sheet_map = build_sheet_map(app_cfg['SS_AS'], sheet_map, 'SS_AS', run_dir)
    sheet_map = build_sheet_map(app_cfg['SS_SAAS'], sheet_map, 'SS_SAAS', run_dir)

    #
    # Get dict updates from linked sheets CX/AS/SAAS
    #
    cx_dict = get_linked_sheet_update(sheet_map, 'SS_CX', sheet_keys)
    as_dict = get_linked_sheet_update(sheet_map, 'SS_AS', sheet_keys)
    saas_dict = get_linked_sheet_update(sheet_map, 'SS_SAAS', sheet_keys)

    print()
    print('We have CX Updates: ', len(cx_dict))
    print('We have AS Updates: ', len(as_dict))
    print('We have SAAS Updates: ', len(saas_dict))
    print()

    # Create Platform dict for platform lookup
    tmp_dict = build_sku_dict()
    platform_dict = {}
    for key, val in tmp_dict.items():
        if val[0] == 'Product' or val[0] == 'SaaS':
            platform_dict[key] = val[1]

    #
    # Init Main Loop Variables
    #
    new_rows = []
    new_row = []
    bookings_col_num = -1
    sensor_col_num = -1
    svc_bookings_col_num = -1
    platform_type_col_num = -1
    sku_col_num = -1
    my_col_idx = {}
    no_as_match = []

    # Create top row for the dashboard
    # also make a dict (my_col_idx) of {column names : column number}
    for col_idx, col in enumerate(sheet_map):
        new_row.append(col[0])
        my_col_idx[col[0]] = col_idx
    new_rows.append(new_row)

    #
    # Main loop
    #
    for customer, orders in customer_order_dict.items():
        new_row = []
        order = []
        orders_found = len(orders)

        # Default Values
        bookings_total = 0
        sensor_count = 0
        service_bookings = 0
        platform_type = 'Not Identified'

        saas_status = 'No Status'
        cx_contact = 'None assigned'
        cx_status = 'No Update'
        as_pm = ''
        as_cse1 = ''
        as_cse2 = ''
        as_complete = ''  # 'Project Status/PM Completion'
        as_comments = ''  # 'Delivery Comments'

        #
        # Get update from linked sheets (if any)
        #
        if customer in saas_dict:
            saas_status = saas_dict[customer][0]
            if saas_status is True:
                saas_status = 'Provision Complete'
            else:
                saas_status = 'Provision NOT Complete'
        else:
            saas_status = 'No Status'

        if customer in cx_dict:
            cx_contact = cx_dict[customer][0]
            cx_status = cx_dict[customer][1]
        else:
            cx_contact = 'None assigned'
            cx_status = 'No Update'

        if customer in as_dict:
            if as_dict[customer][0] == '':
                as_pm = 'None Assigned'
            else:
                as_pm = as_dict[customer][0]

            if as_dict[customer][1] == '':
                as_cse1 = 'None Assigned'
            else:
                as_cse1 = as_dict[customer][1]

            if as_dict[customer][2] == '':
                as_cse2 = 'None Assigned'
            else:
                as_cse2 = as_dict[customer][2]

            if as_dict[customer][3] == '':
                as_complete = 'No Update'
            else:
                # 'Project Status/PM Completion'
                as_complete = as_dict[customer][3]

            if as_dict[customer][4] == '':
                as_comments = 'No Comments'
            else:
                as_comments = as_dict[customer][4]
        else:
            no_as_match.append([customer])


        #
        # Loop over this customers orders
        # Create one summary row for this customer
        # Total things
        # Build a list of things that may change order to order (ie Renewal Dates, Customer Names)
        #
        platform_count = 0
        for order_idx, order in enumerate(orders):
            # calculate totals in this loop (ie total_books, sensor count etc)
            bookings_total = bookings_total + order[my_col_idx['Total Bookings']]
            sensor_count = sensor_count + order[my_col_idx['Sensor Count']]

            if order[my_col_idx['Product Type']] == 'Service':
                service_bookings = service_bookings + order[my_col_idx['Total Bookings']]

            if order[my_col_idx['Bundle Product ID']] in platform_dict:
                platform_count += 1
                platform_type = platform_dict[order[my_col_idx['Bundle Product ID']]]
                if platform_count > 1:
                    platform_type = platform_type + ' plus ' + str(platform_count-1)

        #
        # Modify/Update this record as needed and then add to the new_rows
        #
        order[my_col_idx['Total Bookings']] = bookings_total
        order[my_col_idx['Sensor Count']] = sensor_count
        order[my_col_idx['Service Bookings']] = service_bookings

        order[my_col_idx['CSM']] = cx_contact
        order[my_col_idx['Comments']] = cx_status

        order[my_col_idx['Project Manager']] = as_pm
        order[my_col_idx['AS Engineer 1']] = as_cse1
        order[my_col_idx['AS Engineer 2']] = as_cse2
        order[my_col_idx['Project Status/PM Completion']] = as_complete  # 'Project Status/PM Completion'
        order[my_col_idx['Delivery Comments']] = as_comments

        order[my_col_idx['Provisioning completed']] = saas_status

        order[my_col_idx['Product Description']] = platform_type

        order[my_col_idx['Orders Found']] = orders_found

        new_rows.append(order)

    push_list_to_xls(no_as_match, 'AS customer name misses.xlsx')
    #
    # End of main loop
    #

    # Do some clean up and ready for output
    #
    # Rename the columns as per the sheet map
    cols_to_delete = []
    for idx, map_info in enumerate(sheet_map):
        if map_info[3] != '':
            if map_info[3] == '*DELETE*':
                # Put the columns to delete in a list
                cols_to_delete.append(idx)
            else:
                # Rename to the new column name
                new_rows[0][idx] = map_info[3]

    # Loop over the new_rows and
    # delete columns we don't need as per the sheet_map
    for col_idx in sorted(cols_to_delete, reverse=True):
        for row_idx, my_row in enumerate(new_rows):
            del new_rows[row_idx][col_idx]

    #
    # Write the Dashboard to an Excel File
    #
    push_list_to_xls(new_rows, app_cfg['XLS_DASHBOARD'], run_dir,'ta_dashboard')
    # push_xls_to_ss(app_cfg['XLS_DASHBOARD']+'_as_of_01_31_2019.xlsx', 'jims dash')

    return

##################
# End of Phase 3
##################

def get_as_skus(bookings):
    # Build a SKU dict as a filter
    tmp_dict = build_sku_dict()
    sku_dict = {}
    header_row = bookings[0]
    header_vals = []
    for my_cell in header_row:
        header_vals.append(my_cell.value)

    # Strip out all but Service sku's
    for sku_key, sku_val in tmp_dict.items():
        if sku_val[0] == 'Service':
            sku_dict[sku_key] = sku_val

    sku_col_header = 'Bundle Product ID'
    sku_col_num = 0
    as_bookings = [header_row]

    # Get the col number that has the SKU's
    for idx, val in enumerate(header_vals):
        if val == sku_col_header:
            sku_col_num = idx
            break

    # Gather all the rows with AS skus
    for booking in bookings:
        if booking[sku_col_num].value in sku_dict:
            as_bookings.append(booking)

    print('All AS SKUs have been extracted from the current data!')
    return as_bookings


if __name__ == "__main__" and __package__ is None:
    print('Package Name:', __package__)
    print('running check_update_files')
    # phase_1(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
    # phase_2(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
    # phase_3(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
    phase_1()
    phase_2()
    phase_3()

    #file_checks(os.path.join(app_cfg['UPDATES_DIR']))
    # file_checks()
    # file_checks(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
