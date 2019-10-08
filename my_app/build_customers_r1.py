from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.get_list_from_ss import get_list_from_ss
from my_app.func_lib.build_sku_dict import build_sku_dict
from my_app.func_lib.build_coverage_dict import build_coverage_dict
from my_app.settings import app_cfg
from my_app.Customer import Customer
from my_app.func_lib.find_team import find_team
from fuzzywuzzy import fuzz
import xlrd
from datetime import datetime
import time


def process_sub_info(subs_list):
    # Lets find all subscriptions records
    # We will match the Subscription by the AS Customer Name
    today = datetime.today()
    sub_summary = ''
    next_renewal_date = datetime(2050, 1, 1)
    next_renewal_rev = 0
    billing_start_date = datetime(2050, 1, 1)
    current_status = 'NONE'
    sub_renew_status = ''
    days_to_renew = ''
    sub_start_date = ''
    sub_renew_date = ''
    sorted_idx = []

    # Create a sorted index list of renewal dates
    subs_list.sort(key=lambda x: x[3], reverse=True)

    for sub_info in subs_list:
        sub_id = sub_info[0]
        sub_start_date = sub_info[2]
        sub_renew_date = sub_info[3]
        sub_renew_status = sub_info[4]
        sub_monthly_rev = sub_info[5]

        if current_status == 'NONE':
            current_status = sub_renew_status

        # Grab the next renewal date (lowest date in list)
        if sub_renew_date < next_renewal_date and sub_renew_status == 'ACTIVE':
            next_renewal_date = sub_renew_date
            next_renewal_rev = sub_monthly_rev
            days_to_renew = (next_renewal_date - today).days
            current_status = sub_renew_status

        # Find the the earliest date we ever billed this customer
        if sub_start_date < billing_start_date:
            billing_start_date = sub_start_date

        sub_summary = sub_id + ' - ' + \
            sub_renew_status + ' - ' + \
            sub_start_date.strftime("%m/%d/%Y") + '\t - ' + \
            sub_renew_date.strftime("%m/%d/%Y") + ' - ' + \
            str(days_to_renew) + ' - ' + \
            '${:,}'.format(round(sub_monthly_rev * 12)) + \
            ' \n' + sub_summary

    sub_summary = sub_summary[:-1]

    # Blank these fields if there is no subs
    if next_renewal_date == datetime(2050, 1, 1):
        next_renewal_date = ''
        days_to_renew = ''
    if billing_start_date == datetime(2050, 1, 1):
        billing_start_date = ''

    if len(subs_list) == 0:
        sub_summary = "No Subscription Info Found"
        current_status = ''

    return sub_summary, billing_start_date, next_renewal_date, days_to_renew, next_renewal_rev, current_status


def main():
    as_wb, as_ws = open_wb(app_cfg['XLS_AS_DELIVERY_STATUS'])
    cust_wb, cust_ws = open_wb(app_cfg['XLS_BOOKINGS'])
    sub_wb, sub_ws = open_wb(app_cfg['XLS_SUBSCRIPTIONS'])

    print()
    print('RAW Input Data')
    print("\tAS Fixed SKUs Rows:", as_ws.nrows)
    print('\tBookings Rows:', cust_ws.nrows)
    print('\tSubscription Rows:', sub_ws.nrows)

    #
    # Build a Team Dict
    #
    team_dict = build_coverage_dict()

    #
    # Create a SKU Filter
    #
    # Options Are: Product / Software / Service / SaaS / All SKUs
    sku_filter_val = 'All SKUs'
    tmp_dict = build_sku_dict()
    sku_filter_dict = {}

    for key, val in tmp_dict.items():
        if val[0] == sku_filter_val:
            sku_filter_dict[key] = val
        elif sku_filter_val == 'All SKUs':
            # Selects ALL Interesting SKUs
            sku_filter_dict[key] = val

    print()
    print('SKU Filter set to:', sku_filter_val)
    print()

    #
    # Build a xref dict of valid customer ids for lookup by SO and ERP Name
    #
    xref_cust_name = {}
    xref_so = {}
    for row_num in range(1, cust_ws.nrows):
        cust_id = cust_ws.cell_value(row_num, 15)
        cust_erp_name = cust_ws.cell_value(row_num, 13)
        cust_so = cust_ws.cell_value(row_num, 11)

        # Only add valid ID/Name Pairs to the reference
        if cust_id == '-999' or cust_id == '':
            continue

        if cust_erp_name not in xref_cust_name:
            xref_cust_name[cust_erp_name] = cust_id
            if (cust_so, cust_erp_name) not in xref_so:
                xref_so[(cust_so, cust_erp_name)] = cust_id

    #
    # Process Main Bookings File
    #
    cntr = 0
    cust_db = {}
    cust_alias_db = {}
    so_dict = {}

    #
    # Main loop over the bookings data starts here
    #
    for row_num in range(1, cust_ws.nrows):
        # Gather the fields we want
        cust_id = cust_ws.cell_value(row_num, 15)
        cust_erp_name = cust_ws.cell_value(row_num, 13)
        cust_ultimate_name = cust_ws.cell_value(row_num, 14)
        cust_so = cust_ws.cell_value(row_num, 11)
        cust_sku = cust_ws.cell_value(row_num, 19)
        cust_sales_lev_1 = cust_ws.cell_value(row_num, 3)
        cust_sales_lev_2 = cust_ws.cell_value(row_num, 4)
        cust_sales_lev_3 = cust_ws.cell_value(row_num, 5)
        cust_sales_lev_4 = cust_ws.cell_value(row_num, 6)
        cust_sales_lev_5 = cust_ws.cell_value(row_num, 7)
        cust_sales_lev_6 = cust_ws.cell_value(row_num, 8)
        cust_acct_mgr = cust_ws.cell_value(row_num, 9)

        # Grab this SO number in a simple dict {so:(cust_id, cust_id)
        if cust_so not in so_dict:
            # so_dict[cust_so] = ((cust_id, cust_erp_name),)
            so_dict[cust_so] = ((cust_id, cust_erp_name, cust_sku),)
        else:
            # so_dict[cust_so] = so_dict[cust_so] + ((cust_id, cust_erp_name),)
            so_dict[cust_so] = so_dict[cust_so] + ((cust_id, cust_erp_name, cust_sku),)

        # Do we have a missing or bad cust_id try to look one up
        if cust_id == '' or cust_id == '-999':
            if cust_erp_name in xref_cust_name:
                cust_id = xref_cust_name[cust_erp_name]

            if (cust_so, cust_erp_name) in xref_so:
                cust_id = xref_so[(cust_so, cust_erp_name)]

            # If id is still bad flag cust_id as UNKNOWN
            if cust_id == '' or cust_id == '-999':
                cust_id = 'UNKNOWN'

        #
        # Check cust_db
        # {cust_id: Customer_obj}
        #
        # Is this a new cust_id ?
        if cust_id not in cust_db:
            # Create a new cust_id object and basic record
            cust_db[cust_id] = Customer(cust_id)
            cust_db[cust_id].sales_lev_1 = cust_sales_lev_1
            cust_db[cust_id].sales_lev_2 = cust_sales_lev_2
            cust_db[cust_id].sales_lev_3 = cust_sales_lev_3
            cust_db[cust_id].sales_lev_4 = cust_sales_lev_4
            cust_db[cust_id].sales_lev_5 = cust_sales_lev_5
            cust_db[cust_id].sales_lev_6 = cust_sales_lev_6
            sales_level = cust_sales_lev_1 + ',' + cust_sales_lev_2 + ',' + cust_sales_lev_3 + ',' + \
                cust_sales_lev_4 + ',' + cust_sales_lev_5 + ',' + cust_sales_lev_6
            sales_team = find_team(team_dict, sales_level)
            pss = sales_team[0]
            tsa = sales_team[1]
            cust_db[cust_id].pss = pss
            cust_db[cust_id].tsa = tsa
            cust_db[cust_id].am = cust_acct_mgr

        # Is this a SKU we want if so add_order
        if cust_sku in sku_filter_dict:
            cust_db[cust_id].add_order(cust_so, cust_sku)

        # Add this customer_erp_name as an alias to the customer object
        cust_db[cust_id].add_alias(cust_erp_name)

        # Add this name to an easy alias lookup dict
        if cust_erp_name not in cust_alias_db:
            cust_alias_db[cust_erp_name] = cust_id

    print('Unique Customer IDs with filter of', " '" + sku_filter_val+"' :", len(cust_db))
    print("Customer Unique Customer Names: ", len(cust_alias_db))
    print("Unique Sales Order Numbers: ", len(so_dict))

    # A quick check on customer ids -
    id_list = [['Customer ID', 'Customer Aliases']]
    for cust_id, cust_obj in cust_db.items():
        alias_list = []
        alias_str = ''
        cust_aliases = cust_obj.aliases
        for cust_alias in cust_aliases:
            alias_list.append(cust_alias)
            alias_str = alias_str + cust_alias + ' : '
        alias_str = alias_str[:-3]
        id_list.append([cust_id, alias_str])

    push_list_to_xls(id_list, 'log_Unique_Cust_IDs.xlsx')

    #
    # Get SAAS Data from the BU Sheet
    #
    # saas_rows = get_list_from_ss(app_cfg['SS_SAAS'])
    # test_list = []
    #
    # for x in saas_rows:
    #     for y in x:
    #         print(type(y), y)
    #     time.sleep(.4)
    # exit()
    #
    # for row_num in range(1, len(saas_rows)):
    #     try:
    #         tmp_val = [saas_rows[row_num][1], str(int(saas_rows[row_num][2]))]
    #     except ValueError:
    #         tmp_val = ['Bad Data in row ' + str(row_num), saas_rows[row_num][1]]
    #
    #     test_list.append(tmp_val)
    # push_list_to_xls(test_list, 'saas_status.xlsx')
    #
    # saas_status_list = [['status', 'cust name', 'saas so', 'cust id']]
    # for row in test_list:
    #     saas_name = row[0]
    #     saas_so = row[1]
    #     saas_status = ''
    #     saas_cust_id = ''
    #
    #     if saas_name.find('Bad Data') != -1:
    #         saas_status = 'Bad Data in SaaS Sheet', saas_name, saas_so, saas_cust_id
    #     else:
    #         if saas_name in cust_alias_db:
    #             saas_cust_id = cust_alias_db[saas_name]
    #             saas_status = 'Matched Data with SaaS Sheet', saas_name, saas_so, saas_cust_id
    #         else:
    #             saas_status = 'No Matching Data from SaaS Sheet', saas_name, saas_so, saas_cust_id
    #
    #     saas_status_list.append(saas_status)
    # push_list_to_xls(saas_status_list, 'log_saas_data_matches.xlsx')


    # # Display Customer IDs and Aliases
    # for cust_id, cust_obj in cust_db.items():
    #     if len(cust_obj.aliases) > 1:
    #         print()
    #         print('Customer ID', cust_id, ' has the following aliases')
    #         for name in cust_obj.aliases:
    #             print('\t\t', name)
    #             time.sleep(1)

    # # Display Sales Order info
    # for cust_id, cust_obj in cust_db.items():
    #     if len(cust_obj.orders) > 1:
    #         print()
    #         print('Customer ID', cust_id, cust_obj.aliases, ' has the following orders')
    #         for my_order, my_skus in cust_obj.orders.items():
    #             print('\t', 'SO Num:', my_order, 'SKUs', my_skus)
    #             time.sleep(1)

    #
    # Process AS AS-F SKU File - match bookings SO and (AS SO / PID) numbers
    # and make a list of tuples for each cust_id
    #
    as_db = {}
    so_status_list = [['AS SO Number', 'AS Customer Name', "AS PID", 'Duplicate ?', 'Match in Booking ?']]
    as_zombie_so = []
    as_so_found_cntr = 0
    as_so_not_found_cntr = 0
    as_so_duplicate_cntr = 0
    as_so_unique_cntr = 0
    for row_num in range(1, as_ws.nrows):
        my_as_info_list = []
        # Gather the fields we want
        as_pid = as_ws.cell_value(row_num, 0)
        as_cust_name = as_ws.cell_value(row_num, 2)
        as_so = as_ws.cell_value(row_num, 19)

        # Just a check
        if as_so in as_db:
            dupe = 'Duplicate SO'
            as_so_duplicate_cntr += 1
        else:
            dupe = 'Unique SO'
            as_so_unique_cntr += 1

        if as_so not in as_db:
            my_as_info_list.append((as_pid, as_cust_name))
            as_db[as_so] = my_as_info_list
        else:
            my_as_info_list = as_db[as_so]
            add_it = True
            for info in my_as_info_list:
                if info == (as_pid, as_cust_name):
                    add_it = False
                    break
            if add_it:
                my_as_info_list.append((as_pid, as_cust_name))
                as_db[as_so] = my_as_info_list

        # Checks
        if as_so not in so_dict:
            so_status_list.append([as_so, as_cust_name, as_pid, dupe, 'NOT in Bookings'])
            as_zombie_so.append([as_so, as_cust_name, as_pid])
            as_so_not_found_cntr += 1
        else:
            so_status_list.append([as_so, as_cust_name, as_pid, dupe, 'FOUND in Bookings'])
            as_so_found_cntr += 1

    push_list_to_xls(so_status_list, 'log_AS SO_Status_List.xlsx')
    print('AS SO NOT Found (Zombies):', as_so_not_found_cntr)
    print('AS SO Found:', as_so_found_cntr)
    print('\t AS SO Totals:', as_so_found_cntr + as_so_not_found_cntr)
    print()
    print('AS SO Duplicate:', as_so_duplicate_cntr)
    print('AS SO Unique:', as_so_unique_cntr)
    print('len of as_db',len(as_db))

    #
    # Update the cust_db objects with the AS data from as_db
    #
    found_list = 0
    as_zombies = [['AS SO', 'AS PID', 'AS Customer Name', 'Possible Match', 'Ratio']]
    for as_so, as_info in as_db.items():
        # as_info is [so #:[(as_pid, as_cust_name),()]]
        as_cust_name = as_info[0][1]

        if as_so in so_dict:
            cust_id = so_dict[as_so][0][0]
            cust_obj = cust_db[cust_id]
            found_list = found_list + len(as_info)
            cust_obj.add_as_pid(as_so, as_info)
        else:
            # OK this AS_SO is NOT in the Main so_dict
            # We need to attempt to match on as_cust_name in the customer alias dict
            # We need to find the customer_id
            if as_cust_name in cust_alias_db:
                cust_id = cust_alias_db[as_cust_name]
                cust_obj = cust_db[cust_id]
                found_list = found_list + len(as_info)
                cust_obj.add_as_pid(as_so, as_info)
            else:
                # do a fuzzy match search against all customer aliases
                best_match = 0
                for k, v in cust_alias_db.items():
                    match_ratio = fuzz.ratio(as_cust_name, k)
                    if match_ratio > best_match:
                        possible_cust = k
                        best_match = match_ratio

                cust_id = cust_alias_db[possible_cust]
                cust_obj = cust_db[cust_id]
                found_list = found_list + len(as_info)
                cust_obj.add_as_pid(as_so, as_info)

                cust_obj.add_as_pid(as_so, as_info)

                as_zombies.append([as_so, as_info[0][0], as_info[0][1], possible_cust, best_match])
                print('\tNOT FOUND Customer ID for: ', as_cust_name)

    push_list_to_xls(as_zombies, 'tmp_zombies.xlsx')
    print('Updated cust_db with: ', found_list, ' AS SOs')

    #
    # Process Subscriptions and add to Customer Objects
    #
    for row_num in range(1, sub_ws.nrows):
        # Gather the fields we want
        sub_cust_name = sub_ws.cell_value(row_num, 2)
        sub_id = sub_ws.cell_value(row_num, 4)
        sub_start_date = sub_ws.cell_value(row_num, 6)
        sub_renew_date = sub_ws.cell_value(row_num, 8)
        sub_renew_status = sub_ws.cell_value(row_num, 5)
        sub_monthly_rev = sub_ws.cell_value(row_num, 10)

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_start_date, sub_wb.datemode)
        sub_start_date = datetime(year, month, day)

        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_renew_date, sub_wb.datemode)
        sub_renew_date = datetime(year, month, day)

        if sub_cust_name in cust_alias_db:
            cust_id = cust_alias_db[sub_cust_name]
            cust_obj = cust_db[cust_id]
            sub_info = [sub_id, sub_cust_name, sub_start_date, sub_renew_date, sub_renew_status, sub_monthly_rev]
            cust_obj.add_sub_id(sub_info)

    #
    # Make the Magic List
    #
    magic_list = []
    header_row = ['Customer ID', 'AS SO', 'AS PID', 'AS Customer Name', 'Sales Level 1', 'Sales Level 2', 'PSS', 'TSA',
                  'AM', 'Subscription History' + ' \n' +
                  'Sub ID - Start Date - Renewal Date - Days to Renew - Annual Rev',
                  'Sub 1st Billing Date', 'Next Renewal Date',
                  'Days to Renew', 'Next Renewal Monthly Rev', 'Sub Current Status', 'AS Delivery Mgr', 'AS Tracking Status',
                  'AS Tracking Sub Status', 'AS Tracking Comments', 'AS SKU',
                  'AS Project Creation Date', 'AS Project Start Date', 'AS Scheduled End Date',
                  'Days from 1st Sub Billing to AS Project Start']
    magic_list.append(header_row)
    print (magic_list)
    x = 0
    today = datetime.today()

    for cust_id, cust_obj in cust_db.items():
        cust_aliases = cust_obj.aliases
        as_pids = cust_obj.as_pids
        sub_ids = cust_obj.subs
        pss = cust_obj.pss
        tsa = cust_obj.tsa
        am = cust_obj.am
        sales_lev1 = cust_obj.sales_lev_1
        sales_lev2 = cust_obj.sales_lev_2

        if len(as_pids) == 0:
            # No AS PID info available
            sub_summary, billing_start_date, next_renewal_date, days_to_renew, renewal_rev, sub_renew_status = process_sub_info(cust_obj.subs)
            magic_row = [cust_id, '', 'AS Info Unavailable', cust_aliases[0], sales_lev1, sales_lev2, pss, tsa, am,
                         sub_summary, billing_start_date, next_renewal_date,
                         days_to_renew, renewal_rev, sub_renew_status, '', '', '', '', '', '', '', '', '']
            magic_list.append(magic_row)
        else:
            # Let's look at the AS PIDs in cust_obj
            for so, as_pid_info in as_pids.items():
                # We will make a row for each AS SO
                for as_detail in as_pid_info:
                    magic_row = []
                    as_so = so
                    as_pid = as_detail[0]
                    as_cust_name = as_detail[1]

                    sub_summary, billing_start_date, next_renewal_date, days_to_renew, renewal_rev, sub_renew_status = process_sub_info(cust_obj.subs)

                    # Go get additional AS Info
                    as_tracking_status = ''
                    as_tracking_sub_status = ''
                    as_tracking_comments = ''
                    as_dm = ''
                    as_project_start = ''
                    as_scheduled_end = ''
                    as_project_created = ''
                    as_sku = ''

                    for row_num in range(1, as_ws.nrows):
                        if as_pid == as_ws.cell_value(row_num, 0):
                            as_dm = as_ws.cell_value(row_num, 1)
                            as_tracking_status = as_ws.cell_value(row_num, 7)
                            as_tracking_sub_status = as_ws.cell_value(row_num, 8)
                            as_tracking_comments = as_ws.cell_value(row_num, 9)
                            as_sku = as_ws.cell_value(row_num, 14)
                            as_project_start = as_ws.cell_value(row_num, 26)
                            as_scheduled_end = as_ws.cell_value(row_num, 27)
                            as_project_created = as_ws.cell_value(row_num, 28)

                            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(as_project_start, as_wb.datemode)
                            as_project_start = datetime(year, month, day)

                            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(as_scheduled_end, as_wb.datemode)
                            as_scheduled_end = datetime(year, month, day)

                            year, month, day, hour, minute, second = xlrd.xldate_as_tuple(as_project_created, as_wb.datemode)
                            as_project_created = datetime(year, month, day)
                            break

                    if isinstance(billing_start_date, datetime) and isinstance(as_project_start, datetime):
                        time_to_service = billing_start_date - as_project_start
                    else:
                        time_to_service = ''

                    magic_row = [cust_id, so, as_pid, as_cust_name, sales_lev1, sales_lev2, pss, tsa, am,
                                 sub_summary, billing_start_date, next_renewal_date,
                                 days_to_renew, renewal_rev, sub_renew_status, as_dm, as_tracking_status,
                                 as_tracking_sub_status, as_tracking_comments, as_sku,
                                 as_project_created, as_project_start, as_scheduled_end, time_to_service]

                    magic_list.append(magic_row)

    # print(len(magic_list))
    # print(x)
    # for my_row in magic_list:
    #     for my_col in my_row:
    #         print (my_col, type(my_col))
    #     time.sleep(.1)
    # exit()
    push_list_to_xls(magic_list, 'magic.xlsx')


    #
    # Make a NEW customer list
    #
    cust_as_of = 201910
    new_cust_dict = {}
    new_cust_list = [['Booking Period', 'Customer ID', 'Customer Name', 'PSS', 'PSS email', 'TSA', 'TSA email', 'AM']]
    for row_num in range(1, cust_ws.nrows):
        cust_id = cust_ws.cell_value(row_num, 15)
        if cust_id in cust_db:
            booking_period = cust_ws.cell_value(row_num, 2)
            cust_name = cust_ws.cell_value(row_num, 13)
            pss = cust_db[cust_id].pss
            tsa = cust_db[cust_id].tsa
            am = cust_db[cust_id].am
            if int(cust_ws.cell_value(row_num, 2)) >= cust_as_of:
                new_cust_list.append([booking_period, cust_id, cust_name, pss, tsa, am])

    push_list_to_xls(new_cust_list, 'tmp_New_Customer_list.xlsx')

    return


if __name__ == "__main__" :
    main()




#
#
# today = datetime.today()
# expired = []
# thirty_days = []
# sixty_days = []
# ninety_days = []
# ninety_plus = []
#
#
# for row_num in range(1, sub_ws.nrows):
#     # Gather the fields we want
#     sub_cust_name = sub_ws.cell_value(row_num, 2)
#     sub_id = sub_ws.cell_value(row_num, 4)
#     sub_status = sub_ws.cell_value(row_num, 5)
#     sub_start_date = sub_ws.cell_value(row_num, 6)
#     sub_renew_date = sub_ws.cell_value(row_num, 8)
#
#     if sub_cust_name in cust_id_db:
#         # Get the cust_id that matches this subscription name
#         sub_cust_id = cust_id_db[sub_cust_name]
#
#         # Go get a list of SOs for this cust_id
#         # Use this to find and AS engagements
#         my_so_dict = cust_db[sub_cust_id]
#         my_so_list = []
#         for so, skus in my_so_dict.items():
#             my_so_list.append(so)
#
#         # Go get a list of AS PIDs for these SO's
#         my_as_pids = []
#         for so in my_so_list:
#             if so in as_db:
#                 # Found an AS record
#                 as_info = as_db[so]
#                 as_pid = as_info[0][0]
#                 as_cust_name = as_info[0][1]
#                 my_as_pids.append(as_pid)
#             # else:
#             #     my_as_pids.append("NO AS Engagements Found !")
#     else:
#         # We can't find a match on this customer name
#         # Maybe check aliases ?
#         sub_cust_id = 'Unknown'
#         my_as_pids = ''
#         my_so_list = ''
#
#     print(sub_cust_id, sub_cust_name, 'have ', len(my_as_pids), ' PIDS')
#     print('\t\t',my_as_pids)
#     # print('\tSOs',my_so_list)
#     # print('\tAS PIDS', my_as_pids)
#     # print()
#     time.sleep(1)
#
#     year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_start_date, sub_wb.datemode)
#     sub_start_date = datetime(year, month, day)
#
#     year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_renew_date, sub_wb.datemode)
#     sub_renew_date = datetime(year, month, day)
#
#     days_to_renew = (sub_renew_date - today).days
#
#     #
#     # Bucket this customer renewal by age
#     #
#     if days_to_renew < 0:
#         expired.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
#     elif days_to_renew <= 30:
#         thirty_days.append([sub_cust_id, sub_cust_name, sub_id, sub_renew_date, days_to_renew, sub_status])
#     elif days_to_renew <= 60:
#         sixty_days.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
#     elif days_to_renew <= 90:
#         ninety_days.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
#     elif days_to_renew > 90:
#         ninety_plus.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
#         # print(ninety_plus)
#         # time.sleep(1)
#
# subs_total = len(expired)+len(thirty_days)+len(sixty_days)+len(ninety_days)+len(ninety_plus)
# print()
# print('Total Subscriptions: ',subs_total)
# print('\tExpired:', len(expired))
# print('\t30 days:', len(thirty_days))
# print('\t60 days:', len(sixty_days))
# print('\t90 days:', len(ninety_days))
# print('\t90+ days:', len(ninety_plus))
# print()
#
# print(header_row)
# thirty_days.insert(0, header_row)
#
#
#
#
# push_list_to_xls(thirty_days,'jim_subs.xlsx')
# print('sub hits', hit)
# print('sub miss', miss)
#
#
# cust_id_db

