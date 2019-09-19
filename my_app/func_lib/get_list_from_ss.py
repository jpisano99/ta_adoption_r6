from my_app.ss_lib.Ssheet_class import Ssheet


#
# Get the SmartSheet Coverage as a simple list
#
def get_list_from_ss(my_ss):

    my_coverage = Ssheet(my_ss, False)

    col_names = []
    for column in my_coverage.columns:
        col_names.append(column['title'])

    # Add column names as first row
    my_rows = [col_names]
    # print(col_names)

    # Now add each row to the list
    for row in my_coverage.rows:
        idx = 0
        my_row = []
        for col_data in row['cells']:
            # this_col_name = col_names[idx]
            if 'value' in col_data:
                my_row.append(col_data['value'])
            else:
                my_row.append('')

            idx += 1

        my_rows.append(my_row)

    return my_rows
