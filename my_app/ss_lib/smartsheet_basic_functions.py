import smartsheet
import my_app.my_secrets
import json


def ss_get_sheet(ss, sheet_name):
    # Get Sheet meta-data
    response = ss.Sheets.list_sheets(include_all=True)
    sheets = response.data
    # print(json.dumps(response.to_dict(), indent=2))
    sheet_info_dict = {}
    for sheet in sheets:
        if sheet.name == sheet_name:
            sheet_info_dict.update(sheet.to_dict())

    # if the sheet_name was NOT found sheet_info_dict will be empty
    # so we return a -1 for the sheet id
    if not sheet_info_dict:
        sheet_info_dict.update({'id': -1})

    return sheet_info_dict


def ss_create_sheet(ss, sheet_name, col_dict):
    # All columns are now defined
    # Send off to Smartsheets to create the sheet
    sheet_spec = ss.models.Sheet({'name': sheet_name, 'columns': col_dict})
    response = ss.Home.create_sheet(sheet_spec)
    return response.result.id

# def ss_delete_sheet(ss, sheet_id):
#     response = ss.Sheets.delete_sheet(sheet_id)
#     return response

def ss_delete_sheet(ss, sheet_name):
    # Get Sheet meta-data
    response = ss.Sheets.list_sheets(include_all=True)
    sheets = response.data
    sheets_deleted = 0
    for sheet in sheets:
        if sheet.name == sheet_name:
            print('Deleted ', sheet.name, sheet.id)
            sheets_deleted += 1
            response = ss.Sheets.delete_sheet(sheet.id)
    return sheets_deleted


def ss_get_template(ss, template_name):
    # Get template meta-data
    response = ss.Templates.list_user_created_templates()
    templates = response.data
    template_info_dict = {}
    for template in templates:
        if template.name == template_name:
            template_info_dict = template.to_dict()
    return template_info_dict


def ss_get_col_data(ss, sheet_id):
    # Get column data from sheet_id
    tmp = ss.Sheets.get_sheet(sheet_id, include='rowIds')
    # tmp = ss.Sheets.get_sheet(sheet_id, include='format')
    sheet_dict = tmp.to_dict()
    columns = sheet_dict.get('columns', {})
    # print(json.dumps(tmp.to_dict(), indent=2))
    return columns


def ss_get_ws(ss, ws_name):
    response = ss.Workspaces.list_workspaces(include_all=True)
    workspaces = response.data
    # print(json.dumps(response.to_dict(), indent=2))

    workspace_info_dict = {}
    for workspace in workspaces:
        if workspace.name == ws_name:
            workspace_info_dict.update(workspace.to_dict())
    return workspace_info_dict


def ss_move_sheet(ss, id, dest_id):
    sheet = ss.Sheets.move_sheet(id,
                ss.models.ContainerDestination({
                    'destination_type': 'workspace',       # folder, workspace, or home
                    'destination_id': dest_id  # folder_id
                  }))
    return


def ss_col_name_idx(ss, columns):
    # Return a dict of {col_name:col_id}
    col_name_idx = {}
    for column in columns:
        # key = column['title']
        # value = column['id']
        col_name_idx[column['title']] = column['id']
    return col_name_idx


def ss_col_id_idx(ss, columns):
    # Return a dict of {col_id:col_name}
    col_id_idx = {}
    for column in columns:
        col_id_idx[column['id']] = column['title']
    return col_id_idx


def ss_get_row_data(ss, sheet_id):
    # Get row data from sheet_id
    tmp = ss.Sheets.get_sheet(sheet_id, include='rowIds')
    sheet_dict = tmp.to_dict()
    rows = sheet_dict.get('rows', {})
    return rows


def ss_del_column(ss, sheet_id, my_cols):
    ss.Sheets.delete_column(sheet_id, my_cols)
    return


def ss_add_column(ss, sheet_id, my_cols):
    # Add columns to the sheet
    col_list = []
    for col in my_cols:
        col_list.append(ss.models.Column(col))

    response = ss.Sheets.add_columns(sheet_id, col_list)
    return


def ss_del_rows(ss, sheet_id, my_rows):
    ss.Sheets.delete_rows(sheet_id, my_rows)
    return


def ss_add_rows(ss, sheet_id, my_rows):
    # my_rows is a list of rows to be added
    # { [{"strict": false, "columnId": 1, "value": "jim"}] }
    rows_to_add = []
    for row in my_rows:
        # Create a row object
        row_next = ss.models.Row()
        row_next.to_top = True
        for cell in row:
            # Gather all the cells in __this__ row
            row_next.cells.append(cell)

        # Add to the list of row objects to send to SS
        rows_to_add.append(row_next)

    # Send to SS to ADD the rows
    response = ss.Sheets.add_rows(sheet_id, rows_to_add)
    return


def ss_mod_cell(ss, sheet_id, col_id, my_row_dict):
    # Take a dict of row_ids:new_val
    # and replace in the col_id provided
    new_rows =[]
    for row_id, row_val in my_row_dict.items():
        new_cell = ss.models.Cell()
        new_cell.column_id = col_id
        new_cell.value = row_val
        new_cell.strict = False

        # Build the row to update
        new_row = ss.models.Row()
        new_row.id = row_id
        new_row.cells.append(new_cell)
        new_rows.append(new_row)

    # Update rows
    updated_row = ss.Sheets.update_rows(sheet_id, new_rows)

    return


if __name__ == "__main__":
    ss_token = my_secrets.passwords['SS_TOKEN']
    ss = smartsheet.Smartsheet(ss_token)
    # sheet_dict = ss_get_sheet(ss, 'Tetration On-Demand POV Status')
    # print(sheet_dict)
    # print('TIME MOD:', sheet_dict['modifiedAt'])

    sheet_name = 'jim'
    col_dict = [{
            'title': 'col1',
            'type': 'CHECKBOX',
            'symbol': 'STAR'
        }, {
            'title': 'col2',
            'primary': True,
            'type': 'TEXT_NUMBER'
        }
        ]
    print (col_dict)

    ss_create_sheet(ss, sheet_name, col_dict)