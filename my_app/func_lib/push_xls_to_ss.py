import smartsheet
from my_app.ss_lib.smartsheet_basic_functions import ss_get_ws, ss_move_sheet
from my_app.settings import app_cfg, passwords
import os


def push_xls_to_ss(my_file, ss_name, run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir, my_file))
    print('PUSHING to ', path_to_run_dir, ' to SmartSheets>>>>>> ', ss_name)

    ws = app_cfg['SS_WORKSPACE']

    ss_token = passwords['SS_TOKEN']
    ss = smartsheet.Smartsheet(ss_token)

    # Import as an Excel sheet
    response = ss.Sheets.import_xlsx_sheet(
        path_to_run_dir,
        ss_name,  # sheet_name
        header_row_index=0)
    response_dict = response.to_dict()
    my_sheet_id = response_dict['data']['id']

    # Get the workspace id
    # Then move this sheet to the configured workspace
    ws_info = ss_get_ws(ss, ws)
    ss_move_sheet(ss, my_sheet_id, ws_info['id'])

    return
