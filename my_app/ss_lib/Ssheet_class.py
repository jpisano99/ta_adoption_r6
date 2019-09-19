from .smartsheet_basic_functions import *


ss_config = dict(
    SS_TOKEN=my_app.my_secrets.passwords["SS_TOKEN"]
    )


class Ssheet:

    ss_token = ss_config['SS_TOKEN']
    ss = smartsheet.Smartsheet(ss_token)


    def __init__(self, name, meta_data_only=False):
        self.name = name
        self.sheet = {}
        self.id = 0
        self.columns = {}
        self.rows = {}
        self.col_name_idx = {}
        self.col_id_idx = {}
        self.refresh(meta_data_only)

    def refresh(self, meta_data_only):
        self.sheet = ss_get_sheet(self.ss, self.name)
        self.id = self.sheet['id']

        # If this self.sheet does NOT exist the self.id will equal -1
        # so skip refreshing the following to avoid an error
        if self.id != -1:
            self.columns = ss_get_col_data(self.ss, self.id)
            # self.rows = ss_get_row_data(self.ss, self.id)
            self.col_name_idx = ss_col_name_idx(self.ss, self.columns)
            self.col_id_idx = ss_col_id_idx(self.ss, self.columns)
        if not meta_data_only:
            self.rows = ss_get_row_data(self.ss, self.id)

    def row_lookup(self, col_name, row_value):
        # Return a list of all row_ids
        # Where col_name contains row_value
        row_ids = []
        col_id = self.col_name_idx[col_name]
        for row in self.rows:
            cells = row['cells']
            for cell in cells:
                if cell['columnId'] == col_id:
                    cell_val = cell['value'] if 'value' in cell else ''  # In case cell has no value assign null
                    if cell_val == row_value:
                        row_ids.append(row['id'])
        return row_ids

    def get_rows(self):
        row_dict = {}
        for row in self.rows:
            row_record = {}
            for cell in row['cells']:
                raw_cell_val = cell['value'] if 'value' in cell else ''
                raw_col_name = self.col_id_idx[cell['columnId']]
                row_record[raw_col_name] = raw_cell_val
            row_dict[row['rowNumber']] = row_record
        return row_dict

    def add_rows(self, add_rows):
        ss_add_rows(self.ss, self.id, add_rows)
        return

    def del_rows(self, del_rows):
        ss_del_rows(self.ss, self.id, del_rows)
        return

    def add_cols(self, add_cols):
        ss_add_column(self.ss, self.id, add_cols)
        return

    def del_cols(self, del_cols):
        ss_del_column(self.ss, self.id, del_cols)
        return

    def mod_cell(self, col_id, row_dict):
        ss_mod_cell(self.ss, self.id, col_id, row_dict)
        return

    def create_sheet(self, name, _col_dict):
        # If our self.id was a -1 we can create a sheet
        # Just pass in a 'name' and a col_dict
        # Example(s):
        # my_columns = [{'primary': True, 'title': 'ERP Customer Name', 'type': 'TEXT_NUMBER'},
        #               {'title': 'End Customer Ultimate Name', 'type': 'TEXT_NUMBER'},
        #               {'title': 'col1', 'type': 'CHECKBOX', 'symbol': 'STAR'}]

        ss_create_sheet(self.ss, name, _col_dict)
        self.name = name
        self.refresh()
        return

    def __repr__(self):
        return "Ssheet('{}')".format(self.name)

    # def __iter__(self):
    #     return self
    #
    # def __next__(self):
    #     self.index += 1
    #     if self.index == len(self.sql_to_ss):
    #         raise StopIteration
    #     return self.sql_to_ss[self.index]


if __name__ == "__main__":
    my_ss = Ssheet('Tetration On-Demand POV Raw Data')
    print(my_ss)

    print("Sheet Data: ", my_ss.sheet)
    print(my_ss.sheet['id'])
    print("Columns: ", my_ss.columns)
    print("Column Dict: ", my_ss.col_dict)
    print("Row Data: ", my_ss.rows)
    exit()
    print('Row IDs:  ', my_ss.row_lookup('cisco_owner_name', 'Chris McHenry'))

    # Add Columns Example
    # my_cols = []
    # my_cols.append({
    #     'title': 'New Picklist Column 2',
    #     'type': 'PICKLIST',
    #     'options': [
    #         'First',
    #         'Second',
    #         'Third'],
    #     'index': 4})
    # my_cols.append({
    #     'title': 'New Date Column1',
    #     'type': 'DATE',
    #     'validation': True,
    #     'index': 4})
    # my_ss.add_cols(my_cols)
    # my_ss.refresh()

    # Delete Column Example
    # my_col_id = my_ss.col_dict['New Date Column']
    # my_ss.del_cols(my_col_id)
    # my_ss.refresh()

    # # Add Rows Example
    # my_col_id = my_ss.col_dict['cisco_owner_name']
    # my_col1_id = my_ss.col_dict['cisco_owner']
    #
    # cell_data = []
    # my_rows = []
    # cell_data.append({
    #     'column_id': my_col_id,
    #     'value': 'blanche',
    #     'strict': False})
    # cell_data.append({
    #     'column_id': my_col1_id,
    #     'value': 'stan',
    #     'strict': False})
    # my_rows.append(cell_data)
    #
    # cell_data = []
    # cell_data.append({
    #     'column_id': my_col1_id,
    #     'value': 'blanche',
    #     'strict': False})
    # my_rows.append(cell_data)
    #
    # my_ss.add_rows(my_rows)
    # # Call this to update our sheet object
    # my_ss.refresh()
    # print("Added Rows: ", len(my_ss.rows))

    # Delete Rows Example
    # print("Deleted # of Rows BEFORE: ", len(my_ss.rows))
    # print('Row IDs:  ', my_ss.row_lookup('cisco_owner', 'stan'))
    # rows_to_delete = my_ss.row_lookup('cisco_owner', 'stan')
    # my_ss.del_rows(rows_to_delete)
    # my_ss.refresh()
    # print("Deleted # of Rows AFTER: ", len(my_ss.rows))

    # Modify Cells Example
    my_row_ids = my_ss.row_lookup('cisco_owner_name', 'ang')
    my_col_id = my_ss.col_dict['company_name']
    new_val = 'client director'
    my_row_dict = {}
    for id in my_row_ids:
        my_row_dict[id] = new_val
    my_ss.mod_cell(my_col_id, my_row_dict)
    my_ss.refresh()

    exit()

    # RESPONSE DEBUG CODE - DO NOT DELETE
    # print(json.dumps(my_ss.rows, indent=2))
