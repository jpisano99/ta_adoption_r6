from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.settings import app_cfg


#
# Create a dict called sku_dict from the ta_skus smartsheet
#
def build_sku_dict():
    # Create an object from SmartSheets
    skus = app_cfg['SS_SKU']
    my_skus = Ssheet(skus)

    sku_dict = {}

    for row in my_skus.rows:
        sku = ''
        sku_type = ''
        sku_desc = ''
        sensor_count = 0
        x = 0
        for row_data in row['cells']:
            x += 1
            if x == 1:
                sku = row_data['value']
            elif x == 2:
                sku_type = row_data['value']
            elif x == 3:
                sku_desc = row_data['value']
            elif x == 4:
                sensor_count = row_data['value']

        sku_dict[sku] = [sku_type, sku_desc, sensor_count]

    #
    # the sku_dict is now created !!!
    #

    return sku_dict
