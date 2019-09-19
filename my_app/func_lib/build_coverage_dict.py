from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.settings import app_cfg

#
# Create a dict called team_dict from the coverage smartsheet
#


def build_coverage_dict():
    # Create an object from SmartSheets
    coverage = app_cfg['SS_COVERAGE']
    my_coverage = Ssheet(coverage)

    team_dict = {}

    for row in my_coverage.rows:
        pss = ''
        tsa = ''
        x = 0
        key = ''
        for row_data in row['cells']:
            x += 1
            if x == 1:
                continue
            elif x == 2:
                pss = row_data['value']
            elif x == 3:
                tsa = row_data['value']
            elif x >= 4:
                key = key + row_data['value']+','

        key = key.replace('*', '')

        if len(key) == 0:
            key = '*'
        key = key.rstrip(',')

        # # Create a Dict of sales_levels with a List of team(s) covering each territory
        info = team_dict.get(key, [])
        info.append((pss, tsa))
        team_dict[key] = info

    #
    # the team_dict is now created !!!
    #

    return team_dict

