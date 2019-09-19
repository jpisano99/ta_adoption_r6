#
# Get the Acct team in team_dict covering sales_level
#


def find_team(team_dict, sales_level):
    # Look for the team(s) with the longest match on the territory
    # sales_level formatted as:
    # 'Americas,US COMMERCIAL,COMMERCIAL SOUTH AREA,South West Select Operation,
    #                                           GULF STATES SELECT,GULF STATES SELECT 6,'
    #
    longest_match = 0
    team = []
    for k, v in team_dict.items():
        if sales_level.startswith(k, 0, len(k)):
            if len(k) >= longest_match:
                pss = []
                tsa = []
                longest_match = len(k)
                team = v
                for team in v:
                    pss.append(team[0])
                    tsa.append(team[1])

    if not team:
        team = ('None assigned', 'None assigned')

    return team
