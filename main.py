import os
import xlsxwriter
import pandas as pd

teamDict = {
    "Angels": "laa",
    "Astros": "hou",
    "Athletics": "oak",
    "Blue Jays": "tor",
    "Braves": "atl",
    "Brewers": "mil",
    "Cardinals": "stl",
    "Cubs": "chc",
    "Diamondbacks": "ari",
    "Dodgers": "lad",
    "Giants": "sf",
    "Guardians": "cle",
    "Mariners": "sea",
    "Marlins": "mia",
    "Mets": "nym",
    "Nationals": "wsh",
    "Orioles": "bal",
    "Padres": "sd",
    "Phillies": "phi",
    "Pirates": "pit",
    "Rangers": "tex",
    "Rays": "tb",
    "Red Sox": "bos",
    "Reds": "cin",
    "Rockies": "col",
    "Royals": "kc",
    "Tigers": "det",
    "Twins": "min",
    "White Sox": "chw",
    "Yankees": "nyy"
}
pitcherCols = ['Name', 'Games Played', 'Games Started', 'Quality Starts', 'Wins', 'Losses', 'Saves',
               'Holds', 'Innings Pitched', 'Hits', 'Earned Runs', 'Home Runs', 'Walks', 'Strikeouts',
               'Strikeouts per 9 Innings', 'Pitches per Start', 'Wins Above Replacement',
               'Walks + Hits per Innings Pitched', 'Earned Run Average']
batterCols = ['Name', 'Games Played', 'At Bats', 'Runs', 'Hits', 'Doubles', 'Triples',
              'Home Runs', 'Runs Batted In', 'Total Bases', 'Walks', 'Strikeouts', 'Stolen Bases', 'Batting Average',
              'On Base Percentage', 'Slugging Percentage', 'On Base + Slugging Percentage',
              'Wins Above Replacement']

print("This program will grab player stats and save them to an Excel Spreadsheet for each team.")
direct = input("Copy and paste the desired directory to save to: ")


for x in teamDict:
    print(x)
    writer = pd.ExcelWriter(direct + "\\" + x + '.xlsx', engine='xlsxwriter')
    workbook = writer.book

    pitchers2022 = pd.read_html('https://www.espn.com/mlb/team/stats/_/type/pitching/name/' + teamDict[x])
    batters2022 = pd.read_html('https://www.espn.com/mlb/team/stats/_/name/' + teamDict[x])

    for i, df in enumerate(pitchers2022):
        if i == 0:
            df.to_excel('p.xlsx', index=False)
        elif i == 1:
            df.to_excel('ps.xlsx', index=False)
    df = pd.concat(map(pd.read_excel, ['p.xlsx', 'ps.xlsx']), axis=1, ignore_index=True)
    df.columns = pitcherCols
    df.to_excel(writer, sheet_name='Pitchers', index=False)
    os.remove('p.xlsx')
    os.remove('ps.xlsx')

    for column in df:
        column_length = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Pitchers'].set_column(col_idx, col_idx, column_length * 1.20)

    for i, df2 in enumerate(batters2022):
        if i == 0:
            df2.to_excel('b.xlsx', index=False)
        elif i == 1:
            df2.to_excel('bs.xlsx', index=False)
    df2 = pd.concat(map(pd.read_excel, ['b.xlsx', 'bs.xlsx']), axis=1, ignore_index=True)
    df2.columns = batterCols
    df2.to_excel(writer, sheet_name='Batters', index=False)
    os.remove('b.xlsx')
    os.remove('bs.xlsx')

    for column in df2:
        column_length = max(df2[column].astype(str).map(len).max(), len(column))
        col_idx = df2.columns.get_loc(column)
        writer.sheets['Batters'].set_column(col_idx, col_idx, column_length * 1.2)
    writer.save()