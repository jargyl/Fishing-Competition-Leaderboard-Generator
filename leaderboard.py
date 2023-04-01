import pandas as pd

# read the 'leaderboard.xlsx' file into a Pandas DataFrame
df = pd.read_excel('groups.xlsx', sheet_name=None)

# create an empty dictionary to hold the scoreboards per group
scoreboards = {}

# iterate over the sheets (i.e., groups) in the DataFrame
for sheet_name, sheet_data in df.items():
    # group the participants by gender if necessary
    if (sheet_data['Gender'] == 'F').sum() > 2:
        sheet_data = sheet_data.sort_values(by='Gender')
        men = sheet_data[sheet_data['Gender'] == 'M']
        women = sheet_data[sheet_data['Gender'] == 'F']
        sheet_data = pd.concat([men, women], ignore_index=True)

    # sort the participants by their KG values in descending order
    sheet_data = sheet_data.sort_values(by='KG', ascending=False)

    # create a scoreboard for the male participants
    male_scoreboard = []
    for i, row in sheet_data.iterrows():
        if row['Gender'] == 'M':
            male_scoreboard.append((row['Name'], row['KG']))

    # create a scoreboard for the female participants if applicable
    if (sheet_data['Gender'] == 'F').sum() > 0:
        female_scoreboard = []
        for i, row in sheet_data.iterrows():
            if row['Gender'] == 'F':
                female_scoreboard.append((row['Name'], row['KG']))
    else:
        female_scoreboard = None

    # add the scoreboards to the dictionary for the group
    scoreboards[sheet_name] = {'male': male_scoreboard, 'female': female_scoreboard}

# create a new Excel file to store the scoreboards
writer = pd.ExcelWriter('scoreboards.xlsx', engine='xlsxwriter')

# iterate over the scoreboards dictionary and write each scoreboard to a sheet
for group, sb_dict in scoreboards.items():
    # create a new sheet for the group
    sheet = writer.book.add_worksheet(group)

    # write the male scoreboard to the sheet
    sheet.write(0, 0, 'Male')
    sheet.write(1, 0, 'Ranking')
    sheet.write(1, 1, 'Participant')
    sheet.write(1, 2, 'KG')
    for i, (name, kg) in enumerate(sb_dict['male']):
        sheet.write(i+2, 0, f'{i+1}')
        sheet.write(i+2, 1, f'{name}')
        sheet.write(i+2, 2, f'{kg} KG')

    # write the female scoreboard to the sheet if applicable
    if sb_dict['female']:
        sheet.write(len(sb_dict['male']) + 3, 0, 'Female')
        sheet.write(len(sb_dict['male']) + 4, 0, 'Ranking')
        sheet.write(len(sb_dict['male']) + 4, 1, 'Participant')
        sheet.write(len(sb_dict['male']) + 4, 2, 'KG')
        for i, (name, kg) in enumerate(sb_dict['female']):
            sheet.write(len(sb_dict['male']) + i + 5, 0, f'{i+1}')
            sheet.write(len(sb_dict['male']) + i + 5, 1, f'{name}')
            sheet.write(len(sb_dict['male']) + i + 5, 2, f'{kg} KG')

writer.close()
