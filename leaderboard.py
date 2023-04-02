import pandas as pd
from decimal import Decimal

# read the 'leaderboard.xlsx' file into a Pandas DataFrame
df = pd.read_excel('groups.xlsx', sheet_name=None)

# create an empty dictionary to hold the leaderboards per group
leaderboards = {}

# iterate over the sheets (i.e., groups) in the DataFrame
for sheet_name, sheet_data in df.items():
    # group the participants by gender if necessary
    if (sheet_data['Gender'] == 'F').sum() > 2:
        sheet_data = sheet_data.sort_values(by='Gender')
        men = sheet_data[sheet_data['Gender'] == 'M']
        women = sheet_data[sheet_data['Gender'] == 'F']
        sheet_data = pd.concat([men, women], ignore_index=True)

        # convert the KG values to Decimal objects
    sheet_data['KG'] = sheet_data['KG'].apply(lambda x: Decimal(str(x).replace(',', '.')))

    # sort the participants by their KG values in descending order
    sheet_data = sheet_data.sort_values(by='KG', ascending=False)

    # create a leaderboard for the male participants
    male_leaderboard = []
    for i, row in sheet_data.iterrows():
        if row['Gender'] == 'M':
            male_leaderboard.append((row['Name'], row['KG']))

    # create a leaderboard for the female participants if applicable
    if (sheet_data['Gender'] == 'F').sum() > 0:
        female_leaderboard = []
        for i, row in sheet_data.iterrows():
            if row['Gender'] == 'F':
                female_leaderboard.append((row['Name'], row['KG']))
    else:
        female_leaderboard = None

    # add the leaderboards to the dictionary for the group
    leaderboards[sheet_name] = {'male': male_leaderboard, 'female': female_leaderboard}

# create a new Excel file to store the leaderboards
writer = pd.ExcelWriter('leaderboards.xlsx', engine='xlsxwriter')

# iterate over the leaderboards dictionary and write each leaderboard to a sheet
for group, sb_dict in leaderboards.items():
    # create a new sheet for the group
    sheet = writer.book.add_worksheet(group)

    # write the male leaderboard to the sheet
    sheet.write(0, 0, 'Male')
    sheet.write(1, 0, 'Ranking')
    sheet.write(1, 1, 'Participant')
    sheet.write(1, 2, 'KG')
    for i, (name, kg) in enumerate(sb_dict['male']):
        sheet.write(i + 2, 0, f'{i + 1}')
        sheet.write(i + 2, 1, f'{name}')
        sheet.write(i + 2, 2, f'{kg} KG')

    # write the female leaderboard to the sheet if applicable
    if sb_dict['female']:
        sheet.write(len(sb_dict['male']) + 3, 0, 'Female')
        sheet.write(len(sb_dict['male']) + 4, 0, 'Ranking')
        sheet.write(len(sb_dict['male']) + 4, 1, 'Participant')
        sheet.write(len(sb_dict['male']) + 4, 2, 'KG')
        for i, (name, kg) in enumerate(sb_dict['female']):
            sheet.write(len(sb_dict['male']) + i + 5, 0, f'{i + 1}')
            sheet.write(len(sb_dict['male']) + i + 5, 1, f'{name}')
            sheet.write(len(sb_dict['male']) + i + 5, 2, f'{kg} KG')

# create an empty DataFrame to hold the global leaderboard
global_leaderboard = pd.DataFrame(columns=['Participant', 'KG', 'Group'])

# iterate over the leaderboards dictionary and add each participant to the global leaderboard
for group, sb_dict in leaderboards.items():
    for name, kg in sb_dict['male'] + (sb_dict['female'] or []):
        global_leaderboard = pd.concat(
            [global_leaderboard, pd.DataFrame({'Participant': [name], 'KG': [kg], 'Group': [group]})],
            ignore_index=True)

# sort the global leaderboard by KG in descending order
global_leaderboard = global_leaderboard.sort_values(by='KG', ascending=False)

# write the global leaderboard to a sheet in the Excel file
global_leaderboard.to_excel(writer, sheet_name='Global Leaderboard', index=False)

writer.close()