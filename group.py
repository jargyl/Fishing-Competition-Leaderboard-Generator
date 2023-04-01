import random
import pandas as pd

# list of participants with their names and genders
participants = [("John", "M"), ("Samantha", "F"), ("Mike", "M"), ("Emma", "F"), ("David", "M"), ("Lisa", "F"), ("Jacob", "M"), ("Olivia", "F"), ("William", "M"), ("Ava", "F"), ("Ethan", "M"), ("Sophia", "F"), ("Michael", "M"), ("Isabella", "F"), ("Alexander", "M"), ("Mia", "F"), ("Daniel", "M"), ("Charlotte", "F"), ("Matthew", "M"), ("Amelia", "F"), ("Christopher", "M"), ("Emily", "F"), ("Joseph", "M"), ("Abigail", "F"), ("Andrew", "M"), ("Harper", "F"), ("Joshua", "M"), ("Madison", "F"), ("Benjamin", "M"), ("Elizabeth", "F"), ("William", "M"), ("Ella", "F"), ("David", "M"), ("Grace", "F"), ("Ryan", "M"), ("Chloe", "F"), ("James", "M"), ("Victoria", "F"), ("Samuel", "M"), ("Avery", "F")]

# shuffle the list of participants to get a random order
random.shuffle(participants)

# divide the participants into four groups
groups = [[] for _ in range(4)]
for i, (name, gender) in enumerate(participants):
    group_idx = i % 4  # assign participant to a group based on its index
    groups[group_idx].append((name, gender))

# divide groups with more than two women by gender
for i, group in enumerate(groups):
    num_women = sum(gender == 'F' for _, gender in group)
    if num_women > 2:
        group[:] = sorted(group, key=lambda x: x[1])  # sort by gender
        men = [(name, gender) for name, gender in group if gender == 'M']
        women = [(name, gender) for name, gender in group if gender == 'F']
        group[:] = men + women

# create an Excel spreadsheet with four sheets, one for each group
with pd.ExcelWriter('groups.xlsx') as writer:
    for i, group in enumerate(groups):
        df = pd.DataFrame(group, columns=['Name', 'Gender'])
        df['KG'] = ''  # add empty 'KG' column
        df.to_excel(writer, sheet_name=f'Group {i+1}', index=False)
