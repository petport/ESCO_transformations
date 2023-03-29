# import excel files from the folder "data"

import os

import openpyxl
import pandas as pd
import numpy as np
import math

os.chdir("data")

# open all files in the folder and all the sheet of each file
ESCO_skills = []
ESCO_knowledge = []
for file in os.listdir():
    if "~$" in file:
        continue
    try:
        df = pd.read_excel(file, sheet_name=None)
    except ValueError:
        print("Error: ", file)

    ESCO_level_1 = []
    ESCO_level_2 = []
    ESCO_level_3 = []

    for sheet in df:
        # get sheet name
        sheet_name = sheet
        print(sheet_name)
        if "Matrix 1" in sheet_name:
            ESCO_level_1.append(df[sheet_name])
        elif "Matrix 2" in sheet_name:
            ESCO_level_2.append(df[sheet_name])
        elif "Matrix 3" in sheet_name:
            ESCO_level_3.append(df[sheet_name])

        # Get sheet data with df[sheet_name]
        pass
    if "Skills" in file:
        ESCO_skills.append(ESCO_level_1)
        ESCO_skills.append(ESCO_level_2)
        ESCO_skills.append(ESCO_level_3)
        ESCO_skills = pd.DataFrame(ESCO_skills)
    elif "Knowledge" in file:
        ESCO_knowledge.append(ESCO_level_1)
        ESCO_knowledge.append(ESCO_level_2)
        ESCO_knowledge.append(ESCO_level_3)
        ESCO_knowledge = pd.DataFrame(ESCO_knowledge)
a = 1
pass


def get_all_roles_from_skills(ESCO_skills):
    return


def filter_skill_by_occupation_across_all_levels(ESCO_skills):
    # To get "technical director" occupation
    # ESCO_skills[0][n]['Unnamed: 1'][1] with n=0,1,2
    # or ESCO_skills[0][0].iloc[1,1]
    return


# # Get occupation URI
# ESCO_skills[0][0].values[1][0]
#
# # Get occupation name
# ESCO_skills[0][0].values[1][1]
#
# # Get index of non-zero items from the technical director occupation
# index_of_non_zero = np.nonzero(ESCO_skills[0][0].values[1][2:])[0]
#
# # Get the values of non-zero items
# non_zero_values = [ESCO_skills[0][0].values[1][2:][i] for i in index_of_non_zero]
#
# # Get skill URI for the non-zero skills
# non_zero_skills_URI = [ESCO_skills[0][0].keys()[2:][i] for i in index_of_non_zero]
#
# # Get skill name for the non-zero skills
# non_zero_skills_names = [ESCO_skills[0][0][i][0] for i in non_zero_skills_URI]
# a = 1

def get_data_across_all_levels_for_one_occupation(data, occupation_index, ESCO_level, ISCO_level):
    index_of_non_zero = np.nonzero(data[ISCO_level][ESCO_level].values[occupation_index][2:])[0]
    non_zero_values = [data[ISCO_level][ESCO_level].values[occupation_index][2:][i] for i in index_of_non_zero]
    non_zero_skills_URI = [data[ISCO_level][ESCO_level].keys()[2:][i] for i in index_of_non_zero]
    # temp_arr = [ESCO_skills[0][ESCO_level].values[1][0], ESCO_skills[0][ESCO_level].values[1][1]]
    # temp_skills = [ESCO_skills[0][ESCO_level][i][0] for i in non_zero_skills_URI]

    first_row = [" ", " "]
    first_row_URI = non_zero_skills_URI
    first_row.extend(first_row_URI)

    second_row = [data[ISCO_level][ESCO_level].values[occupation_index][0],
                  data[ISCO_level][ESCO_level].values[occupation_index][1]]
    second_row_skills_names = [data[ISCO_level][ESCO_level][i][0] for i in non_zero_skills_URI]
    second_row.extend(second_row_skills_names)

    third_row = [" ", " "]
    third_row.extend(non_zero_values)
    final_item_list = [first_row, second_row, third_row]

    return final_item_list


def get_data_across_all_levels_for_all_occupations(data, ESCO_level, ISCO_level):
    all_occupations = []
    for i in range(1, len(data[ISCO_level][ESCO_level].values)):
        all_occupations.append(get_data_across_all_levels_for_one_occupation(data, i, ESCO_level, ISCO_level))
    return all_occupations


all_occupations_level_1_0 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 0, 0)
all_occupations_level_2_0 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 1, 0)
all_occupations_level_3_0 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 2, 0)

all_occupations_level_1_1 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 0, 1)
all_occupations_level_2_1 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 1, 1)
all_occupations_level_3_1 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 2, 1)

all_occupations_level_1_2 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 0, 2)
all_occupations_level_2_2 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 1, 2)
all_occupations_level_3_2 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 2, 2)

all_occupations_level_1_3 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 0, 3)
all_occupations_level_2_3 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 1, 3)
all_occupations_level_3_3 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 2, 3)

all_occupations_level_1_4 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 0, 4)
all_occupations_level_2_4 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 1, 4)
all_occupations_level_3_4 = get_data_across_all_levels_for_all_occupations(ESCO_skills, 2, 4)

# Same for knowledge
all_occupations_level_1_0_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 0, 0)
all_occupations_level_2_0_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 1, 0)
all_occupations_level_3_0_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 2, 0)

all_occupations_level_1_1_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 0, 1)
all_occupations_level_2_1_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 1, 1)
all_occupations_level_3_1_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 2, 1)

all_occupations_level_1_2_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 0, 2)
all_occupations_level_2_2_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 1, 2)
all_occupations_level_3_2_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 2, 2)

all_occupations_level_1_3_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 0, 3)
all_occupations_level_2_3_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 1, 3)
all_occupations_level_3_3_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 2, 3)

all_occupations_level_1_4_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 0, 4)
all_occupations_level_2_4_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 1, 4)
all_occupations_level_3_4_knowledge = get_data_across_all_levels_for_all_occupations(ESCO_knowledge, 2, 4)

pass

a = 1

os.chdir("..")
os.chdir("output")

writer = pd.ExcelWriter('Skills' + '.xlsx', options={'strings_to_urls': False})

for index, sheet_data in enumerate([all_occupations_level_1_0, all_occupations_level_2_0, all_occupations_level_3_0,
                                    all_occupations_level_1_1, all_occupations_level_2_1, all_occupations_level_3_1,
                                    all_occupations_level_1_2, all_occupations_level_2_2, all_occupations_level_3_2,
                                    all_occupations_level_1_3, all_occupations_level_2_3, all_occupations_level_3_3,
                                    all_occupations_level_1_4, all_occupations_level_2_4, all_occupations_level_3_4]):

    a = pd.DataFrame(pd.DataFrame(sheet_data).to_numpy().flatten())
    for i in range(2, len(a), 3):
        a[0][i] = [str(i) for i in a[0][i]]

    b = pd.DataFrame(np.zeros(len(a)), dtype="string")
    for idx, i in enumerate(range(len(a))):
        try:
            b[0][idx] = "''".join(a[0][i])
        except TypeError:
            if math.isnan(a[0][i][1]):
                a[0][i][1] = "N/A (dev comment)"
                b[0][idx] = ",".join(a[0][i])
        DBG_var = 1337

    # print(idx)
    # b = [",".join(item) for item in a[0]]

    # open excel file with pandas in order to add a sheet
    # open existing excel file with pandas in order to add a sheet

    # add sheet
    # pd.DataFrame(all_occupations_level_1_0).to_excel(writer, sheet_name='Test')
    b.to_excel(writer, sheet_name=str(index % 3 + 1) + '_' + str(index // 3))

    # save file csv
    # b.to_csv(str(index % 3 + 1) + '_' + str(index // 3) + '.csv', index=False, header=False

    # close writer
    # writer.save()

a = 1
writer.close()
# for i in range(2,10,3):
#   a[0][i] = [str(i) for i in a[0][i]]


writer = pd.ExcelWriter('Knowledge' + '.xlsx', options={'strings_to_urls': False})

for index, sheet_data in enumerate(
        [all_occupations_level_1_0_knowledge, all_occupations_level_2_0_knowledge, all_occupations_level_3_0_knowledge,
         all_occupations_level_1_1_knowledge, all_occupations_level_2_1_knowledge, all_occupations_level_3_1_knowledge,
         all_occupations_level_1_2_knowledge, all_occupations_level_2_2_knowledge, all_occupations_level_3_2_knowledge,
         all_occupations_level_1_3_knowledge, all_occupations_level_2_3_knowledge, all_occupations_level_3_3_knowledge,
         all_occupations_level_1_4_knowledge, all_occupations_level_2_4_knowledge,
         all_occupations_level_3_4_knowledge]):

    a = pd.DataFrame(pd.DataFrame(sheet_data).to_numpy().flatten())
    for i in range(2, len(a), 3):
        a[0][i] = [str(i) for i in a[0][i]]

    b = pd.DataFrame(np.zeros(len(a)), dtype="string")
    for idx, i in enumerate(range(len(a))):
        try:
            b[0][idx] = "''".join(a[0][i])
        except TypeError:
            if math.isnan(a[0][i][1]):
                a[0][i][1] = "N/A (dev comment)"
                b[0][idx] = ",".join(a[0][i])
        DBG_var = 1337

    # print(idx)
    # b = [",".join(item) for item in a[0]]

    # open excel file with pandas in order to add a sheet
    # open existing excel file with pandas in order to add a sheet

    # add sheet
    # pd.DataFrame(all_occupations_level_1_0).to_excel(writer, sheet_name='Test')
    b.to_excel(writer, sheet_name=str(index % 3 + 1) + '_' + str(index // 3))

    # save file csv
    # b.to_csv(str(index % 3 + 1) + '_' + str(index // 3) + '.csv', index=False, header=False

    # close writer
    # writer.save()

a = 1
writer.close()
