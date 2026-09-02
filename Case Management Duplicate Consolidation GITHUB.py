'''
PART 1
The Duplicate Entity model calculates the probability of first name, last name, and date of birth to help users
assess risks of updating and consolidating client entities.

See documentation for more details
'''
import os
import pandas as pd
import openpyxl as ox

ref_dir = 'C:/Users/ema/Documents/Dupe Probability Model/Reference'
dupe_dir = 'C:/Users/ema/Documents/Dupe Probability Model/Inputs'
write_dir = 'C:/Users/ema/Documents/Dupe Probability Model/Outputs'


######VVVVVV### Input File Names Below ###VVVVVV######
ref_file = 'Ref Name and DOB Benchmark 8.20.26.csv'
dupe_file = 'Dupe NYSID Consolidation 8.20.26.xlsx'
write_file = 'Dupe Consolidation Entry Worksheet 8.20.26.xlsx'

ref_path = os.path.join(ref_dir, ref_file)
dupe_path = os.path.join(dupe_dir, dupe_file)
write_path = os.path.join(write_dir, write_file)

test_path = 'C:/Users/ema/Documents/Dupe Probability Model/Outputs/Dupe_test.xlsx'
######/\/\/\### Input File Names Above ###/\/\/\######


###########################################################
##################### Format Ref Data #####################
###########################################################
#read file
name_dob_data = pd.read_csv(
    ref_path,
    header=None,
    names=['Name_DOB','Freq_Name_DOB']
)

#regroup and sum standardized names
name_dob_data = (
    name_dob_data
    .groupby('Name_DOB', as_index=False)['Freq_Name_DOB']
    .sum()
)

# calculate name_dob denominator for probability
name_dob_total = name_dob_data['Freq_Name_DOB'].sum()

name_dob_data['Probability_Occurrence'] = (name_dob_data['Freq_Name_DOB']/name_dob_total)

#f_name_data.to_excel(test_path)

###################################################################
##################### Format Duplicate Entity #####################
###################################################################

#read file
dupe_data = pd.read_excel(dupe_path)

#strip upper case name
dupe_data['client_intake_name'] = (
    dupe_data['client_intake_name']
    .str.strip()
    .str.upper()
)

#split names
dupe_data[['l_name','f_name']] = (
    dupe_data['client_intake_name']
    .str.split(', ', n=1, expand=True)
)

#standardize DOB
dupe_data['date_of_birth'] = pd.to_datetime(
    dupe_data['date_of_birth'].replace('NULL', pd.NA),
    errors="coerce"
)

#fill NA
dupe_data['date_of_birth'] = (
    dupe_data['date_of_birth']
    .dt.strftime('%Y-%m-%d')
    .fillna('NULL')
)

#Create new field to join with reference
dupe_data['Name_DOB'] = dupe_data['l_name'] + ', ' + dupe_data['f_name'] + ', ' + dupe_data['date_of_birth']

##################################################################
##################### Attach to DE Worksheet #####################
##################################################################

#merge name_DOB probability
dupe_data_merge = dupe_data.merge(name_dob_data, how='left', on='Name_DOB')

#create odds
dupe_data_merge['One_in_N'] = (
    1 / dupe_data_merge['Probability_Occurrence']
)

dupe_data_merge['Occurrence'] = (
    dupe_data_merge['Probability_Occurrence']
    .apply(
        lambda p: f"1 in {1 / p:,.0f}"
        if pd.notna(p) and p > 0
        else ''
    )
)

#Create Filter to remove returning matches with same entities as reference, or null values
def filter_list(row):
    client_key = row['client_entity_key']
    match_key = row['Name_DOB_Match_Entity_key']

    if pd.isna(match_key):
        return 'remove'

    if str(match_key).strip().upper() == 'NULL':
        return 'remove'

    if str(client_key).strip() == str(match_key).strip():
        return 'remove'

    return 'keep'

dupe_data_merge['Filter'] = dupe_data_merge.apply(
    filter_list,
    axis=1
)

#apply filter and subset data
dupe_data_merge = dupe_data_merge.loc[dupe_data_merge['Filter']=='keep']

#Trim data file
dupe_data_merge = pd.DataFrame(data=dupe_data_merge,
                               columns=[
                                   'client_nysid_nbr',
                                   'Name_DOB',
                                   'client_entity_key',
                                   'docket_number',
                                   'practice_office_name',
                                   'arrest_number',
                                   'init_top_charge',
                                   'Probability_Occurrence',
                                   'Occurrence',
                                   'Name_DOB_Match_Entity_key',
                                   'name_dob_nysid',
                                   'User_Confirm NYSID',
                                   'User_Confirm Entity',
                                   'User_Notes',
                                   'Additional Entities'
                               ]
)

#Write File
dupe_data_merge.to_excel(write_path, index=False)



'''
PART 2
This script formats and merges files for account consolidation and summarizes results in a pivot table.
From here, the user can choose which entity key to survive after consolidation

1) Run the SQL reporting script
2) Place the file in the inputs folder
3) Refer to the data entry worksheet to merge odds ratio data
4)
'''

import os
import pandas as pd
import openpyxl as ox

#Open and read file paths
new_dir = 'C:/Users/ema/Documents/Cleaning DUP Consolidation/Inputs'
write_dir = 'C:/Users/ema/Documents/Cleaning DUP Consolidation/Outputs'
prob_dir = 'C:/Users/ema/Documents/Dupe Probability Model/Outputs'

###VVVVVV### Input File Names Below ###VVVVVV###
new_file = 'Dupe NYSID Consolidation 8.20.26.xlsx'
prob_file = 'Dupe Consolidation Entry Worksheet 8.20.26.xlsx'
write_file = 'Cleaning Dup Nysid Report 8.20.26v1.xlsx'

new_file_path = os.path.join(new_dir, new_file)
prob_path = os.path.join(prob_dir, prob_file)
write_path = os.path.join(write_dir, write_file)
###/\/\/\### Input File Names Above ###/\/\/\###

#######################################################
##################### Format Data #####################
#######################################################

#Read File
new_df = pd.read_excel(new_file_path)
prob_data = pd.read_excel(prob_path)


#format name
new_df['client_name'] = new_df['client_name'].str.upper()

#replace Nulls
new_df['practice_area'] = new_df['practice_area'].replace('NULL', 'Other Practice').fillna('Other Practice')
new_df['matched_client_dob'] = new_df['matched_client_dob'].replace('NULL', 'Unknown').fillna('Unknown')

#read probability set from worksheet
prob_data = pd.read_excel(prob_path)

#trim worksheet data
prob_data = pd.DataFrame(data=prob_data, columns=['client_entity_key','Occurrence'])

#remove dupes
prob_data = prob_data.drop_duplicates(keep='first')

#merge in the odds ratio
new_df = new_df.merge(prob_data, how='left', left_on='matching_entity', right_on='client_entity_key')

group_columns = [
    'client_name',
    'matching_entity',
    'practice_area',
    'matched_client_dob',
    'Occurrence'
]

new_df_pivot = (
    new_df
    .groupby(
        group_columns,
        dropna=False,
        observed=True
    )['case_count']
    .sum()
    .to_frame()
)

#write
new_df_pivot.to_excel(
    write_path,
    merge_cells=True
)

