import pandas as pd
from datetime import datetime
import re
# pip3 install openpyxl
import os
import argparse

BASE_DIR = os.getcwd()
OUTPUT_PATH = BASE_DIR + '/Output'
CURRENT_TIMESTAMP = str(datetime.now()).replace(" ", "_")
parser = argparse.ArgumentParser(description='Take Excel workbook name as input.')

# Create folder if it does not exist
all_files_and_folders = [i for i in os.listdir(BASE_DIR)]
if 'Output' not in all_files_and_folders:
    os.mkdir(OUTPUT_PATH)

"""
Assumptions so far:
- No column
"""
# READ DATA FROM EXCEL SHEET
parser.add_argument('--filename', type=str, required=True)
args = parser.parse_args()
file_name_cleaned = args.filename.replace(" ", "_")
print("Processing file :", file_name_cleaned)
excel_workbook_name = os.getcwd() + '/' + args.filename
excel_workbook = excel_workbook_name
pd_excel_file = pd.ExcelFile(excel_workbook)
excel_sheet_names = pd_excel_file.sheet_names
# filtered_sheet_names = list(filter(lambda x: x.endswith(" consolidation"), excel_sheet_names))

#########################################################################################
#########################################################################################
#########################################################################################

df_plan = pd_excel_file.parse("Plan consolidation", header=[0, 1],
                              skipinitialspace=True, tupleize_cols=True)
df_plan.columns = df_plan.columns.map(lambda x: " | ".join(tuple(map(str, x))))
rename_by_plan = {}
for col in df_plan.columns:
    if 'Plan Item id' in col:
        rename_by_plan[col] = 'Item_ID'
    if 'Plan Item Name' in col:
        rename_by_plan[col] = 'Item_Name'
    if '#' in col:
        rename_by_plan[col] = '#'
df_plan.rename(columns=rename_by_plan, inplace=True)
df_plan['item_type'] = 'plan'

#########################################################################################
#########################################################################################
#########################################################################################

df_addon = pd_excel_file.parse("Addon consolidation", header=[0, 1],
                               skipinitialspace=True, tupleize_cols=True)
df_addon.columns = df_addon.columns.map(lambda x: " | ".join(tuple(map(str, x))))
rename_by_addon = {}
for col in df_addon.columns:
    if 'Addon Item id' in col:
        rename_by_addon[col] = 'Item_ID'
    if 'Addon Item Name' in col:
        rename_by_addon[col] = 'Item_Name'
    if '#' in col:
        rename_by_addon[col] = '#'
df_addon.rename(columns=rename_by_addon,
                inplace=True)
df_addon['item_type'] = 'addon'

#########################################################################################
#########################################################################################
#########################################################################################
df_charge = pd_excel_file.parse("Charge consolidation", header=[0, 1],
                                skipinitialspace=True, tupleize_cols=True)
df_charge.columns = df_charge.columns.map(lambda x: " | ".join(tuple(map(str, x))))
rename_by_charge = {}
for col in df_charge.columns:
    if 'Addon Item id' in col:
        rename_by_charge[col] = 'Item_ID'
    if 'Addon Item Name' in col:
        rename_by_charge[col] = 'Item_Name'
    if '#' in col:
        rename_by_charge[col] = '#'
df_charge.rename(columns=rename_by_charge,
                 inplace=True)
df_charge['item_type'] = 'charge'

# Get only the duplicates by Item id
df_by_itemid = pd.concat([
    df_plan.loc[:, ['#', 'Item_ID', 'item_type']],
    df_addon.loc[:, ['#', 'Item_ID', 'item_type']],
    df_charge.loc[:, ['#', 'Item_ID', 'item_type']],
])
duplicate_byid_df = df_by_itemid[df_by_itemid.duplicated(subset=["Item_ID"], keep=False)]
duplicate_byid_df = duplicate_byid_df[~duplicate_byid_df['Item_ID'].isna()].sort_values("Item_ID")
duplicate_byid_df.to_csv(OUTPUT_PATH + f"/Item_ID_duplicates_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                         index=False)

# Get only the duplicates by Item Name
df_by_itemname = pd.concat([
    df_plan.loc[:, ['#', 'Item_Name', 'item_type']],
    df_addon.loc[:, ['#', 'Item_Name', 'item_type']],
    df_charge.loc[:, ['#', 'Item_Name', 'item_type']],
])
duplicate_byname_df = df_by_itemname[df_by_itemname.duplicated(subset=["Item_Name"], keep=False)]
duplicate_byname_df = duplicate_byname_df[~duplicate_byname_df['Item_Name'].isna()].sort_values("Item_Name")
duplicate_byname_df.to_csv(OUTPUT_PATH + f"/Item_Name_duplicates_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                           index=False)

#########################################################################################
#########################################################################################
#########################################################################################
df_plan_cols = list(
    filter(
        lambda x: re.search(r"([A-Z]{3}|[A-Z]{3}\s)\-", x) is not None,
        df_plan.columns
    )
)

df_plan["Row_num"] = [i for i in range(df_plan.shape[0])]
df_plan_pc1_stack_raw = df_plan[df_plan_cols].stack().reset_index()
df_plan_pc1_stack_raw.columns = ['Row_num', 'Currency_frequency', 'Duplicate_Values']
df_plan_pc1_temp = df_plan_pc1_stack_raw.merge(df_plan[['Row_num', '#', 'item_type']], on='Row_num')
df_plan_pc1 = df_plan_pc1_temp[['#', 'Currency_frequency', 'Duplicate_Values', 'item_type']]

#########################################################################################
#########################################################################################
#########################################################################################
df_addon_cols = list(
    filter(
        lambda x: re.search(r"([A-Z]{3}|[A-Z]{3}\s)\-", x) is not None,
        df_addon.columns
    )
)
df_addon["Row_num"] = [i for i in range(df_addon.shape[0])]
df_addon_pc1_stack_raw = df_addon[df_addon_cols].stack().reset_index()
df_addon_pc1_stack_raw.columns = ['Row_num', 'Currency_frequency', 'Duplicate_Values']
df_addon_pc1_temp = df_addon_pc1_stack_raw.merge(df_addon[['Row_num', '#', 'item_type']], on='Row_num')
df_addon_pc1 = df_addon_pc1_temp[['#', 'item_type', 'Currency_frequency', 'Duplicate_Values']]

#########################################################################################
#########################################################################################
#########################################################################################
df_charge_cols = list(
    filter(
        lambda x: re.search(r"[A-Z]{3}", x) is not None,
        df_charge.columns
    )
)
df_charge["Row_num"] = [i for i in range(df_charge.shape[0])]
df_charge_pc1_stack_raw = df_charge[df_charge_cols].stack().reset_index()
df_charge_pc1_stack_raw.columns = ['Row_num', 'Currency_frequency', 'Duplicate_Values']
df_charge_pc1_temp = df_charge_pc1_stack_raw.merge(df_charge[['Row_num', '#', 'item_type']], on='Row_num')
df_charge_pc1 = df_charge_pc1_temp[['#', 'item_type', 'Currency_frequency', 'Duplicate_Values']]

# Calculating the total number of entities
calDF = pd.DataFrame([
    ['Plan Consolidation', df_plan_pc1.shape[0]],
    ['Addon Consolidation', df_addon_pc1.shape[0]],
    ['Charge Consolidation', df_charge_pc1.shape[0]]])
calDF.columns = ['Sheet Name', 'Count of entity IDs in sheet']
calDF.to_csv(OUTPUT_PATH + f"/Count_of_items_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv", index=False)

# Get duplicates of the entities
mergePC1DF = pd.concat([df_plan_pc1, df_addon_pc1, df_charge_pc1])
entityDuplicatesDF = mergePC1DF[mergePC1DF.duplicated(subset=["Duplicate_Values"], keep=False)].sort_values(
    "Duplicate_Values")
entityDuplicatesDF.loc[len(entityDuplicatesDF.index)] = ['#', 'Count of entity ids',
                                                         'Total number of entities for Plan Consolidation',
                                                         df_plan_pc1.shape[0]]
entityDuplicatesDF.loc[len(entityDuplicatesDF.index)] = ['#', 'Count of entity ids',
                                                         'Total number of entities for Addon Consolidation',
                                                         df_addon_pc1.shape[0]]
entityDuplicatesDF.loc[len(entityDuplicatesDF.index)] = ['#', 'Count of entity ids',
                                                         'Total number of entities for Charge Consolidation',
                                                         df_charge_pc1.shape[0]]
entityDuplicatesDF.to_csv(OUTPUT_PATH + f"/EntityID_duplicates_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                          index=False)
print("File has been processed. :)")
