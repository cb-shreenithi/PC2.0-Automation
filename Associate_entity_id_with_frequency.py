from datetime import datetime
import os
import argparse
import chargebee
import json
import pandas as pd
import warnings

warnings.simplefilter(action='ignore')

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
parser.add_argument('--instance_key', type=str, required=True)
parser.add_argument('--instance_name', type=str, required=True)
parser.add_argument('--filename', type=str, required=True)
args = parser.parse_args()
file_name_cleaned = args.filename.replace(" ", "_")
print("Processing file :", file_name_cleaned)
excel_workbook_name = os.getcwd() + '/' + args.filename
excel_workbook = excel_workbook_name
pd_excel_file = pd.ExcelFile(excel_workbook)
excel_sheet_names = pd_excel_file.sheet_names
# filtered_sheet_names = list(filter(lambda x: x.endswith(" consolidation"), excel_sheet_names))
chargebee.configure(args.instance_key, args.instance_name)

#########################################################################################
#########################################################################################
#########################################################################################
"""
Generate Plans DF from Instance
"""

entries_plan = chargebee.Plan.list()
plans_instance_list = [json.loads(str(i.plan)) for i in entries_plan if i.plan]
plans_instance_df = pd.DataFrame(plans_instance_list)
plans_instance_df_cleaned = plans_instance_df[plans_instance_df['status'].isin(["active", "archived"])]
plans_instance_subset_df = plans_instance_df_cleaned[["id", "period_unit", "period", "currency_code"]]
plans_instance_subset_df.columns = ["id_instance", "period_unit_instance",
                                    "period_instance", "currency_code_instance"]
plans_instance_filtered_DF = plans_instance_subset_df
plans_instance_filtered_DF['frequency_instance'] = plans_instance_filtered_DF['currency_code_instance'] + '-' \
                                                   + plans_instance_filtered_DF['period_instance'].apply(
    lambda x: str(x)) + '-' \
                                                   + plans_instance_filtered_DF['period_unit_instance']. \
                                                       apply(lambda x: x.capitalize())
plans_instance_DF = plans_instance_filtered_DF[['id_instance', 'frequency_instance']]

#########################################################################################

"""
Generate Plans DF from Excel sheet
"""

plans_excel_df = pd_excel_file.parse("Plan consolidation", header=[0, 1],
                                     skipinitialspace=True, tupleize_cols=True)
plan_excel_cons_currency_freq = plans_excel_df[["Item price points ('Plan id's in catalog 1.0)"]].stack().reset_index()
plan_excel_cons_currency_freq.columns = ['row_num_excel', 'currency_frequency_excel', 'price_entity_id_excel']
plan_excel_cons_currency_freq['currency_frequency_excel'] = plan_excel_cons_currency_freq['currency_frequency_excel']. \
    apply(lambda x: ''.join(x.split()))
plans_excel_curr_df = plan_excel_cons_currency_freq[['price_entity_id_excel', 'currency_frequency_excel']]

#########################################################################################

"""
Associate Plans DF from Instance with Excel DF
"""

plans_matched_rows = plans_instance_DF.merge(plans_excel_curr_df, left_on="id_instance",
                                             right_on="price_entity_id_excel", how="inner")
plans_matched_rows.to_csv(OUTPUT_PATH + f"/Matched_entity_ids_plan_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                          index=False)
plans_instance_left_excel = plans_instance_DF.merge(plans_excel_curr_df, left_on="id_instance",
                                                    right_on="price_entity_id_excel", how="left")
plans_unmatched_from_instance = plans_instance_left_excel[plans_instance_left_excel['price_entity_id_excel'].isna()]. \
    reset_index().drop('index', axis=1)
plans_instance_right_excel = plans_instance_DF.merge(plans_excel_curr_df, left_on="id_instance",
                                                     right_on="price_entity_id_excel", how="right")
plans_unmatched_from_excel = plans_instance_right_excel[plans_instance_right_excel['id_instance'].isna()].reset_index(). \
    drop('index', axis=1)
plans_unmatched = pd.concat([plans_unmatched_from_instance, plans_unmatched_from_excel]).reset_index(). \
    drop('index', axis=1)
plans_unmatched.to_csv(OUTPUT_PATH + f"/Unmatched_entity_ids_plan_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                       index=False)

#########################################################################################


#########################################################################################
#########################################################################################
#########################################################################################
"""
Generate Addons DF from Instance
"""

entries_addon = chargebee.Addon.list()
addons_instance_list = [json.loads(str(i.addon)) for i in entries_addon if i.addon]
addons_instance_df = pd.DataFrame(addons_instance_list)
addons_instance_df_cleaned = addons_instance_df[addons_instance_df['status'].isin(["active", "archived"])]
addons_instance_subset_df = addons_instance_df_cleaned[["id", "period_unit", "period", "currency_code"]]
addons_instance_subset_df.columns = ["id_instance", "period_unit_instance",
                                     "period_instance", "currency_code_instance"]
addons_instance_filtered_DF = addons_instance_subset_df
addons_instance_filtered_DF['frequency_instance'] = addons_instance_filtered_DF['currency_code_instance'] + '-' \
                                                    + addons_instance_filtered_DF['period_instance'].apply(
    lambda x: str(x)) + '-' \
                                                    + addons_instance_filtered_DF['period_unit_instance']. \
                                                        apply(lambda x: x.capitalize())
addons_instance_DF = addons_instance_filtered_DF[['id_instance', 'frequency_instance']]

#########################################################################################

"""
Generate Addons DF from Excel sheet
"""

addons_excel_df = pd_excel_file.parse("Addon consolidation", header=[0, 1],
                                      skipinitialspace=True, tupleize_cols=True)
addons_excel_cons_currency_freq = addons_excel_df[["Item price points ('Addon id's in catalog 1.0)"]].stack(). \
    reset_index()
addons_excel_cons_currency_freq.columns = ['row_num_excel', 'currency_frequency_excel', 'price_entity_id_excel']
addons_excel_cons_currency_freq['currency_frequency_excel'] = addons_excel_cons_currency_freq[
    'currency_frequency_excel']. \
    apply(lambda x: ''.join(x.split()))
addons_excel_curr_df = addons_excel_cons_currency_freq[['price_entity_id_excel', 'currency_frequency_excel']]

#########################################################################################

"""
Associate Addons DF from Instance with Excel DF
"""

addons_matched_rows = addons_instance_DF.merge(addons_excel_curr_df, left_on="id_instance",
                                               right_on="price_entity_id_excel", how="inner")
addons_matched_rows.to_csv(OUTPUT_PATH + f"/Matched_entity_ids_addon_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                           index=False)
addons_instance_left_excel = addons_instance_DF.merge(addons_excel_curr_df, left_on="id_instance",
                                                      right_on="price_entity_id_excel", how="left")
addons_unmatched_from_instance = addons_instance_left_excel[addons_instance_left_excel['price_entity_id_excel'].isna()]. \
    reset_index().drop('index', axis=1)
addons_instance_right_excel = addons_instance_DF.merge(addons_excel_curr_df, left_on="id_instance",
                                                       right_on="price_entity_id_excel", how="right")
addons_unmatched_from_excel = addons_instance_right_excel[
    addons_instance_right_excel['id_instance'].isna()].reset_index(). \
    drop('index', axis=1)
addons_unmatched = pd.concat([addons_unmatched_from_instance, addons_unmatched_from_excel]).reset_index(). \
    drop('index', axis=1)
addons_unmatched.to_csv(OUTPUT_PATH + f"/Unmatched_entity_ids_addon_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                        index=False)

#########################################################################################


#########################################################################################
#########################################################################################
#########################################################################################
"""
Generate Charges DF from Instance
"""

charges_instance_filtered_DF = addons_instance_df_cleaned[
    addons_instance_df_cleaned['charge_type'] == 'non_recurring']
charges_instance_subset_df = charges_instance_filtered_DF[["id", "period_unit", "period", "currency_code"]]
charges_instance_subset_df.columns = ["id_instance", "period_unit_instance",
                                      "period_instance", "currency_code_instance"]
charges_instance_filtered_DF = charges_instance_subset_df
charges_instance_filtered_DF['frequency_instance'] = charges_instance_filtered_DF['currency_code_instance'] + '-' \
                                                     + charges_instance_filtered_DF['period_instance'].apply(
    lambda x: str(x)) + '-' \
                                                     + charges_instance_filtered_DF['period_unit_instance']. \
                                                         apply(lambda x: x.capitalize())
charges_instance_DF = charges_instance_filtered_DF[['id_instance', 'frequency_instance']]

#########################################################################################

"""
Generate Charges DF from Excel sheet
"""

charges_excel_df = pd_excel_file.parse("Charge consolidation", header=[0, 1],
                                       skipinitialspace=True, tupleize_cols=True)
charges_excel_cons_currency_freq = plans_excel_df[
    ["Item price points ('Charge id's in catalog 1.0)"]].stack().reset_index()
charges_excel_cons_currency_freq.columns = ['row_num_excel', 'currency_frequency_excel', 'price_entity_id_excel']
charges_excel_cons_currency_freq['currency_frequency_excel'] = charges_excel_cons_currency_freq[
    'currency_frequency_excel']. \
    apply(lambda x: ''.join(x.split()))
charges_excel_curr_df = charges_excel_cons_currency_freq[['price_entity_id_excel', 'currency_frequency_excel']]

#########################################################################################

"""
Associate Charges DF from Instance with Excel DF
"""

charges_matched_rows = charges_instance_DF.merge(charges_excel_curr_df, left_on="id_instance",
                                                 right_on="price_entity_id_excel", how="inner")
charges_matched_rows.to_csv(OUTPUT_PATH + f"/Matched_entity_ids_charge_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                            index=False)
charges_instance_left_excel = charges_instance_DF.merge(charges_excel_curr_df, left_on="id_instance",
                                                        right_on="price_entity_id_excel", how="left")
charges_unmatched_from_instance = charges_instance_left_excel[
    charges_instance_left_excel['price_entity_id_excel'].isna()]. \
    reset_index().drop('index', axis=1)
charges_instance_right_excel = charges_instance_DF.merge(charges_excel_curr_df, left_on="id_instance",
                                                         right_on="price_entity_id_excel", how="right")
charges_unmatched_from_excel = charges_instance_right_excel[
    charges_instance_right_excel['id_instance'].isna()].reset_index(). \
    drop('index', axis=1)
charges_unmatched = pd.concat([charges_unmatched_from_instance, charges_unmatched_from_excel]).reset_index(). \
    drop('index', axis=1)
charges_unmatched.to_csv(OUTPUT_PATH + f"/Unmatched_entity_ids_charge_{file_name_cleaned}_{CURRENT_TIMESTAMP}.csv",
                         index=False)

#########################################################################################

print("Files have been processed. :)")