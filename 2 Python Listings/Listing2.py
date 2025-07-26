
import pandas as pd
import os  
from datetime import datetime

# Input files
ae_file = "AE_DEMO.xlsx"
cm_file = "CM_DEMO.xlsx"

# Output directory and file
output_dir = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
output_file = os.path.join(output_dir, "Listing1.xlsx")

# Step 1: Define the required columns and their labels
columns_with_labels = {
    'STUDYID': 'STUDYID',
    'SITEID': 'SITEID',
    'SITENAME': 'Site Name',
    'SUBJID': 'SUBJID',
    'VISITID': 'VISITID',
    'VISITSEQ': 'VISITSEQ',
    'FORMID': 'FORMID',
    'FORMSEQ': 'FORMSEQ',
    'CHKREF': 'CHKREF',
    'EXEC_DTTM': 'Execution Date/Time',
    'AETERM': 'What is the adverse event term',
    'AESTDAT': 'Start Date',
    'AEENDTC': 'End Date',
    'AESER': 'IS AE Serious?',
    'AEOUT': 'Outcome',
    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message',
}

# Step 2: Read and merge AE and CM datasets
ae_df = pd.read_excel(ae_file, engine="openpyxl")
cm_df = pd.read_excel(cm_file, engine="openpyxl")
merged_df = pd.concat([ae_df, cm_df], ignore_index=True)

# Step 3: Ensure all required columns exist
for col in columns_with_labels:
    if col not in merged_df.columns:
        merged_df[col] = ""

# Step 4: Assign default or calculated values
merged_df['SITEID'] = "A000"
merged_df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d')
merged_df['FORMID'] = merged_df['DOMAIN']
merged_df['FORMSEQ'] = merged_df['AESEQ']
merged_df['CHKREF'] = "Data Dump"
merged_df['VISITSEQ'] = "NA"
merged_df['SITENAME'] = "Remote"
merged_df['SUBJID'] = merged_df['USUBJID']
merged_df['VISITID'] = "AESAE"
merged_df['REVINST'] = "Review data for completeness "
merged_df['MSG'] = "Some data is missing, please review and update. Else Clarify with the site."

# Step 5: Reorder and rename columns
df_final = merged_df[list(columns_with_labels.keys())].rename(columns=columns_with_labels)

# Step 6: Save to a new Excel file
df_final.to_excel(output_file, index=False)

print(f"{output_file} has been created with merged data and formatted columns.")
