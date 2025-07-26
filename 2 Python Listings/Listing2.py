
import pandas as pd
import os  
from datetime import datetime

# Input files
ae_file = "DEMO_AE.xlsx"
cm_file = "DEMO_CM.xlsx"

# Output directory and file
output_dir = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
output_file = os.path.join(output_dir, "Listing09.xlsx")

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
    'AETERM': 'Adverse event term',
    'AESTDAT': 'Start Date',
    'AEENDTC': 'End Date',
    'AECONTRT':'Concomitant treatment given for AE',
#CM Variables
    'CMTRT': 'Medication or Therapy ',
    'CMINDC':'Indication',
    'CMSTDAT':'CM Start Date',
    'AEDSL1':'AE log line start date and term 1',
    'AEDSL2':'AE log line start date and term 2',
    'AEDSL3':'AE log line start date and term 3',
    'AEDSL4':'AE log line start date and term 4',
    'AEDSL5':'AE log line start date and term 5',     
    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message',
}

# Step 2: Read and merge AE and CM datasets
ae_df = pd.read_excel(ae_file, engine="openpyxl")
cm_df = pd.read_excel(cm_file, engine="openpyxl")

merged_df = pd.merge(ae_df, cm_df, on="USUBJID", how="outer")
#"outer" ensures all records from both datasets are included
#"inner" would only include records that match in both datasets

# Step 3: Ensure all required columns exist
for col in columns_with_labels:
    if col not in merged_df.columns:
        merged_df[col] = ""

# Step 4: Assign default or calculated values
merged_df['SITEID'] = "A000"
merged_df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d')
merged_df['FORMID'] = DEMO_AE['DOMAIN']
merged_df['FORMSEQ'] = DEMO_AE['AESEQ']
merged_df['CHKREF'] = "AE-CM Recon"
merged_df['VISITSEQ'] = "NA"
merged_df['SITENAME'] = "Remote"
merged_df['SUBJID'] = DEMO_AE['USUBJID']
merged_df['VISITID'] = "AESAE"
merged_df['REVINST'] = "If Indication for this CM is 'Adverse Event', there must be an Adverse Event recorded where the question 'concomitant treatment?' is answered yes. "
merged_df['MSG'] = "Indication for this CM is 'Adverse Event', but there is no Adverse Event recorded where the question 'concomitant treatment?' is answered yes. Please check and confirm. "


# Step 5: Reorder and rename columns
df_final = merged_df[list(columns_with_labels.keys())].rename(columns=columns_with_labels)

# Step 6: Save to a new Excel file
df_final.to_excel(output_file, index=False)

print(f"{output_file} has been created with merged data and formatted columns.")
