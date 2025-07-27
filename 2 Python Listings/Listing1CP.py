
import pandas as pd
import os  
from datetime import datetime

# Input and output file names
input_file = "DEMO_AE.xlsx"
#output_file = "Listing1_26JUL2025.xlsx"//if in the same folder
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

# Step 2: Read the Excel file
 AE_df= pd.read_excel(input_file, engine="openpyxl")

# Step 3: Ensure all required columns exist (create if missing)
for col in columns_with_labels:
    if col not in AE_df.columns:
        df[col] = ""

# Step 4: Assign default or calculated values
df['SITEID'] = "A000"
df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d, %H:%M')
df['FORMID'] = df['DOMAIN']
df['FORMSEQ'] = df['AESEQ']
df['CHKREF'] = "Data Dump"
df['VISITSEQ'] = "NA"
df['SITENAME'] = "Remote"
df['SUBJID']=df['USUBJID']
df['VISITID'] = "AESAE"
df['REVINST'] = "Review data for completeness "
df['MSG'] = "Some data is missing, please review and update. Else Clarify with the site."

# Step 5: Reorder and rename columns
#df_final = AE_df[list(columns_with_labels.keys())].rename(columns=columns_with_labels)
df_temp=AE_df[list(columns_with_labels.keys())]
df_final=df_temp.rename(columns=columns_with_labels)

# Step 6: Save to a new Excel file
df_final.to_excel(output_file, index=False)

print(f"{output_file} has been created with renamed columns and assigned values.")
