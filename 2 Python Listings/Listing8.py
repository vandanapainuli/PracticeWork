
import pandas as pd
import os  
from datetime import datetime


AE_FILE = "DEMO_AE.xlsx"
DM_FILE = "DEMO_DM.xlsx"
OUTPUT_DIR = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Listing08.xlsx")

COLUMN_WITH_LABELS = {
    'STUDYID': 'STUDYID',
    'SITEID': 'SITENUMBER',
    'SITENAME': 'SITE',
    'SUBJID': 'SUBJECT',
    'VISITID': 'FOLDERNAME',
    'VISITSEQ': 'VISITSEQ',
    'FORMID': 'FORMID',
    'FORMSEQ': 'FORMSEQ',
    'CHKREF': 'CHKREF',
    'EXEC_DTTM': 'Execution Date/Time',
    'AETERM': 'What is the adverse event term',
    'AESTDAT': 'Start Date',
    'AESER': 'Serious AE',
    'AEAGE': 'Age at onset of SAE',
    'AGEU': 'Age unit',
    'RFICDAT': 'Informed consent date',
    'AGE': "What is the subject's age",

    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message'
}

def pull_ae_data():
    return pd.read_excel(AE_FILE, engine="openpyxl")
def pull_dm_data():
    return pd.read_excel(DM_FILE, engine="openpyxl")


def data_merge(df_ae, df_dm):
    # Merge with suffixes to detect duplicates
    merged_df = pd.merge(df_ae, df_dm, on="USUBJID", how="outer", suffixes=('', '_dup'))

    # Remove all columns that are duplicates (i.e., have the '_dup' suffix)
    merged_df = merged_df.loc[:, ~merged_df.columns.str.endswith('_dup')]

    return merged_df



def processing(df):
    for col in COLUMN_WITH_LABELS:
        if col not in df.columns:
            df[col] = ""
    
    # trying to remove duplicate columns, keeping the first occurrence
    df = df.loc[:, ~df.columns.duplicated()]
    
    df['SITEID'] = "A000"
    df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d')
    df['FORMID'] = df['DOMAIN']
    df['FORMSEQ'] = df['AESEQ']
    df['CHKREF'] = "AE-DM Recon"
    df['VISITSEQ'] = "NA"
    df['SITENAME'] = "Remote"
    df['SUBJID'] = df['USUBJID']
    df['VISITID'] = "AESAE"
    df['REVINST'] = "'Age at onset of SAE' on AE form should be equal to Age reported in Demographics form. "
    df['MSG'] = "Age at onset of SAE on AE form is NOT equal to the Age reported in Demographics form. Please reconcile. "
    
    return df


def logicprocessing(df):
    # Ensure AEAGE and AGE are numeric (if applicable)
    df['AEAGE'] = pd.to_numeric(df['AEAGE'], errors='coerce')
    df['AGE'] = pd.to_numeric(df['AGE'], errors='coerce')

    # Apply the updated logic
    filtered_df = df[
        (df['AESER'].str.upper() == 'Y') &
        (df['AEAGE'].notna()) &
        (
            (df['AGE'].isna()) |  # AGE is null
            (df['AEAGE'] != df['AGE'])  # or values are different
        )
    ]
    
    return filtered_df


def final_processing(df):
    df_temp = df[list(COLUMN_WITH_LABELS.keys())]
    df_final = df_temp.rename(columns=COLUMN_WITH_LABELS)
    return df_final

def write_to_excel(df):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"Data has been written to {OUTPUT_FILE}")

def main():
        ae_df = pull_ae_data()
        dm_df = pull_dm_data()
        
        # Step 1: Merge the dataframes
        merged_df = data_merge(ae_df, dm_df)
        
        # Step 2: Process the merged data
        processed_df = processing(merged_df)
        
        # Step 3: Apply logic to filter rows
        filtered_df = logicprocessing(processed_df)
        
        # Step 4: Final processing and renaming columns
        final_df = final_processing(filtered_df)
        # Step 5: Write to a new Excel file
        write_to_excel(final_df)

if __name__ == "__main__":
    main()
    print("Script executed successfully.")
        
    