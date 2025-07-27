
import pandas as pd
import os  
from datetime import datetime


AE_FILE = "DEMO_AE.xlsx"
CM_FILE = "DEMO_CM.xlsx"
OUTPUT_DIR = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Listing09.xlsx")

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
    'AETERM': 'Adverse Event',
    'AESTDAT': 'Start Date',
    'AEENDTC': 'End Date',
    'AECONTRT': 'Concomitant treatment given for AE',
    'CMTRT': 'Medication or Therapy',
    'CMINDC': 'Indication',
    'CMSTDAT': 'CM Start Date',
    'AEDSL1': 'AE log line start date and term 1',
    'AEDSL2': 'AE log line start date and term 2',
    'AEDSL3': 'AE log line start date and term 3',
    'AEDSL4': 'AE log line start date and term 4',
    'AEDSL5': 'AE log line start date and term 5',
    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message'
}

def pull_ae_data():
    return pd.read_excel(AE_FILE, engine="openpyxl")
def pull_cm_data():
    return pd.read_excel(CM_FILE, engine="openpyxl")


def data_merge(df_ae, df_cm):
    # Merge with suffixes to detect duplicates
    merged_df = pd.merge(df_ae, df_cm, on="USUBJID", how="outer", suffixes=('', '_dup'))

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
    df['CHKREF'] = "AE-CM Recon"
    df['VISITSEQ'] = "NA"
    df['SITENAME'] = "Remote"
    df['SUBJID'] = df['USUBJID']
    df['VISITID'] = "AESAE"
    df['REVINST'] = "If Indication for this CM is 'Adverse Event', there must be an Adverse Event recorded where the question 'concomitant treatment?' is answered yes. "
    df['MSG'] = "Indication for this CM is 'Adverse Event', but there is no Adverse Event recorded where the question 'concomitant treatment?' is answered yes. Please check and confirm. "
    
    return df

def logicprocessing(df):
            aedsl_cols = ['AEDSL1', 'AEDSL2', 'AEDSL3', 'AEDSL4', 'AEDSL5']
            filtered_df = df[
                (df['CMINDC'] == 'ADVERSE EVENT') &
                (df['AECONTRT'] == 'Y') &
                (df[aedsl_cols].isnull() | (df[aedsl_cols] == '')).all(axis=1)
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
        cm_df = pull_cm_data()
        
        # Step 1: Merge the dataframes
        merged_df = data_merge(ae_df, cm_df)
        
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
        
    