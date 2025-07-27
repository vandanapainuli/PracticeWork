
import pandas as pd
import os  
from datetime import datetime


CM_FILE = "DEMO_CM.xlsx"
OUTPUT_DIR = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Listing07.xlsx")

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
    'CMTRT': 'Medication or Therapy',
    'CMINDC': 'Indication',
    'CMINDDSC': "If indication is 'Prophylaxis'",
    'CMSTDAT': 'CM Start Date',
    'CMENDAT': 'CM End Date',
    'CMSEQ': 'Record position',

    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message'
}

def pull_cm_data():
    return pd.read_excel(CM_FILE, engine="openpyxl")

def processing(df):
    for col in COLUMN_WITH_LABELS:
        if col not in df.columns:
            df[col] = ""
    
    # trying to remove duplicate columns, keeping the first occurrence
    df = df.loc[:, ~df.columns.duplicated()]
    
    df['SITEID'] = "A000"
    df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d')
    df['FORMID'] = df['DOMAIN']
    df['FORMSEQ'] = df['CMSEQ']
    df['CHKREF'] = "CM Recon"
    df['VISITSEQ'] = "NA"
    df['SITENAME'] = "Remote"
    df['SUBJID'] = df['USUBJID']
    df['VISITID'] = "CM"
    df['REVINST'] = "If indication is reported as Prophylaxis or Other, then the Indication specification should not contain the term Prophylaxis.  "
    df['MSG'] = "Indication is reported as Prophylaxis or Other; however, Indication specification contains the term Prophylaxis. Please check "
    
    return df

def logicprocessing(df):
    # Normalize text columns to uppercase for consistent comparison
    df['CMINDC'] = df['CMINDC'].str.upper()
    df['CMINDDSC'] = df['CMINDDSC'].str.upper()

    # Apply the new logic
    filtered_df = df[
        (df['CMINDC'].isin(['PROPHYLAXIS', 'OTHER'])) &
        (df['CMINDDSC'].str.contains('PROPHYLAXIS', na=False))
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
        
        cm_df = pull_cm_data()
                     
        # Step 2: Process the merged data
        processed_df = processing(cm_df)
        
        # Step 3: Apply logic to filter rows
        filtered_df = logicprocessing(processed_df)
        
        # Step 4: Final processing and renaming columns
        final_df = final_processing(filtered_df)
        # Step 5: Write to a new Excel file
        write_to_excel(final_df)

if __name__ == "__main__":
    main()
    print("Script executed successfully.")
        
    