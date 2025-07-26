import pandas as pd
import os  
from datetime import datetime


INPUT_FILE="DEMO_AE.xlsx"
OUTPUT_DIR=r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
OUTPUT_FILE=os.path.join(OUTPUT_DIR, "Listing1.xlsx")

COLUMN_WITH_LABELS = {
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

#Pull Data
def pull_data():
    return pd.read_excel(INPUT_FILE, engine="openpyxl")

def processing(df):
    # Step 3: Ensure all required columns exist (create if missing)
    for col in COLUMN_WITH_LABELS:
        if col not in df.columns:
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
    return df

def final_processing(df):
    df_temp=df[list(COLUMN_WITH_LABELS.keys())]
    df_final=df_temp.rename(columns=COLUMN_WITH_LABELS)
    return df_final

def write_to_excel(df):
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")

    print(f"Data has been written to {OUTPUT_FILE}")

def main():
    # Step 1: Read the Excel file
    AE_df = pull_data()

    # Step 2: Process the data
    processed_df = processing(AE_df)

    # Step 3: Final processing and renaming columns
    final_df = final_processing(processed_df)

    # Step 4: Write to a new Excel file
    write_to_excel(final_df)

if __name__ == "__main__":
    main()
    print("Script executed successfully.")

