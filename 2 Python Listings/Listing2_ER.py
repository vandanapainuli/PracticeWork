import pandas as pd
import os  
from datetime import datetime


MH_FILE = "DEMO_MH.xlsx"
CM_FILE = "DEMO_CM.xlsx"
OUTPUT_DIR = r"C:\Users\u819970\Downloads\TEST\PracticeWork\2 Python Listings\Listings Output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Listing02.xlsx")

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
    'CMSEQ': 'CM Recordposition',
    'CMTRT': 'Medication or Therapy',
    'CMINDC': 'Indication',
    'CMSTDAT': 'CM Start Date',
    'CMENDAT': 'End Date',
    'MHDSL1': 'MH#-Term_1',
    'MHDSL2': 'MH#-Term_2',
    'MHDSL3': 'MH#-Term_3',
    'MHDSL4': 'MH#-Term_4',
    'MHDSL5': 'MH#-Term_5',
    'MHSEQ': 'MH Recordposition',
    'MHCAT': 'MH Category',
    'MHSTDAT': 'MH Start Date',
    'MHENDAT': 'End Date',
    'REVINST': 'Reviewer Instructions',
    'MSG': 'Message'
    }
def pull_cm_data():
    return pd.read_excel(CM_FILE, engine="openpyxl")
def pull_mh_data():
    return pd.read_excel(MH_FILE, engine="openpyxl")

def data_merge(df_cm, df_mh):
    # Merge with suffixes to detect duplicates
    merged_df = pd.merge(df_cm, df_mh, on="USUBJID", how="outer", suffixes=('', '_dup'))

    # Remove all columns that are duplicates (i.e., have the '_dup' suffix)
    merged_df = merged_df.loc[:, ~merged_df.columns.str.endswith('_dup')]

    return merged_df

def processing(df):
    for col in COLUMN_WITH_LABELS:
        if col not in df.columns:
            df[col] = ""

    df = df.loc[:, ~df.columns.duplicated()]
    df['SITEID'] = "A000"
    df['EXEC_DTTM'] = datetime.today().strftime('%Y-%m-%d')
    df['FORMID'] = df['DOMAIN']
    df['FORMSEQ'] = df['CMSEQ']
    df['CHKREF'] = "CM-MH Recon"
    df['VISITSEQ'] = "NA"
    df['SITENAME'] = "Remote"
    df['SUBJID'] = df['USUBJID']
    df['VISITID'] = "GENERAL CONMED"
    df['REVINST'] = "If Indication for the Medication is reported as 'Medical History', then the  Medication Start and End  Date/Time should be consistent with the corresponding Medical History Start and End Date. "
    msg_str="Indication in Concomitant Therapy form is reported as 'Medical History'. However, concomitant therapy start date ({CMSTDAT}) is prior to the Medical history date ({MHSTDAT}) or the concomitant therapy end date({CMENDAT}) is after the Medical history end date. Please correct else clarify. "
    df['MSG'] = df.apply(lambda x: msg_str.format(CMSTDAT=x['CMSTDAT'], MHSTDAT=x['MHSTDAT'], CMENDAT=x['CMENDAT']), axis=1)
    
    return df

def logicprocessing(df):
    mhdsl_cols = ['MHDSL1', 'MHDSL2', 'MHDSL3', 'MHDSL4', 'MHDSL5']
                                    
        # Ensure all date columns are converted to datetime
    date_cols = ['CMSTDAT', 'CMENDAT', 'MHSTDAT', 'MHENDAT']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors='coerce')

                # Condition 1: CMINDC is 'Medical History'
    condition_1 = df['CMINDC'].str.upper() == 'MEDICAL HISTORY'
                
                # Condition 2: At least one MHDSL column is not null or blank
    condition_2 = ~((df[mhdsl_cols].isnull() | (df[mhdsl_cols] == '')).all(axis=1))
                
                # Condition 3: Date logic
    condition_3 = (
                    (df['CMSTDAT'] < df['MHSTDAT']) |
                    (df['CMSTDAT'] > df['MHENDAT']) |
                    (df['CMENDAT'] > df['MHENDAT'])
                    )
                
                # Combine all conditions
    filtered_df = df[condition_1 & condition_2 & condition_3]
                
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
        mh_df = pull_mh_data()
        cm_df = pull_cm_data()
        
        # Step 1: Merge the dataframes
        merged_df = data_merge(cm_df, mh_df)
        
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
        
    