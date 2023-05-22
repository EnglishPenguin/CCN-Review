import datetime as dt
import os
import win32com.client as win32
import pandas as pd
import shutil

def run():
    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'
    CROSSWALK_FILE = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/RetrievalDescriptionCrosswalk.csv'

    today = dt.date.today()
    # file_date = today - dt.timedelta(days=1)
    # if today is Monday
    if today.weekday() == 0:
        # set FILE_DATE to today - 3
        file_date = today - dt.timedelta(days=3)
    else:
        # set FILE_DATE to today - 1
        file_date = today - dt.timedelta(days=1)

    last_day_of_month = (file_date.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    if file_date == last_day_of_month:
        # backtrack to the previous day until it's a business day
        while True:
            file_date -= dt.timedelta(days=1)
            if file_date.weekday() < 5:
                break

    fd_mmddyyyy = file_date.strftime('%m%d%Y')
    fd_mm_yyyy = file_date.strftime('%m %Y')
    file_year = file_date.strftime('%Y')

    # Source file locations for template files
    file_location_1 = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/'
    file1 = file_location_1 + 'CCN Output Checker.xlsb'
    file2 = file_location_1 + 'DP Comments Template.xlsx'

    # Destination file location for template files to be copied to
    file_location_2 = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/'

    # Copy template files from source to destination
    shutil.copy2(file1, file_location_2)
    shutil.copy2(file2, file_location_2)

    file_to_review = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/DP Comments Template.xlsx'
    output_file = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/Northwell_ChargeCorrection_Output_{fd_mmddyyyy}.xls'
    rep_submission_file = f'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Inputs/{file_year}/{fd_mm_yyyy} Inputs.xlsx'

    df_dp_export = pd.read_excel(file_to_review, sheet_name="export", engine="openpyxl")
    df_cross = pd.read_csv(CROSSWALK_FILE)
    df_output = pd.read_excel(output_file, sheet_name="export", engine="xlrd")
    df_output = df_output[['INVNUM', 'PAYER', 'CRN#', 'InvBal', 'CPT', 'RevisedCPTList', 'InvoiceDOS', 'OriginalLocation', 'NewLocation', 'OriginalDX','NewDX', 'DxPointers', 'OriginalModifier', 'NewModifier', 'TXN', 'ActionAddRemoveReplace', 'StatusID', 'RetrievalStatus', 'RetrievalDescription']]
    df_rep_list = pd.read_excel(rep_submission_file, sheet_name="Sheet1", engine="openpyxl")

    df_dp_export_final = pd.DataFrame(columns=df_dp_export.columns)
    df_dp_export_final = pd.concat([df_dp_export, df_output], ignore_index=True)
    df_output = pd.merge(df_output,df_cross, on="RetrievalDescription", how="left")
    df_rep_list = df_rep_list.rename(columns={'Invoice': 'INVNUM'})
    df_output= pd.merge(df_output,df_rep_list[['INVNUM', 'Rep Username', 'Rep Name','Supervisor','Department']],on='INVNUM',how='left')
    df_output.fillna(value="", inplace=True)

    with pd.ExcelWriter(file_to_review) as writer:
        df_dp_export_final.to_excel(writer, sheet_name="export", engine="openpyxl", index=False)
        df_output.to_excel(writer, sheet_name="Sheet1", engine="openpyxl", index=False)


if __name__ == '__main__':
    run()