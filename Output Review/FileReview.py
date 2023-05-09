import datetime as dt
import os
import win32com.client as win32
import pandas as pd

def run():
    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'
    CROSSWALK_FILE = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/RetrievalDescriptionCrosswalk.csv'

    today = dt.date.today()
    file_date = today - dt.timedelta(days=1)
    fd_mmddyyyy = file_date.strftime('%m%d%Y')
    fd_mm_yyyy = file_date.strftime('%m %Y')
    file_year = file_date.strftime('%Y')
    fd_mm_dd = file_date.strftime('%m/%d')

    file_to_review = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/DP Comments Template.xlsx'
    output_file = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/Northwell_ChargeCorrection_Output_{fd_mmddyyyy}.xls'
    ccn_checker = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/CCN Output Checker.xlsb'
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
    df_output= pd.merge(df_output,df_rep_list[['INVNUM','Rep Name','Supervisor','Department']],on='INVNUM',how='left')
    df_output.fillna(value="", inplace=True)

    with pd.ExcelWriter(file_to_review) as writer:
        df_dp_export_final.to_excel(writer, sheet_name="export", engine="openpyxl", index=False)
        df_output.to_excel(writer, sheet_name="Sheet1", engine="openpyxl", index=False)


if __name__ == '__main__':
    run()