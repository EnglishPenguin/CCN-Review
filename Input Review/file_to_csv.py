import pandas as pd
import glob
from datetime import date, timedelta
from pathlib import Path


def run():
    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Audits - Files Sent to Bot'
    OUT_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'

    today = date.today()

    # if today is Friday
    if today.weekday() == 4:
        # set FILE_DATE to today + 3
        file_date = today + timedelta(days=3)
    else:
        # set FILE_DATE to today + 1
        file_date = today + timedelta(days=1)

    last_day_of_month = (file_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    if file_date == last_day_of_month:
        # advance to the next day until it's a business day
        while True:
            file_date += timedelta(days=1)
            if file_date.weekday() < 5:
                break

    fd_mmddyyyy = file_date.strftime('%m%d%Y')
    f = f"{FILE_PATH}/Northwell_ChargeCorrection_Input_{fd_mmddyyyy}.xlsx"
    df = pd.read_excel(f, sheet_name="Sheet3", engine="openpyxl")
    df = df[df['Exclude'] != 'Exclude']

    df = df.drop(columns= [
        'BAR_B_INV.SER_DT,', 
        'BAR_B_INV.TOT_CHG,', 
        'INV_BAL,', 
        'BAR_B_INV.ORIG_FSC__5,', 
        'BAR_B_INV.CORR_INV_NUM', 
        'Exclude DOS > 1 Year', 
        'FSC Review', 
        'Valid FSC', 
        'Modifier Review', 
        'Original DX Count', 
        'New DX Count', 
        'DxPointers Count', 
        'Max Pointer', 
        'DxPointers Null', 
        'DxPointers String', 
        'DxPointer Review', 
        'BAR_B_TXN_LI_PAY.PAY_CODE__2',
        'unique_paycode_count',
        'Exclude Multi Paycode',
        'Exclude'
        ])
    
    f_csv = f"Northwell_ChargeCorrection_Input_{fd_mmddyyyy}.csv"
    df.to_csv(f'{OUT_PATH}/{f_csv}', index=False)

if __name__ == '__main__':
    run()