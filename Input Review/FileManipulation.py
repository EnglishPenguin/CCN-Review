import pandas as pd
import numpy as np
from datetime import datetime
from dateutil.relativedelta import relativedelta


def run():
    def count_delimiter(delim, ref_column, new_column):
        df3[new_column] = 0
        # Iterate over each row of the dataframe
        for index, row in df3.iterrows():
            # Count the number of commas in the value of Column C for this row
            delim_in_row = str(row[ref_column]).count(delim)
            if delim_in_row >= 1:
                delim_in_row += 1
            # Add 1 to the count if the value is not empty but has zero commas
            if delim_in_row == 0 and isinstance(row[ref_column], str) and delim not in row[ref_column] and row[ref_column] != "":
                delim_in_row += 1
            # Add the count to the new column
            df3.at[index, new_column] = delim_in_row


    def is_orig_equal_new(orig_field, new_field):
        """
        Compares Orig_field to New_field and if they are equal, it clears the Orig_Field value.
        Otherwise, it keeps the orig_field value
        :param orig_field:
        :param new_field:
        """
        df3[f'{orig_field}'] = np.where(
            (df3[f'{orig_field}'] == df3[f'{new_field}']),
            "",
            df3[f'{orig_field}']
        )


    def is_orig_blank(new_field, orig_field):
        """
        If the orig_field is blank, it will clear the new_field value.
        Otherwise, it keeps the new_field value already present
        :param new_field:
        :param orig_field:
        """
        df3[f'{new_field}'] = np.where(
            (df3[f'{orig_field}'] == ""),
            "",
            df3[f'{new_field}']
        )

    # def count_delim( delim1, input3, input4, input5):


    currentTimeDate = datetime.today()
    currentFileDate = currentTimeDate + relativedelta(days=1)
    fd_mmddyyyy = currentFileDate.strftime('%m%d%Y')
    currentFileDate = currentFileDate.strftime('%m/%d/%Y')
    one_year_ago = datetime.today() - relativedelta(years=1) + relativedelta(days=1)
    one_year_ago = one_year_ago.strftime('%m/%d/%Y')
    print(currentFileDate)
    print(one_year_ago)

    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Audits - Files Sent to Bot'
    f = f"{FILE_PATH}/Northwell_ChargeCorrection_Input_{fd_mmddyyyy}.xlsx"
    df1 = pd.read_excel(f, sheet_name="Sheet1", engine="openpyxl")
    df2 = pd.read_excel(f, sheet_name="Sheet2", engine="openpyxl")
    df2 = df2.drop(
        [
            "BAR_B_TXN.SER_DT,",
            "PROV__1,",
            "LOC__2,",
            "BAR_B_INV.DX_ONE__3,",
            "DX_TWO__3,",
            "DX_THREE__3,",
            "DX_FOUR__3,",
            "DX_FIVE__3,",
            "DX_SIX__3,",
            "DX_SEVEN__3,",
            "DX_EIGHT__3,",
            "DX_NINE__3,",
            "DX_TEN__3,",
            "BAR_B_INV.DX_ELEVEN__3,",
            "BAR_B_INV.DX_TWELVE__3,",
            "TXN_NUM,",
            "PROC__2,",
            "MOD,",
            "BAR_B_TXN.DX_NUM,",
            "BAR_B_INV.CHG_CORR_FLAG,"
        ], axis=1
    )
    df3 = df1.merge(df2, how='left', on='Invoice')
    df3 = df3.drop_duplicates(subset='Invoice', keep='last')
    df3 = df3.astype(
        {
            'InvoiceBalance': 'float64',
            'BAR_B_INV.TOT_CHG,': 'float64',
            'INV_BAL,': 'float64',
            'STEP': 'int64'
        }
    )


    # if Input Invoice Balance != Query Inv Bal; Set Input Invoice Balance to Query Inv Bal value
    df3['InvoiceBalance'] = np.where(
        (df3['InvoiceBalance'] != df3['INV_BAL,']),
        df3['INV_BAL,'],
        df3['InvoiceBalance']
    )

    # If Invoice Balance = total charges and STEP = 3; set STEP = 2
    df3['STEP'] = np.where(
        (df3['InvoiceBalance'] == df3['BAR_B_INV.TOT_CHG,']) &
        (df3['STEP'] == 3),
        2,
        df3['STEP']
    )

    # If invoice balance != total charges and STEP = 2; set STEP = 3
    df3['STEP'] = np.where(
        (df3['InvoiceBalance'] != df3['BAR_B_INV.TOT_CHG,']) &
        (df3['STEP'] == 2),
        3,
        df3['STEP']
    )


    df3['Exclude DOS > 1 Year'] = np.where(
        (df3['BAR_B_INV.SER_DT,'] <= one_year_ago),
        "Exclude",
        ""
    )

    # If Insurance != Orig FSC and Step != 2; set Review to True
    df3['FSC Review'] = np.where(
        (df3['Insurance'] != df3['BAR_B_INV.ORIG_FSC__5,']) &
        (df3['STEP'] != 2),
        "Review",
        ""
    )

    df3['STEP'] = df3.apply(
        lambda row: 4 if (len(str(row['OriginalCPT'])) > len(str(row['NewCPT']))) and (row['STEP'] != 4) else row['STEP'],
        axis=1
    )

    # If Original Modifier is not null & New Modifier is null; set Review to True
    df3['Modifier Review'] = np.where(
        (df3['OriginalModifier'].notnull()) &
        (df3['NewModifier'].isnull()),
        "Review",
        ""
    )

    # If Original = Original CPT and no other LI changes are being made; Clear Original CPT
    df3['OriginalCPT'] = np.where(
        (df3['OriginalCPT'] == df3['NewCPT']) &
        (df3['DxPointers'].isnull()) &
        (df3['OriginalModifier'].isnull()) &
        (df3['NewModifier'].isnull()) &
        (df3['NewDOS'].isnull()),
        "",
        df3['OriginalCPT']
    )

    is_orig_equal_new('ProviderName', 'NewProvider')
    is_orig_blank('NewProvider', 'ProviderName')
    is_orig_equal_new('OriginalLocation', 'NewLocation')
    is_orig_blank('NewLocation', 'OriginalLocation')
    is_orig_blank('NewCPT', 'OriginalCPT')
    is_orig_equal_new('OriginalDX', 'NewDX')
    is_orig_blank('NewDX', 'OriginalDX')

    count_delimiter(delim=',', ref_column='OriginalDX', new_column='Original DX Count')
    count_delimiter(delim=',', ref_column='NewDX', new_column='New DX Count')

    df3['Max Pointer'] = df3['DxPointers'].apply(
        lambda x: int(max([int(i) for i in str(x).replace('|', ',').split(',') if i.isdigit()]))
        if any(i.isdigit() for i in str(x))
        and len([int(i) for i in str(x).replace('|', ',').split(',') if i.isdigit()]) > 0
        else 0
    )

    df3['DxPointer Review'] = df3.apply(lambda row:
                                '' if pd.isna(row['DxPointers']) or
                                        isinstance(row['DxPointers'], int) or
                                        '|' not in str(row['DxPointers'])
                                else 'Review' if str(row['DxPointers']).split('|')[0] == ''
                                    or row['Original DX Count'] != row['New DX Count']
                                else '',
                                axis=1)

    df3['Review Diagnosis'] = np.where(
        (df3['Max Pointer'] > df3['New DX Count']) |
        ((df3['New DX Count'] != df3['Original DX Count']) & (df3['Max Pointer'] == "")),
        "Review",
        ""
    )

    df3.to_clipboard(index=False)

if __name__ == '__main__':
    run()
