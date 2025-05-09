import pandas as pd
import glob
from datetime import date, timedelta
from pathlib import Path
import os
import win32com.client as win32
import numpy as np
from dateutil.relativedelta import relativedelta
from pathlib import Path
from input_logger_setup import logger
from tkinter import simpledialog as sd
from tkinter import messagebox as mb

# Global Variables
FILE_PATH = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Audits - Files Sent to Bot'
OUT_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'
ARS_FILE_PATH = "M:/CPP-Data/AR SUPPORT/SPECIAL PROJECTS/CHARGE CORRECTION BOT/SPREADSHEETS TO SEND TO BOT"
P1_FILE_PATH = "M:/CPP-Data/Payor 1/Bot CCN"
P2_FILE_PATH = "M:/CPP-Data/Payer 2/BOTS/Charge Correction Files"
REP_SUPE_CROSSWALK = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/Report To.xlsx"
INPUTS_PATH = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Inputs"
FSC_GRID = 'M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/FSCs that accept electronic CCL.xlsx'

class Date_Functions:
    def __init__(self) -> None:
        self.today = date.today()
        # self.today = date(2024, 12, 24)

    def run(self):
        self.get_file_date()
        self.fd_mmddyyy = self.get_mmddyyyy()
        self.fd_mm = self.get_mm()
        self.fd_yyyy = self.get_yyyy()
        self.fd_slashes = self.get_mm_dd_yyyy_slashes()
        self.fd_spaces = self.get_mm_dd_yyyy_spaces() 

    def get_file_date(self):
        logger.info(f"today is {self.today}")
        if self.today.weekday() == 4:
            logger.info(f"today is a friday, setting date to coming monday")
            self.file_date = self.today + timedelta(days=3)
        else:
            logger.info(f"today is not a friday, advancing date by 1")
            self.file_date = self.today + timedelta(days=1)

        last_day_of_month = (self.file_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        if self.file_date == last_day_of_month:
            while True:
                self.file_date += timedelta(days=1)
                if self.file_date.weekday() < 5:
                    break
        logger.info(f"file date is set to: {self.file_date}")
        return self.file_date
        
    def get_mmddyyyy(self):
        logger.info(f"mmddyyyy: {self.file_date.strftime('%m%d%Y')}")
        return self.file_date.strftime('%m%d%Y')
    
    def get_mm(self):
        logger.info(f"mm: {self.file_date.strftime('%m')}")
        return self.file_date.strftime('%m')
    
    def get_yyyy(self):
        logger.info(f"yyyy: {self.file_date.strftime('%Y')}")
        return self.file_date.strftime('%Y')
    
    def get_mm_dd_yyyy_slashes(self):
        logger.info(f"mm/dd/yyyy: {self.file_date.strftime('%m/%d/%Y')}")
        return self.file_date.strftime('%m/%d/%Y')
    
    def get_mm_dd_yyyy_spaces(self):
        logger.info(f"mm dd yyyy: {self.file_date.strftime('%m %d %Y')}")
        return self.file_date.strftime('%m %d %Y')


class File_Combine(Date_Functions):
    def __init__(self, fd_mmddyyy, fd_mm, fd_yyyy, fd_spaces, fd_slashes):
        super().__init__()
        global FILE_PATH, OUT_PATH, ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH, REP_SUPE_CROSSWALK, INPUTS_PATH, FSC_GRID
        # variables
        self.fd_no_spaces = fd_mmddyyy
        self.month = fd_mm
        self.year = fd_yyyy
        self.fd_spaces = fd_spaces
        self.fd_slashes = fd_slashes
        self.outlook = win32.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNameSpace("MAPI")
        # strings
        self.B16_FILE_PATH = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/New vs Established/Formatted Inputs/{month} {year}"
        # lists
        self.open_file_list = []
        self.open_file_name_list = []
        self.user_name_list = []
        self.df_list = []
        self.rep_name_list = []
        self.trunc_file_list = []
        self.empty_file_list = []
        self.empty_f_user_list = []
        self.supe_list = None
        self.email_list = []
        self.supe_email_list = []
        self.input_columns = [
            "Invoice",
            "ClaimReferenceNumber",
            "InvoiceDOS",
            "OriginalDOS",
            "NewDOS",
            "Charge",
            "TotalChg",
            "InvoiceBalance",
            "ProviderName",
            "NewProvider",
            "BillingArea",
            "NewBillingArea",
            "OriginalLocation",
            "NewLocation",
            "Insurance",
            "TXN",
            "OriginalCPT",
            "NewCPT",
            "OriginalDX",
            "NewDX",
            "DxPointers",
            "OriginalModifier",
            "NewModifier",
            "ActionAddRemoveReplace",
            "Reason",
            "STEP",
            "Data",
            "Retrieval_Status",
            "Retrieval_Description"
        ]
        self.ref_columns = [
        'Invoice',
        'BAR_B_INV.SER_DT,',
        'BAR_B_TXN.SER_DT,',
        'BAR_B_INV.TOT_CHG,',
        'INV_BAL,',
        'PROV__1,',
        'LOC__2,',
        'BAR_B_INV.ORIG_FSC__5,',
        'BAR_B_INV.DX_ONE__3,',
        'DX_TWO__3,',
        'DX_THREE__3,',
        'DX_FOUR__3,',
        'DX_FIVE__3,',
        'DX_SIX__3,',
        'DX_SEVEN__3,',
        'DX_EIGHT__3,',
        'DX_NINE__3,',
        'DX_TEN__3,',
        'BAR_B_INV.DX_ELEVEN__3,',
        'BAR_B_INV.DX_TWELVE__3,',
        'TXN_NUM,',
        'PROC__2,',
        'MOD,',
        'BAR_B_TXN.DX_NUM,',
        'BAR_B_INV.CHG_CORR_FLAG,',
        'BAR_B_INV.CORR_INV_NUM',
        'BAR_B_TXN_LI_PAY.PAY_CODE__2',
        'BAR_B_TXN.U_CPTCODE_LIST',
        ]
        self.input_ref_columns = [
            "Invoice", 
            "ActionAddRemoveReplace", 
            "Reason", 
            "STEP", 
            "File Name", 
            "Rep Name", 
            "File Date",
            ]
        # dataframes
        self.final_df = None
        self.input_df = None
        self.input_ref_df = None
        self.ref_df = None
        self.rep_supe_df = None
        self.user_name_df = None
        self.df = None
        
        
    
    # Methods
    def run(self):
        self.B16_FP_UPDT = self.get_b16_file_path()
        self.file_path_ref = [ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH, self.B16_FP_UPDT]
        self.glob_files()
        self.iterate_files()
        self.generate_dataframes()
        self.supe_list = self.user_name_df["Report to"].tolist()
        self.create_input_file(df1= self.input_df, df2= self.ref_df, path= FILE_PATH, file_date= self.fd_no_spaces)
        self.create_input_ref_file(df= self.input_df)
        self.create_email()

    # replace {month} and {year} with self.month and self.year in self.B16_FILE_PATH
    def get_b16_file_path(self):
        logger.info("Getting B16 File Path")
        return self.B16_FILE_PATH.format(month=self.month, year=self.year)
    
    # Glob all files in the file_path_ref list if the file contains fd_spaces
    def glob_files(self):
        logger.info("Globbing Files")
        self.files = []
        logger.info(f"{self.file_path_ref}")
        for fp in self.file_path_ref:
            self.files += glob.glob(f'{fp}/*{self.fd_no_spaces}/*')
        logger.info(f"Total # of files globbed: {len(self.files)}")
    
    # check if the file exists, if not raise FileNotFoundError
    def check_files(self, file_path):
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(file_path)
        try:
            path.rename(path)
        except PermissionError:
            logger.error(f"File {file_path} is in use")
            return True
        else:
            return False

    # truncate the file name by replacing P1_FILE_PATH, P2_FILE_PATH, ARS_FILE_PATH and self.B16_FP_UPDT with ''
    def truncate_file_name(self, file_path):
        short_name = file_path.replace(P1_FILE_PATH, '').replace(P2_FILE_PATH, '').replace(ARS_FILE_PATH, '').replace(self.B16_FP_UPDT, '')
        short_name = short_name.lstrip(f'\\{self.fd_no_spaces}\\ ')
        return short_name
    
    def iterate_files(self):
        logger.info(f"length of globbed files: {len(self.files)}")
        logger.info(f"Truncating file names")
        logger.info(f"Checking if files exist and are not open/empty")
        for f in self.files:
            truncated_f = self.truncate_file_name(f)
            if '.xlsm' in truncated_f:
                trunc_f_no_ext = truncated_f.replace(f' {self.fd_spaces}.xlsm', '')
            else:
                trunc_f_no_ext = truncated_f.replace(f' {self.fd_spaces}.xlsx', '')
            if self.check_files(f) and '~$' not in f:
                self.open_file_list.append(truncated_f)
                self.open_file_name_list.append(trunc_f_no_ext)
                self.user_name_list.append(trunc_f_no_ext)
            elif '~$' not in f and '.tmp' not in f and '.db' not in f:
                df = pd.read_excel(f, engine='openpyxl')
                if not df.empty:
                    self.df_list.append(df)
                    for _ in range(df.shape[0]):
                        self.trunc_file_list.append(truncated_f)
                        self.rep_name_list.append(trunc_f_no_ext)
                else:
                    self.empty_file_list.append(truncated_f)
                    self.empty_f_user_list.append(trunc_f_no_ext)
                    self.user_name_list.append(trunc_f_no_ext)
    
    def create_df(self, df= None, col= None, id= str):
        logger.info(f"Creating DataFrame for {id}")
        return pd.DataFrame(df, columns=col)
    
    def generate_dataframes(self):
        self.final_df = pd.concat(self.df_list, ignore_index=True)
        self.input_df = self.create_df(df= self.final_df, col= self.input_columns, id= 'Input DataFrame')
        self.ref_df = self.create_df(col= self.ref_columns, id= 'Reference DataFrame')
        self.rep_supe_df = pd.read_excel(REP_SUPE_CROSSWALK, engine='openpyxl', sheet_name="GECB Usernames and Report Tos")
        self.user_name_df = self.create_df(df=self.user_name_list, col=['User Name'], id='User Name DataFrame')
        self.user_name_df = pd.merge(self.user_name_df, self.rep_supe_df, how='left', left_on='User Name', right_on='GECB Username')
        
    def create_input_file(self, df1, df2, path, file_date):
        logger.info("Creating Input File")
        df1.to_excel(f'{path}/Northwell_ChargeCorrection_Input_{file_date}.xlsx', index=False)
        with pd.ExcelWriter(f'{path}/Northwell_ChargeCorrection_Input_{file_date}.xlsx', mode='a', engine='openpyxl') as writer:
            # Write the DataFrame to a new sheet
            df2.to_excel(writer, sheet_name='Sheet2', index=False)

    def create_input_ref_file(self, df):
        logger.info("Creating Input Reference File")
        self.input_ref_df = df[["Invoice", "ActionAddRemoveReplace", "Reason", "STEP"]].copy()
        self.input_ref_df['File Name'] = self.trunc_file_list
        self.input_ref_df['Rep Name'] = self.rep_name_list
        self.input_ref_df['File Date'] = self.fd_slashes
        self.input_ref_df.to_excel(f'{INPUTS_PATH}/{self.year}/{self.month} {self.year}/Invoice Numbers and Rep Names {self.fd_spaces}.xlsx', index=False)

    def add_recipient_email(self, user_list, email, id=str):
        logger.info(f"attempting to add {id} user names to email list")
        for user in user_list:
            recipient = self.namespace.CreateRecipient(user)
            recipient.Resolve()
            if recipient.Resolved:
                email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                email.append(email_address)
            elif user == 'Edgar Santiago':
                email.append('esantiago5@northwell.edu')
            else:
                logger.info(f"No email address  found for alias or display name: {user}")

    def iterate_recipient_emails(self):
        logger.info("iterating through email lists")
        self.add_recipient_email(user_list= self.empty_f_user_list, email=self.email_list, id="Empty File")
        self.add_recipient_email(user_list=self.open_file_name_list, email=self.email_list, id="Open File")
        self.add_recipient_email(user_list=self.supe_list, email=self.supe_email_list, id="Supervisor")
        self.supe_email_list.append('dpashayan@northwell.edu')
    
    def create_email(self):
        self.html_body = f"""
            <p>Good Afternoon,</p>
            <p>In attempting to combine the charge correction input file, I was notified that your file is currently open or contains no invoices.</p>
            <p><strong><u>The Files belonging to the following users are currently open and WILL NOT be sent to the bot. Please re-add any corrections to your file for the next business day and remember to close your file on time:</u></strong>.<br>
            {self.open_file_name_list}</p>
            <p><strong><u>The files belonging to the following users have no invoices. While no further action is required, please do not open a file until you have an invoice to correct:</u></strong><br>
            {self.empty_f_user_list}</p>
            <p>
                <strong>Thank you,<br>
                ORCCA Team</strong><br>
                <span style="font-size: 9pt;">
                    Optimizing Revenue Cycle with Cognitive Automation (ORCCA) Team<br>
                    Northwell Health Physician Partners<br>
                    1111 Marcus Avenue, Ste. M04<br>
                    Lake Success, NY 11042
                </span><br><br>
                <span style="font-family: Arial; font-size: 11pt; color: #002060;"><strong>Northwell Health</strong></span><br>
            </p>

            """
        self.iterate_recipient_emails()
        logger.info('creating email')
        mail = self.outlook.CreateItem(0)
        mail.Subject = f"CCN Input - Open or Blank | File Date - {self.fd_spaces}"
        mail.HTMLBody = self.html_body
        self.email_list = list(set(self.email_list))
        self.supe_email_list = list(set(self.supe_email_list))
        for email in self.email_list:
            mail.Recipients.Add(email)
        for email in self.supe_email_list:
            recipient = mail.Recipients.Add(email)
            recipient.Type = 2
        mail.display(False)


class Input_Review(Date_Functions):
    def __init__(self, fd_mmddyyy, fd_mm, fd_yyyy, fd_spaces, fd_slashes):
        super().__init__()
        global FILE_PATH, OUT_PATH, ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH, REP_SUPE_CROSSWALK, INPUTS_PATH, FSC_GRID
        self.fd_no_spaces = fd_mmddyyy
        self.fd_spaces = fd_spaces
        self.fd_slashes = fd_slashes
        self.month = fd_mm
        self.year = fd_yyyy
        self.file = f"{FILE_PATH}/Northwell_ChargeCorrection_Input_{self.fd_no_spaces}.xlsx"
        self.fsc_grid_df = pd.read_excel(FSC_GRID, engine='openpyxl', sheet_name='Updated 08 04 2023')
        self.fsc_values = self.fsc_grid_df['FSC'].tolist()
        self.initial_sub_df = pd.read_excel(self.file, engine='openpyxl', sheet_name='Sheet1')
        self.query_df = pd.read_excel(self.file, engine='openpyxl', sheet_name='Sheet2')
        self.query_drop_col = [
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
            "BAR_B_INV.CHG_CORR_FLAG,",
        ]
        self.review_df = None
        self.outlook = win32.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNameSpace("MAPI")
        self.review_sheet_name = 'Sheet3'
        
    def run(self):
        self.calc_one_year_ago()
        self.prep_query_df()
        self.prep_review_df()
        self.fsc_review()
        self.review_cpt()
        self.update_inv_bal()
        self.update_step()
        self.exclude_multi_paycode()
        self.exclude_dos_over_one_year()
        self.review_mod()
        self.review_orig_cpt()
        self.blank_and_equal()
        self.count_delim(delim=',', ref_col='OriginalDX', new_col='Original DX Count')
        self.count_delim(delim=',', ref_col='NewDX', new_col='New DX Count')
        self.find_max_pointer_value()
        self.review_dx_pointers()
        self.evaluate_dx_pointers()
        self.invalid_bie()
        self.add_rep_supe()
        self.value_exclude_col()
        self.create_exclusion_df()
        self.order_col()
        self.write_to_excel()

    def calc_one_year_ago(self):
        logger.info("Calculating One Year Ago")
        self.one_year_ago = self.today - relativedelta(years=1) + relativedelta(days=1)
        self.one_year_ago = self.one_year_ago.strftime('%m/%d/%Y')
        logger.info(f"One year ago is {self.one_year_ago}")
    
    def create_email(self, body = str, e_type = 'error', inv_num = int, rep = [], supe = []):
        self.html_body = body
        logger.debug(inv_num)
        logger.debug(rep)
        logger.debug(supe)
        logger.info('creating email')
        mail = self.outlook.CreateItem(0)            
        if e_type == "error":
            mail.Subject = f"CCN Input - Error - {inv_num}"
        else:
            mail.Subject = f"CCN Input - Return - {inv_num}"
        for r in rep:
            try:
                recipient = self.namespace.CreateRecipient(r)
                recipient.Resolve()
                if recipient.Resolved:
                    email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                mail.Recipients.Add(email_address)
            except UnboundLocalError as e:
                logger.error(f'invalid rep email address for {rep}')
        for s in supe:
            if s == "Edgar Santiago":
                email_address = 'esantiago5@northwell.edu'
            else:
                recipient = self.namespace.CreateRecipient(s)
                recipient.Resolve()
                if recipient.Resolved:
                    email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            try:
                recipient = mail.Recipients.Add(email_address)
            except UnboundLocalError as e:
                continue
            recipient.Type = 2
        mail.HTMLBody = self.html_body
        mail.display(False)

    def prep_query_df(self):
        logger.info("Preparing Query DataFrame")
        self.query_cpts = self.query_df.groupby('Invoice')['PROC__2,'].apply(lambda x: list(map(str, x.unique()))).to_dict()
        self.query_df.drop(columns=self.query_drop_col)
        self.pay_code_counts = self.query_df.groupby('Invoice')['BAR_B_TXN_LI_PAY.PAY_CODE__2'].nunique().reset_index(name='unique_paycode_count')
        self.query_df = self.query_df.merge(self.pay_code_counts, on='Invoice', how='left')
        logger.debug(f'Length of Query DataFrame: {len(self.query_df)}')

    def prep_review_df(self):
        logger.info("Preparing Review DataFrame")
        self.review_df = self.initial_sub_df.merge(
                self.query_df, 
                on='Invoice', 
                how='left'
            ).drop_duplicates(
                subset='Invoice', 
                keep='last'
            ).astype(
                {
                'BAR_B_INV.TOT_CHG,': 'float64',
                'STEP': 'int64',
                'INV_BAL,': 'float64',
                }
            )
        try:
            self.review_df['InvoiceBalance'] = self.review_df['InvoiceBalance'].astype('float64')
        except ValueError as e:
            logger.info('Error converting InvoiceBalance to float')
            for index, value in self.review_df['InvoiceBalance'].items():
                try:
                    float(value)
                except ValueError as e:
                    logger.info(f'Error converting {value} to float at index: {index}')
                    self.review_df.at[index, 'InvoiceBalance'] = 0.00
            self.review_df['InvoiceBalance'] = self.review_df['InvoiceBalance'].astype('float64')
        logger.debug(f'Length of Review DataFrame: {len(self.review_df)}')

    def fsc_review(self):
        logger.info("Reviewing FSCs")
        self.review_df['FSC REVIEW'] = np.where(
            self.initial_sub_df['Insurance'].isin(self.fsc_values), 
            '', 
            'FSC not in grid'
            )
        
    def update_inv_bal(self):
        self.review_df['InvoiceBalance'] = np.where(
            (self.review_df['InvoiceBalance'] != self.review_df['INV_BAL,']) &
            (self.review_df['INV_BAL,'].notnull()),
            self.review_df['INV_BAL,'],
            self.review_df['InvoiceBalance']
        )

    def update_step(self):
        # If Invoice Balance = total charges and STEP = 3; set STEP = 2
        logger.info("Updating Step")
        logger.info("Checking if Invoice Balance = Total Charges and Step = 3; if true set Step = 2")
        self.review_df['STEP'] = np.where(
            (self.review_df['InvoiceBalance'] == self.review_df['BAR_B_INV.TOT_CHG,']) &
            (self.review_df['STEP'] == 3),
            2,
            self.review_df['STEP']
        )
        logger.info('Checking if Invoice Balance != Total Charges and Step = 2; if true set Step = 3')
        # If invoice balance != total charges and STEP = 2; set STEP = 3
        self.review_df['STEP'] = np.where(
            (self.review_df['InvoiceBalance'] != self.review_df['BAR_B_INV.TOT_CHG,']) &
            (self.review_df['STEP'] == 2) & 
            self.review_df['BAR_B_INV.TOT_CHG,'].notnull(),
            3,
            self.review_df['STEP']
        )
        logger.info('Checking length of OriginalCPT and NewCPT; if OriginalCPT > NewCPT and Step != 4; set Step = 4')
        # If a CPT has been removed from the 'OriginalCPT' list, then set the step to 4, otherwise keep it as is
        self.review_df['STEP'] = self.review_df.apply(
            lambda row: 4 if (len(str(row['OriginalCPT'])) > len(str(row['NewCPT']))) and (row['STEP'] != 4) else row['STEP'],
            axis=1
        )

    def exclude_multi_paycode(self):
        logger.info("Excluding Multi Paycode Invoices")
        self.review_df['Exclude Multi Paycode'] = np.where(
            (self.review_df['unique_paycode_count'] > 1) &
            (self.review_df['STEP'] != 2),
            "Exclude",
            ""
        )
    
    def exclude_dos_over_one_year(self):
        logger.info("Excluding DOS over One Year")
        self.review_df['Exclude DOS > 1 Year'] = np.where(
            (self.review_df['BAR_B_INV.SER_DT,'] <= self.one_year_ago),
            "Exclude",
            ""
        )

    def review_mod(self):
        logger.info("Identifying modifiers that need to be reviewed")
        self.review_df['Modifier Review'] = np.where(
            (self.review_df['OriginalModifier'].notnull()) &
            (self.review_df['NewModifier'].isnull()),
            "Review",
            ""
        )

    def review_orig_cpt(self):
        logger.info("Valuing Original CPT Field based on CPT/DX/Modifier/Date of Service")
        self.review_df['OriginalCPT'] = np.where(
            (self.review_df['OriginalCPT'] == self.review_df['NewCPT']) &
            (self.review_df['DxPointers'].isnull()) &
            (self.review_df['OriginalModifier'].isnull()) &
            (self.review_df['NewModifier'].isnull()) &
            (self.review_df['NewDOS'].isnull()),
            "",
            self.review_df['OriginalCPT']
        )
    
    def is_orig_blank(self, new_field, orig_field):
        self.review_df[f'{new_field}'] = np.where(
            (self.review_df[f'{orig_field}'] == ""),
            "",
            self.review_df[f'{new_field}']
        )

    def is_orig_eq_new(self, orig_field, new_field):
        self.review_df[f'{orig_field}'] = np.where(
            (self.review_df[f'{orig_field}'] == self.review_df[f'{new_field}']),
            "",
            self.review_df[f'{orig_field}']
        )

    def clear_date_field(self):
        self.review_df['OriginalDOS'] = np.where(
            (self.review_df['NewDOS'].isnull()) | 
            (self.review_df['NewDOS'] == ""),
            "",
            self.review_df['OriginalDOS']
        )

    def blank_and_equal(self):
        logger.info("Blanking and Equaling Fields for Provider")
        self.is_orig_eq_new('ProviderName', 'NewProvider')
        self.is_orig_blank('NewProvider', 'ProviderName')
        logger.info("Blanking and Equaling Fields for Location")
        self.is_orig_eq_new('OriginalLocation', 'NewLocation')
        self.is_orig_blank('NewLocation', 'OriginalLocation')
        logger.info("Blanking and Equaling Fields for CPT")
        self.is_orig_blank('NewCPT', 'OriginalCPT')
        logger.info("Blanking and Equaling Fields for DX")
        self.is_orig_eq_new('OriginalDX', 'NewDX')
        self.is_orig_blank('NewDX', 'OriginalDX')
        logger.info("Blanking and Equaling Fields for DOS")
        self.is_orig_eq_new('NewDOS','OriginalDOS')
        self.clear_date_field()
        
    def count_delim(self, delim, ref_col, new_col):
        logger.info(f"Counting delimiter {delim} in column {ref_col} and adding to new column {new_col}")
        self.review_df[new_col] = 0
        # Iterate over each row of the dataframe
        for index, row in self.review_df.iterrows():
            # Count the number of commas in the value of Column C for this row
            delim_in_row = str(row[ref_col]).count(delim)
            if delim_in_row >= 1:
                delim_in_row += 1
            # Add 1 to the count if the value is not empty but has zero commas
            if delim_in_row == 0 and isinstance(row[ref_col], str) and delim not in row[ref_col] and row[ref_col] != "":
                delim_in_row += 1
            # Add the count to the new column
            self.review_df.at[index, new_col] = delim_in_row

    def find_max_pointer_value(self):
        logger.info("Finding Max Pointer Value")
        self.review_df['Max Pointer'] = self.review_df['DxPointers'].apply(
            lambda x: int(max([int(i) for i in str(x).replace('|', ',').split(',') if i.isdigit()]))
            if any(i.isdigit() for i in str(x))
            and len([int(i) for i in str(x).replace('|', ',').split(',') if i.isdigit()]) > 0
            else 0
        )

    def review_dx_pointers(self):
        logger.info("Reviewing Dx Pointers for null values and empty strings")
        # Appends 'True' if "DxPointers" is null, otherwise 'False'
        self.review_df['DxPointers Null'] = self.review_df.apply(lambda row: True if pd.isna(row['DxPointers']) else False, axis=1)

        # Appends 'False' if "DxPointers Null" is true, Otherwise 'True' if any string before the "|" is a null string
        self.review_df['DxPointers String'] = self.review_df.apply(lambda row: False if row['DxPointers Null'] else 
                                    (True if any(val.strip() == '' for val in str(row['DxPointers']).split('|')) else False),
                                    axis=1)

    def evaluate_dx_pointers(self):
        logger.info("Evaluating Dx Pointers")
        # Evaluates if DxPointers need to be reviewed. 
        # 1) Orig Count > New Count & DxPointers String == True
        # 2) Orig Count != New Count & Max Pointer == 0
        # 3) Max Pointer > New Count & Orig Count != 0 & New Count != 0
        self.review_df['DxPointer Review'] = np.where(
        (
            (self.review_df['Original DX Count'] > self.review_df['New DX Count']) & 
            (self.review_df['DxPointers String'])
        ) |
        (
            (self.review_df['Original DX Count'] != self.review_df['New DX Count']) & 
            (self.review_df['Max Pointer'] == 0)
        ) |
        (
            (self.review_df['Max Pointer'] > self.review_df['New DX Count']) & 
            (self.review_df['Original DX Count'] != 0) & 
            (self.review_df['New DX Count'] != 0)
        ),
        'Review',
        '')

    def invalid_bie(self):
        logger.info("Evaluating for Invalid BIE")
        # Create a new column 'Invalid BIE' that looks at the 'OriginalCPT' column for a string containing '|'. If there is no "|" then review the 'NewCPT' column and if there is a null value, evaluate to 'Exclude'
        self.review_df['Invalid BIE'] = np.where(
            (self.review_df['OriginalCPT'].apply(lambda x: '|' not in str(x))) &
            (self.review_df['OriginalCPT'].isnull() == False) &
            (self.review_df['NewCPT'].isnull() == True),
            'Exclude',
            ''
        )

    def review_FSC(self):
        logger.info
        self.review_df['Review FSC'] = np.where(
            self.review_df['Insurance'] != self.review_df['BAR_B_INV.ORIG_FSC__5,'],
            'Review',
            ''
        )

    def review_cpt(self):
        logger.info("Reviewing CPTs")
        # 99205|19038R|38505|10005 or |93770|94761|93040|77003| or J0491|||| or 99292
        # columns to use 'OriginalCPT List', 'OriginalCPT List Count', 'BAR_B_TXN.U_CPTCODE_LIST', 'QueryCPT', 'QueryCPT Count', 'Review CPT'

        self.review_df.rename(columns={'BAR_B_TXN.U_CPTCODE_LIST': 'QueryCPT'}, inplace=True)
        self.review_df['OriginalCPT'] = self.review_df['OriginalCPT'].apply(lambda x:str(x))
        self.review_df['QueryCPT']= self.review_df['QueryCPT'].apply(lambda x:str(x))
        
        self.review_df['OriginalCPT List'] = self.review_df['OriginalCPT'].apply(lambda x: x.split('|') if '|' in x else [x])
        self.review_df['QueryCPT'] = self.review_df['QueryCPT'].apply(lambda x: x.split('|') if '|' in x else [x])

        self.review_df['Review CPT'] = ''

        for index, row in self.review_df.iterrows():
            query_list = row['QueryCPT']
            orig_list = row['OriginalCPT List']

            if len(query_list) != len(orig_list):
                self.review_df.at[index, 'Review CPT'] = 'Count Mismatch'
                continue

            for i, o in enumerate(orig_list):
                q = query_list[i]
                if o and o != q:
                    self.review_df.at[index, 'Review CPT'] = 'Comparison Error'
                    break

        

    def value_exclude_col(self):
        logger.info("Valuing Exclude Column")
        self.review_df['Exclude'] = np.where(
            (self.review_df['Exclude Multi Paycode'] == 'Exclude') | 
            (self.review_df['Exclude DOS > 1 Year'] == 'Exclude') |
            (self.review_df['Invalid BIE'] == 'Exclude') |
            (self.review_df['BAR_B_INV.CORR_INV_NUM'] >= 1),
            'Exclude',
            ''
        )
        for index, row in self.review_df.iterrows():
            if row['Rep Name'] == 'HCOB16Electronic':
                self.review_df.at[index, 'Exclude'] = ''
            elif row['Rep Name'] == 'HCOB16Paper':
                self.review_df.at[index, 'Exclude'] = ''
            elif row['Rep Name'] == 'MBPROJECT':
                self.review_df.at[index, 'Exclude'] = ''
            elif row['Rep Name'] == 'MBProject':
                self.review_df.at[index, 'Exclude'] = ''
            else:
                continue

    def order_col(self):
        logger.info("Ordering Columns")
        col_order = [
            "Invoice",
            "ClaimReferenceNumber",
            "InvoiceDOS",
            "OriginalDOS",
            "NewDOS",
            "Charge",
            "TotalChg",
            "InvoiceBalance",
            "ProviderName",
            "NewProvider",
            "BillingArea",
            "NewBillingArea",
            "OriginalLocation",
            "NewLocation",
            "Insurance",
            "TXN",
            "OriginalCPT",
            "NewCPT",
            "OriginalDX",
            "NewDX",
            "DxPointers",
            "OriginalModifier",
            "NewModifier",
            "ActionAddRemoveReplace",
            "Reason",
            "STEP",
            "Data",
            "Retrieval_Status",
            "Retrieval_Description",
            "BAR_B_INV.SER_DT,",
            "BAR_B_INV.TOT_CHG,",
            "INV_BAL,",
            "BAR_B_INV.ORIG_FSC__5,",
            "FSC REVIEW",
            "QueryCPT",
            "OriginalCPT List",
            "Review CPT",
            "BAR_B_INV.CORR_INV_NUM",
            "unique_paycode_count",
            "Exclude Multi Paycode",
            "Exclude DOS > 1 Year",
            "Invalid BIE",
            "Modifier Review",
            "DxPointer Review",
            "Original DX Count",
            "New DX Count",
            "Max Pointer",
            "Exclude",
            "Rep Name",
            "Report to",
        ]
        # set the column order of self.review_df to col_order
        self.review_df = self.review_df[col_order]


    def add_rep_supe(self):
        logger.info("Adding Rep and Supervisor Columns")
        self.rep_supe_df = pd.read_excel(REP_SUPE_CROSSWALK, engine='openpyxl', sheet_name="GECB Usernames and Report Tos")
        self.rep_supe_df = self.rep_supe_df[['GECB Username', 'Report to']]
        self.user_name_df = pd.read_excel(f'{INPUTS_PATH}/{self.year}/{self.month} {self.year}/Invoice Numbers and Rep Names {self.fd_spaces}.xlsx', engine='openpyxl')
        self.user_name_df = self.user_name_df[['Invoice', 'Rep Name']]
        self.user_name_df = pd.merge(self.user_name_df, self.rep_supe_df, left_on='Rep Name', right_on='GECB Username', how='left')
        self.user_name_df = self.user_name_df[['Invoice', 'Rep Name', 'Report to']]
        self.review_df = pd.merge(left=self.review_df, right=self.user_name_df, left_on="Invoice", right_on='Invoice')

    def create_exclusion_df(self):
        logger.info("Creating dataframe of exclusions")
        self.excl_df = self.review_df[self.review_df['Exclude'] == 'Exclude']
        # iterate through each invoice in the exclusion dataframe, review the columns 'Exclude Multi Paycode', 'Exclude DOS > 1 Year', 'Corrected Invoice Number'. If there is a value in any of these columns, create an email using the create_email method
        for index, row in self.excl_df.iterrows():
            if 'HCOB16' in row['Rep Name']:
                continue
            elif 'MBPROJECT' in row['Rep Name']:
                continue
            elif 'MBProject' in row['Rep Name']:
                continue
            else:
                if row['Exclude Multi Paycode'] == 'Exclude':
                    logger.debug('Multi Paycode')
                    self.create_email(
                        body=f"""
                            <p>Good Afternoon,</p>
                            <p>Please be advised that invoice {row['Invoice']} is being removed from the input file as the Invoice Detail has multiple payments from different carriers.<br>
                            The bot is unable to repost payments when multiple different paycodes are used in payment posting. Please have this invoice corrected manually.</p>
                            <p><strong>Invoice Number:</strong> {row['Invoice']}</p>
                            <p><strong>File Date:</strong> {self.fd_slashes}</p>
                            <p>
                                <strong>Thank you,<br>
                                ORCCA Team</strong><br>
                                <span style="font-size: 9pt;">
                                    Optimizing Revenue Cycle with Cognitive Automation (ORCCA) Team<br>
                                    Northwell Health Physician Partners<br>
                                    1111 Marcus Avenue, Ste. M04<br>
                                    Lake Success, NY 11042
                                </span><br><br>
                                <span style="font-family: Arial; font-size: 11pt; color: #002060;"><strong>Northwell Health</strong></span><br>
                            </p>
                        """,
                        e_type='return',
                        inv_num=row['Invoice'],
                        rep=[row['Rep Name']],
                        supe=[row['Report to']]
                    )
                elif row['Exclude DOS > 1 Year'] == 'Exclude':
                    logger.debug('DOS > 1 year')
                    self.create_email(
                        body=f"""
                            <p>Good Afternoon,</p>
                            <p>Please be advised that invoice {row['Invoice']} is being removed from the input file as the Invoice DOS is over a year ago.<br> 
                            Please complete this charge correction manually. If you need assistance please reach out to your Senior Representative or Supervisor for help with that process.<br></p>
                            <p><strong>Invoice Number:</strong> {row['Invoice']}</p>
                            <p><strong>File Date:</strong> {self.fd_slashes}</p>
                            <p>
                                <strong>Thank you,<br>
                                ORCCA Team</strong><br>
                                <span style="font-size: 9pt;">
                                    Optimizing Revenue Cycle with Cognitive Automation (ORCCA) Team<br>
                                    Northwell Health Physician Partners<br>
                                    1111 Marcus Avenue, Ste. M04<br>
                                    Lake Success, NY 11042
                                </span><br><br>
                                <span style="font-family: Arial; font-size: 11pt; color
                        """,
                        e_type='error',
                        inv_num=row['Invoice'],
                        rep=[row['Rep Name']],
                        supe=[row['Report to']]
                    )
                elif row['Invalid BIE'] == 'Exclude':
                    logger.debug('Invalid BIE')
                    self.create_email(
                        body=f"""
                            <p>Good Afternoon,</p>
                            <p>Please be advised that invoice {row['Invoice']} is being removed from the input file as the of all CPTs as BIE is not a valid correction for the CCN Bot.<br> 
                            Please complete this charge correction manually. If you need assistance please reach out to your Senior Representative or Supervisor for help with that process.<br></p>
                            <p><strong>Invoice Number:</strong> {row['Invoice']}</p>
                            <p><strong>File Date:</strong> {self.fd_slashes}</p>
                            <p>
                                <strong>Thank you,<br>
                                ORCCA Team</strong><br>
                                <span style="font-size: 9pt;">
                                    Optimizing Revenue Cycle with Cognitive Automation (ORCCA) Team<br>
                                    Northwell Health Physician Partners<br>
                                    1111 Marcus Avenue, Ste. M04<br>
                                    Lake Success, NY 11042
                                </span><br><br>
                                <span style="font-family: Arial; font-size: 11pt; color
                        """,
                        e_type='error',
                        rep=[row['Rep Name']], 
                        supe=[row['Report to']]
                    )
                elif row['BAR_B_INV.CORR_INV_NUM'] != '':
                    logger.debug('Inv already CCN')
                    self.create_email(
                        body=f"""
                            <p>Good Afternoon,</p>
                            <p>Please be advised that invoice {row['Invoice']} is being removed from the input file as the invoice has already been corrected.<br>
                            Please review new invoice to determine if correction is still warranted.</p>
                            <p><strong>Invoice Number:</strong> {row['Invoice']}</p>
                            <p><strong>File Date:</strong> {self.fd_slashes}</p>
                            <p>
                                <strong>Thank you,<br>
                                ORCCA Team</strong><br>
                                <span style="font-size: 9pt;">
                                    Optimizing Revenue Cycle with Cognitive Automation (ORCCA) Team<br>
                                    Northwell Health Physician Partners<br>
                                    1111 Marcus Avenue, Ste. M04<br>
                                    Lake Success, NY 11042
                                </span><br><br>
                                <span style="font-family: Arial; font-size: 11pt; color
                        """,
                        e_type='return',
                        inv_num=row['Invoice'],
                        rep=[row['Rep Name']],
                        supe=[row['Report to']]
                    )
    
    def write_to_excel(self):
        logger.info("Writing to Excel")
        with pd.ExcelWriter(self.file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            # Write the DataFrame to a new sheet
            self.review_df.to_excel(writer, sheet_name=self.review_sheet_name, index=False)


class File_To_CSV:
    def __init__(self, fd_mmddyyy, fd_mm, fd_yyyy, fd_spaces, fd_slashes) -> None:
        global FILE_PATH, OUT_PATH, ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH, REP_SUPE_CROSSWALK, INPUTS_PATH, FSC_GRID
        self.fd_no_spaces = fd_mmddyyy
        self.fd_spaces = fd_spaces
        self.fd_slashes = fd_slashes
        self.month = fd_mm
        self.year = fd_yyyy
        self.export_file = f"{FILE_PATH}/Northwell_ChargeCorrection_Input_{self.fd_no_spaces}.xlsx"
        self.export_df = pd.read_excel(self.export_file, sheet_name="Sheet3", engine='openpyxl')
        self.drop_col = [
            "BAR_B_INV.SER_DT,",
            "BAR_B_INV.TOT_CHG,",
            "INV_BAL,",
            "BAR_B_INV.ORIG_FSC__5,",
            "FSC REVIEW",
            "QueryCPT",
            "OriginalCPT List",
            "Review CPT",
            "BAR_B_INV.CORR_INV_NUM",
            "unique_paycode_count",
            "Exclude Multi Paycode",
            "Exclude DOS > 1 Year",
            "Invalid BIE",
            "Modifier Review",
            "DxPointer Review",
            "Original DX Count",
            "New DX Count",
            "Max Pointer",
            "Exclude",
            "Rep Name",
            "Report to",
        ]
        self.file_csv = f"Northwell_ChargeCorrection_Input_{self.fd_no_spaces}.csv"

    def run(self):
        self.export_df = self.rem_excl_drop_col(df= self.export_df, col=self.drop_col)
        self.export_df['ClaimReferenceNumber'] = self.strip_white_space(self.export_df, 'ClaimReferenceNumber')
        self.to_csv(self.export_df)

    def rem_excl_drop_col(self, df, col):
        logger.info("Dropping excluded lines")
        df_no_excl = df[df['Exclude'] != 'Exclude']
        logger.info("Dropping unnecessary columns")
        df_col_drop = df_no_excl.drop(columns=col)
        return df_col_drop
    
    def strip_white_space(self, df, col):
        logger.info("Stripping white space")
        return df[col].str.strip()
    
    def to_csv(self, df):
        logger.info(f"Saving csv to {OUT_PATH}")
        df.to_csv(f'{OUT_PATH}/{self.file_csv}', index=False)
        