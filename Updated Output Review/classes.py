import datetime as dt
import win32com.client as win32
import pandas as pd
import shutil
from logger_setup import logger
from tkinter import simpledialog as sd
from tkinter import messagebox as mb
import mappings
import pathlib
from pathlib import Path
from glob import glob
import os
import re


class Output_Review():
    def __init__(self):     
        self.comments_template_fp = mappings.ccn_dict["comments_template_fp"]
        self.output_checker_fp = mappings.ccn_dict["output_checker_fp"]
        self.columns = mappings.ccn_dict["columns"]
        self.column_rename = mappings.ccn_dict["column_rename"]
        self.rd_reason_crosswalk = mappings.ccn_dict["rd_reason_cross"]
        self.destination_fp = mappings.ccn_dict["destination_fp"]
        self.file_to_review_fp = mappings.ccn_dict["file_to_review_fp"]
        self.sutherland_output_fp = mappings.ccn_dict["sutherland_output_fp"]
        self.rep_submission_fp = mappings.ccn_dict["rep_submission_file_fp"] 
        self.cleanup_dir = Path('//NT2KWB972SRV03/SHAREDATA/CPP-Data/Sutherland RPA/ChargeCorrection')
        

    def prep_and_export_file(self):
        logger.info('isolating needed columns')
        self.df_formatted_exp = self.df_sutherland_exp[self.columns].copy()
        logger.info('renaming columns')
        self.df_formatted_exp.rename(columns=self.column_rename, inplace=True)
        logger.info('exporting final to destination')
        self.df_formatted_exp.to_excel(self.formatted_file_to_review, sheet_name="export", engine="openpyxl", index=False) 

    def get_rep_submissions(self):
        logger.info('adding rep name, super name, dept')
        # logger.debug(f"rep submission invoice {self.df_rep_submission['INVNUM']} sutherland exp invoice {self.df_sutherland_exp['INVNUM']}")
        try:
            self.df_sutherland_exp = pd.merge(self.df_sutherland_exp, self.df_rep_submission[['INVNUM', 'Rep Username', 'Rep Name','Supervisor','Department']],on='INVNUM',how='left')
        except ValueError:
            logger.error(f"Error merging DataFrames, likely due to an invald row in Northwell_ChargeCorrection_Output_{self.fd_mmddyyyy}.xls")

    def populate_stat_comm_cat_act(self):
        logger.info('populating dp status column')
        self.df_sutherland_exp['DP Status'] = self.df_sutherland_exp.apply(lambda row: self.get_crosswalk_values(row, sub_cat='DP Status'), axis=1)
        logger.info('populating dp comments column')
        self.df_sutherland_exp['DP Comments'] = self.df_sutherland_exp.apply(lambda row: self.get_crosswalk_values(row, sub_cat='DP Comments'), axis=1)
        logger.info('populating dp category column')
        self.df_sutherland_exp['DP Category'] = self.df_sutherland_exp.apply(lambda row: self.get_crosswalk_values(row, sub_cat='DP Category'), axis=1)
        logger.info('populating action column')
        self.df_sutherland_exp['Action'] = self.df_sutherland_exp.apply(lambda row: self.get_crosswalk_values(row, sub_cat='Action'), axis=1)

    def get_crosswalk_values(self, row, sub_cat):
        rd_reason_entry = self.rd_reason_crosswalk.get(row['RD + Reason'])
        if rd_reason_entry is not None:
            sub_value = rd_reason_entry.get(sub_cat, 'Unknown')
            return sub_value
        else:
            return 'Unknown'

    def create_data_frames(self):
        logger.info('creating df from sutherland output')
        self.df_sutherland_exp = pd.read_excel(self.formatted_sutherland_output, sheet_name="export", na_values="")
        logger.info('creating df of rep submission list')
        self.df_rep_submission = pd.read_excel(self.formatted_rep_submission, sheet_name="Sheet1", engine="openpyxl")
        self.df_rep_submission = self.df_rep_submission.rename(columns={'Invoice': 'INVNUM'})
        logger.info('populating rd + reason column')
        self.df_sutherland_exp["RD + Reason"] = self.df_sutherland_exp["RetrievalDescription"]+" - "+self.df_sutherland_exp["Reason"].fillna('')

    def move_templates_to_detination(self):
        logger.info("moving dp comments template")
        shutil.copy2(self.comments_template_fp, self.formatted_destination)
        logger.info("moving ccn checker template")
        shutil.copy2(self.output_checker_fp, self.formatted_destination)

    def replace_strings_in_fp(self, fp_str):
        logger.info(f"formatting file path string for {fp_str}")
        file_var_str = fp_str.format(yyyy=self.file_year, mm_yyyy=self.fd_mm_yyyy, mmddyyyy=self.fd_mmddyyyy)
        return(file_var_str)

    def get_file_date(self):
        logger.info("Retrieving today's date")
        self.today = dt.date.today()
        logger.info("determining file date")
        if self.today.weekday() == 0:
            self.file_date = self.today - dt.timedelta(days=3)
        else:
            self.file_date = self.today - dt.timedelta(days=1)

        if self.file_date == self.get_last_day_of_month():
            logger.info(f"{self.file_date} is the same as the last business day of the month")
            while self.file_date.weekday() >= 4:
                self.file_date -= dt.timedelta(days=1)
        logger.info(f"latest file date is {self.file_date}")
    
    def get_last_day_of_month(self):
        logger.info("Calculating last business day of month")
        self.ldom = (self.file_date.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
        while self.ldom.weekday() >= 5:
                logger.info("last day of the month is a weekend")
                self.ldom -= dt.timedelta(days=1)
        return(self.ldom)

    def convert_dates_to_strings(self):
        logger.info(f"converting {self.file_date} to strings: MMDDYYYY, MM/DD/YYYY, MM YYYY, YYYY")
        self.fd_mmddyyyy = self.file_date.strftime('%m%d%Y')
        self.fd_mm_dd_yyyy = self.file_date.strftime('%m/%d/%Y')
        self.fd_mm_yyyy = self.file_date.strftime('%m %Y')
        self.file_year = self.file_date.strftime('%Y')
        self.formatted_destination = self.replace_strings_in_fp(fp_str = self.destination_fp)
        self.formatted_file_to_review = self.replace_strings_in_fp(fp_str = self.file_to_review_fp)
        self.formatted_sutherland_output = self.replace_strings_in_fp(fp_str = self.sutherland_output_fp)
        self.formatted_rep_submission = self.replace_strings_in_fp(fp_str = self.rep_submission_fp)   

    def ask_if_correct_date(self):
        answer = mb.askyesno(f"Check Date", f"Do you want to run for the file date of {self.fd_mm_dd_yyyy}?")
        if not answer:
            self.day = sd.askinteger("Follow Up", "Please enter day of the month as number : ", minvalue=1, maxvalue=31)
            logger.debug(f"{self.day} was entered")
            self.month = sd.askinteger("Follow Up", "Please enter a month number (e.g. March = 3): ", minvalue=1, maxvalue=12)
            logger.debug(f"{self.month} was entered")
            self.year = sd.askinteger("Follow Up", "Please enter a year: ", minvalue=2020, maxvalue=2030)
            logger.debug(f"{self.year} was entered")
            try:
                self.file_date = self.today.replace(day=self.day, month=self.month, year=self.year)
                logger.info(f"new file date is {self.file_date}")
                self.convert_dates_to_strings()
            except TypeError:
                logger.critical("No date selected")
                logger.info("Stopping Process")
                exit()


class Email_Prep(Output_Review):
    def __init__(self):
        super().__init__()
        self.file_to_review_fp = mappings.ccn_dict["file_to_review_fp"]
        self.drop_columns = mappings.ccn_dict["drop_columns"]
        self.error_columns = mappings.ccn_dict["error_columns"]
        self.user_list = []
        self.email_list = []
        self.sheet_name = "Sheet2"
        
    
    def create_email_dataframes(self):
        self.formatted_file_to_review = self.replace_strings_in_fp(fp_str = self.file_to_review_fp)
        logger.info('creating dataframe from reviewed file')
        self.df_file_to_review = pd.read_excel(self.formatted_file_to_review, sheet_name='export', engine='openpyxl')
        logger.info('isolating errors into their own dataframe')
        self.df_errors = self.df_file_to_review[self.df_file_to_review['DP Category'] != 'Success']
        self.df_errors = self.df_errors[self.df_errors['Action'] != 'No Action Needed']

    def get_users_to_email(self):
        logger.info('creating list of names for email')
        for rep in self.df_errors['Rep Username'].unique():
            self.user_list.append(rep)
        for rep in self.df_errors['Rep Name'].unique():
            self.user_list.append(rep)
    
    def drop_columns_reorder_write_to_file(self):
        logger.info('dropping extra columns')
        self.df_errors = self.df_errors.drop(columns=self.drop_columns).sort_values(by='Rep Name')
        self.df_errors = self.df_errors[self.error_columns]
        logger.info('saving errors dataframe to original file')
        with pd.ExcelWriter(self.formatted_file_to_review, mode='a', engine='openpyxl') as writer:
            self.df_errors.to_excel(writer, sheet_name=self.sheet_name, index=False)
    
    def prep_email_list(self):
        logger.info('prepping the list of email addresses')
        self.outlook = win32.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")

        for user in self.user_list:
            recipient = self.namespace.CreateRecipient(user)
            recipient.Resolve()
            if recipient.Resolved:
                email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                self.email_list.append(email_address)
            else:
                logger.error(f'no email address found for: {user}')
    
    def create_email(self):
        logger.info('creating the email')
        mail = self.outlook.CreateItem(0)
        mail.Subject = f'The {self.fd_mm_dd_yyyy} CCN Output File is ready for review'
        html_table = self.df_errors.to_html(index=False, classes='dataframe', border=2, justify='center')
        html_body = f"""
            <p>Greetings,</p>
            <p>The {self.fd_mm_dd_yyyy} CCN output file is ready for review.</p>
            <p><strong><u>Supervisors</u></strong> - The complete file can be found in the following location: <a href="{self.formatted_destination}">Click Here</a>. The hyperlink will take you to the specific folder for the File Date being reviewed, review the file marked <strong><u>“DP Comments”</u></strong>.</p>
            <p><strong><u>Representatives</u></strong> - Please see the table for follow-up. Reference Column labeled with <strong><u>“Actions”</u></strong> as to next steps based on our review.</p>
            {html_table}
            <br>
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
        mail.HTMLBody = html_body
        self.email_list = list(set(self.email_list))
        for email in self.email_list:
            mail.Recipients.Add(email)
        cc_list = ["dpashayan@northwell.edu", "vlombardi2@northwell.edu", "rjohnson16@northwell.edu", "alang@northwell.edu", "jmullen3@northwell.edu", "AR Supervisors"]
        for cc in cc_list:
            recipient = mail.Recipients.Add(cc)
            recipient.Type = 2
        
        mail.Display(False)

    def cleanup_directory(self):
        
        files = glob(str(self.cleanup_dir / '*.csv'))
        file_dates = []
        for file in files:
            # get the filename
            filename = os.path.basename(file)
            # extract the date from the filename it is in MMDDYYYY format
            date = re.search(r'\d{8}', filename).group()
            # convert the date to a datetime object
            date = dt.datetime.strptime(date, '%m%d%Y')
            file_dates.append(date)

        # get todays date
        today = dt.datetime.now()
        # if date in the file is more than 5 days old, delete the file date
        for file, date in zip(files, file_dates):
            if (today - date).days > 5:
                os.remove(file)
            else:
                continue        
