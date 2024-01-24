import pandas as pd
import glob
from datetime import date, timedelta
from pathlib import Path
import os
import win32com.client as win32

def run():
    def is_file_in_use(file_path):
        """
        Determines if the file is open and Returns boolean. Raises FileNotFoundError if the file does not exist
        :param file_path:
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError

        try:
            path.rename(path)
        except PermissionError:
            return True
        else:
            return False


    def truncate_file_name(file_name):
        """
        Takes the full file path and Returns a truncated file name
        :param file_name:
        """
        short_name = file_name.replace(f"{P1_FILE_PATH}", "").replace(f"{P2_FILE_PATH}", "").replace(f"{ARS_FILE_PATH}", "").replace(f"{B16_FILE_PATH}", "")
        # short_name = file_name.replace(f"{P1_FILE_PATH}", "").replace(f"{P2_FILE_PATH}", "").replace(f"{ARS_FILE_PATH}", "")
        short_name = short_name.lstrip(f"\\{file_date}\\ ")
        return short_name

    def add_recipient_email(user_list, email_list):
        """
        Takes the indicated list of users and checks if the email address exists. If it does, it will add it to the list of email addresses. If no email address exists for that user, it will notify in the script
        """
        for user in user_list:
            # Search for the user in the Outlook address book
            recipient = namespace.CreateRecipient(user)
            recipient.Resolve()
            if recipient.Resolved:
                # Retrieve the user's email address from the resolved recipient object
                email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                email_list.append(email_address)
            else:
                print(f"No email address found for alias or display name: '{user}'")


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

    # format file_date into desired formats
    file_date_plain = file_date.strftime('%m%d%Y')
    file_date_w_spaces = file_date.strftime('%m %d %Y')
    file_date_w_slashes = file_date.strftime('%m/%d/%Y')

    # extract day and month from file_date
    day = file_date.strftime('%d')
    month = file_date.strftime('%m')
    year = file_date.strftime('%Y')
    
    ARS_FILE_PATH = "M:/CPP-Data/AR SUPPORT/SPECIAL PROJECTS/CHARGE CORRECTION BOT/SPREADSHEETS TO SEND TO BOT"
    P1_FILE_PATH = "M:/CPP-Data/Payor 1/Bot CCN"
    P2_FILE_PATH = "M:/CPP-Data/Payer 2/BOTS/Charge Correction Files"
    B16_FILE_PATH = f"M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/New vs Established/Formatted Inputs/{month} {year}"
    REP_SUPE_CROSSWALK = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/References/Report To.xlsx"

    file_paths = [ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH, B16_FILE_PATH]
    # file_paths = [ARS_FILE_PATH, P1_FILE_PATH, P2_FILE_PATH]

    open_file_list = []
    empty_file_list = []
    files = []
    trunc_file_list = []
    final = []
    rep_name_list = []
    empty_file_name_list = []
    open_file_name_list = []
    email_list = []
    user_name_list = []
    supe_email_list = []

    # Creates a LIST containing the file paths for a given file date
    for f in file_paths:
        files += glob.glob(f'{f}/*{file_date_plain}/*')
    # print(files)

    # Goes through the File List. Checks if the file is open.
    # If it's not open it attempts to create a list of Pandas DataFrames
    for f in files:
        # print(f)
        trunc_file = truncate_file_name(f)
        if '.xlsm' in trunc_file:
            trunc_file_no_ext = trunc_file.replace(f" {file_date_w_spaces}.xlsm", "")
        else:
            trunc_file_no_ext = trunc_file.replace(f" {file_date_w_spaces}.xlsx", "")
        if is_file_in_use(f) and '~$' not in f:
            open_file_list.append(trunc_file)
            open_file_name_list.append(trunc_file_no_ext)
            user_name_list.append(trunc_file_no_ext)
        elif '~$' not in f and '.tmp' not in f and '.db' not in f:
            # print(f)
            df = pd.read_excel(f, engine="openpyxl")
            if not df.empty:
                final.append(df)
                for _ in range(df.shape[0]):
                    trunc_file_list.append(trunc_file)
                    rep_name_list.append(trunc_file_no_ext)
            else:
                empty_file_list.append(trunc_file)
                empty_file_name_list.append(trunc_file_no_ext)
                user_name_list.append(trunc_file_no_ext)
                # print(f"the file {trunc_file} is empty")
        else:
            continue

    # Concatenates the list of DataFrames into a single DataFrame
    final = pd.concat(final)

    # Sets the columns for the DataFrame
    df1 = pd.DataFrame(
        final,
        columns=[
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
    )

    df3 = pd.DataFrame(
        columns=[
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
        'BAR_B_TXN_LI_PAY.PAY_CODE__2'
        ]
        )

    # Dataframe for Report To's crosswalk
    df4 = pd.read_excel(REP_SUPE_CROSSWALK, engine="openpyxl", sheet_name="GECB Usernames and Report Tos")

    # Dataframe to merge with df4 to get supe name for CC list
    df5 = pd.DataFrame(user_name_list, columns= ["User Name"])

    df5 = pd.merge(df5, df4, on=None, left_on="User Name", right_on="GECB Username", how='left')

    supe_list = df5["Report to"].tolist()

    # Writes the finale DataFrame to a new Excel sheet this becomes the Input File
    OUT_PATH_1 = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Audits - Files Sent to Bot/"
    OUT_PATH_2 = "M:/CPP-Data/Sutherland RPA/Northwell Process Automation ETM Files/Monthly Reports/Charge Correction/Inputs/"
    
    df1.to_excel(f'{OUT_PATH_1}Northwell_ChargeCorrection_Input_{file_date_plain}.xlsx', index=False)
    with pd.ExcelWriter(f'{OUT_PATH_1}Northwell_ChargeCorrection_Input_{file_date_plain}.xlsx', mode='a', engine='openpyxl') as writer:
        # Write the DataFrame to a new sheet
        df3.to_excel(writer, sheet_name='Sheet2', index=False)
    
    # Create the input reference file
    df2 = df1[["Invoice", "ActionAddRemoveReplace", "Reason", "STEP"]].copy()
    df2["File Name"] = trunc_file_list
    df2["Rep Name"] = rep_name_list
    df2["File Date"] = file_date_w_slashes

    # df2.to_excel(writer2, index=False)
    df2.to_excel(f"{OUT_PATH_2}{year}/{month} {year}/Invoice Numbers and Rep Names {file_date_w_spaces}.xlsx", index=False)

    # output_path = 'C:/Users/denglish2/Desktop/output.txt'
    # Get the path to the desktop directory of the current user
    desktop_path = os.path.expanduser("~/Desktop")

    # Specify the output file name
    output_filename = "OpenOrInUseFilesOutput.txt"

    # Create the full output file path
    output_path = os.path.join(desktop_path, output_filename)
    # open the file in write mode and write the output to it
    with open(output_path, 'w') as f:
        f.write(f"These are the files that are still open: \n{open_file_list}\n")
        f.write(f"These are the files without any entries: \n{empty_file_list}\n")
        
    # open the file for reading and print its contents to the console
    with open(output_path, 'r') as f:
        print(f.read())

    # Create an instance of the Outlook application
    outlook = win32.Dispatch("Outlook.Application")
    # Get the MAPI namespace of the Outlook application
    namespace = outlook.GetNamespace("MAPI")

    add_recipient_email(user_list=empty_file_name_list, email_list=email_list)
    add_recipient_email(user_list=open_file_name_list, email_list=email_list)
    add_recipient_email(user_list=supe_list, email_list=supe_email_list)
    supe_email_list.append('dpashayan@northwell.edu')

    # Create a new email message
    mail = outlook.CreateItem(0)

    mail.Subject = f"CCN Input - Open or Blank | File Date - {file_date_w_spaces}"

    html_body = f"""
    <p>Good Afternoon,</p>
    <p>In attempting to combine the charge correction input file, I was notified that your file is currently open or contains no invoices.</p>
    <p><strong><u>The Files belonging to the following users are currently open and WILL NOT be sent to the bot. Please re-add any corrections to your file for the next business day and remember to close your file on time:</u></strong>.<br>
    {open_file_name_list}</p>
    <p><strong><u>The files belonging to the following users have no invoices. While no further action is required, please do not open a file until you have an invoice to correct:</u></strong><br>
    {empty_file_name_list}</p>
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
    email_list = list(set(email_list))
    supe_email_list = list(set(supe_email_list))
    # Set the To: field of the email message
    for email_address in email_list:
        mail.Recipients.Add(email_address)
    for email in supe_email_list:
        recipient = mail.Recipients.Add(email)   
        recipient.Type = 2

    # Display the email message (leave it open for editing)
    mail.Display(False)


if __name__ == '__main__':
    run()
