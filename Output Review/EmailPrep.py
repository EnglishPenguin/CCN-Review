import pandas as pd
import datetime as dt
import win32com.client as win32


def run():
    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'

    today = dt.date.today()
    file_date = today - dt.timedelta(days=3)
    fd_mmddyyyy = file_date.strftime('%m%d%Y')
    fd_mm_yyyy = file_date.strftime('%m %Y')
    file_year = file_date.strftime('%Y')
    fd_mm_dd = file_date.strftime('%m/%d')

    file_to_review = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/DP Comments Template.xlsx'

    df_email = pd.read_excel(file_to_review, sheet_name="Sheet3", engine="openpyxl")

    user_list = []
    email_list = []
    for rep in df_email['Rep Name'].unique():
        user_list.append(rep)

    # Create an instance of the Outlook application
    outlook = win32.Dispatch("Outlook.Application")

    # Get the MAPI namespace of the Outlook application
    namespace = outlook.GetNamespace("MAPI")

    # Iterate through the user list and retrieve the email address for each user
    for user in user_list:
        # Search for the user in the Outlook address book
        recipient = namespace.CreateRecipient(user)
        recipient.Resolve()
        if recipient.Resolved:
            # Retrieve the user's email address from the resolved recipient object
            email_address = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            email_list.append(email_address)
        else:
            print(f"No user found with alias or display name '{user}'")

    # Create an instance of the Outlook application
    outlook = win32.Dispatch("Outlook.Application")

    # Create a new email message
    mail = outlook.CreateItem(0)

    mail.Subject = f"The {fd_mm_dd} CCN output file is ready for review."
    mail.Body = ""

    # Set the To: field of the email message
    for email_address in email_list:
        mail.Recipients.Add(email_address)

    # Set the CC: field of the email message
    cc_list = ["dpashayan@northwell.edu", "vlombardi2@northwell.edu", "rjohnson16@northwell.edu", "alang@northwell.edu", "jmullen3@northwell.edu", "AR Supervisors"]
    for distribution_list in cc_list:
        recipient = mail.Recipients.Add(distribution_list)
        recipient.Type = 2    

    # Display the email message (leave it open for editing)
    mail.Display(False)

if __name__ == '__main__':
    run()