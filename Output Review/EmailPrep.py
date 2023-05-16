import pandas as pd
import datetime as dt
import win32com.client as win32


def run():
    FILE_PATH = 'M:/CPP-Data/Sutherland RPA/ChargeCorrection'

    today = dt.date.today()
    # file_date = today - dt.timedelta(days=1)
    # if today is Monday
    if today.weekday() == 0:
        # set FILE_DATE to today - 3
        file_date = today - dt.timedelta(days=3)
    else:
        # set FILE_DATE to today - 1
        file_date = today - dt.timedelta(days=1)

    last_day_of_month = (file_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    if file_date == last_day_of_month:
        # backtrack to the previous day until it's a business day
        while True:
            file_date -= timedelta(days=1)
            if file_date.weekday() < 5:
                break

    fd_mmddyyyy = file_date.strftime('%m%d%Y')
    fd_mm_yyyy = file_date.strftime('%m %Y')
    file_year = file_date.strftime('%Y')
    fd_mm_dd = file_date.strftime('%m/%d')

    file_to_review = f'{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}/DP Comments Northwell_ChargeCorrection_Output_{fd_mmddyyyy}.xlsx'

    df_email = pd.read_excel(file_to_review, sheet_name="Sheet3", engine="openpyxl")

    user_list = []
    email_list = []
    for rep in df_email['Rep Name'].unique():
        user_list.append(rep)

    # Create an instance of the Outlook application
    outlook = win32.Dispatch("Outlook.Application")

    # Get the MAPI namespace of the Outlook application
    namespace = outlook.GetNamespace("MAPI")

    output_path = 'C:/Users/denglish2/Desktop/emailoutput.txt'
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
            # open the file in write mode and write the output to it
            with open(output_path, 'w') as f:
                f.write(f"No user found with alias or display name '{user}'")                
            # open the file for reading and print its contents to the console
            with open(output_path, 'r') as f:
                print(f.read())

    # Create an instance of the Outlook application
    outlook = win32.Dispatch("Outlook.Application")

    # Create a new email message
    mail = outlook.CreateItem(0)

    mail.Subject = f"The {fd_mm_dd} CCN output file is ready for review."

    # Convert the DataFrame to an HTML table
    html_table = df_email.to_html(index=False, classes="dataframe", border=2, justify="center")

    html_body = f"""
    <p>Greetings,</p>
    <p>The {fd_mm_dd} CCN output file is ready for review.</p>
    <p><strong><u>Supervisors</u></strong> - The complete file can be found in the following location: <a href="file:///{FILE_PATH}/{file_year}/{fd_mm_yyyy}/{fd_mmddyyyy}">Click Here</a>. The hyperlink will take you to the specific folder for the File Date being reviewed, review the file marked <strong><u>“DP Comments”</u></strong>.</p>
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