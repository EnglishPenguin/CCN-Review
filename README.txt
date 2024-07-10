 _____ _                                           
/  __ \ |                                          
| /  \/ |__   __ _ _ __ __ _  ___                  
| |   | '_ \ / _` | '__/ _` |/ _ \                 
| \__/\ | | | (_| | | | (_| |  __/                 
 \____/_| |_|\__,_|_|  \__, |\___|                 
                        __/ |                      
                       |___/                       
 _____                          _   _              
/  __ \                        | | (_)             
| /  \/ ___  _ __ _ __ ___  ___| |_ _  ___  _ __   
| |    / _ \| '__| '__/ _ \/ __| __| |/ _ \| '_ \  
| \__/\ (_) | |  | | |  __/ (__| |_| | (_) | | | | 
 \____/\___/|_|  |_|  \___|\___|\__|_|\___/|_| |_| 
                                                   
                                                   
______           _                                 
| ___ \         (_)                                
| |_/ /_____   ___  _____      __                  
|    // _ \ \ / / |/ _ \ \ /\ / /                  
| |\ \  __/\ V /| |  __/\ V  V /                   
\_| \_\___| \_/ |_|\___| \_/\_/                    
                                                   
                                                   
The Scripts in this project will perform many of the repetive steps for the Charge Correction Combine and Output Review Process.
This document will describe the order the files should be used, what each script should accomplish and also identify any other steps that need to be taken.


 _____                  _        
|_   _|                | |       
  | | _ __  _ __  _   _| |_      
  | || '_ \| '_ \| | | | __|     
 _| || | | | |_) | |_| | |_      
 \___/_| |_| .__/ \__,_|\__|     
           | |                   
           |_|                   
______           _               
| ___ \         (_)              
| |_/ /_____   ___  _____      __
|    // _ \ \ / / |/ _ \ \ /\ / /
| |\ \  __/\ V /| |  __/\ V  V / 
\_| \_\___| \_/ |_|\___| \_/\_/  
                                 
                                 
                                
Step 1 - Run input_main.py with Option 1
This will iterate through the each user's CCN Input files for AR Support, Payer 1, Payer 2 and Payer 5 and combine them into one file.
It will be saved at the following location then separated specifically by date. 
M:\CPP-Data\Sutherland RPA\Northwell Process Automation ETM Files\Monthly Reports\Charge Correction\Audits - Files Sent to Bot

Step 2 - Remove the duplicate invoices from Sheet1 on this new file created

Step 3 - Take the invoice Numbers to Athena IDX and load them into your custom table. Then run the RPA_CCN_VERIFY_PAYCODE query. 
If you need to copy the query the full name is: DENGLISH2_RPA_CCN_VERIFY_PAYCODE. Be sure to update the query to use your specific custom table

Step 4 - Open the file and save the output of the query to Sheet2 on the file created in Step1. Save and close the file before proceeding

Step 5. Rune input_main.py with Option 2
This will output the manipulations to a new sheet named 'Sheet3' on the same file. It will automatically populate Emails for invoices that 
will be excluded due to established business rules

Step 6. Review 'Sheet3' and the various columns. Any row that needs to be excluded must have the word 'Exclude' in the last column.
Once the Manual review is complete, save the file

Step 7. Run input_main.py with Option 3
This will remove any row that has a value of 'Exclude' in the final column. It will also drop all of the validation columns used to help with the manual review.
This will save the .csv version of the file for Sutherland to pick up at the M:\CPP-Data\Sutherland RPA\ChargeCorrection file path




 _____       _               _   
|  _  |     | |             | |  
| | | |_   _| |_ _ __  _   _| |_ 
| | | | | | | __| '_ \| | | | __|
\ \_/ / |_| | |_| |_) | |_| | |_ 
 \___/ \__,_|\__| .__/ \__,_|\__|
                | |              
                |_|              
______           _               
| ___ \         (_)              
| |_/ /_____   ___  _____      __
|    // _ \ \ / / |/ _ \ \ /\ / /
| |\ \  __/\ V /| |  __/\ V  V / 
\_| \_\___| \_/ |_|\___| \_/\_/  
                                 

Step 1 - Run output_main.py with Option 1
This will make a copy of the DP Comments Template and CCN Checker files to the appropriate folder for the file date to be reviewed
This will copy the data from the output file from Sutherland to the DP Comments Template file. 
It will also assign the DP Status/Comment/Category/Action for the most likely Retrieval Descriptions.

Step 2 - Take the invoice Numbers to Athena IDX and load them into your custom table. Then run the RPA_CCN_VERIFY query. 
If you need to copy the query the full name is: DENGLISH2_RPA_CCN_VERIFY. Be sure to update the query to use your specific custom table

Step 3 - Paste the results from the IDX query onto the "Original - DBMS" tab of the CCN Output Checker.xlsb file

Step 4 - Perform manual review of the CCN Output. 
Invoices labeled with C00 are almost always successful. You will need to review these to ensure the rep did not use the wrong step which would leave a credit bal on the original.
Invoices with an E## retrieval code will need to have the screenshots reviewed. Especially the ones which have "REVIEW" listed in the cell for the DP Comment/DP Category
Invoices with a P00 error will need to be manually viewed to determine the error.
If you want to exclude the invoice from being sent to the Rep, DP Category needs to = 'Success' or Action needs to = 'No Action Needed'
Perform a VLOOKUP against the CCN Checker for negative balances and check the step. If step 2 = Partial Success, rep needs to carry forward payments
Perform a VLOOKUP against the CCN Checker for Corrected Invoice Number. If the invoice has been corrected and not a Rep Error, put "No Action Needed" in the "Action" column

Step 5 - Create a pivot table of the DP Comments Template 'Sheet1'. Use Department, Supervisor, Representative for the Rows. Use DP Category for the columns.
For Values use the Count of DP Category. Rename the file

Step 6 - Run output_main.py with Option 2
This will grab the user names from 'Sheet1', filter out any Successes and any line listed with 'No Action Needed' in the action column and collect their email addresses. 
It will create a new 'Sheet3' on the file with this information
It will create an email to the Representatives that need to fix their errors.
It will also populate the AR Supervisors Group, Danielle, Rich Johnson, John Mullen, Amanda Lang, Vincent to the CC field. 
It will give the email a subject and paste in the body of the email which will include the information on 'Sheet3'

Step 7 - Send the email.