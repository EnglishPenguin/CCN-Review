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
                                 
                                 
                                
Step 1 - Rune FileCombine.py
This will iterate through the each user's CCN Input files for AR Support, Payer 1 and Payer 2 and combine them into one file.
It will be saved at the following location then separated specifically by date. 
M:\CPP-Data\Sutherland RPA\Northwell Process Automation ETM Files\Monthly Reports\Charge Correction\Audits - Files Sent to Bot

Step 2 - Take the invoice Numbers to Athena IDX and load them into your custom table. Then run the RPA_CCN_VERIFY query. 
If you need to copy the query the full name is: DENGLISH2_RPA_CCN_VERIFY. Be sure to update the query to use your specific custom table

Step 3 - Open the file and create a new sheet. Ensure it is named 'Sheet2'. Add the following as columns to that sheet. There will be a future enhancement that
will do this automatically. Do not change the name of any of the columns in this list. Be sure to transpose the list so they are the column headers. 
Save the file before moving to Step 4.

Invoice
BAR_B_INV.SER_DT,
BAR_B_TXN.SER_DT,
BAR_B_INV.TOT_CHG,
INV_BAL,
PROV__1,
LOC__2,
BAR_B_INV.ORIG_FSC__5,
BAR_B_INV.DX_ONE__3,
DX_TWO__3,
DX_THREE__3,
DX_FOUR__3,
DX_FIVE__3,
DX_SIX__3,
DX_SEVEN__3,
DX_EIGHT__3,
DX_NINE__3,
DX_TEN__3,
BAR_B_INV.DX_ELEVEN__3,
BAR_B_INV.DX_TWELVE__3,
TXN_NUM,
PROC__2,
MOD,
BAR_B_TXN.DX_NUM,
BAR_B_INV.CHG_CORR_FLAG,
BAR_B_INV.CORR_INV_NUM

Step 4. Rune FileManipulation.py
This will output the manipulations to your clipboard

Step 5. Create a new sheet named 'Sheet3' and paste the values from Step 4.

Step 6. Review 'Sheet3' and the various columns. Any row that needs to be excluded must have the word 'Exclude' in the last column.
Once the Manual review is complete, save the file

Step 7. Run file_to_csv.py
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
                                 

Step 1 - Copy the CCN Output Checker.xlsb and DP Comments Template.xlsx from the following file path:
M:\CPP-Data\Sutherland RPA\Northwell Process Automation ETM Files\Monthly Reports\Charge Correction\References
Paste the two files to the appropriate file folder by date

Step 2 - Run file_review.py
This will copy the data from the output file from Sutherland to the DP Comments Template file. 
It will also assign the DP Status/Comment/Category/Action for the most likely Retrieval Descriptions.

Step 3 - Take the invoice Numbers to Athena IDX and load them into your custom table. Then run the RPA_CCN_VERIFY query. 
If you need to copy the query the full name is: DENGLISH2_RPA_CCN_VERIFY. Be sure to update the query to use your specific custom table

Step 4 - Paste the results from the IDX query onto the "Original - DBMS" tab of the CCN Output Checker.xlsb file

Step 5 - Perform manual review of the CCN Output. 
Invoices labeled with C00 are almost always successful. You will need to review these to ensure the rep did not use the wrong step which would leave a credit bal on the original.
Invoices with an E## retrieval code will need to have the screenshots reviewed. Especially the ones which have "REVIEW" listed in the cell for the DP Comment/DP Category
Invoices with a P00 error will need to be manually viewed to determine the error.

Step 6 - Create a pivot table of the DP Commnets Template 'Sheet1'. Use Department, Supervisor, Representative for the Rows. Use DP Category for the columns.
For Values use the Count of DP Category.

Step 7 - Create a new sheet 'Sheet3'. 

Step 8 - Filter 'Sheet1' for any errors. Copy the Representative Column to 'Sheet3'. Copy the Invoice Number coluymn to 'Sheet3'. 
Copy the DP Category and Action columns to 'Sheet3'. Save the DP Comments Template file

Step 9 - Run EmailPrep.py
This will grab the user names from 'Sheet3' and collect their email addresses. It will create an email to the Representatives that need to fix their errors.
It will also populate the AR Supervisors, Danielle, Rich Johnson, John Mullen, Amanda Lang, Vincent to the CC field. 
It will give the email a subject and paste in the body of the email which will include the information on 'Sheet3'

Step 10 - Send the email.