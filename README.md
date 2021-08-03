# ecometry-POSys-step1-ImportSuggestedBuyData
Automatically runs three Red Prairie Ecometry recommended buy reports, parses the resulting text files, enters the parsed items into a working database, and emails managers the results.

ImportSuggestedBuyData (this program)
Automatically runs three Ecometry recommended buy reports for A1, C3 and R1 items.
Parses the resulting text files, 
Enters the parsed items into a working database
Emails managers the results of the run, so they can use the RecmdBuys to preapprove items for the Purchasing department personnel.

RecmdBuys main features
Restrict running of the program to authorized managers
Allows greater access and visibility to the COO
Allows managers to change the recommended purchase quantities or remove unwanted items.
Displays items, styles, sales, activity, and advertisement reports on demand


POMaker main features
Restrict running of the program to authorized users.
If the COO runs the program, he is prompted to select a run-as user
Displays the approved items to the purchasing clerk, the 
Allows selecting different vendors
Allows adjusting item cost
It prompts for manager's password for price increases of over 2%
Creates Quote email with professionally formatted Excel spreadsheet attachment for vendor and prefills vendor’s name and email from information in Ecometry system.  Email subject contains a pre-validated PO number. Email body included company standard language and purchasing clerk company standard-compliant signature.
Displays PO as html documents before submitting to Ecometry
Records email sent, Excel document, PO html file on file system
Enters new PO in Ecometry, so it appears on the items' histories.
