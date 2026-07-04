==============================================================================
  AT&T  ->  SALESFORCE  ASSET IMPORT PREP
  Quick-start guide
==============================================================================

WHAT THIS DOES
--------------
It takes the two reports you download each week (the AT&T Premier export and the
Salesforce "AT&T Client Assets" report), compares them, and produces a CSV that
contains ONLY the assets that are NOT already in Salesforce -- ready to feed into
the Salesforce import wizard. It does NOT touch Salesforce itself; you still run
the import wizard by hand at the end.

WHERE THE TWO SOURCE REPORTS COME FROM
---------------------------------------
  Report 1 -- AT&T Premier  (att.com/premier)
              AT&T's business account portal. Go to the Inventory / Asset section,
              export the full device list. Save as  .xlsx  (NOT .csv).
              File downloads as:  detail_report.xlsx

  Report 2 -- Salesforce  (your Salesforce org)
              In Salesforce go to: App Launcher -> [your Operations app] -> Reports tab
              -> All Folders -> AT&T Assets folder -> "AT&T Client Assets" report.
              (Adjust this path to wherever your org keeps the AT&T assets report.)
              Download it. File is named like:
              All Active ATT Clients Assets-<date>.xlsx


ONE-TIME SETUP  (do these steps in order)
------------------------------------------
STEP 1 -- Double-click  "Install_Requirements.bat"
          This installs Python and all required packages.
          (If Python is already on your computer it just installs the packages.)
          Follow the on-screen instructions. You only need to do this once.

STEP 2 -- Double-click  "Setup_Weekly_Schedule.bat"
          This schedules the tool to run automatically every MONDAY at 8:30 AM.
          You'll see a "SUCCESS!" message. That's it.


EVERY WEEK  (your only recurring task)
--------------------------------------
1. Download the two reports (see "WHERE THE TWO SOURCE REPORTS COME FROM" above).
2. Drop BOTH files into the  "input"  folder inside this tool folder.
   Do this before Monday 8:30 AM.

That's all. On Monday morning the tool runs on its own.


WHERE THE RESULTS GO
--------------------
In the  "output"  folder you'll find:
   ATT_Import_<date>.csv          <- upload THIS one in the Salesforce wizard
   ATT_Import_<date>_REVIEW.csv   <- only appears if some rows need a human look
                                     (usually a bad/missing Universal ID -- these
                                     rows were held back to avoid a failed import).

A file called  last_run_log.txt  (in the main tool folder) shows exactly what
happened on the last run. Check it if anything seems off.


WANT TO RUN IT RIGHT NOW (not wait for Monday)?
-----------------------------------------------
Put the two files in the "input" folder, then double-click  "Run_Now.bat".
A window opens, shows the results, and waits for you to press a key.


SAFETY FEATURES (so it never makes a mess)
-------------------------------------------
- If a required file is missing from "input" it stops and says so clearly.
- If the files in "input" look old (more than 6 days), it assumes you forgot to
  download this week's reports and STOPS -- it will not silently re-use last
  week's files.
- After each successful run, the files you dropped in "input" are moved into
  "input\_processed\<date>\" automatically. You never have to clean up after it,
  and old files can never be accidentally re-processed.


TO TURN OFF the automatic weekly run
--------------------------------------
Double-click  "Remove_Weekly_Schedule.bat".


QUESTIONS?
----------
Reach out to your IT/automation contact.
==============================================================================
