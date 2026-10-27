# Salesforce Email Drafts

Download and extract this private repository. Follow START HERE.txt.
Double-click Setup Once.bat on first use, then Prepare Drafts.bat for daily use.
The setup checks for Python 3.11 or newer even when Python is missing from PATH. If no compatible interpreter is found, it tries to install Python 3.12 with Windows Package Manager (winget), creates the project .venv, and installs the required packages. If winget or company policy blocks installation, setup shows the official Python download link. Setup can be rerun safely to repair the environment.

Requires Windows, Python 3.11+, and a working Classic Outlook profile for direct drafts.

Choose Excel and the specific Outlook mailbox in the window. Test two dummy drafts before a full batch. Messages save sequentially to the selected account's Drafts folder without opening windows or sending mail. The signature banner is embedded.

If a save cannot be confirmed, the app stops so you can check Outlook before continuing.

The first worksheet must use the supplied column headers. Rows group by Case ID; different Case IDs remain separate. Conflicting identity/address fields and missing required information are held for review. The sheet's Duplicate flag does not block drafting or produce a review error; flagged cases are drafted like normal cases. Identical asset descriptions are deduplicated.

Every run is independent: the same Excel file can be prepared again whenever needed. Previously sent manual emails are not detected: select only eligible new cases in the daily export. Follow-ups are outside this initial-email workflow.

Output contains dated runs with Results.txt. The desktop window creates drafts directly in the selected Classic Outlook mailbox; it does not create EML test files.

Automated tests use a fake Outlook writer and dummy data. Outlook 2013 rendering, permissions and actual saves still require a two-case test on the target desktop. Compatibility is not a claim of current Microsoft support.

Developer checks: python -m unittest discover -s tests -v

Use the desktop launcher for direct drafts.
