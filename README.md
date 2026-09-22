# Salesforce Email Drafts

Download and extract this private repository. Follow START HERE.txt.
Run Setup Once.bat once, then Prepare Drafts.bat for daily use.
Requires Windows, Python 3.11+, and a working Classic Outlook profile for direct drafts.

Choose Excel and the specific Outlook mailbox in the window. Test two dummy drafts before a full batch. Messages save sequentially to the selected account's Drafts folder without opening windows or sending mail. The signature banner is embedded.

If a save cannot be confirmed, the app stops so you can check Outlook before continuing.

The first worksheet must use the supplied column headers. Rows group by Case ID; different Case IDs remain separate. Conflicting identity/address fields, missing required information and duplicate flags are held for review. Duplicate flags are not blindly discarded: they may mark repeated cases containing different assets. Identical asset descriptions are deduplicated.

Every run is independent: the same Excel file can be prepared again whenever needed. Previously sent manual emails are not detected: select only eligible new cases in the daily export. Follow-ups are outside this initial-email workflow.

Output contains dated runs with Results.txt. The desktop window creates drafts directly in the selected Classic Outlook mailbox; it does not create EML test files.

Automated tests use a fake Outlook writer and dummy data. Outlook 2013 rendering, permissions and actual saves still require a two-case test on the target desktop. Compatibility is not a claim of current Microsoft support.

Developer checks: python -m unittest discover -s tests -v

Use the desktop launcher for direct drafts.
