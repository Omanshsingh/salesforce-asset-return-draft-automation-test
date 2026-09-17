# Salesforce asset-return draft automation

This is the desktop-first version. It reads the copied Salesforce workbook, ignores duplicate rows, groups repeated Case IDs, combines all distinct assets into one message, validates required fields, and creates Outlook drafts in the exact supplied format.

## Safe test

Open PowerShell in this folder and run:

```powershell
python -m pip install -r requirements.txt
python desktop_draft_automation.py "C:\path\to\dummy\Omansh_Test_Dummy_100_Cases.xlsx" --output test_output
```

The default is a dry run. It creates HTML previews, `manifest.json`, and `review.csv`. It cannot send an email.

## Create Outlook drafts

After reviewing the dry run, open classic Outlook, then run:

```powershell
python desktop_draft_automation.py "C:\path\to\copied\live_sheet.xlsx" --output draft_output --create-outlook-drafts
```

The script saves drafts only. The admin reviews and sends them manually from Outlook. Do not use this switch until the dry-run results are correct.

## Multiple assets and duplicate protection

- Rows are grouped by `Case ID`.
- Rows marked `Duplicate`, `Yes`, `Y`, `True`, or `1` are skipped.
- Distinct asset values are combined into one email.
- One asset produces the normal single-asset email; two or more produce one combined multiple-asset email.
- Missing name, email, address, or asset data is marked `Needs Review` and does not create a draft.
- A deterministic `draft_key` is written to the manifest for future duplicate protection.

The current version deliberately does not write back to the live/copied workbook. That should be added after the draft process is approved, using a separate processing log with Case ID, draft key, Outlook EntryID, timestamp, and status.
