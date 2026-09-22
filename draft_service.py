"""Sequential draft writer. No Send or Display calls."""
import time
from pathlib import Path
from desktop_draft_automation import build_html, recipients

class OutlookWriter:
    def __init__(self, smtp):
        import win32com.client
        self.app = win32com.client.Dispatch('Outlook.Application')
        accounts = [a for a in self.app.Session.Accounts if str(a.SmtpAddress).lower() == smtp.lower()]
        if len(accounts) != 1:
            raise RuntimeError('The selected Outlook account is unavailable. Open Classic Outlook and try again.')
        self.account = accounts[0]
        self.folder = self.account.DeliveryStore.GetDefaultFolder(16)
        self.mailbox = smtp.lower() + '|' + str(self.folder.StoreID)

    def save(self, case, body, attachments):
        # Create the item in the selected mailbox's Drafts folder. This avoids
        # Outlook 2013 saving a generic item against the wrong store.
        message = self.folder.Items.Add(0)
        try:
            # The folder already selects the mailbox. Some older/IMAP Outlook
            # profiles reject SendUsingAccount even though the draft is valid.
            try:
                message.SendUsingAccount = self.account
            except Exception:
                pass
            message.BodyFormat = 2
            message.To = '; '.join(recipients(case['email']))
            message.Subject = case['subject']
            for cid, path in attachments.items():
                attachment = message.Attachments.Add(str(Path(path).resolve()), 1, 0)
                props = attachment.PropertyAccessor
                props.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', cid)
                attachment = props = None
            message.HTMLBody = body
            message.Save()
            entry = str(message.EntryID)
            if not entry:
                raise RuntimeError('Outlook returned no draft ID after Save.')
            return entry
        finally:
            message = None

def run_cases(cases, writer, banner, report, progress, cancelled, limit=None, delay=0.25):
    created = 0
    with Path(report).open('w', encoding='utf-8') as log:
        log.write('Email draft results\n\n')
        for i, case in enumerate(cases, 1):
            if cancelled():
                log.write('Stopped by user. Saved drafts remain in Outlook.\n')
                break
            if limit is not None and created >= limit: break
            reason = '; '.join(case['missing_fields'])
            if reason:
                status = 'Needs review: ' + reason
            else:
                body, attachments = build_html(case, banner)
                try:
                    writer.save(case, body, attachments)
                except Exception as exc:
                    log.write(case['case_id'] + ': Save could not be confirmed. Check Outlook. Batch stopped.\n')
                    log.write('Technical detail: ' + repr(exc) + '\n')
                    log.flush()
                    raise RuntimeError('Outlook save could not be confirmed. Processing stopped. Check Outlook and Results.txt for the technical detail.') from exc
                created += 1
                status = 'Draft created'
                time.sleep(delay)
            log.write(case['case_id'] + ': ' + status + '\n')
            log.flush()
            progress(f'{i}/{len(cases)} cases checked. {created} drafts created.')
    return created
