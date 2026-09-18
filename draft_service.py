"""Sequential draft writer. No Send or Display calls. SQLite survives restarts."""
import hashlib
import json
import sqlite3
import time
from pathlib import Path
from desktop_draft_automation import build_html

class History:
    def __init__(self, path):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(path)
        self.db.execute('CREATE TABLE IF NOT EXISTS drafts (mailbox TEXT, case_id TEXT, fingerprint TEXT, status TEXT, entry_id TEXT, PRIMARY KEY(mailbox,case_id))')
        self.db.commit()

    def reserve(self, mailbox, case):
        content = {k: case[k] for k in ('name','email','address1','address2','city','state','country','postal')}
        content['assets'] = sorted(case['assets'])
        fingerprint = hashlib.sha256(json.dumps(content, sort_keys=True).encode()).hexdigest()
        self.db.execute('BEGIN IMMEDIATE')
        try:
            prior = self.db.execute('SELECT fingerprint,status FROM drafts WHERE mailbox=? AND case_id=?', (mailbox,case['case_id'])).fetchone()
            if prior:
                self.db.commit()
                if prior[1] == 'pending': return 'Needs review: earlier save interrupted; check Outlook before retrying'
                return 'Already prepared' if prior[0] == fingerprint else 'Needs review: case changed since earlier draft'
            self.db.execute('INSERT INTO drafts VALUES (?,?,?,?,?)', (mailbox,case['case_id'],fingerprint,'pending',''))
            self.db.commit()
            return None
        except Exception:
            self.db.rollback()
            raise

    def complete(self, mailbox, case, entry):
        self.db.execute('UPDATE drafts SET status=?, entry_id=? WHERE mailbox=? AND case_id=?', ('saved',entry,mailbox,case['case_id']))
        self.db.commit()

    def close(self):
        self.db.close()

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
        message = self.folder.Items.Add('IPM.Note')
        try:
            message.SendUsingAccount = self.account
            message.BodyFormat = 2
            message.To = case['email']
            message.Subject = case['subject']
            message.UserProperties.Add('AssetReturnCase', 1).Value = case['case_id']
            for cid, path in attachments.items():
                attachment = message.Attachments.Add(str(Path(path).resolve()), 1, 0)
                props = attachment.PropertyAccessor
                props.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x3712001F', cid)
                props.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x370E001F', 'image/png')
                props.SetProperty('http://schemas.microsoft.com/mapi/proptag/0x7FFE000B', True)
                attachment = props = None
            message.HTMLBody = body
            message.Save()
            entry = str(message.EntryID)
            message.Close(0)
            return entry
        finally:
            message = None

def run_cases(cases, writer, history, banner, report, progress, cancelled, limit=None, delay=0.25):
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
                status = history.reserve(writer.mailbox, case)
                if status is None:
                    body, attachments = build_html(case, banner)
                    try:
                        entry = writer.save(case, body, attachments)
                        history.complete(writer.mailbox, case, entry)
                    except Exception:
                        log.write(case['case_id'] + ': Save could not be confirmed. Check Outlook. Batch stopped.\n')
                        log.flush()
                        raise RuntimeError('Outlook save could not be confirmed. Processing stopped. Check Outlook and the results report; this case will not be retried automatically.') from None
                    created += 1
                    status = 'Draft created'
                    time.sleep(delay)
            log.write(case['case_id'] + ': ' + status + '\n')
            log.flush()
            progress(f'{i}/{len(cases)} cases checked. {created} drafts created.')
    return created
