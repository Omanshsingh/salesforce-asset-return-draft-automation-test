import copy
import tempfile
import unittest
from unittest.mock import MagicMock
from pathlib import Path
from email import policy
from email.parser import BytesParser
from desktop_draft_automation import read_cases, build_html, create_eml
from draft_service import History, run_cases, OutlookWriter

BASE = Path(__file__).resolve().parents[1]

class FakeWriter:
    mailbox = 'test-mailbox'
    def __init__(self): self.saved = []; self.fail = False
    def save(self, case, body, attachments):
        if self.fail: raise RuntimeError('ambiguous failure')
        self.saved.append(case['case_id']); return str(len(self.saved))

class Tests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.folder = Path(self.tmp.name)
        self.history = History(self.folder/'history.sqlite')
        self.cases, _ = read_cases(BASE/'Omansh_Test_Dummy_100_Cases.xlsx')
        self.writer = FakeWriter()
    def tearDown(self): self.history.close(); self.tmp.cleanup()
    def run_batch(self, **kwargs):
        return run_cases(self.cases,self.writer,self.history,BASE/'signature-banner.png',self.folder/'Results.txt',lambda t:None,lambda:False,delay=0,**kwargs)
    def test_100_cases_and_rerun(self):
        self.assertEqual(len(self.cases),100)
        self.assertEqual(sum(len(c['assets']) for c in self.cases),157)
        self.assertTrue(all(c['email'].endswith('@example.test') for c in self.cases))
        self.assertEqual(self.run_batch(),100)
        self.assertEqual(self.run_batch(),0)
    def test_limit_and_changes(self):
        self.assertEqual(self.run_batch(limit=2),2)
        changed = copy.deepcopy(self.cases[0]); changed['address1'] += ' changed'
        self.assertIn('changed',self.history.reserve(self.writer.mailbox,changed))
    def test_uncertain_save_blocks_retry(self):
        self.writer.fail = True
        with self.assertRaises(RuntimeError): self.run_batch(limit=2)
        self.assertIn('interrupted',self.history.reserve(self.writer.mailbox,self.cases[0]))
        self.assertFalse(self.writer.saved)
    def test_eml_image_and_grouping(self):
        case = next(c for c in self.cases if len(c['assets'])==3)
        body, images = build_html(case,BASE/'signature-banner.png')
        p = create_eml(case,body,self.folder,images)
        msg = BytesParser(policy=policy.default).parsebytes(p.read_bytes())
        self.assertEqual(msg.get_content_type(),'multipart/related')
        image = next(p for p in msg.walk() if p.get_content_type()=='image/png')
        self.assertEqual(image.get_payload(decode=True),(BASE/'signature-banner.png').read_bytes())
        for asset in case['assets']: self.assertIn(asset,msg.get_body(('html',)).get_content())
        self.assertIn('width="655"',body)
        self.assertFalse(list(self.folder.glob('*.html')))
    def test_outlook_adapter_saves_without_opening_or_sending(self):
        writer = OutlookWriter.__new__(OutlookWriter)
        writer.account = object(); writer.folder = MagicMock()
        message = writer.folder.Items.Add.return_value
        message.EntryID = 'saved-id'
        body, images = build_html(self.cases[0], BASE/'signature-banner.png')
        self.assertEqual(writer.save(self.cases[0],body,images),'saved-id')
        message.Save.assert_called_once()
        message.Display.assert_not_called()
        message.Send.assert_not_called()
        self.assertIs(message.SendUsingAccount,writer.account)
        message.Attachments.Add.assert_called_once()
    def test_history_survives_reopen(self):
        self.run_batch(limit=2)
        self.history.close()
        self.history = History(self.folder/'history.sqlite')
        self.assertEqual(self.run_batch(),98)

if __name__ == '__main__': unittest.main()
