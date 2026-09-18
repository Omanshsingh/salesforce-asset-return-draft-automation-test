import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
from desktop_draft_automation import read_cases, build_html, create_eml, recipients
from draft_service import History, OutlookWriter, run_cases

BASE = Path(__file__).resolve().parent

class App:
    def __init__(self, root):
        self.root = root
        root.title('Salesforce Email Drafts'); root.geometry('680x420')
        self.events = queue.Queue(); self.stop = threading.Event()
        self.busy = False; self.cases = []; self.controls = []
        self.report_dir = BASE / 'Output'
        box = ttk.Frame(root, padding=22); box.pack(fill='both', expand=True)
        ttk.Label(box, text='Prepare email drafts', font=('Segoe UI', 18)).pack(anchor='w')
        ttk.Label(box, text='Choose Excel, select the mailbox, then prepare drafts. Review and send in Outlook.\nMultiple assets under the same Case ID go into one email.', wraplength=630).pack(anchor='w', pady=10)
        self.button(box, '1. Choose Excel file', self.choose)
        self.file_label = ttk.Label(box, text='No file selected', wraplength=630)
        self.file_label.pack(anchor='w', pady=6)
        row = ttk.Frame(box); row.pack(fill='x', pady=8)
        self.button(row, '2. Load Outlook accounts', self.accounts, side='left')
        self.account = ttk.Combobox(row, state='readonly', width=37)
        self.account.pack(side='left', padx=8); self.controls.append(self.account)
        row2 = ttk.Frame(box); row2.pack(fill='x', pady=10)
        self.button(row2, 'Test 2 dummy drafts', lambda: self.start('test'), side='left')
        self.button(row2, '3. Prepare Outlook drafts', lambda: self.start('outlook'), side='left')
        row3 = ttk.Frame(box); row3.pack(fill='x')
        self.button(row3, 'Save EML test files', lambda: self.start('eml'), side='left')
        ttk.Button(row3, text='Stop after current case', command=self.stop.set).pack(side='left', padx=6)
        ttk.Button(row3, text='Open results folder', command=self.open_results).pack(side='left')
        self.status = ttk.Label(box, text='Ready. Nothing is sent automatically.', wraplength=630)
        self.status.pack(anchor='w', pady=18)
        root.protocol('WM_DELETE_WINDOW', self.close); root.after(100, self.poll)

    def button(self, parent, label, command, **pack):
        b = ttk.Button(parent, text=label, command=command)
        b.pack(anchor='w', **pack); self.controls.append(b)

    def choose(self):
        path = filedialog.askopenfilename(title='Choose Excel file', initialdir=str(BASE), filetypes=[('Excel workbook', '*.xlsx')])
        if not path: return
        try:
            cases, skipped = read_cases(Path(path))
            if not cases: raise ValueError('No cases found on the first worksheet.')
            self.cases = cases; ready = sum(not c['missing_fields'] for c in cases)
            self.file_label.config(text=f'{Path(path).name}\n{len(cases)} cases: {ready} ready, {len(cases)-ready} need review. {len(skipped)} rows without Case ID.')
        except Exception as e:
            self.cases = []; messagebox.showerror('Check Excel', str(e))

    def accounts(self):
        try:
            import win32com.client
            app = win32com.client.Dispatch('Outlook.Application')
            values = [str(a.SmtpAddress) for a in app.Session.Accounts if str(a.SmtpAddress)]
            if not values: raise RuntimeError('No accounts')
            self.account['values'] = values
            if len(values) == 1: self.account.current(0)
            self.status.config(text='Select the mailbox where drafts should be saved.')
        except Exception:
            messagebox.showerror('Classic Outlook required', 'Open Classic Outlook with your working email account. New Outlook does not support this connection. You can still use Save EML test files.')

    def start(self, mode):
        if self.busy: return
        if not self.cases:
            messagebox.showinfo('Choose Excel', 'Select your Excel file first.'); return
        if mode != 'eml' and not self.account.get():
            messagebox.showinfo('Select mailbox', 'Load Outlook accounts and select a mailbox first.'); return
        cases = list(self.cases)
        if mode == 'test':
            if any(not a.lower().endswith('@example.test') for c in cases for a in recipients(c['email'])):
                messagebox.showerror('Use dummy data', 'Test 2 dummy drafts requires only @example.test recipients.'); return
            cases.sort(key=lambda c: (len(c['assets']) < 2, c['case_id']))
        count = sum(not c['missing_fields'] for c in cases)
        if mode != 'eml' and not messagebox.askokcancel('Prepare drafts', f'Create up to {min(count,2) if mode == "test" else count} drafts in {self.account.get()}?\nPreviously processed cases will be skipped. Nothing will be sent.'): return
        self.busy = True; self.stop.clear()
        for w in self.controls: w.config(state='disabled')
        self.status.config(text='Preparing. Please keep Outlook open.')
        threading.Thread(target=self.work, args=(mode,cases,self.account.get()), daemon=True).start()

    def work(self, mode, cases, account):
        history = None; com = None
        try:
            folder = BASE / 'Output' / datetime.now().strftime('%Y-%m-%d_%H-%M-%S-%f')
            folder.mkdir(parents=True); self.report_dir = folder
            banner = BASE / 'signature-banner.png'
            if not banner.is_file(): raise RuntimeError('Signature image missing. Restore signature-banner.png.')
            if mode == 'eml':
                count = 0
                with (folder / 'Results.txt').open('w', encoding='utf-8') as report:
                    for c in cases:
                        if self.stop.is_set(): break
                        if c['missing_fields']:
                            report.write(c['case_id'] + ': Needs review: ' + '; '.join(c['missing_fields']) + '\n'); continue
                        body, attachments = build_html(c,banner)
                        create_eml(c,body,folder / 'Email files',attachments); count += 1
                        report.write(c['case_id'] + ': Email file created\n')
                result = f'{count} email files saved. Open results folder and test one file at a time.'
            else:
                import pythoncom
                com = pythoncom; com.CoInitialize()
                writer = OutlookWriter(account)
                history = History(BASE / 'Data' / 'draft-history.sqlite')
                count = run_cases(cases,writer,history,banner,folder / 'Results.txt',lambda t:self.events.put(('progress',t)),self.stop.is_set,limit=2 if mode == 'test' else None)
                result = f'{count} new drafts saved in {account}. Review Outlook Drafts. See Results.txt for skipped cases.'
                writer = None
            self.events.put(('done',result))
        except Exception as e: self.events.put(('error',str(e)))
        finally:
            if history: history.close()
            if com: com.CoUninitialize()

    def poll(self):
        while not self.events.empty():
            kind, msg = self.events.get(); self.status.config(text=msg)
            if kind != 'progress':
                self.busy = False
                for w in self.controls: w.config(state='readonly' if w is self.account else 'normal')
                if kind == 'error': messagebox.showerror('Processing stopped',msg)
        self.root.after(100,self.poll)

    def open_results(self):
        self.report_dir.mkdir(parents=True,exist_ok=True); os.startfile(str(self.report_dir))

    def close(self):
        if self.busy:
            self.stop.set(); self.status.config(text='Stopping after current case. Wait before closing.')
        else: self.root.destroy()

if __name__ == '__main__':
    root = tk.Tk(); App(root); root.mainloop()
