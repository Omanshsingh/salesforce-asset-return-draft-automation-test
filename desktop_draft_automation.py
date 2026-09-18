"""Salesforce asset-return draft generator.

Desktop-first, draft-only automation for the copied Salesforce Excel workbook.
Default behavior is a dry run: it reads the workbook, groups assets by Case ID,
and writes HTML previews plus a manifest. Use --create-outlook-drafts only after
testing; it saves drafts in the currently logged-in classic Outlook profile and
does not send them.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import html
import json
import re
from collections import OrderedDict
from datetime import datetime
from email.message import EmailMessage
from email.utils import formatdate
from pathlib import Path
from typing import Any

import openpyxl


def text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return str(value).strip()


def first(row: dict[str, Any], *names: str) -> str:
    for name in names:
        value = text(row.get(name))
        if value:
            return value
    return ""


def is_yes(value: str) -> bool:
    return text(value).lower() in {"yes", "y", "true", "1", "duplicate"}


def valid_email(value: str) -> bool:
    # Supports the workbook's semicolon-separated multiple-recipient format.
    addresses = [part.strip() for part in re.split(r"[;,]", value) if part.strip()]
    return bool(addresses) and all(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", a) for a in addresses)


def recipients(value: str) -> list[str]:
    return [part.strip() for part in re.split(r"[;,]", value) if part.strip()]


def build_html(case: dict[str, Any], banner_path: Path | None = None) -> tuple[str, dict[str, str]]:
    esc = lambda value: html.escape(text(value), quote=True)
    address_parts = [case["address1"], case["address2"]]
    locality = ", ".join(part for part in [case["city"], case["state"], case["country"], case["postal"]] if part)
    if locality:
        address_parts.append(locality)
    address = "<br>".join(esc(part) for part in address_parts if part)
    assets = "".join(f"<div>{esc(asset)}</div>" for asset in case["assets"])
    cid = ""
    banner = ""
    attachments: dict[str, str] = {}
    if banner_path and banner_path.exists():
        cid = "ecoreco-banner@draft-automation"
        banner = f'<p><img src="cid:{cid}" style="max-width:820px;width:100%;height:auto" alt="EcoReco"></p>'
        attachments[cid] = str(banner_path)

    body = f"""<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><style>
p.MsoNormal,li.MsoNormal,div.MsoNormal {{ margin:0cm; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri",sans-serif; }}
p.wordsection1,li.wordsection1,div.wordsection1 {{ margin-right:0cm; margin-left:0cm; font-size:12.0pt; font-family:"Times New Roman",serif; }}
a:link {{ color:#0563C1; text-decoration:underline; }}
</style></head><body lang=EN-US link="#0563C1" vlink="#954F72"><div class=WordSection1>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Hello,<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Greetings from Ecoreco !!<o:p></o:p></span></p>
<p class=wordsection1 style='margin-bottom:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Eco Recycling Ltd, is authorized by Salesforce for the collection of assets from your below mentioned address, <o:p></o:p></span></b></p>
<p class=wordsection1 style='margin-bottom:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></b></p>
<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0 style='width:517.0pt;border-collapse:collapse'><tr style='height:40.9pt'><td width=228 style='width:171.0pt;border:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt;height:40.9pt'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;text-align:center'><span style='mso-fareast-language:EN-IN'>Pickup Address<o:p></o:p></span></p></td><td width=461 style='width:346.0pt;border:solid windowtext 1.0pt;border-left:none;padding:0cm 5.4pt 0cm 5.4pt;height:40.9pt'><p class=MsoNormal align=center style='text-align:center'><span style='color:black'>{address}<o:p></o:p></span></p></td></tr><tr style='height:66.05pt'><td width=228 style='width:171.0pt;border:solid windowtext 1.0pt;border-top:none;padding:0cm 5.4pt 0cm 5.4pt;height:66.05pt'><p class=MsoNormal align=center style='mso-margin-top-alt:auto;text-align:center'><span style='mso-fareast-language:EN-IN'>Asset Type &amp; Serial Number<o:p></o:p></span></p></td><td width=461 style='width:346.0pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0cm 5.4pt 0cm 5.4pt;height:66.05pt'><p class=MsoNormal><span style='color:black'>{assets}<o:p></o:p></span></p></td></tr></table>
<p class=wordsection1><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Kindly confirm the pickup address and asset serial no details so that we can initiate the reverse collection process.<br><br><b><span style='background:yellow;mso-highlight:yellow'>If there are any discrepancies, please inform us at your earliest convenience.</span></b><o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>We look forward to your response to proceed further.<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='color:black'>Thanks &amp; Regards,</span></b><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'><o:p></o:p></span></b></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='color:black'>CRM Executive</span></b><span lang=EN-IN style='color:black'>| +91-22-4005 2951/+91-9004149714| </span><a href="http://www.ecoreco.com/"><span style='color:black'>www.ecoreco.com</span></a><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'><o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='color:black'>Eco Recycling Limited</span></b><span lang=EN-IN style='color:black'> </span><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'><o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='color:black'>422, The Summit Business Park | Andheri Kurla road | Andheri (East), Mumbai 400093</span><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:black'><o:p></o:p></span></p>
{banner}
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>"""
    return body, attachments


def read_cases(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("The workbook is empty.")
    headers = [text(value) for value in rows[0]]
    grouped: OrderedDict[str, list[dict[str, Any]]] = OrderedDict()
    skipped: list[dict[str, Any]] = []
    for row_number, values in enumerate(rows[1:], start=2):
        row = {headers[i]: values[i] if i < len(values) else None for i in range(len(headers)) if headers[i]}
        case_id = first(row, "Case ID")
        if not case_id:
            skipped.append({"row": row_number, "reason": "Missing Case ID"})
            continue
        if is_yes(first(row, "Duplicate")):
            skipped.append({"row": row_number, "case_id": case_id, "reason": "Marked duplicate"})
            continue
        grouped.setdefault(case_id, []).append(row)

    cases: list[dict[str, Any]] = []
    for case_id, case_rows in grouped.items():
        row = case_rows[0]
        asset_values = []
        for item in case_rows:
            asset = first(item, "Asset Serial Number", "Asset Type & Serial Number", "Asset Details")
            if asset and asset not in asset_values:
                asset_values.append(asset)
        case = {
            "case_id": case_id,
            "name": first(row, "User's Name"),
            "email": first(row, "User's Email"),
            "phone": first(row, "User's Phone Number"),
            "address1": first(row, "Line Address 1"),
            "address2": first(row, "Line Address 2"),
            "city": first(row, "City"),
            "state": first(row, "State/Province"),
            "country": first(row, "Country"),
            "postal": first(row, "Zipcode/Postal Code"),
            "assets": asset_values,
            "source_row_count": len(case_rows),
        }
        missing = []
        if not case["name"]: missing.append("User's Name")
        if not valid_email(case["email"]): missing.append("User's Email")
        if not case["address1"] or not case["city"] or not case["country"]: missing.append("Pickup address")
        if not case["assets"]: missing.append("Asset details")
        case["missing_fields"] = missing
        case["status"] = "Needs Review" if missing else "Ready for Draft"
        case["subject"] = f"Salesforce Asset Return - {case['name'] or 'Customer'} - {case_id}"
        case["draft_key"] = hashlib.sha256(f"{case_id}|{case['email']}|{ '|'.join(case['assets'])}".encode()).hexdigest()[:20]
        cases.append(case)
    return cases, skipped


def create_outlook_draft(case: dict[str, Any], body: str, attachments: dict[str, str]) -> str:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Outlook draft mode needs pywin32. Install requirements.txt first.") from exc
    outlook = win32com.client.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = "; ".join(recipients(case["email"]))
    message.Subject = case["subject"]
    message.HTMLBody = body
    for cid, file_path in attachments.items():
        attachment = message.Attachments.Add(file_path, 1, 0, Path(file_path).name)
        attachment.PropertyAccessor.SetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
        )
    message.Save()
    return str(message.EntryID)


def create_eml(case: dict[str, Any], body: str, output: Path, attachments: dict[str, str]) -> Path:
    """Write an unsent RFC 822 email file that can be opened for review."""
    message = EmailMessage()
    message["To"] = "; ".join(recipients(case["email"]))
    message["Subject"] = case["subject"]
    message["Date"] = formatdate(localtime=True)
    message.set_content("Please open this message in an HTML-capable mail application.")
    message.add_alternative(body, subtype="html")
    for cid, file_path in attachments.items():
        data = Path(file_path).read_bytes()
        suffix = Path(file_path).suffix.lower()
        maintype, subtype = ("image", "png") if suffix == ".png" else ("image", "jpeg")
        message.get_payload()[-1].add_related(data, maintype=maintype, subtype=subtype, cid=f"<{cid}>", filename=Path(file_path).name)
    output.mkdir(parents=True, exist_ok=True)
    target = output / f"{case['case_id']}.eml"
    target.write_bytes(message.as_bytes())
    return target


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Salesforce asset-return Outlook drafts from Excel.")
    parser.add_argument("input", type=Path, help="Copied Salesforce workbook (.xlsx)")
    parser.add_argument("--output", type=Path, default=Path("draft_output"), help="Output folder for previews and manifest")
    parser.add_argument("--banner", type=Path, help="Optional EcoReco banner image to embed in drafts")
    parser.add_argument("--create-outlook-drafts", action="store_true", help="Save drafts in the logged-in classic Outlook profile")
    parser.add_argument("--create-eml", action="store_true", help="Write unsent .eml files for manual review")
    args = parser.parse_args()

    if args.banner is None:
        default_banner = Path(__file__).with_name("signature-banner.png")
        if default_banner.exists():
            args.banner = default_banner

    args.output.mkdir(parents=True, exist_ok=True)
    cases, skipped = read_cases(args.input)
    manifest = []
    for case in cases:
        if case["status"] == "Needs Review":
            manifest.append({**case, "draft_id": "", "preview": "", "action": "Review required"})
            continue
        body, attachments = build_html(case, args.banner)
        preview = args.output / f"{case['case_id']}.html"
        preview.write_text(body, encoding="utf-8")
        draft_id = create_outlook_draft(case, body, attachments) if args.create_outlook_drafts else ""
        eml_path = create_eml(case, body, args.output / "eml", attachments) if args.create_eml else None
        action = "Outlook draft created" if draft_id else ("EML created" if eml_path else "Preview only")
        manifest.append({**case, "draft_id": draft_id, "eml": str(eml_path) if eml_path else "", "preview": str(preview), "action": action})

    (args.output / "manifest.json").write_text(json.dumps({"source": str(args.input), "cases": manifest, "skipped_rows": skipped}, indent=2, default=str), encoding="utf-8")
    with (args.output / "review.csv").open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=["case_id", "name", "email", "asset_count", "status", "missing_fields", "action"])
        writer.writeheader()
        for item in manifest:
            writer.writerow({"case_id": item.get("case_id", ""), "name": item.get("name", ""), "email": item.get("email", ""), "asset_count": len(item.get("assets", [])), "status": item.get("status", ""), "missing_fields": "; ".join(item.get("missing_fields", [])), "action": item.get("action", "")})
    ready = sum(item.get("status") == "Ready for Draft" for item in manifest)
    review = sum(item.get("status") == "Needs Review" for item in manifest)
    print(f"Cases found: {len(manifest)} | ready: {ready} | needs review: {review} | skipped rows: {len(skipped)}")
    print(f"Output: {args.output.resolve()}")
    if not args.create_outlook_drafts:
        print("No Outlook drafts were created.")


if __name__ == "__main__":
    main()
