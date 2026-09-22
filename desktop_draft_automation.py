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
    # Salesforce exports may separate recipients with semicolons, commas, or
    # line breaks. A period is never a separator because it belongs in domains.
    addresses = recipients(value)
    return bool(addresses) and all(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", a) for a in addresses)


def recipients(value: str) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    # Salesforce exports may separate multiple recipients with commas,
    # semicolons, spaces, or line breaks. Email addresses themselves cannot
    # contain whitespace, so a space is safe to treat as a separator here.
    for part in re.split(r"[;,\s]+", text(value)):
        address = part.strip()
        key = address.casefold()
        if address and key not in seen:
            result.append(address)
            seen.add(key)
    return result


def address_case(value: str) -> str:
    def format_word(match: re.Match[str]) -> str:
        word = match.group(0)
        if word.isupper() and len(word) <= 4:
            return word
        if len(word) > 1 and word[0].isupper() and word[1:].islower():
            return word
        return word[:1].upper() + word[1:].lower()

    return re.sub(r"[A-Za-z]+", format_word, text(value))


def address_tokens(value: str) -> list[str]:
    return re.findall(r"[^\W_]+", text(value).casefold(), flags=re.UNICODE)


def contains_address_field(existing: str, value: str, postal: bool = False) -> bool:
    tokens = address_tokens(existing)
    wanted = address_tokens(value)
    if not wanted:
        return True
    # Match complete tokens, never substrings such as York inside Yorkshire.
    if any(tokens[i:i + len(wanted)] == wanted for i in range(len(tokens))):
        return True
    if postal:
        # Postal codes may contain spaces/hyphens: 411 045 or SW1A-1AA.
        code = "".join(wanted)
        return any(
            "".join(tokens[i:j]) == code
            for i in range(len(tokens))
            for j in range(i + 1, min(len(tokens), i + len(code)) + 1)
        )
    return False


def address_lines(case: dict[str, Any]) -> list[str]:
    lines = [address_case(case.get(key)) for key in ("address1", "address2") if text(case.get(key))]
    extra = []
    for key in ("city", "state", "country", "postal"):
        value = text(case.get(key))
        if value and not contains_address_field(" ".join(lines + extra), value, postal=key == "postal"):
            extra.append(address_case(value))
    if extra:
        lines.append(", ".join(extra))
    return lines


def address_value(case: dict[str, Any]) -> str:
    # Keep one flowing address so Outlook wraps only when the cell is full.
    # Normalize punctuation at the same time so source cells with line breaks,
    # spaces before commas, or repeated commas render consistently.
    parts = []
    for line in address_lines(case):
        value = re.sub(r"[\r\n]+", ", ", text(line))
        value = re.sub(r"\s+", " ", value)
        value = re.sub(r"\s*,\s*", ", ", value)
        value = re.sub(r",\s*,+", ", ", value)
        value = value.strip(" ,")
        if value:
            parts.append(value)
    return ", ".join(parts)


def postal_conflict(case: dict[str, Any]) -> bool:
    # Only clearly labelled Indian PINs are checked; an arbitrary number may
    # be a building number or phone number and must not be guessed as a PIN.
    if text(case.get("country")).casefold() != "india":
        return False
    expected = "".join(address_tokens(case.get("postal", "")))
    if not re.fullmatch(r"[1-9][0-9]{5}", expected):
        return False
    address = " ".join(text(case.get(key)) for key in ("address1", "address2"))
    found = re.findall(r"\bpin(?:\s*code)?\s*[:.\-]?\s*([1-9][0-9]{2}[ -]?[0-9]{3})\b", address, re.I)
    return any(re.sub(r"\D", "", value) != expected for value in found)


def build_html(case: dict[str, Any], banner_path: Path | None = None) -> tuple[str, dict[str, str]]:
    esc = lambda value: html.escape(text(value), quote=True)
    greeting = f"Hello {esc(case.get('name'))}," if text(case.get('name')) else "Hello,"
    address_parts = address_lines(case)
    address = esc(address_value(case))
    # Keep each asset on one visual line. A single long asset must not be split
    # across lines, while <br> still gives each asset its own line.
    assets = "<br>".join(
        f'<span style="white-space:nowrap;word-break:keep-all">{esc(asset)}</span>'
        for asset in case["assets"]
    )
    cid = ""
    banner_image = ""
    attachments: dict[str, str] = {}
    if banner_path and banner_path.exists():
        cid = "ecoreco-banner@draft-automation"
        banner_image = f'<img border="0" width="655" height="106" src="cid:{cid}" alt="EcoReco">'
        attachments[cid] = str(banner_path)

    body = f"""<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=us-ascii"><style>
p.MsoNormal,li.MsoNormal,div.MsoNormal {{ margin:0cm; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri",sans-serif; }}
p.wordsection1,li.wordsection1,div.wordsection1 {{ margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif; }}
p.MsoNormal {{ margin:0cm; font-size:11.0pt; font-family:"Calibri",sans-serif; line-height:1.15; }}
body, table, td, p, div, span, b, strong, a {{ font-size:11.0pt; font-family:"Calibri",sans-serif; }}
td {{ vertical-align:middle; }}
a:link {{ color:#0563C1; text-decoration:underline; }}
</style></head><body lang=EN-US link="#0563C1" vlink="#954F72" style='font-size:11.0pt;font-family:"Calibri",sans-serif'><div class=WordSection1 style='font-size:11.0pt;font-family:"Calibri",sans-serif'>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>{greeting}<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Greetings from Ecoreco !!<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin-bottom:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Eco Recycling Ltd, is authorized by Salesforce for the collection of assets from your below mentioned address, <o:p></o:p></span></b></p>
<p class=wordsection1 style='margin-bottom:0cm;margin-bottom:.0001pt'><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></b></p>
<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="540" style="width:405pt;border-collapse:collapse;table-layout:fixed;mso-table-lspace:0pt;mso-table-rspace:0pt">
<tr><td width="200" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;word-wrap:break-word;overflow-wrap:break-word;width:150pt"><p class="MsoNormal" align="center" style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;text-align:center"><b>Pickup Address</b></p></td><td width="340" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;word-wrap:break-word;overflow-wrap:break-word;width:255pt"><p class="MsoNormal" align="center" style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;text-align:center">{address}</p></td></tr>
<tr><td width="200" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;word-wrap:break-word;overflow-wrap:break-word;width:150pt"><p class="MsoNormal" align="center" style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;text-align:center"><b>Asset Type &amp;<br>Serial Number</b></p></td><td width="340" valign="middle" style="border:1px solid #000;padding:4pt 8pt;vertical-align:middle;font-family:Calibri,Arial,sans-serif;font-size:11pt;word-wrap:normal;overflow-wrap:normal;white-space:nowrap;width:255pt"><p class="MsoNormal" align="center" style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:13pt;mso-line-height-rule:exactly;text-align:center;white-space:nowrap">{assets}</p></td></tr>
</table>
<p class="MsoNormal" style="margin:0;font-size:11pt;line-height:10pt"><o:p>&nbsp;</o:p></p>
<p class=wordsection1><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>Kindly confirm the pickup address and asset serial no details so that we can initiate the reverse collection process.<br><br><b><span style='background:yellow;mso-highlight:yellow'>If there are any discrepancies, please inform us at your earliest convenience.</span></b><br><br><o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'>We look forward to your response to proceed further.<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:12.0pt'><b><span lang=EN-IN style='color:#0F243E'>Thanks &amp; Regards,</span></b><b><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:#0F243E'><o:p></o:p></span></b></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:12.0pt'><b><span lang=EN-IN style='color:#0F243E'>CRM Executive</span></b><span lang=EN-IN style='color:#0F243E'> | +91-22-4005 2951/+91-9004149714 | </span><a href="http://www.ecoreco.com/"><span style='color:#0F243E'>www.ecoreco.com</span></a><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif;color:#0F243E'><br>Eco Recycling Limited | 422, The Summit Business Park | Andheri Kurla road | Andheri (East), Mumbai 400093<o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:12.0pt'><b><span lang=EN-IN style='color:#0F243E'>{banner_image}</span></b><span lang=EN-IN style='color:#0F243E'><o:p></o:p></span></p>
<p class=wordsection1 style='margin:0cm;margin-bottom:.0001pt'><span lang=EN-IN style='font-size:11.0pt;font-family:"Calibri",sans-serif'><o:p>&nbsp;</o:p></span></p></div></body></html>"""
    return body, attachments


def read_cases(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet = workbook[workbook.sheetnames[0]]
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    if not rows:
        raise ValueError("The workbook is empty.")
    headers = [text(value) for value in rows[0]]
    required = {'Case ID', "User's Name", "User's Email", 'Line Address 1', 'City', 'Country', 'Asset Serial Number'}
    if required - set(headers):
        raise ValueError('Missing columns: ' + ', '.join(sorted(required - set(headers))))
    grouped: OrderedDict[str, list[dict[str, Any]]] = OrderedDict()
    skipped: list[dict[str, Any]] = []
    for row_number, values in enumerate(rows[1:], start=2):
        row = {headers[i]: values[i] if i < len(values) else None for i in range(len(headers)) if headers[i]}
        case_id = first(row, "Case ID")
        if not case_id:
            skipped.append({"row": row_number, "reason": "Missing Case ID"})
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
        if postal_conflict(case):
            missing.append("Conflicting PIN code in address and postal column")
        if not re.fullmatch(r'[A-Za-z0-9_-]+', case_id): missing.append('Invalid Case ID')
        for field in ["User's Name", "User's Email", 'Line Address 1', 'Line Address 2', 'City', 'State/Province', 'Country', 'Zipcode/Postal Code']:
            if len({first(r, field).casefold() for r in case_rows}) > 1:
                missing.append('Conflicting ' + field)
        if any(is_yes(first(r, 'Duplicate')) for r in case_rows):
            missing.append('Duplicate flag: confirm case before drafting')
        if any(not first(r, 'Asset Serial Number') for r in case_rows):
            missing.append('Asset missing on a case row')
        if any('#NAME?' in first(r, 'Asset Serial Number') or '#REF!' in first(r, 'Asset Serial Number') for r in case_rows):
            missing.append('Asset contains Excel error')
        if not case["name"]: missing.append("User's Name")
        if not valid_email(case["email"]): missing.append("User's Email")
        if any(c in case['name'] for c in '\r\n'): missing.append('Invalid name')
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Salesforce asset-return Outlook drafts from Excel.")
    parser.add_argument("input", type=Path, help="Copied Salesforce workbook (.xlsx)")
    parser.add_argument("--output", type=Path, default=Path("draft_output"), help="Output folder for previews and manifest")
    parser.add_argument("--banner", type=Path, help="Optional EcoReco banner image to embed in drafts")
    parser.add_argument("--create-outlook-drafts", action="store_true", help="Save drafts in the logged-in classic Outlook profile")
    parser.add_argument("--create-html-preview", action="store_true", help="Also write browser-viewable HTML previews")
    args = parser.parse_args()
    if args.create_outlook_drafts:
        parser.error('Use Prepare Drafts.bat for Outlook drafts with mailbox selection and duplicate protection.')

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
        preview = None
        if args.create_html_preview:
            preview = args.output / f"{case['case_id']}.html"
            preview.write_text(body, encoding="utf-8")
        draft_id = create_outlook_draft(case, body, attachments) if args.create_outlook_drafts else ""
        action = "Outlook draft created" if draft_id else "Preview only"
        manifest.append({**case, "draft_id": draft_id, "preview": str(preview) if preview else "", "action": action})

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
