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

    body = f"""<div style="font-family:Calibri,Arial,sans-serif;font-size:15px;color:#0b1f3a;line-height:1.5;max-width:900px">
<p>Hello,</p>
<p>Greetings from Ecoreco !!</p>
<p><strong>Eco Recycling Ltd, is authorized by Salesforce for the collection of assets from your below mentioned address,</strong></p>
<table style="border-collapse:collapse;width:100%;max-width:860px" border="1" cellpadding="12" cellspacing="0">
<tr><td style="width:115px;text-align:center;vertical-align:middle">Pickup<br>Address</td><td>{address}</td></tr>
<tr><td style="width:115px;text-align:center;vertical-align:middle">Asset Type &amp;<br>Serial<br>Number</td><td>{assets}</td></tr>
</table>
<p>Kindly confirm the pickup address and asset serial no details so that we can initiate the reverse collection process.</p>
<p><strong style="background:#ffff00">If there are any discrepancies, please inform us at your earliest convenience.</strong></p>
<p>We look forward to your response to proceed further.</p>
<p><strong>Thanks &amp; Regards,</strong></p>
<p><strong>CRM Executive</strong> | <a href="tel:+912240052951">+91-22-4005 2951</a>/<a href="tel:+919004149714">+91-9004149714</a> | <a href="https://www.ecoreco.com/">www.ecoreco.com</a><br>
<strong>Eco Recycling Limited</strong><br>
422, The Summit Business Park | <a href="https://www.google.com/maps/search/Andheri+Kurla+road+%7C+Andheri+(East),+Mumbai+400093">Andheri Kurla road | Andheri (East), Mumbai 400093</a></p>
{banner}</div>"""
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Create Salesforce asset-return Outlook drafts from Excel.")
    parser.add_argument("input", type=Path, help="Copied Salesforce workbook (.xlsx)")
    parser.add_argument("--output", type=Path, default=Path("draft_output"), help="Output folder for previews and manifest")
    parser.add_argument("--banner", type=Path, help="Optional EcoReco banner image to embed in drafts")
    parser.add_argument("--create-outlook-drafts", action="store_true", help="Save drafts in the logged-in classic Outlook profile")
    args = parser.parse_args()

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
        manifest.append({**case, "draft_id": draft_id, "preview": str(preview), "action": "Draft created" if draft_id else "Preview only"})

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
        print("Dry run only: no Outlook drafts were created.")


if __name__ == "__main__":
    main()
