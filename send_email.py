import html
import os
import secrets
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import gspread
from gspread.utils import rowcol_to_a1


def required_setting(name):
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Set the {name} environment variable before running this script.")
    return value


def get_pending_recipients():
    credentials_file = required_setting("GOOGLE_APPLICATION_CREDENTIALS")
    sheet_id = required_setting("GOOGLE_SHEET_ID")
    worksheet_name = os.environ.get("GOOGLE_WORKSHEET_NAME", "").strip()
    spreadsheet = gspread.service_account(filename=credentials_file).open_by_key(sheet_id)
    worksheet = (
        spreadsheet.worksheet(worksheet_name)
        if worksheet_name
        else spreadsheet.sheet1
    )

    values = worksheet.get_all_values()
    if not values:
        raise RuntimeError("The configured Google Sheet is empty.")

    headers = [value.strip().casefold() for value in values[0]]

    def find_column(names):
        for name in names:
            if name in headers:
                return headers.index(name) + 1
        return None

    email_column = find_column(("email", "email id", "email address"))
    status_column = find_column(("status",))
    if email_column is None or status_column is None:
        raise RuntimeError(
            "The worksheet must have an Email (or Email ID) column and a Status column."
        )

    token_column = find_column(("tracking token",))
    if token_column is None:
        token_column = len(values[0]) + 1
        worksheet.update_cell(1, token_column, "Tracking Token")

    recipients = []
    tokens_seen = set()
    token_index = token_column - 1
    for row_number, row in enumerate(values[1:], start=2):
        token = row[token_index].strip() if token_index < len(row) else ""
        if token and token in tokens_seen:
            raise RuntimeError(f"Duplicate tracking token found in worksheet row {row_number}.")
        if token:
            tokens_seen.add(token)

    token_updates = []
    for row_number, row in enumerate(values[1:], start=2):
        email_index = email_column - 1
        status_index = status_column - 1
        recipient = row[email_index].strip() if email_index < len(row) else ""
        status = row[status_index].strip() if status_index < len(row) else ""
        token = row[token_index].strip() if token_index < len(row) else ""

        if status.casefold() != "pending" or "@" not in recipient:
            continue

        if not token:
            while not token or token in tokens_seen:
                token = secrets.token_urlsafe(32)
            tokens_seen.add(token)
            token_updates.append(
                {
                    "range": rowcol_to_a1(row_number, token_column),
                    "values": [[token]],
                }
            )

        tokens_seen.add(token)
        recipients.append((recipient, token))

    if token_updates:
        worksheet.batch_update(token_updates)
    return recipients


def make_tracking_url(base_url, token):
    parts = urlsplit(base_url)
    if parts.scheme not in ("http", "https") or not parts.netloc:
        raise RuntimeError("BASE_URL must be a fully qualified http:// or https:// URL.")
    query = urlencode(
        [(key, value) for key, value in parse_qsl(parts.query) if key != "token"]
        + [("token", token)]
    )
    return urlunsplit((parts.scheme, parts.netloc, parts.path, query, parts.fragment))


def run_campaign():
    sender_email = required_setting("SENDER_EMAIL")
    sender_password = required_setting("SENDER_PASSWORD")
    base_url = required_setting("BASE_URL")

    recipients = get_pending_recipients()
    if not recipients:
        print("No employees with Pending status and a valid email address were found.")
        return

    print(f"Prepared {len(recipients)} tracking link(s) for Pending employees:")
    tracking_links = []
    for recipient, token in recipients:
        tracking_link = make_tracking_url(base_url, token)
        tracking_links.append((recipient, token, tracking_link))
        print(f"{recipient}: {tracking_link}")

    print("\nConnecting to Gmail SMTP...")
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        for recipient, _token, tracking_link in tracking_links:
            escaped_url = html.escape(tracking_link, quote=True)
            message = MIMEMultipart("alternative")
            message["From"] = sender_email
            message["To"] = recipient
            message["Subject"] = "Mandatory IT Security Verification"
            message.attach(
                MIMEText(
                    f"""
                    <p>Dear User,</p>
                    <p>Please click the link below to perform mandatory verification for your corporate profile:</p>
                    <p><a href="{escaped_url}" style="background-color: #0d6efd; color: white; padding: 10px 15px; text-decoration: none; border-radius: 5px;">Verify System Profile</a></p>
                    <p>Regards,<br>IT Infrastructure Team</p>
                    """,
                    "html",
                )
            )
            server.sendmail(sender_email, recipient, message.as_string())
            print(f"Email sent to {recipient}.")

    print("Campaign dispatch completed.")


if __name__ == "__main__":
    run_campaign()
