import os
import sys
import datetime
import random
import time
import smtplib
import pdfkit
import shutil
import re
import usaddress
from email.message import EmailMessage
from email.utils import formataddr

# === CONFIG ===
path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
pdfkit_config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

base_directory = os.path.dirname(os.path.realpath(__file__))

# === TRACKING ===
success_count = 0
fail_count = 0
attempted = 0

# === HELPERS ===
def smart_format_address(raw_address, html=False):
    if not raw_address.strip():
        return "United States"
    try:
        parsed = usaddress.tag(raw_address)[0]
    except usaddress.RepeatedLabelError:
        return raw_address  # fallback

    street_parts = []
    for key in ['AddressNumber', 'StreetNamePreDirectional', 'StreetName', 'StreetNamePostType', 'OccupancyType', 'OccupancyIdentifier']:
        if key in parsed:
            street_parts.append(parsed[key])
    street = " ".join(street_parts)

    city = parsed.get('PlaceName', '')
    state = parsed.get('StateName', '')
    zip_code = parsed.get('ZipCode', '')

    city_line = f"{city}, {state} {zip_code}".strip()
    sep = "<br>" if html else "\n"
    return f"{street}{sep}{city_line}" if street and city_line else raw_address

def print_inline_progress(email, index, total):
    bar_width = shutil.get_terminal_size((80, 20)).columns - 60
    filled = int(bar_width * index / total)
    percent = (index / total) * 100
    bar = "█" * filled + "-" * (bar_width - filled)
    print(f"\rMail to = {email:<30} [{index}/{total}] {percent:5.1f}% {bar}", end="", flush=True)

def send_mail(sender_name, sender_email, receiver, subject, message_content, cc_email, html, company_name, reply_to, smtp_server, smtp_port, smtp_user, smtp_password, pdf_path=None, static_file_path=None, send_attachment=False, add_cc=False):
    global success_count, fail_count
    try:
        message = EmailMessage()
        if html:
            message.add_alternative(message_content, subtype='html')
        else:
            message.set_content(message_content)

        message['From'] = formataddr((sender_name, sender_email))
        if add_cc:
            message['CC'] = cc_email
        message['Reply-To'] = reply_to
        message['To'] = receiver
        message['Subject'] = subject

        if send_attachment:
            if pdf_path and os.path.exists(pdf_path):
                with open(pdf_path, 'rb') as f:
                    pdf_data = f.read()
                message.add_attachment(pdf_data, maintype='application', subtype='pdf', filename=os.path.basename(pdf_path))

            if static_file_path and os.path.exists(static_file_path):
                with open(static_file_path, 'rb') as f:
                    static_data = f.read()
                message.add_attachment(static_data, maintype='application', subtype='pdf', filename=os.path.basename(static_file_path))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(message)

        with open("emails_send.txt", "a", encoding="utf-8") as f:
            f.write(f"{receiver} | {sender_name} | {company_name} | {datetime.datetime.now().date().strftime('%d-%m-%Y')}\n")

        success_count += 1
        return True

    except smtplib.SMTPException as smtp_error:
        print(f"\n[SMTP ERROR] Sending stopped at line {attempted} ({receiver} {company_name}) due to SMTP error: {smtp_error}\nPlease change SMTP credentials or server and try again.")
        return False
    except Exception as error:
        print(f"\n[X] Failed to send to {receiver} | Error: {error}")
        fail_count += 1
        return True  # continue for non-SMTP errors

# === STATIC CONFIG ===
reply_to = "Clifton Davidson <clifton@grabenllc.com>"
sender_name = "Clifton Davidson"
subjects = ["INV22748 – Revised INV for {company_name} (Q1 Advisory Services)"]

smtp_server = "easypacklogistics.com"
smtp_port = 587
smtp_user = sender_email = "alex@easypacklogistics.com"
smtp_password = "easypack2021"

box = "grabenllc.com"
pdf_filename = "Inv_22748 Q1 2025.pdf"
static_file_path = os.path.join(base_directory, "W9.pdf")
send_attachment = True
add_cc = True

# === LOAD TEMPLATES ===
messages = []
for file in os.listdir(os.path.join(base_directory, "messages")):
    if file.endswith(".txt"):
        file_path = os.path.join(base_directory, "messages", file)
        with open(file_path, encoding="utf-8") as f:
            messages.append(f.read().strip())

html_template_path = os.path.join(base_directory, "template.html")
if os.path.exists(html_template_path):
    with open(html_template_path, encoding="utf-8") as f:
        html_template = f.read()
else:
    html_template = None

# === INPUT ===
input_filename = sys.argv[1] if len(sys.argv) > 1 else "input.txt"
print(f"Using input file: {input_filename}")
lines = open(input_filename, encoding="latin-1").readlines()
total_recipients = len(lines)

for i, x in enumerate(lines):
    attempted += 1
    x = x.strip()
    parts = x.split(" | ")
    email, name, company_name, domain, username, ceo = parts[:6]
    address_raw = parts[6] if len(parts) >= 7 else "United States"
    address_text = smart_format_address(address_raw, html=False)
    address_html = smart_format_address(address_raw, html=True)

    text = random.choice(messages)
    text = text.replace("{email}", email)
    text = text.replace("{name}", name)
    text = text.replace("{company_name}", company_name)
    text = text.replace("{domain}", domain)
    text = text.replace("{username}", username)
    text = text.replace("{ceo}", ceo)
    text = text.replace("{box}", box)
    text = text.replace("{address}", address_text)

    html = text.startswith("<")
    subject = random.choice(subjects).replace("{company_name}", company_name).replace("{address}", address_text)
    cc_email = f"Accounting <accounting@{box}>"

    if html_template:
        personalized_html = html_template.replace("{name}", name)
        personalized_html = personalized_html.replace("{email}", email)
        personalized_html = personalized_html.replace("{company_name}", company_name)
        personalized_html = personalized_html.replace("{domain}", domain)
        personalized_html = personalized_html.replace("{username}", username)
        personalized_html = personalized_html.replace("{ceo}", ceo)
        personalized_html = personalized_html.replace("{address}", address_html)

        pdf_path = os.path.join(base_directory, pdf_filename)
        pdfkit.from_string(personalized_html, pdf_path, configuration=pdfkit_config)
    else:
        pdf_path = None

    print_inline_progress(email, i + 1, total_recipients)
    continue_sending = send_mail(sender_name, sender_email, email, subject, text, cc_email, html, company_name, reply_to, smtp_server, smtp_port, smtp_user, smtp_password, pdf_path, static_file_path, send_attachment, add_cc)
    if not continue_sending:
        break

    time.sleep(10)

    if pdf_path and os.path.exists(pdf_path):
        os.remove(pdf_path)

print("\n=== Summary ===")
print(f"Total attempted: {attempted}")
print(f"Successfully sent: {success_count}")
print(f"Failed to send: {fail_count}")
