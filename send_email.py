import os
import smtplib
import pandas as pd
from email.message import EmailMessage
from dotenv import load_dotenv

# Load email credentials
load_dotenv()

EMAIL_HOST = os.getenv("EMAIL_HOST")
EMAIL_PORT = int(os.getenv("EMAIL_PORT"))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

# Paths
DATA_PATH = "../data/cleaned_master.csv"
REPORTS_DIR = "reports"

# Load cleaned data
df = pd.read_csv(DATA_PATH)
df.columns = df.columns.str.strip().str.lower()

# Ensure expected columns exist
if "company_name" not in df.columns or "email_address" not in df.columns:
    raise ValueError("❌ Missing 'company_name' or 'email_address' columns in cleaned_master.csv")

# Build list of (company, email) pairs
targets = df[["company_name", "email_address"]].dropna().drop_duplicates()

# Send email with attachment
def send_email(recipient, subject, body, attachment_path):
    msg = EmailMessage()
    msg["From"] = EMAIL_USER
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    # Attach the PDF
    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="pdf",
            filename=os.path.basename(attachment_path)
        )

    with smtplib.SMTP(EMAIL_HOST, EMAIL_PORT) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

# Loop through companies
for _, row in targets.iterrows():
    company = row["company_name"]
    email = row["email_address"]
    filename = "".join(c if c.isalnum() else "_" for c in company) + ".pdf"
    report_path = os.path.join(REPORTS_DIR, filename)

    if not os.path.exists(report_path):
        print(f"❌ Report not found for {company} — skipping.")
        continue

    try:
        print(f"📤 Sending report for {company} to {email}")
        send_email(
            recipient=email,
            subject=f"Resilience Report – {company}",
            body=f"Dear team,\n\nAttached is your resilience scan report for {company}.\n\nBest regards,\nResilience Team",
            attachment_path=report_path
        )
        print(f"✅ Sent to {email}")
    except Exception as e:
        print(f"❌ Failed to send to {email}: {e}")
