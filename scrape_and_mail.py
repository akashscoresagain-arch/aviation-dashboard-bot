import os
import pandas as pd
import smtplib
from email.message import EmailMessage

# demo Excel data
data = [
    ["Domestic flights", "3217"],
    ["International flights", "632"]
]

df = pd.DataFrame(data, columns=["Metric", "Value"])
file = "dashboard.xlsx"
df.to_excel(file, index=False)

# read secrets
EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASS = os.environ["EMAIL_PASS"]

# email setup
msg = EmailMessage()
msg["Subject"] = "Daily Aviation Dashboard"
msg["From"] = EMAIL_USER
msg["To"] = EMAIL_USER
msg.set_content("Dashboard attached")

with open(file, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="xlsx",
        filename=file
    )

# send email
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL_USER, EMAIL_PASS)
    smtp.send_message(msg)

print("✅ Email sent!")
