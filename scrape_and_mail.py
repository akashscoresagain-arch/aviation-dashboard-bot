import pandas as pd
import smtplib
from email.message import EmailMessage

# demo Excel file creation
data = [
    ["Domestic flights", "3217"],
    ["International flights", "632"]
]

df = pd.DataFrame(data, columns=["Metric", "Value"])
file = "dashboard.xlsx"
df.to_excel(file, index=False)

# email setup
msg = EmailMessage()
msg["Subject"] = "Daily Aviation Dashboard"
msg["From"] = "YOUR_EMAIL@gmail.com"
msg["To"] = "YOUR_EMAIL@gmail.com"
msg.set_content("Dashboard attached")

with open(file, "rb") as f:
    msg.add_attachment(f.read(),
                       maintype="application",
                       subtype="xlsx",
                       filename=file)

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login("YOUR_EMAIL@gmail.com", "APP_PASSWORD")
    smtp.send_message(msg)

print("Email sent!")
