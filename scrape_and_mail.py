import os
import time
import pandas as pd
import smtplib
from email.message import EmailMessage

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager


print("Opening dashboard...")

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

driver.get("https://www.civilaviation.gov.in/")

# IMPORTANT — wait for full JS render
time.sleep(25)

print("Reading dashboard text...")

body_text = driver.find_element("tag name", "body").text

driver.quit()

lines = [line.strip() for line in body_text.split("\n") if line.strip()]

print(f"Lines captured: {len(lines)}")

# crude metric/value pairing
data = []

for i in range(0, len(lines)-1, 2):
    metric = lines[i]
    value = lines[i+1]
    data.append([metric, value])

df = pd.DataFrame(data, columns=["Metric", "Value"])

file_name = "dashboard.xlsx"
df.to_excel(file_name, index=False)

print("Excel created")


# ================= EMAIL =================

EMAIL = os.environ.get("EMAIL_USER")
PASSWORD = os.environ.get("EMAIL_PASS")

msg = EmailMessage()
msg["Subject"] = "Daily Aviation Dashboard"
msg["From"] = EMAIL
msg["To"] = EMAIL
msg.set_content("Dashboard attached")

with open(file_name, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="xlsx",
        filename=file_name
    )

print("Sending email...")

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASSWORD)
    smtp.send_message(msg)

print("Email sent ✅")
