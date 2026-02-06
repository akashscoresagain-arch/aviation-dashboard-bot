import time
import pandas as pd
import smtplib
from email.message import EmailMessage
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# ---------------------------
# SCRAPE DASHBOARD
# ---------------------------

options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)

driver.get("https://www.civilaviation.gov.in/")
time.sleep(10)

data = []

cards = driver.find_elements(By.CLASS_NAME, "card")

for card in cards:
    text = card.text.strip()
    if text:
        lines = text.split("\n")

        if len(lines) >= 2:
            metric = lines[0]
            value = lines[-1]
            data.append([metric, value])

driver.quit()

# ---------------------------
# CREATE EXCEL
# ---------------------------

df = pd.DataFrame(data, columns=["Metric", "Value"])
file = "dashboard.xlsx"
df.to_excel(file, index=False)

# ---------------------------
# EMAIL
# ---------------------------

EMAIL = EMAIL_USER
PASS = EMAIL_PASS

msg = EmailMessage()
msg["Subject"] = "Daily Aviation Dashboard"
msg["From"] = EMAIL
msg["To"] = EMAIL
msg.set_content("Dashboard attached")

with open(file, "rb") as f:
    msg.add_attachment(
        f.read(),
        maintype="application",
        subtype="xlsx",
        filename=file
    )

with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL, PASS)
    smtp.send_message(msg)

print("✅ Dashboard emailed!")
