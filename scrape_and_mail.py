import os
import pandas as pd
import smtplib
from email.message import EmailMessage

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ===============================
# SCRAPE CIVIL AVIATION DASHBOARD
# ===============================

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

print("Opening dashboard...")

driver.get("https://www.civilaviation.gov.in/")

# wait for dashboard cards to load
wait = WebDriverWait(driver, 30)
wait.until(EC.presence_of_element_located((By.CLASS_NAME, "card")))

cards = driver.find_elements(By.CLASS_NAME, "card")

print("Cards found:", len(cards))

data = []

for card in cards:
    text = card.text.strip()

    if text:
        lines = text.split("\n")

        if len(lines) >= 2:
            metric = lines[0]
            value = lines[-1]

            data.append([metric, value])
            print(metric, "->", value)

driver.quit()


# ===============================
# CREATE EXCEL
# ===============================

df = pd.DataFrame(data, columns=["Metric", "Value"])

file = "dashboard.xlsx"
df.to_excel(file, index=False)

print("Excel created")


# ===============================
# EMAIL DASHBOARD
# ===============================

EMAIL = os.environ["EMAIL_USER"]
PASS = os.environ["EMAIL_PASS"]

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

print("✅ Dashboard emailed successfully!")
