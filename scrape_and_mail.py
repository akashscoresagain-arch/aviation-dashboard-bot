import os
import time
import pandas as pd
import smtplib
from email.message import EmailMessage

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ======================
# SCRAPE DASHBOARD
# ======================

print("Opening dashboard...")

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=chrome_options
)

driver.get("https://www.civilaviation.gov.in/")

wait = WebDriverWait(driver, 30)

# wait for page load
wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

# extra wait for JS dashboard render
time.sleep(12)

cards = driver.find_elements(By.CLASS_NAME, "card")

print(f"Cards found: {len(cards)}")

data = []

for card in cards:
    text = card.text.strip()
    if text:
        lines = text.split("\n")
        if len(lines) >= 2:
            metric = lines[0]
            value = lines[1]
            data.append([metric, value])

driver.quit()


# ======================
# CREATE EXCEL
# ======================

df = pd.DataFrame(data, columns=["Metric", "Value"])

file_name = "dashboard.xlsx"
df.to_excel(file_name, index=False)

print("Excel created")


# ======================
# EMAIL FILE
# ======================

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

print("Email sent successfully ✅")
