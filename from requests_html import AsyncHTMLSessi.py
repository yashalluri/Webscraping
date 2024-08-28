import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

def scrape_with_selenium_and_email(recipient_email):
    url = 'https://www.legis.ga.gov/schedule/all'
    
    # Setup Chrome options
    options = Options()
    options.headless = True
    driver = webdriver.Chrome(options=options)

    # Navigate to the page
    driver.get(url)

    try:
        # Wait until the table-responsive elements are present
        WebDriverWait(driver, 20).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "table-responsive"))
        )
    except TimeoutException:
        print("Timed out waiting for page elements to load")
        driver.quit()
        return

    # Use BeautifulSoup to parse the HTML
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    tables = soup.find_all('div', class_='table-responsive')

    # Prepare to collect data
    data = []

    # Extract data from each 'td' within the found 'div' elements
    for table in tables:
        rows = table.find_all('tr')
        for row in rows:
            cols_data = []
            cols = row.find_all('td')
            for col in cols:
                text = col.text.strip()
                link = col.find('a')
                href = link['href'] if link else None
                cols_data.append((text, href))  # Tuple of text and href
            if cols_data:
                data.append(cols_data)

    driver.quit()

    # Convert list to DataFrame
    # Flatten the data and separate text and href into different columns
    flat_data = []
    for row in data:
        flat_row = []
        for text, href in row:
            flat_row.extend([text, href])
        flat_data.append(flat_row)
    df = pd.DataFrame(flat_data)
    # Define column headers if necessary
    # df.columns = ['Column1', 'Link1', 'Column2', 'Link2', ...]


    # Save to Excel
    excel_file = 'scraped_data.xlsx'
    df.to_excel(excel_file, index=False)
    print("Data scraped and saved to Excel.")

    # Email the file
    email_data(recipient_email, excel_file)

def email_data(recipient_email, attachment_file):
    # Email settings
    sender_email = "yashwanthalluri26@gmail.com"
    sender_password = "vcez eamv kkfl gomu"
    
    # Create the MIME message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = "Information of the Meeting held by the Georgia Assembly"
    body = "Attached Excel file contains the information and times of the meetings. For more information please visit: https://www.legis.ga.gov/schedule/all."
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach the file
    with open(attachment_file, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f"attachment; filename= {os.path.basename(attachment_file)}",
        )
        msg.attach(part)
    
    # Connect and send email
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_email, sender_password)
    text = msg.as_string()
    server.sendmail(sender_email, recipient_email, text)
    server.quit()
    print("Email sent!")

# Call the function with the email of the recipient
scrape_with_selenium_and_email("mparkerson1@gmail.com")
