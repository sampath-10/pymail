#!/usr/bin/env python
import openpyxl
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests  # Import the requests library to fetch the GitHub file
from io import BytesIO  # Import BytesIO to work with file content in memory
github_raw_url = 'https://github.com/sampath-10/pymail/blob/379ecb03a2a56cbd964f90c7c85d54de78ab034a/Book12.xlsx'
response = requests.get(github_raw_url)
file_content = BytesIO(response.content)
workbook = openpyxl.load_workbook(file_content)
sheet = workbook['Sheet1']
today = datetime.today().strftime('%m-%d')
from_email = 'trailidsam@gmail.com'
password = 'sufapdhwpmytxyla'
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(from_email, password)
for row in sheet.iter_rows(values_only=True):
    name, dob_str, email = row
    if today == dob_str:
        subject = 'Happy Birthday!'
        message = f"Dear {name},\n\nHappy Birthday! 🎉🎂\n\nBest wishes, Your Name"
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'plain'))
        server.sendmail(from_email, email, msg.as_string())
        print(f"Birthday email sent to {name} ({email})")
        for row2 in sheet.iter_rows(values_only=True):
            name2, _, email2 = row2
            if email2 and email2 != email:
                subject2 = f"Today is {name}'s birthday!"
                message2 = f"Hi {name2},\n\nJust a reminder that today is {name}'s birthday. Don't forget to send your warm wishes!"                
                msg2 = MIMEMultipart()
                msg2['From'] = from_email
                msg2['To'] = email2
                msg2['Subject'] = subject2
                msg2.attach(MIMEText(message2, 'plain'))
                server.sendmail(from_email, email2, msg2.as_string())
                print(f"Reminder email sent to {name2} ({email2})")
server.quit()
workbook.close()
