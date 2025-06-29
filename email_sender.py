
import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

df = pd.read_excel("recipients.xlsx")

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login('your_email@gmail.com', 'your_password')

for index, row in df.iterrows():
    msg = MIMEMultipart()
    msg['From'] = 'your_email@gmail.com'
    msg['To'] = row['Email']
    msg['Subject'] = "Certificate of Participation"
    body = f"Dear {row['Name']},\n\nThank you for participating. Find your certificate attached."
    msg.attach(MIMEText(body, 'plain'))
    server.send_message(msg)

server.quit()
print("Emails sent.")
