import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

username = "youruseremail@domain.com"
mail_from = "youruseremail@domain.com"
mail_to = "mailto@domain.com"
mail_subject = "Custom email from Python"
mail_body = """
This is a message sent by Python

Cheers.
"""

password = input("Type your password and press enter: ")

mimemsg = MIMEMultipart()
mimemsg['From'] = mail_from
mimemsg['To'] = mail_to
mimemsg['Subject'] = mail_subject
mimemsg.attach(MIMEText(mail_body, 'plain'))
connection = smtplib.SMTP(host='smtp.office365.com', port=587)
connection.starttls()
connection.login(username, password)
connection.send_message(mimemsg)
connection.quit()

