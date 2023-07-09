import smtplib
import pandas as pd
from email.mime.text import MIMEText
import argparse

# Create the argument parser
parser = argparse.ArgumentParser()

# Add cli arguments
parser.add_argument("--email", help="Email address")
parser.add_argument("--password", help="Email password")
parser.add_argument("--file", help="Path to excel file")
parser.add_argument("--sheet", help="Sheet name")
parser.add_argument("--subject", help="Subject of the email")
parser.add_argument("--body", help="Body of the email")


def enviar_email(receiver, subject, body):
    message = MIMEText(body)
    message['Subject'] = subject
    message['From'] = email_sender
    message['To'] = receiver

    smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_server.starttls()
    smtp_server.login(email_sender, password)
    smtp_server.send_message(message)
    smtp_server.quit()


def read_excel(file_path, sheet_name, subject, body):
    dataframe = pd.read_excel(file_path, sheet_name, index_col=None, na_values=["NA"])
    for index, row in dataframe.iterrows():
        # Acessar os valores das colunas para cada linha
        date = row['data']
        email = row['email']
        enviar_email(email, subject, body + " " + str(date))


if __name__ == "__main__":
    args = parser.parse_args()
    global email_sender
    global password
    email_sender = args.email
    password = args.password

    file = args.file
    sheet = args.sheet
    subject = args.subject
    body = args.body

    read_excel(file, sheet, subject, body)



