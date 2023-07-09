import tkinter as tk
from tkinter import messagebox
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


def enviar_email(email_receiver, date):
    email_sender = entry_email.get()
    password = entry_password.get()
    subject = entry_subject.get()
    body = entry_body.get()

    message = MIMEText(body + str(date))
    message['Subject'] = subject
    message['From'] = email_sender
    message['To'] = email_receiver

    try:
        smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
        smtp_server.starttls()
        smtp_server.login(email_sender, password)
        smtp_server.send_message(message)
        smtp_server.quit()
    except Exception as e:
        messagebox.showerror("Error", str(e))


def read_excel():
    file_path = entry_file.get()
    sheet_name = entry_sheet.get()

    try:
        dataframe = pd.read_excel(file_path, sheet_name, index_col=None, na_values=["NA"])
        for index, row in dataframe.iterrows():
            # Acessar os valores das colunas para cada linha
            date = row['data']
            email = row['email']
            enviar_email(email, date)
    except Exception as e:
        messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    # Criar a janela
    window = tk.Tk()
    window.title('Enviar Email')
    window.geometry("400x300")  # Define o tamanho da janela

    # Campo de entrada de email
    label_email = tk.Label(window, text="Email address:")
    label_email.pack()
    entry_email = tk.Entry(window)
    entry_email.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Campo de entrada de password
    label_password = tk.Label(window, text="Password:")
    label_password.pack()
    entry_password = tk.Entry(window, show="*")
    entry_password.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Campo de entrada de file
    label_file = tk.Label(window, text="Path to excel file:")
    label_file.pack()
    entry_file = tk.Entry(window)
    entry_file.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Campo de entrada de sheet
    label_sheet = tk.Label(window, text="Sheet name:")
    label_sheet.pack()
    entry_sheet = tk.Entry(window)
    entry_sheet.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Campo de entrada de subject
    label_subject = tk.Label(window, text="Subject:")
    label_subject.pack()
    entry_subject = tk.Entry(window)
    entry_subject.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Campo de entrada de body
    label_body = tk.Label(window, text="Body:")
    label_body.pack()
    entry_body = tk.Entry(window)
    entry_body.pack(fill=tk.X)  # Aumenta o tamanho horizontalmente

    # Botão de enviar
    btn_enviar = tk.Button(window, text='Enviar', command=read_excel)
    btn_enviar.pack()

    # Iniciar o loop da janela
    window.mainloop()
