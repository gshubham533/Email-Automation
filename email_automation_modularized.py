import csv
import datetime
from email.message import EmailMessage
import json
import ssl
import smtplib
import re


class EmailSender:
    def __init__(self, credentials_file, email_list_file, subject_file, body_file, sent_log_file):
        self.credentials_file = credentials_file
        self.email_list_file = email_list_file
        self.subject_file = subject_file
        self.body_file = body_file
        self.sent_log_file = sent_log_file
        self.smtp = None

    def load_credentials(self):
        with open(self.credentials_file) as file:
            creds = json.load(file)
        self.email_sender = creds["email"]
        self.password = creds["password"]
        self.smtp_url = creds.get("smtp", "smtp.office365.com")

    def login(self):
        context = ssl.create_default_context()
        try:
            self.smtp = smtplib.SMTP(self.smtp_url, 587)
            self.smtp.starttls(context=context)
            self.smtp.login(self.email_sender, self.password)
        except smtplib.SMTPAuthenticationError:
            print("Invalid credentials")
            return False
        except Exception as e:
            print("Something went wrong: " + str(e))
            return False
        return True

    def send_email(self, email_subject, email_body, to_email, is_html=False):
        msg = EmailMessage()
        msg["From"] = self.email_sender
        msg["Subject"] = email_subject
        msg["To"] = to_email
        if is_html:
            msg.add_alternative(email_body, subtype='html')
        else:
            msg.set_content(email_body)
        try:
            self.smtp.sendmail(self.email_sender, to_email, msg.as_string())
            return True
        except Exception as e:
            print("Failed to send email: " + str(e))
            return False

    def get_email_list(self):
        rows = []
        with open(self.email_list_file, 'r') as file:
            csv_reader = csv.reader(file)
            header = next(csv_reader)
            for row in csv_reader:
                rows.append(row)
        return header, rows

    def log_email_sent(self, to_email, mail_sent, subject, body, date_time):
        status = "Yes" if mail_sent else "No"
        log_entry = [to_email, status, self.email_sender, subject, str(body), str(date_time)]
        with open(self.sent_log_file, 'a', newline='\n') as file:
            writer = csv.writer(file)
            writer.writerow(log_entry)

    def load_email_content(self):
        with open(self.subject_file, 'r') as file:
            subject = file.read()
        with open(self.body_file, 'r') as file:
            body = file.read()
        return subject, body

    @staticmethod
    def extract_vars(text):
        return re.findall(r'\{\{.*?\}\}', text)

    @staticmethod
    def replace_vars(text, var_map, values):
        for var, index in var_map.items():
            text = text.replace(var, values[index])
        return text

    def truncate_and_add_header(self, header_vars):
        with open(self.email_list_file, 'w') as file:
            file.truncate()
        header = ["Email"] + [var.replace('{', '').replace('}', '') for var in header_vars]
        with open(self.email_list_file, 'a') as file:
            writer = csv.writer(file)
            writer.writerow(header)

    def run(self):
        try:
            self.load_credentials()
            subject_template, body_template = self.load_email_content()
            header, email_data = self.get_email_list()
            all_vars = self.extract_vars(subject_template) + self.extract_vars(body_template)
            var_map = {var: index for index, var in enumerate(all_vars)}

            if not self.login():
                input("Press enter to exit.")
                return

            count = 0
            for row in email_data:
                to_email = row[0]
                values = row[1:]
                subject = self.replace_vars(subject_template, var_map, values)
                body = self.replace_vars(body_template, var_map, values)
                mail_sent = self.send_email(subject, body, to_email, is_html=True)
                self.log_email_sent(to_email, mail_sent, subject, body.replace("\n", "\\n"), datetime.datetime.now())
                print(f"{count + 1}. Sent to {to_email}")
                count += 1

            print(f"Successfully sent {count} emails.")
            self.truncate_and_add_header(all_vars)
            input("Press enter to exit.")
        except Exception as e:
            input(f"Error: {e}. Contact support. Press enter to exit.")


if __name__ == "__main__":
    email_sender = EmailSender(
        credentials_file='credentials/email_password.json',
        email_list_file='list_of_emails.csv',
        subject_file='Setup Email Format/subject.txt',
        body_file='Setup Email Format/body.txt',
        sent_log_file='email_sent_list.csv'
    )
    email_sender.run()
