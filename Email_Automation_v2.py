import csv
import datetime
from email.message import EmailMessage
import json
import ssl
import smtplib
import re

def get_sender_email():
    with open('credentials/email_password.json') as file:
        creds = json.load(file)
    email_sender = creds["email"]
    return email_sender

def login():
    with open('credentials/email_password.json') as file:
        creds = json.load(file)
    email_sender = creds["email"]
    password = creds["password"]
    smtp_url = creds.get("smtp", "smtp.office365.com")

    context = ssl.create_default_context()

    try:
        smtp = smtplib.SMTP(smtp_url, 587)  # Using STARTTLS on port 587
        smtp.starttls(context=context)
        smtp.login(email_sender, password)
    except smtplib.SMTPAuthenticationError:
        print("xxxxxxxxxxxxxxxxxxxxx       INVALID CREDENTIALS             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        return False
    except Exception as e:
        print("xxxxxxxxxxxxxxxxxxxxx       SOMETHING WENT WRONG (PLEASE CONTACT SHUBHAM)             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        print("xxxxxxxxxxxxxxxxxxxxx       ERROR MESSAGE : " + str(e) + "             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        return False

    return smtp

def send_email(smtp, email_subject, email_body, to_email, is_html=False):
    email_sender = get_sender_email()
    gm = EmailMessage()
    gm["From"] = email_sender
    gm["Subject"] = email_subject
    gm["To"] = to_email
    if is_html:
        gm.add_alternative(email_body, subtype='html')
    else:
        gm.set_content(email_body)

    mail_sent_status = smtp.sendmail(email_sender, to_email, gm.as_string())
    return mail_sent_status

def get_all_emails():
    rows = []
    header = []
    with open('list_of_emails.csv', 'r') as email_list_file:
        csv_reader = csv.reader(email_list_file)
        header = next(csv_reader)
        for row in csv_reader:
            rows.append(row)
    return header, rows

def truncate_and_add_header(var_header):
    with open('list_of_emails.csv', 'w') as email_list_file:
        email_list_file.truncate()

    write_this_header = ["Email"]
    for var in var_header:
        write_this_header.append(var.replace('{', '').replace('}', ''))
    with open('list_of_emails.csv', 'a') as email_list_file:
        writer_obj = csv.writer(email_list_file)
        writer_obj.writerow(write_this_header)
    return True

def mail_sent_update_csv(to_email="", mail_sent=False, from_email="", subject="", body="", date_time=""):
    email_sent = "Yes" if mail_sent else "No"
    write_this_value = [to_email, email_sent, from_email, subject, str(body), str(date_time)]
    with open('email_sent_list.csv', 'a', newline='\n') as email_sent_list:
        writer_obj = csv.writer(email_sent_list)
        writer_obj.writerow(write_this_value)
    return True

def get_subject_and_body():
    with open('Setup Email Format/subject.txt', 'r') as subject_file:
        subject = subject_file.read()
    with open('Setup Email Format/body.txt', 'r') as body_file:
        body = body_file.read()
    return subject, body

def get_all_vars(text):
    vars = re.findall(r'\{\{.*?\}\}', text)
    return vars

def replace_vars(text, var_column_name, variable_values):
    for column, index in var_column_name.items():
        text = text.replace(column, variable_values[index])
    return text

def press_enter_to_exit():
    input("\n\n/-/-/-/-/-/-/-/  Press enter to exit.    /-/-/-/-/-/-/-/-/-/")

if __name__ == "__main__":
    try:
        email_sender = get_sender_email()
        subject, im_body = get_subject_and_body()
        header, data = get_all_emails()
        count = 0
        email_var_with_index = dict()
        vars_body = get_all_vars(im_body)
        vars_subject = get_all_vars(subject)
        
        all_vars = vars_body + vars_subject
        for index, k in enumerate(all_vars):
            email_var_with_index[k] = index
        smtp = login()
        if smtp == False:
            press_enter_to_exit()
            exit()

        try:
            print("------------------------    Process Started     --------------------------------")

            for i in data:
                to_email = i[0]
                variable = []
                for var_name, index in email_var_with_index.items():
                    variable.append(i[index + 1])
                body = replace_vars(im_body, email_var_with_index, variable)
                email_subject = replace_vars(subject, email_var_with_index, variable)
                try:
                    mail_sent_status = send_email(smtp, email_subject, body, to_email, is_html=True)  # Set is_html to True if body is HTML
                    mail_sent_status = True
                except Exception as e:
                    mail_sent_status = False
                print(str(count + 1) + ". Sent to " + to_email)
                date_time = datetime.datetime.now()
                save_logs = mail_sent_update_csv(to_email, mail_sent_status, email_sender, email_subject, body.replace("\n", "\\n"), date_time)
                count = count + 1
            print("------------------------    Successfully sent to " + str(count) + " emails     --------------------------------")
            is_done = truncate_and_add_header(email_var_with_index)
            press_enter_to_exit()
        except IndexError:
            print("\n\nxxxxxxxxxxxxxxxxxxx         Email list is Empty Or Provide Variables               xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
            press_enter_to_exit()
    except Exception as e:
        input("/-/-/-/-/-/-/-/  Error Message : " + str(e) + " ---> Contact Shubham   /-/-/-/-/-/-/-/-/-/")
        press_enter_to_exit()
