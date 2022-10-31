import csv,datetime   
from email.message import EmailMessage
import json , ssl, smtplib, re

def get_sender_email():
    file = open('credentials/email_password.json')
    creds =  json.load(file)
    email_sender        = creds["email"]
    file.close()
    return email_sender

def login():
    file = open('credentials/email_password.json')
    creds =  json.load(file)
    email_sender        = creds["email"]
    password            = creds["password"]
    file.close()

    context = ssl.create_default_context()
    smtp = smtplib.SMTP_SSL("smtp.gmail.com", 465 , context=context)
    try:
        login = smtp.login(email_sender,password)
    except smtplib.SMTPAuthenticationError:
        print("xxxxxxxxxxxxxxxxxxxxx       INVALID CREDENTIALS             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
    except Exception as e:
        print("xxxxxxxxxxxxxxxxxxxxx       SOMETHING WENT WRONG (PLEASE CONTACT SHUBHAM)             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        print("xxxxxxxxxxxxxxxxxxxxx       ERROR MESSAGE : "+e+"             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")

    return smtp

def send_email(email_subject,email_body,to_email):
    gm = EmailMessage()
    gm["From"]      =  get_sender_email()
    gm["Subject"]   = email_subject
    gm["To"]   = to_email
    gm.set_content(email_body)

    smtp = login()
    mail_sent_status = smtp.sendmail(email_sender,to_email,gm.as_string())
    return mail_sent_status

def get_all_emails():
    rows = []
    header = []
    with open('list_of_emails.csv', 'r') as email_list_file:
        csv_reader = csv.reader(email_list_file)
        header = next(csv_reader)
        for row in csv_reader:
            rows.append(row)
        
    return header,rows

def truncate_and_add_header(var_header):
    with open('list_of_emails.csv', 'w') as email_list_file:
        email_list_file.truncate()

    write_this_header = ["Email"]
    for var in var_header:
        write_this_header.append(var.replace('{','').replace('}',''))
    print(write_this_header)
    with open('list_of_emails.csv', 'a') as email_list_file:
        writer_obj = csv.writer(email_list_file)
        writer_obj.writerow(write_this_header)
        email_list_file.close()       
 
    return True

def mail_sent_update_csv(to_email="",mail_sent=False,from_email="",subject="",body="",date_time=""):
    if mail_sent == True:
        email_sent = "Yes"
    else:
        email_sent = "No"
    write_this_value = [to_email,email_sent,from_email,to_email,subject,str(body),str(date_time)]
    with open('email_sent_list.csv', 'a', newline='\n') as email_sent_list:
        writer_obj = csv.writer(email_sent_list)
        writer_obj.writerow(write_this_value)
        email_sent_list.close()

    return True

def get_subject_and_body():
    subject = ""
    body = "" 
    with open('Setup Email Format/subject.txt', 'r') as subject_file:
        subject = subject_file.readlines()
        subject = "".join(subject)
        subject_file.close()
    with open('Setup Email Format/body.txt', 'r') as body_file:
        body = body_file.readlines()
        body = "".join(body)
        body_file.close()

    return subject,body

def get_all_vars(body):
    vars = re.findall('\{.*?\}', body)
    return vars

def replace_vars_in_body(body,var_column_name,variable_values):
    for column,index in var_column_name.items():
        body = body.replace(column,variable_values[index])
    return body

if __name__ == "__main__":
    try:

        email_sender = get_sender_email()
        subject,im_body = get_subject_and_body()
        header,data = get_all_emails()
        count = 0
        # print(header)
        print("------------------------    Process Started     --------------------------------")
        email_var_with_index = dict()
        vars = get_all_vars(im_body)
        
        # print(vars)
        for index,k in enumerate(vars):
            # k = k.replace('{','').replace('}','')
            email_var_with_index[k] = index

        try:
            for i in data:
                to_email = i[0]
                variable = []
                for var_name,index in email_var_with_index.items():
                    variable.append(i[index+1])
                body = replace_vars_in_body(im_body,email_var_with_index,variable)
                try:
                    mail_sent_status = send_email(subject,body,to_email) # returns {} (empty dictionary)
                    mail_sent_status = True
                except Exception as e:
                    mail_sent_status = False
                print(str(count + 1) + ". Sent to " + to_email)
                date_time = datetime.datetime.now()
                save_logs = mail_sent_update_csv(to_email,mail_sent_status,email_sender,subject,body.replace("\n","\\n"),date_time)
                count = count + 1
            print("------------------------    Successfully sent to "+str(count)+" emails     --------------------------------")
            is_done = truncate_and_add_header(email_var_with_index)
            print(input("/-/-/-/-/-/-/-/  Press enter to exit.    /-/-/-/-/-/-/-/-/-/"))
        except IndexError:
            print("\n\nxxxxxxxxxxxxxxxxxxx         Email list is Empty Or Provide Variables               xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
    except Exception as e:
        print(input("/-/-/-/-/-/-/-/  Error Message : "+str(e)+" ---> Contact Shubham   /-/-/-/-/-/-/-/-/-/"))
        