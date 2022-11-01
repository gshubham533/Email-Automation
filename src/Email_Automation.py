import csv,datetime   
from email.message import EmailMessage
import json , ssl, smtplib, re


"""
Returns the sender/owner/from/Receipient Email address
"""
def get_sender_email():
    file = open('src/credentials/email_password.json')
    creds =  json.load(file)
    email_sender        = creds["email"]
    file.close()
    return email_sender


"""
Authenticated the gmail account and returns smtp instance to send emails.
"""
def login():
    file = open('src/credentials/email_password.json')
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
        return False
    except Exception as e:
        print("xxxxxxxxxxxxxxxxxxxxx       SOMETHING WENT WRONG (PLEASE CONTACT SHUBHAM)             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        print("xxxxxxxxxxxxxxxxxxxxx       ERROR MESSAGE : "+str(e)+"             xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        return False

    return smtp

"""
Args : Email Subject , Email Body , To Send

Returns empty dictionary on success ex: {}
"""
def send_email(smtp,email_subject,email_body,to_email):
    email_sender = get_sender_email()
    gm = EmailMessage()
    gm["From"]      =  email_sender
    gm["Subject"]   = email_subject
    gm["To"]   = to_email
    gm.set_content(email_body)

    mail_sent_status = smtp.sendmail(email_sender,to_email,gm.as_string())
    return mail_sent_status


"""
Returns the header and the rows from the file list_of_emails

Ex : header = ["email","name"]
     body = [["abc@hotmail.com","abc"],["test@gamil.com","test"]]
"""
def get_all_emails():
    rows = []
    header = []
    with open('src/list_of_emails.csv', 'r') as email_list_file:
        csv_reader = csv.reader(email_list_file)
        header = next(csv_reader)
        for row in csv_reader:
            rows.append(row)
        
    return header,rows

"""
truncates the file list_of_emails and then adds the headers.

1 Arg : List of variables present in body.txt 
Ex. var_header = ["name", "company name"]


default header at index 0 is Email
Then it searches for varibale name in the body.txt and adds them as header as well
Ex. : Email,name,company name
"""
def truncate_and_add_header(var_header):
    with open('src/list_of_emails.csv', 'w') as email_list_file:
        email_list_file.truncate()

    write_this_header = ["Email"]
    for var in var_header:
        write_this_header.append(var.replace('{','').replace('}',''))
    with open('src/list_of_emails.csv', 'a') as email_list_file:
        writer_obj = csv.writer(email_list_file)
        writer_obj.writerow(write_this_header)
        email_list_file.close()       
 
    return True


"""
Updates the logs file named email_sent_list.csv with the status

Args : To Email (Receipents email), Email Sent (Yes/No), From Email (senders email), Subject (Email Subject), Body (Email body with next line as \n Ex. Dear Shubham\n\nThanks), DateTime (Email sent date and time) 

returns Boolean.
If the logging was succes then returns True else False
"""
def mail_sent_update_csv(to_email="",mail_sent=False,from_email="",subject="",body="",date_time=""):
    if mail_sent == True:
        email_sent = "Yes"
    else:
        email_sent = "No"
    write_this_value = [to_email,email_sent,from_email,to_email,subject,str(body),str(date_time)]
    with open('src/email_sent_list.csv', 'a', newline='\n') as email_sent_list:
        writer_obj = csv.writer(email_sent_list)
        writer_obj.writerow(write_this_value)
        email_sent_list.close()

    return True

"""
return to variables subject and body from file "Setup Email Format/subject.txt" and "Setup Email Format/body.txt"
"""
def get_subject_and_body():
    subject = ""
    body = "" 
    with open('src/Setup Email Format/subject.txt', 'r') as subject_file:
        subject = subject_file.readlines()
        subject = "".join(subject)
        subject_file.close()
    with open('src/Setup Email Format/body.txt', 'r') as body_file:
        body = body_file.readlines()
        body = "".join(body)
        body_file.close()

    return subject,body

"""
Using regex return list of all variables in body.txt enclosed in braces {}

Ex. Dear {name}
    Thanks,
    {my name}

returns vars = ["{name}","{my name}"]
"""
def get_all_vars(body):
    vars = re.findall('\{.*?\}', body)
    return vars


"""
Args : Email Body (str) , Variables from body with its index no. (dict) , Varibale Value (list) 

replaces the variables with the actual value in email body

Ex. : Email Body = Dear {name}
                    Regards
                    {my name}
      var_column_name  = {"{name}":0,"{my name}":1}
      variable_values = ["Alex","Shubham"]

returns Body 

Ex. Dear Alex,
    Regards
    Shubham
"""
def replace_vars_in_body(body,var_column_name,variable_values):
    for column,index in var_column_name.items():
        body = body.replace(column,variable_values[index])
    return body

def press_enter_to_exit():
    print(input("\n\n/-/-/-/-/-/-/-/  Press enter to exit.    /-/-/-/-/-/-/-/-/-/"))


if __name__ == "__main__":
    try:

        email_sender = get_sender_email()
        subject,im_body = get_subject_and_body()
        header,data = get_all_emails()
        count = 0
        email_var_with_index = dict()
        vars = get_all_vars(im_body)
        for index,k in enumerate(vars):
            # k = k.replace('{','').replace('}','')
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
                for var_name,index in email_var_with_index.items():
                    variable.append(i[index+1])
                body = replace_vars_in_body(im_body,email_var_with_index,variable)
                try:
                    mail_sent_status = send_email(smtp,subject,body,to_email) # returns {} (empty dictionary)
                    mail_sent_status = True
                except Exception as e:
                    mail_sent_status = False
                print(str(count + 1) + ". Sent to " + to_email)
                date_time = datetime.datetime.now()
                save_logs = mail_sent_update_csv(to_email,mail_sent_status,email_sender,subject,body.replace("\n","\\n"),date_time)
                count = count + 1
            print("------------------------    Successfully sent to "+str(count)+" emails     --------------------------------")
            is_done = truncate_and_add_header(email_var_with_index)
            press_enter_to_exit()
        except IndexError:
            print("\n\nxxxxxxxxxxxxxxxxxxx         Email list is Empty Or Provide Variables               xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
            press_enter_to_exit()
    except Exception as e:
        print(input("/-/-/-/-/-/-/-/  Error Message : "+str(e)+" ---> Contact Shubham   /-/-/-/-/-/-/-/-/-/"))
        press_enter_to_exit()
        