# HR-Automation
This Python script will help you in automation like sending bulk emails in a go. To get started follow the steps mentioned below.  


Step 1 (Getting your credentials from gmail to send automated emails) :  
a. Go to http://myaccount.google.com/  
b. If you are not signed in then enter your gmail crendentials to go ahead and if you are already signed in then move to next step.  
c. Go to Security Tab and then turn on you 2-step verification.  
d. Go to security tab and then in "Signing in to Google" section click on "App passwords" (Google will prompt to enter your password so enter you gmail passwords)  
e. Under "select app" select "other (custom name)" and provide any name that you want name. (In my case I will name it "Python")  
d. Then a pop will appear containg your 16 letter password simply copy it and paste it inside the script file inside credentials folder in email_password.json file (credentials/email_password.json)  

[Note] : Password will appear once so copy and paste it properly else you would have to start the process from step d. again.  

example :  
  
{  
    "email":"your email here",  
    "password":"your 16 letter password here"  
}  

Step 2:  
Go inside "Setup Email Format" folder and save the subject of the email in "subject.txt" file and save body of the email in "body.txt" file.  

Step 3:  
Open "list_of_emails.csv" file and enter the email addresses to send in email column.  

step 4:  
Run HR_Automation.py script.  

[Note] : Once you run the script the emails you entered will be clearead from file "list_of_emails.csv"  
[Note] : You can check the activity report in the file "email_sent_list.csv"   

Screenshot
! [Output ScreenShot] (https://raw.githubusercontent.com/gshubham533/HR-Automation/main/screenshot/Email%20Automation%20Screenshot.png)
