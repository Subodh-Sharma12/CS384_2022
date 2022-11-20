from platform import python_version
from datetime import datetime
import smtplib
import pandas as pd
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders



start_time = datetime.now()

data_reg_stud =pd.read_csv("input_registered_students.csv")
df=pd.read_csv("input_attendance.csv")
df_consolidat=pd.DataFrame()

df_consolidat['Roll']=data_reg_stud['Roll No'].copy()
df_consolidat['Name']=data_reg_stud['Name'].copy()

attendance=[]

def attendance_report():
    Roll_number = [str(i) for i in data_reg_stud['Roll No']]
    # We obtain the list of dates in which lectures were taken considering all occurences of 'Monday' and 'Thursday'
    va_dates = list({datetime.strptime(str(i).split(" ")[0],"%d-%m-%Y").date() for i in df['Timestamp']  if datetime.strptime(str(i).split(" ")[0],"%d-%m-%Y").strftime('%a') in ['Mon','Thu']})
    va_dates.sort()
    # We define a dictionary for storing the duplicate entries for each date
    false_attendance=[]
    invalid = {date : {} for date in va_dates}
    invalid_info = {roll_number : {date.strftime('%d-%m-%Y') : 0 for date in va_dates} for roll_number in Roll_number}
    attend_date = {roll_number : [] for roll_number in Roll_number}
    
    false_info = {roll_number : {date.strftime('%d-%m-%Y') : 0 for date in va_dates} for roll_number in Roll_number}
    #for loop for finding the students who have actual attendence, fake attendance, duplicate attendance
    for k in range(len(df['Timestamp'])):
        date_obj = datetime.strptime(str(df['Timestamp'][k]), '%d-%m-%Y %H:%M')
        date = date_obj.date()
        #Check if the person attended the class on monday or thursday
        if date_obj.weekday() == 0 or date_obj.weekday() == 3:
            if(date_obj.hour==14):
                stud_roll_no=(str(df['Attendance'][k])).split(" ")[0]
               
                if stud_roll_no in invalid[date]:
                    invalid[date][stud_roll_no]['entries'].append(date_obj)
                    invalid_info[(str(df['Attendance'][k])).split(" ")[0]][date.strftime('%d-%m-%Y')]+=1
                if stud_roll_no == 'nan' or stud_roll_no not in Roll_number:
                    continue
                else:
                    invalid[date][stud_roll_no] = {'name': df['Attendance'][k].split(' ', 1)[1], 'entries': [date_obj]}
                    attend_date[stud_roll_no].append(date.strftime('%d-%m-%Y'))
                    attendance.append(stud_roll_no)       
            if(date_obj.hour<14 or date_obj.hour>=15):
                false_attendance.append((str(df['Attendance'][k])).split(" ")[0])
                false_info[(str(df['Attendance'][k])).split(" ")[0]][date.strftime('%d-%m-%Y')]+=1
            
        else:
            false_attendance.append((str(df['Attendance'][k])).split(" ")[0])
  
    #Printing the attendance report for each student and also a consolidated report
    for k in range(len(data_reg_stud['Name'])):
        for date in va_dates:
            if date.strftime('%d-%m-%Y') in attend_date[data_reg_stud['Roll No'][k]]:
                df_consolidat.at[k, date.strftime('%d-%m-%Y')]='P'
            else:
                df_consolidat.at[k, date.strftime('%d-%m-%Y')]='A'
        df_consolidat.at[k,'Actual Lecture Taken']=len(va_dates)
        df_consolidat.at[k,'Total Real Attendance']=attendance.count(data_reg_stud['Roll No'][k])
        df_consolidat.at[k,'Percentage (attendance_count_actual/total_lecture_taken) 2 digit decimal']=(round((df_consolidat['Total Real Attendance'][k]/len(va_dates))*100,2))
        individual = pd.DataFrame()
        individual.at[0, 'Date']=''
        for j,date in enumerate(va_dates):
            individual.at[j+1, 'Date'] = date.strftime('%d-%m-%Y')
        individual.at[0,'Roll No'] = data_reg_stud['Roll No'][k]
        individual.at[0,'Name'] = data_reg_stud['Name'][k]
        individual.at[0,'total_attendance_count']=''
        individual.at[0,'Real']=attendance.count(data_reg_stud['Roll No'][k])
        individual.at[0,'Absent']=len(va_dates)-attendance.count(data_reg_stud['Roll No'][k])
        for j,date in enumerate(va_dates):
            individual.at[j+1,'invalid']=false_info[(str(df['Attendance'][k])).split(" ")[0]][date.strftime('%d-%m-%Y')]
            individual.at[j+1, 'duplicate']=invalid_info[data_reg_stud['Roll No'][k]][date.strftime('%d-%m-%Y')]
            if date.strftime('%d-%m-%Y') in attend_date[data_reg_stud['Roll No'][k]]:
                individual.at[j+1, 'Real']=1
                individual.at[j+1, 'Absent']=0
            else:
                individual.at[j+1, 'Absent']=1
                individual.at[j+1, 'Real']=0
            individual.at[j+1, 'total_attendance_count']=individual.at[j+1, 'Rael']+individual.at[j+1,'invalid']+individual.at[j+1, 'duplicate']
        try:
            individual.to_excel('output1/' + data_reg_stud['Roll No'][k] + '.xlsx',index=False)
        except PermissionError:
            print("You don't have the permission to read/write in this directory. Please grant permission or change the working directory")

start_time = datetime.now()

ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")
# We read the input files into a pandas dataframe each

attendance_report()
try:
    df_consolidat.to_excel('./output1/attendance_report_consolidated1.xlsx',index=False)
except:
    print("You don't have the permission to read/write in this directory. Please grant permission or change the working directory")

# Using Gmail
consent = input("Do you want to send email? Enter Yes or No")

print("Output generated\n")

try:
    if( consent == "Yes"):
        server = smtplib.SMTP('smtp.gmail.com',   587)
        fromaddr  = str(input("Enter Your Email\n"))
        frompasswd  = str(input("Enter Your Email Password\n"))
        toaddr = str(input("Enter the destination email\n"))
        # instance of MIMEMultipart
        try:
            msg = MIMEMultipart()
            print("[+] Message Object Created")
        except:
            print("[-] Error in Creating Message Object")
        
        # storing the senders email address  
        msg['From'] =fromaddr
        
        # storing the receivers email address 
        msg['To'] = toaddr
        
        # storing the subject 
        msg['Subject'] = "Attendance Report"
        
        # string to store the body of the mail
        body = "PFA the attendance Report"
        
        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain'))
        
        # open the file to be sent 
        filename = './output1/attendance_report_consolidated1.xlsx'
        
        attachment = open(filename, "rb")
        # print("Attendance File not found")
        
        # instance of MIMEBase and named as p
        p = MIMEBase('application', 'octet-stream')
        
        # To change the payload into encoded form
        p.set_payload((attachment).read())
        
        # encode into base64
        encoders.encode_base64(p)
        
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        
        
        try:
            msg.attach(p)
            print("[+] File Attached")
        except:
            print("[-] Error in Attaching file")

        try:
            #s = smtplib.SMTP('smtp.gmail.com', 587)
            s = smtplib.SMTP('mail.iitp.ac.in', 587)
            s.connect("smtp.example.com",465)
            s.ehlo()
            print("[+] SMTP Session Created")
        except:
            print("[-] Error in creating SMTP session")
        s.ehlo()
        s.starttls()

        try:
            s.login(fromaddr, frompasswd)
            print("[+] Login Successful")
        except:
            print("[-] Login Failed")

        text = msg.as_string()

        try:
            s.sendmail(fromaddr, toaddr, text)
            print("[+] Mail Sent successfully")
        except:
            print('[-] Mail not sent')

        s.quit()
    
            
    else:
        pass     
    
except FileNotFoundError:
    #Print file not found error if the file is not existing the directory
    print("File could not be found in the parent directory")
except ImportError:
    #If the pandas module is not installed, error is shown and program closes
    print("Sorry, module 'Pandas' could not be imported")
except PermissionError:
    print("You don't have the permission to read/write in this directory. Please grant permission or change the working directory")
#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
 
 