import imaplib
import base64
import os,shutil
import email
import re
import pdfplumber
import psycopg2
import smtplib
from exchangelib import DELEGATE, Account, Credentials, EWSTimeZone
from exchangelib.indexed_properties import EmailAddress
from email.utils import make_msgid
from files import mmvl 
from files import pvvnl
import general_functions

import imaplib,email,smtplib
import os
mmvl.mmvl_vendor
pvvnl.pvvnl_vendor

conn=psycopg2.connect(
              host="localhost",
              database="Task_1",
              user="postgres",
              password="12345"
          )



def download_files():
    detach_dir = '.'
    if 'attachments' not in os.listdir(detach_dir):
        os.chdir("D:\\Squeal_String\\TASK\\all_pdf")

    user = 'adityatalegaonkar2997@gmail.com'
    pwd = "Aditya@2997"
    try:
        # imap_url='imap.gmail.com'
        con = imaplib.IMAP4_SSL('imap.gmail.com')
        con.login(user, pwd)
        con.select('INBOX')
        # type, data=con.uid('search', None, '(HEADER Subject "invoice")')
        typ, data = con.search(None, '(SUBJECT "invoice")')
        # type, data = con.search(None, 'ALL')
        mail_ids = data[0]
        # id_list = mail_ids.split()
        # latest_email_id=id_list[-1]

        # print(mail_ids)
        for msgId in data[0].split():
            typ, messageParts = con.fetch(msgId, '(RFC822)')
            if typ != 'OK':
                print('Error fetching mail.')
                # raise

            emailBody = messageParts[0][1]
            # print(emailBody)

            mail = email.message_from_bytes(emailBody)
            for part in mail.walk():
                if part.get_content_maintype() == 'multipart':
                    # print(part.as_string())
                    continue
                if part.get('Content-Disposition') is None:
                    # print (part.as_string())
                    # print("done")
                    continue
                fileName = part.get_filename()
                print('done')
                if bool(fileName):
                    filePath = os.path.join(detach_dir, "D:\\Squeal_String\\TASK\\all_pdf", fileName)
                    if not os.path.isfile(filePath):
                        print(fileName)
                        fp = open(filePath, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
            con.close()
            con.logout()
    except:
        print("unable to download all attachments")

download_files()


path_pdf = r"D:\Squeal_String\TASK\all_pdf"

files = os.listdir(path_pdf)
os.chdir(path_pdf)
def move_file(path_pdf):
    text1 = ""
    with pdfplumber.open(path_pdf) as pdf:
        pg = pdf.pages
        # print(pg)
        for i in range(len(pg)):
            txt = pg[i].extract_text()
            text1 = text1+'\n--------------------------------------------new page------------------------------------------\n'+txt
        # print(text1)
        
    if 'PURVANCHAL VIDYUT VITRAN NIGAM LIMITED' in text1:
        shutil.move(path_pdf,r'D:\Squeal_String\TASK\Vendors-PVVNL')
        # print('file 1 move')     
    
    elif 'MADHYANCHAL VIDYUT VITRAN NIGAM LIMITED' in text1:
        shutil.move(path_pdf,r'D:\Squeal_String\TASK\Vendors-MVVNL')
        # print('file 2 move')


for file in os.listdir(path_pdf):
    new_name=file.replace(" ","_")
    os.rename(r'D:\Squeal_String\TASK\all_pdf\\'+file,r'D:\Squeal_String\TASK\all_pdf\\'+new_name)
    file1=(r'D:\Squeal_String\TASK\all_pdf\\'+new_name)
    print(file1)
    move_file(file1)

path_rename1 = 'D:\Squeal_String\TASK\Vendors-PVVNL'
os.chdir(path_rename1)
general_functions.rename_file(path_rename1)

path_rename2 = 'D:\Squeal_String\TASK\Vendors-MVVNL'
os.chdir(path_rename2)
general_functions.rename_file(path_rename2)


#------------------------------

path1 = r"D:\Squeal_String\TASK\Vendors-PVVNL\\"
files = os.listdir(path1)
os.chdir(path1)
for file in os.listdir(path1):
    new_name = file.replace(" ", "_")
    os.rename(r'D:\Squeal_String\TASK\Vendors-PVVNL\\' + file, r'D:\Squeal_String\TASK\Vendors-PVVNL\\' + new_name)
    file1 = (r'D:\Squeal_String\TASK\Vendors-PVVNL\\' + new_name)
    print(file1)
    pvvnl.pvvnl_vendor(file1) 
    

#--------------------------------------
path2 = r"D:\Squeal_String\TASK\Vendors-MVVNL\\"
files = os.listdir(path2)
os.chdir(path2)
for file in os.listdir(path2):
    new_name = file.replace(" ", "_")
    os.rename(r'D:\Squeal_String\TASK\Vendors-MVVNL\\' + file, r'D:\Squeal_String\TASK\Vendors-MVVNL\\' + new_name)
    file2 = (r'D:\Squeal_String\TASK\Vendors-MVVNL\\' + new_name)
    print(file2)
    mmvl.mmvl_vendor(file2)


# path_zip = 'D:\Squeal_String\TASK\Zip_files'
# os.chdir(path_zip)
general_functions.zipdir()

general_functions.database_to_csv()

general_functions.csv_to_excel()



general_functions.send_mail()

general_functions.send_mail2()

general_functions.send_mail3()

















        
