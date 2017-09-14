''' python tool for reading data from a mysql database and creatimg a excel sheet and than sending it over the mail.'''
''' Author: Laxman Singh ~ laxman.1390@gmail.com '''

import datetime
import mysql.connector
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

cnx = mysql.connector.connect(user='root', host='localhost', database='jpa_onetomany')
cursor = cnx.cursor()
query = ("SELECT id , name, email, phone, address from mail_data where email = %s")
#query = ("SELECT id , name, email, phone, address from mail_data")

query_param = ['abc@xyz.com']

cursor.execute(query, query_param)
#cursor.execute(query)

workbook = xlsxwriter.Workbook('files/mail_attachment.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_paper(1)  # US Letter
worksheet.set_paper(9)  # A4

worksheet.set_column('A:A', 5)
worksheet.set_column('B:C', 20)
worksheet.set_column('C:D', 20)
worksheet.set_column('E:E', 50)
bold = workbook.add_format({'bold': 1})

worksheet.write('A1', 'Sr No.', bold)
worksheet.write('B1', 'Name', bold)
worksheet.write('C1', 'Email', bold)
worksheet.write('D1', 'Phone', bold)
worksheet.write('E1', 'Address', bold)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': 1})

i=2;
for (id, name, email, phone, address) in cursor:
  print("{}, {}, {}, {}, {}".format(
    id, name, email, phone, address))
  worksheet.write('A'+str(i), id)
  worksheet.write('B'+str(i), name)
  worksheet.write('C'+str(i), email)
  worksheet.write('D'+str(i), phone)
  worksheet.write('E'+str(i), address)
  i=i+1

workbook.close();

cursor.close()
cnx.close()


''' send mail '''
fromaddr = "laxman.jboss@gmail.com"
toaddr = "laxman.1390@gmail.com"

# instance of MIMEMultipart
msg = MIMEMultipart()

# storing the senders email address
msg['From'] = fromaddr

# storing the receivers email address
msg['To'] = toaddr

# storing the subject
msg['Subject'] = "Automated mail with excel sheet"

# string to store the body of the mail
body = "Hi All, \n Please find attcahed data for xyx."

# attach the body with the msg instance
msg.attach(MIMEText(body, 'plain'))

# open the file to be sent
filename = "mail_attachment.xlsx"
attachment = open("files/mail_attachment.xlsx", "rb")

# instance of MIMEBase and named as p
p = MIMEBase('application', 'octet-stream')

# To change the payload into encoded form
p.set_payload((attachment).read())

# encode into base64
encoders.encode_base64(p)

p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

# attach the instance 'p' to instance 'msg'
msg.attach(p)

# creates SMTP session
s = smtplib.SMTP('smtp.gmail.com', 587)

# start TLS for security
s.starttls()

# Authentication
s.login(fromaddr, "your_passwd")

# Converts the Multipart msg into a string
text = msg.as_string()

# sending the mail
s.sendmail(fromaddr, toaddr, text)

# terminating the session
s.quit()
