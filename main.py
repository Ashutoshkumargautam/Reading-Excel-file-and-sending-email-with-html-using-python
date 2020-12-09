import pandas as pd
import smtplib
#===================================================
your_email = "xyzxxxaxbxzx123@gmail.com"
your_password = "xxx@123456"
#===================================================
# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
#===================================================
# reading the spreadsheet
email_list = pd.read_excel('C:/Users/Perfect TC
Society/PycharmProjects/New_Project/Email_address_with_location.xlsx')
#===================================================
# getting the names and the emails
names = email_list['NAME']
emails = email_list['EMAIL']
#===================================================
# iterate through the records
for i in range(len(emails)):

   # for every record get the name and the email addresses
   name = names[i]
   email = emails[i]

   # the message to be emailed
   message = ""

   # sending the email
   server.sendmail(your_email, [email], message)

   #print how many email sent to the users
   print(" [+] Sent", emails)

# close the smtp server
server.close()
