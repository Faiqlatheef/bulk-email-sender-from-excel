# Python code to send email to a list of
# emails from a spreadsheet

# import the required libraries
import pandas as pd
import smtplib
import time

# change these as per use
your_email = "fifa61295@gmail.com"
your_password = "fifa2552"

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# reading the spreadsheet
email_list = pd.read_excel('Email.xlsx')

# getting the names and the emails
names = email_list['Message']
emails = email_list['Emails']

while True:
	# iterate through the records
	for i in range(len(emails)):
		# for every record get the name and the email addresses
		name = names[i]
		email = emails[i]

		# the message to be emailed
		message = name

		# sending the email
		server.sendmail(your_email, [email], message)
	time.sleep(60 * 2)  # this is in seconds, so 60 seconds x 30 mins

# close the smtp server
server.close()
