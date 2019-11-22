import os, sys
import smtplib, ssl
import xlrd
import email
import imaplib
import time

def emailSend(useremail, password, Text):
	port = 587  # For starttls
	smtp_server = "smtp.gmail.com"
	To = "" #Report to send out
	Subject = "User cannot make ticket in Jira from email"
	message = 'Subject: {}\n\n{}'.format(Subject, Text)
	context = ssl.create_default_context()
	with smtplib.SMTP(smtp_server, port) as server:
	    server.ehlo()  # Can be omitted
	    server.starttls(context=context)
	    server.ehlo()  # Can be omitted
	    server.login(useremail, password)
	    server.sendmail(useremail, To, message)
	print("Email was sent out")

def loop(useremail, password, sender, subject, Date):
	print(sender)
	num = 1
	loc = ("list.xlsx")
	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)
	while num < 3000:
		try:
			Name = sheet.cell_value(num, 1)
			Emails = sheet.cell_value(num, 3)
			Status = sheet.cell_value(num, 7)
			License = sheet.cell_value(num, 9)
			# print('Name: ' + Name)
			# print('Email: ' + Emails)
			# print('Status: ' + Status)
			if Emails == sender and Status == "Active":
				print('Name: ' + Name)
				print('Email: ' + Emails)
				print('Status: ' + Status)
				print("User exist in Jira")
				print("I finished... :)")
				break
			elif Emails == sender and Status == "Inactive":
				Text = Name + """ account in Jira needs to be reviewed as there is a problem with how his account emails are setup. \n Email: """ + sender + " \n License: " + License + " \n Status: " + Status
				emailSend(useremail, password, Text)
				break
			elif sender in (""):#List of email address to report on from failure due to Zendesk auto-reply header
				Text = """Email comes from Zendesk """ + sender + "\n Subject: " + subject + "\n Date: " + Date
				emailSend(useremail, password, Text)
				break
			else:
				print("Still searching for " + sender)
		except IndexError:
			print("I finished... :)")
			print("Gone thru all Users no match for: " + sender)
			Text = """There is no account in Jira currently for this user: """ + sender
			emailSend(useremail, password, Text)
			print("Email was sent by: " + sender)
			print("Found at: " + time.ctime())
			print()
			break
		num += 1
		print()

useremail = ""#Email channel or mail handler connected to Jira
password = ""#app password
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(useremail, password)
mail.list()
mail.select("inbox")
reach = 1
while reach < 10:
	result, data = mail.search(None, "ALL")
	ids = data[0] # data is a list.
	id_list = ids.split() # ids is a space separated string
	latest_email_id = id_list[-reach] # get the latest
	result, data = mail.fetch(latest_email_id, "(RFC822)") # fetch the email body (RFC822) for the given ID
	raw_email = data[0][1] # here's the body, which is raw text of the whole email
	msg = email.message_from_bytes(raw_email)
	print(msg['From'])
	try:
		term = msg['From'].split("<")[1]
		sender = term.split(">")[0]
		subject = msg['Subject']
		Date = msg['date']
	except IndexError:
		sender = msg['From']
		subject = msg['Subject']
		Date = msg['date']
	loop(useremail, password, sender, subject, Date)
	print("Email count: " + str(reach))
	print("Finished at: " + time.ctime())
	reach += 1
time.sleep(1500)
os.execl(sys.executable, sys.executable, *sys.argv)
