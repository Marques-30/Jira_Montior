import email, email.utils, smtplib, ssl, time, imaplib, os, sys
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText

def gmail(msg, sender, User, msg_data, payload, b, plan, password, useremail):
    you = sender
    bcc = ""
    # Create message container - the correct MIME type is multipart/alternative.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = plan #"Delivery Status Notification (Failure)"
    msg['From'] = me
    msg['To'] = you
    # Create the body of the message (a plain-text and an HTML version).
    text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
    html = """\
    <html>
      <head></head>
      <body>
        <p>Hello %s,<br /><br />
Emails from Zensdesk tickets cced to ask-security@ will bounce; please use one of the following methods to contact us: <br />
   1. Go to Jira Service Desk Portal (drl/ask-security) to document your request<br />     2. Document your request with as much information as you can and send an email (Please do not simply forward an email trail as it can be difficult to decipher what the request is for security)<br /><br />
Regards, <br /> SecOps<br /><br /></p><p><table><tr>%s</tr></table</p>
      </body>
    </html>
    """ % (User, b.get_payload().replace("=", ""))
    # Record the MIME types of both parts - text/plain and text/html.
    part = MIMEText(html, 'html')
    # Attach parts into message container.
    # According to RFC 2046, the last part of a multipart message, in this case
    # the HTML message, is best and preferred.
    msg.attach(part)
    # Send the message via local SMTP server.
    mail = smtplib.SMTP('smtp.gmail.com', 587)
    mail.ehlo()
    mail.starttls()
    mail.login(useremail, password)
    if msg['cc'] != None:
        cc = msg['cc']
        mail.sendmail(useremail, you, cc, msg.as_string())
    else:
        mail.sendmail(useremail, you, msg.as_string())


def loop(msg, sender, Date, User, msg_data, b, payload, password, useremail):
    print('Name: ' + msg['From'])
    print('Subject: ' + msg['subject'])
    print(msg['header'])
    print(User)
    plan = msg['Subject']
    if sender in (""): #List of email address from Zendesk
        gmail(msg, sender, User, msg_data, b, payload, plan, password, useremail)
        print("email sent")


useremail = "" #Email channel ot mail handler connected to Jira
password = "" # app password
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(useremail, password)
mail.list()
mail.select("Zendesk")
num = 0
for i in range(1, 30):
    typ, msg_data = mail.fetch(str(i), '(RFC822)')
    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])
            for header in ['From']:
                print(msg['Date'])
                print('%-8s: %s' % (header.upper(), msg[header]))
                try:
                    term = msg['From'].split("<")[1]
                    sender = term.split(">")[0]
                    User = msg['From'].split("<")[0]
                    Date = msg['Date'].split("+")[0]
                    b = email.message_from_bytes(response_part[1])
                    if b.is_multipart():
                        for payload in b.get_payload():
                            # if payload.is_multipart(): ...
                            print("working")
                    else:
                        print("Email was retrieved successfully")
                except IndexError:
                    sender = msg['From']
                print(msg[header])
                if msg["Subject"] != "Delivery Status Notification (Failure)":
                    loop(msg, sender, Date, User, msg_data, b, payload, password, useremail)
                    typ, response = mail.store(str(i), '+FLAGS', r'(\Deleted)')
                    mail.expunge()
                else:
                    print("Email is a response from previous Zendesk ticket")
                    break
                print("Last completed: " + time.ctime())
                time.sleep(30)
                num += 1
                print("Number of emails sent: " + str(num))
                print()
print()
print("All emails have been checked during this cycle, will check again in 15 minutes")
time.sleep(800)
print()
os.execl(sys.executable, sys.executable, *sys.argv)
