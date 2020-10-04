import os
import imaplib
import email

#attachment file in the email
server = 'outlook.office365.com'
user = 'example@email.com' #Your Email
password = '12345' #Your Password
outputdir = '/Users/jose/Desktop' #The dir to say the attachment from the file
subject = 'test' #subject line of the emails you want to download attachments from
def connect(server, user, password):
    m = imaplib.IMAP4_SSL(server)
    m.login(user, password)
    m.select()
    return m
def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))
def subjectQuery(subject):
    m = connect(server, user, password)
    m.select("Inbox")
    typ, msgs = m.search(None, '(SUBJECT "' + subject + '")')
    msgs = msgs[0].split()
    for emailid in msgs:
        downloaAttachmentsInEmail(m, emailid, outputdir)

subjectQuery(subject)