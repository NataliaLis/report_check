import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(email_from, email_to: list, email_subject, email_cc: list, email_body, email_format):

    message = "From: %s\r\n" % email_from \
        + "To: %s\r\n" % email_to \
        + "CC: %s\r\n" % ",".join(email_cc) \
        + "Subject: %s\r\n" % email_subject \
        + "\r\n" \
        + email_body

    email_tos = email_to + email_cc
    server = smtplib.SMTP('localhost')
    #server.set_debuglevel(1)
    server.sendmail(email_from, email_tos, message)
    server.quit()
