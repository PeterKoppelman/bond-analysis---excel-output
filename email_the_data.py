import smtplib
import mimetypes
import email_reference # list with email addresses to send Excel workbook to
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText


def email_my_data():

    msg = MIMEMultipart()
    msg["From"] = email_reference.emailfrom
    msg["To"] = email_reference.emailto
    msg["Subject"] = "Historical Delta Models"
    body = MIMEText("Hello,\n\nPlease find the relevant data attached.\n\nRegards,\n\nYour Data Team")
    msg.attach(body)

    ctype, encoding = mimetypes.guess_type(email_reference.filetosend)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
        fp = open(email_reference.filetosend)
        # Note: we should handle calculating the charset
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "image":
        fp = open(email_reference.filetosend, "rb")
        attachment = MIMEImage(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "audio":
        fp = open(email_reference.filetosend, "rb")
        attachment = MIMEAudio(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open(email_reference.filetosend, "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=email_reference.filetosend)
    msg.attach(attachment)

    server = smtplib.SMTP("smtp.gmail.com:587")
    server.starttls()
    server.login(email_reference.emailfrom,email_reference.password)
    server.sendmail(email_reference.emailfrom, email_reference.distribution_list, msg.as_string())
    server.quit()
