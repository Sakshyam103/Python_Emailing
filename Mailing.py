import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# future project: add a function to count the number of sent emails
def SendEmail(file, name, emailaddress):
    fromaddr = "from address"
    # password = "password" #password should be updated before running the code
    toaddr = emailaddress
    # <span style="background-color:#ffff33;">have a wonderful break!</span><br>
    subject = "Subject"
    html = """\
   <html>
   <head>
   <style>
   .p1 {
           font-family: Arial, sans-serif; font-size: 13px;}
   </style></head>
   <body>
   <p class = "p1"> Hello first, <br>
   
   Write text here </p></body></html>     
   """
    msg = MIMEMultipart()
    msg["From"] = "Write from name"
    msg["To"] = toaddr
    msg["Subject"] = subject
    msg.attach(MIMEText(html.replace("first", name, 1), "html"))  #replaces the first with name
    filename = file
    fileheader = "abc.pdf"
    attachment = open(filename, "rb")
    p = MIMEBase('application', 'octet-stream')
    p.set_payload(attachment.read())  # this is attaching the document to the email
    encoders.encode_base64(p)  # this encodes the file in base64 format
    p.add_header('Content-Disposition',
                 "attachment; filename=%s" % fileheader)  # this gives the name of the file to receiver
    msg.attach(p)
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()  # this is calling TLS (Transport Layer Security) connection
    s.login(fromaddr,
            "enter password")  # Manage your Google Account- Security-2 step verification- App passwords- wrote "PyCharm" and Google gave me the password
    text = msg.as_string()
    s.sendmail(fromaddr, toaddr, text)
    s.quit()
