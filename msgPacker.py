from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import os.path
# Function createEmail: assembles the Multipart email with:
# Sender, reciepent, cc, subject
# "mixed" encoding for attaments
#   "related" encoding for HTML with embedded images
#       "alternative" encoding for plain-text and html versions of the email
#
#MIMEMultipart structure:
#multipart/mixed                <- Root (because of attachments)
#|-- multipart/related          <- For HTML + embedded images
#|   |-- multipart/alternative  <- For text and HTML versions
#|   |   |-- text/plain         <- Plain text body
#|   |    -- text/html          <- HTML body
#|    -- image/jpeg             <- Embedded image(s)
#|-- application/pdf            <- Attachment(s)
# -- other attachments
#--------------------------------------------------------------
def createEmail(sender,recipient,cc,subject,text,html,attachments,images):

    # Create a multipart message with tag "mixed", to support attachements
    msg = MIMEMultipart("mixed")
    msg["From"] = sender
    msg["To"] = recipient
    if isinstance(cc,str): #check if cc is sting (if no cc then python defaults to float)
        msg["Cc"] = cc
    msg["Subject"] = subject


    # create MIMEMultipart subfields with tags "related" for html with embedded images
    #   and "alternative" for plain text/html versions of the email
    rel = MIMEMultipart("related")
    alt = MIMEMultipart("alternative")

    # attach plain-text and HTML versons of the email ("alternative")
    # add HTML/plain-text parts to MIMEMultipart message with utf-8 encoding for äöü
    #   the recieving email client will try to render the last part first
    alt.attach(MIMEText(text,"plain","utf-8"))
    alt.attach(MIMEText(html, "html","utf-8"))
    rel.attach(alt)

    #add embedded images to the html-part ("related")
    if images:
        i = 1
        for image in images:
            with open(image,"rb") as image:
                imgPart = MIMEImage(image.read())
                cid = "image00"+str(i)
                imgPart.add_header('Content-ID',cid)
                rel.attach(imgPart)
                i+=1
                #print(f"Image added: {image}")
    msg.attach(rel)

    #add attachements ("mixed")
    if attachments:
        for attachment in attachments:
            with open(attachment, 'rb') as f:
                file = MIMEApplication(f.read(), name=os.path.basename(attachment))
                file['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
                msg.attach(file)

    return msg