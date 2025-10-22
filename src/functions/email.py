
#######################################################################################################################################################
## Logging
import logging
logger = logging.getLogger(f"root.email")
#######################################################################################################################################################
import sys
import os
from pathlib import Path

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

################################################################################################################################################################
def sendEmail(sendTo:list, sendFrom:str, subject:str, message_text:str, text_type:str, attachments:list)->str:
    returnMsgs = ''
    # if type(sendTo) == list:
    #     send_to = ", ".join(sendTo)
    # else:
    #     send_to = sendTo
    try:
        msg = MIMEMultipart()
        msg['Subject'] = subject

        msg.attach(MIMEText(message_text, text_type.lower()))
        ## Adds files to email as attachments
        for attachment in attachments or []:
            with open(attachment, "rb") as f:
                file = MIMEApplication(f.read(), name=os.path.basename(attachment))
            
            # After the file is closed
            file['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"' 
            msg.attach(file)
        smtp = smtplib.SMTP('smtp.hdrinc.com')
        print(sendTo)
        smtp.sendmail(sendFrom, sendTo, msg.as_string())
        smtp.close()

        returnMsgs = f'-- Message Sent To --\n{sendTo}'

    except Exception as e:
        returnMsgs = str(e)
    
    return returnMsgs
