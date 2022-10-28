#### importanje PDF-ja
# importing required modules

import fitz
import random
import json
import tkinter as tk
import ctypes
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


input_file = "SKEN.pdf"
output_file = "poslanoPodpisani.pdf"

stamp = "limonera.png"
signature = random.choice(["podpisAmi.png","podpisAmra2.png"])

contactsFile = open("kontakti.json")
contactsData = json.load(contactsFile)

print(signature)
if(signature == "podpisAmra2.png"):
        povecaj = 20
else:
    povecaj = -10

neposlaniMaili = []

doc = fitz.open(input_file)

for page in doc:
    
    # define the position of the stamp
    x1Stamp = random.randrange(400, 410)
    y1Stamp = random.randrange(710, 720)
    x2Stamp = random.randrange(590, 610)
    y2Stamp = random.randrange(750, 760)
    # define the position of the signature
    x1signature = random.randrange(340, 360) - povecaj
    y1signature = random.randrange(650, 700) - povecaj
    x2signature = random.randrange(600, 630) + povecaj
    y2signature = random.randrange(780, 800) + povecaj

    stampRect = fitz.Rect(x1Stamp, y1Stamp, x2Stamp, y2Stamp)
    podpisRect = fitz.Rect(x1signature, y1signature, x2signature, y2signature)
    # add the image
    page.insert_image(stampRect,filename=stamp)
    page.insert_image(podpisRect,filename=signature)
    
doc.save(output_file)
pageCount = len(doc)
doc.close()

for pageNum in range(0,pageCount):
    doc = fitz.open(output_file) 
    doc.select([pageNum])
    docName = "CMR" + str(pageNum) + ".pdf"
    doc.save(docName)
    doc.close() 
    

for pageNum in range(0,pageCount):
    currentFile = "CMR" + str(pageNum) + ".pdf"
    doc = fitz.open(currentFile)
    text = doc.get_page_text(0,"text").replace('\n', ' ').upper()
    
    payload = MIMEBase('application','octate-stream')
    body = '''Hello,
    In the attachment you can find the CMR.
    '''
    foundContact = False
    for contact in contactsData:
        if(contact in text):
            foundContact = True
            contactToSendTo = contact

    if(foundContact == False):  
        contactToSendTo = "Nema nikog"
        neposlaniMaili.append(page.number)

    sender = '2'
    password = '2'
    receiver = contactsData[contact]

    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = ','.join(receiver)
    message['Subject'] = 'This email has an attacment, a pdf file'
    message.attach(MIMEText(body, 'plain'))
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()
    session.login(sender, password)


    payload.set_payload(doc.tobytes())
    encoders.encode_base64(payload)
    payload.add_header('Content-disposition','attachment',filename=currentFile)
    message.attach(payload)
    

    text1 = message.as_string()
    session.sendmail(sender, receiver, text1)
    session.quit()
    doc.close()
    print('Mail Sent')
    if(1==1):
        break
    



neposlaneStrani = listToStr = ' '.join([str(elem) for elem in neposlaniMaili])
ctypes.windll.user32.MessageBoxW(0, neposlaneStrani, "Neposlane strani", 1)
    
