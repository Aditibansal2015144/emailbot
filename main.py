import smtplib
import imaplib
import email
from email.header import decode_header
import os
import speech_recognition as sr
import pyttsx3
from email.message import EmailMessage

# account credentials
username = "aditi@atulkbansal.com"
password = "##########"

recog = sr.Recognizer()
engine = pyttsx3.init()

def speak(text):
    print(text)
    engine.say(text)
    engine.runAndWait()

def listen():
    with sr.Microphone() as source:
        recog.adjust_for_ambient_noise(source)
        speak('speak now...')
        voice = recog.listen(source)
        try:
            info = recog.recognize_google(voice)
            speak(info)
            return info
        except:
            speak('PLEASE TRY AGAIN, COULD NOT UNDERSTAND WHAT YOU ARE SAYING')

email_list = {
    'Aditi': 'aditibansal9540@gmail.com',
    'Anushka': 'anushka@atulkbansal.com',
    'Atul': 'atul@atulkbansal.com',
    'Preeti': 'preetibansal833@gmail.com',
     'Ayush': 'ayush@atulkbansal.com'}

def information():
    speak('Tell receiver name')
    name = listen()
    receiver = email_list[name]
    speak(receiver)
    speak('What is the subject of your email?')
    subject = listen()
    speak('Tell me the text in your email')
    message = listen()
    send(receiver, subject, message)
    speak('your email has been sent successfully')

def send(receiver, subject, message):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    # Make sure to give app access in your Google account
    server.login(username,password)
    email = EmailMessage()
    email['From'] = 'Sender_Email'
    email['To'] = receiver
    email['Subject'] = subject
    email.set_content(message)
    server.send_message(email)

def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)
# create an IMAP4 class with SSL

def read():
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    # authenticate
    imap.login(username, password)
    status, messages = imap.select("INBOX")
    # number of top emails to fetch
    N = 2
    # total number of emails
    messages = int(messages[0])
    for i in range(messages, messages - N, -1):
        # fetch the email message by ID
        res, msg = imap.fetch(str(i), "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                # parse a bytes email into a message object
                msg = email.message_from_bytes(response[1])
                # decode the email subject
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    # if it's a bytes, decode to str
                    subject = subject.decode(encoding)
                # decode email sender
                From, encoding = decode_header(msg.get("From"))[0]
                if isinstance(From, bytes):
                    From = From.decode(encoding)
                speak('SUBJECT : ')
                speak(subject)
                speak('FROM : ')
                speak(From)
                # if the email message is multipart
                if msg.is_multipart():
                    # iterate over email parts
                    for part in msg.walk():
                        # extract content type of email
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        try:
                            # get the email body
                            body = part.get_payload(decode=True).decode()
                        except:
                            pass
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            # print text/plain emails and skip attachments
                            speak('message : ')
                            speak(body)
                        elif "attachment" in content_disposition:
                            # download attachment
                            filename = part.get_filename()
                            if filename:
                                folder_name = clean(subject)
                                if not os.path.isdir(folder_name):
                                    # make a folder for this email (named after the subject)
                                    os.mkdir(folder_name)
                                filepath = os.path.join(folder_name, filename)
                                # download attachment and save it
                                open(filepath, "wb").write(part.get_payload(decode=True))
                else:
                    # extract content type of email
                    content_type = msg.get_content_type()
                    # get the email body
                    body = msg.get_payload(decode=True).decode()
                    if content_type == "text/plain":
                        # print only text email parts
                        speak('message : ')
                        speak(body)
    # close the connection and logout
    imap.close()
    imap.logout()

speak('WELCOME TO VOICE BASED EMAIL SERVICE')
speak('HOW CAN I HELP YOU')
speak('1) SEND EMAIL')
speak('2) READ EMAIL')
speak('3) EXIT')

while(1):
    speak('WHAT IS YOUR CHOICE?')
    choice = listen()

    if (choice == 'send email'):
        speak('you want to send email')
        information()
    elif (choice == 'read email'):
        speak('opening your inbox')
        speak('your emails are')
        read()
        speak('THESE ARE THE TWO FIRST LATEST EMAILS')
    elif(choice == 'exit'):
        speak('are you sure you want to exit? (yes/no)')
        yn = listen()
        if(yn == 'yes'):
            speak('HAVE A NICE DAY!! BYE BYE')
            exit(1)
    else:
        speak('INVALID CHOICE')
