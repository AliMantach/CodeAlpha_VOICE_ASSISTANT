import win32com.client
import os
import speech_recognition as sr
import pyttsx3
import datetime
import re
import pywhatkit
#init
#to recogniize voice
listener = sr.Recognizer() 
speak = pyttsx3.init()
#to set all the voices
voices = speak.getProperty('voices')     
speak.setProperty('voice', voices[1].id)  
#func to speak
def talk(text):
    speak.say(text)
    speak.runAndWait()

#to ans our ques
def take_command():
     # Initialize command 
    command = ""  
    try:
        with sr.Microphone() as source: # use the microphone (i can use ---as mic )
            print("  listening....")
            #listen to the mic 
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            command = command.lower()
            if any(command == variant for variant in ['alexa', 'axla', 'alxa', 'lexa','alex','lexa']):
              talk('how can i help you')
              take_command()
            elif 'alexa' in command:
                command = command.replace('alexa','')
                print(command)    
    except:
        pass
    return command
#function to find the mail in a string
def extract_email(text):
    email_regex = r'[\w\.-]+@[\w\.-]+'
    emails = re.findall(email_regex, text)
    return ', '.join(emails)
#func to find a subject after a string with in the subject is 
def extract_subject(text):
    pattern = r"the subject is\s+(.*)$"
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    else:
        return None
#func to send mails using outlook only     
def send_email(subject=None, body=None, to=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    if subject:
        mail.Subject = subject
    if body:
        mail.Body = body
    if to:
        mail.To = to
    mail.Display()
#recursion function to calculate fibo    
def fibonacci(n):
    if n <= 1:
        return n
    else:
        return fibonacci(n-1) + fibonacci(n-2)    
    
#main function 
def run():
    global running
    #taking the command using speech_recognition 
    command = take_command()
    print(command)
    #test if the command start with a specific tag and run it 
    #tag 1 play a muisc
    if any(command.startswith(variant) for variant in ["play", "plau", "paly", "plsy", "start", "begin", "start the", "begin the", "put on"]):
        song = command.replace("play ", '')
        try:
            talk("playing " + song)
            pywhatkit.playonyt(song)
        except:
         pass
    #tag 2 open google chrome    
    elif any(variant in command for variant in ["chrome", "crom", "chrme", "crhom", "chome", "chorme", "chroem", "chome", "chorm", "chromo"]):
        try :
         os.system("start chrome")
         speak.say("chrome has been open")
         speak.runAndWait()
        except:
         pass 
    #tag 3 open notepad    
    elif any(variant in command for variant in ["notepad", "notpad", "notepadp", "note pad", "note padp", "notepad p", "notepadpp", "notepadd"]):
        try:
            os.system("notepad")
            speak.say("notepad has been open")
            speak.runAndWait()
        except:
            pass  
    #tag 4 get the time      
    elif any(variant in command for variant in ["time", "tim", "tme", "tyme", "tiem"]):
        try:
           time = datetime.datetime.now().strftime('%I %M  %p')
           print(time)
           talk('current time is' + time)
        except:
            pass
    #tag 5 send emails using outlook    
    elif any(variant in command for variant in ["send mail", "send email", "send a email","email","email send","sendmail", "sendemail", "send e-mail", "sendmail", "sendemail"]):
        try:
         subject = extract_subject(command)
         body = ""
         to = extract_email(command)
         send_email(subject, body, to)
         speak.say("the tab is open, please complete it ")
         speak.runAndWait()
        except:
            speak.say("please say sagain!")
            speak.runAndWait()
    #tag 6 calculate the fibonacci
    elif any(variant in command for variant in ["fibonacci", "bonacci","fibonaci", "fibbonaci", "fibonacchi", "fibanacci", "fibonacci sequence", "fibonacci numbers", "fibonacci series"]):
        try:
           talk('calculate the fibonacci please give us the value of n')
           # Taking input from the user
           n = int(input("Enter the value of n: "))
           result = fibonacci(n)
           print(result)
           speak.say('fibonacci of ' + str(n) + ' is ' + str(result))
           speak.runAndWait()
        except:
            pass
    elif any(command == variant for variant in ['close', 'stop', 'clse', 'lose','clos','cls']): 
        talk('closing')      
        running = False
    #there are no function to this tag or error         
    else:
        talk("please say again!")       
 
 
running = True  

while True:  
    run()
    if not running:  
        break



