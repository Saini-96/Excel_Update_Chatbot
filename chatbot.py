import nltk
import warnings
from time import ctime
import time
import os, sys
warnings.filterwarnings("ignore")
import Excel_updation
import getpass


# nltk.download() # for downloading packages

import numpy as np
import random
import string # to process standard python strings


f=open('chatbot.txt','r',errors = 'ignore')
raw=f.read()
raw=raw.lower()# converts to lowercase
nltk.download('punkt') # first-time use only
nltk.download('wordnet') # first-time use only
sent_tokens = nltk.sent_tokenize(raw)# converts to list of sentences 
word_tokens = nltk.word_tokenize(raw)# converts to list of words


sent_tokens[:2]


word_tokens[:5]


lemmer = nltk.stem.WordNetLemmatizer()
def LemTokens(tokens):
    return [lemmer.lemmatize(token) for token in tokens]
remove_punct_dict = dict((ord(punct), None) for punct in string.punctuation)
def LemNormalize(text):
    return LemTokens(nltk.word_tokenize(text.lower().translate(remove_punct_dict)))


GREETING_INPUTS = ("hello", "hi", "greetings", "sup", "what's up","hey",)
GREETING_RESPONSES = ["hi", "hey", "*nods*", "hi there", "hello", "I am glad! You are talking to me"]


# Checking for greetings
def greeting(sentence):
    """If user's input is a greeting, return a greeting response"""
    for word in sentence.split():
        if word.lower() in GREETING_INPUTS:
            return random.choice(GREETING_RESPONSES)
			
			
			
def Chatbot(data):

    data = data.lower()

    if "quit" not in data:
        
        if "how can you help me" in data:
            print("Chatbot: You can update your weekly work status by giving me inputs in format 'status update' in lowercase....")

        elif "status update" in data:
            print("Please enter you name: ")
            user_name = input()
            print("Please enter Excel file name: ")
            Spreadsheet_name = input()
            print("Please enter your work updates in format {'Release': '12/01/2010','User Story/Work Description': 'US00002:Coding','Appl': 'DEV','Analysis': '100%','Doc': '50%','Coding': '50%','Unit Testing': '30%', 'Planned_Del': '12/04/2010', 'Actual_Del': '12/06/2010', 'Testing Support': 'Yes', 'Planned_Imp': '12/15/2010', 'Actual_Imp': '12/15/2010', 'Remarks': 'Coding in Progress'}:")
            user_update = input()
            str_1 = str(user_update)
            update_str = eval(str_1)
            Excel_updation.Excel_update(user_name, Spreadsheet_name, update_str)

        elif "what time is it" in data:
            print("Chatbot: ", ctime())

        elif "help" in data:
            print("We can serve you by: \n\thow can you help me \n\tWhat time is it \n\tstatus update\n")

        elif(data=='thanks' or data=='thank you' ):
            print("Chatbot: You are welcome..")

        elif(greeting(data)!=None):
            print("Chatbot: "+greeting(data))

        else:
            print("Chatbot: I am sorry! I don't understand you")
            
    else:
        print("Chatbot: bye! take care...")
        sys.exit(0)

print("Chatbot: Hey Folk! what can I do for you?. \n If you want to exit, type quit!\n for help type help\n")


# initialization
time.sleep(2)
while 1:
    data = input(getpass.getuser()+": ")
    Chatbot(data)