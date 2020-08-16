# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 12:51:32 2019

@author: pk
"""

import xlrd 
import re
from commonregex import CommonRegex
import nltk
from nltk import ne_chunk, pos_tag, word_tokenize
from nltk.tree import Tree
import usaddress
import docxpy
import PyPDF2 
import csv
import glob
from nltk.corpus import stopwords
stop = stopwords.words('english')
import pytesseract
import pdf2image
import cv2
import pandas as pd
import tkinter as tk
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from sklearn.preprocessing import LabelEncoder,OneHotEncoder
from sklearn.compose import ColumnTransformer, make_column_transformer
import pandas as pd
from sklearn.metrics import classification_report
from sklearn.metrics import confusion_matrix
from sklearn.metrics import accuracy_score
#from sklearn.neighbors import KNeighborsClassifier
from sklearn.svm import SVC
main = tk.Tk()
pytesseract.pytesseract.tesseract_cmd ="C:/Users/pk/AppData/Local/Tesseract-OCR/tesseract.exe"
tessdata_dir_config = "C:/Users/pk/AppData/Local/Tesseract-OCR"
def extract_ssn(values):
    ssn=[]
    for item in values:
       word_item=str(item).split()
       for ssn_item in word_item:
         if bool(re.match(r'^(?!000|.+0{4})(?:\d{9}|\d{3}-\d{2}-\d{4})$', str(ssn_item))):
           ssn.append(ssn_item)
    return ssn 
def extract_creditcard(values):
    creditcard=[]
    for item in values:
       word_item=str(item).split()
       for creditcard_item in word_item:
         if bool(re.match(r'^(?!000|.+0{4})(?:\d{9}|\d{4}-\d{4}-\d{4}-\d{4})$', str(creditcard_item))):
          creditcard.append(creditcard_item)
    return creditcard 
def extract_phonenumbers(values):
    phonelist=[]
    perm_phone=[]
    for item in values:
       phonenumbers=CommonRegex(str(item))
       if phonenumbers.phones:
          phonelist.append(phonenumbers.phones)
    for i in phonelist:
        for j in i:
            perm_phone.append(j)
    return perm_phone
def extract_email(values):
    emaillist=[]
    perm_email=[]
    for item in values:
       Email=CommonRegex(str(item))
       if Email.emails:
          emaillist.append(Email.emails)
    for i in emaillist:
        for j in i:
            perm_email.append(j)
    return perm_email
def extract_dob(values):
    dob_match=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    out_dob=[]
    for item in values:
        word_item = str(item).split()
        for dob_item in word_item:
            for i,c in enumerate(dob_match):
                if str(c) in dob_item:
                    if i<10:
                     dob_item=dob_item.replace(str(c),'0'+str(i+1))
                     #print(dob_item)
                    else:
                         dob_item=dob_item.replace(str(c),str(i+1))
                         #print(dob_item)
            if bool(re.match(r'^(0[1-9]|1[012])[-/.](0[1-9]|[12][0-9]|3[01])[-/.](19|20)\d{2}|(0[1-9]|[12][0-9]|3[01])[-/.](0[1-9]|1[012])[-/.](19|20)\d{2}$',str(dob_item))):
                 out_dob.append(dob_item)
    return out_dob
def extract_passport(values):
     passport=[]
     for item in values:
       word_item=str(item).split()
       for pp_item in word_item:
           if(len(str(pp_item))>=6 and len(str(pp_item))<=9):
             if bool(re.match(r'^[A-Z]?[0-9]$', str(pp_item))):
                 passport.append(pp_item)
     return passport 
def extract_drivinglicense(values):
     drivinglicense=[]
     for item in values:
       word_item=str(item).split()
       for dl_item in word_item:
             if bool(re.match(r'^((A[LKZR])|(C[AOT])|(D[EC])|(FL)|(GA)|(HI)|(I[DLNA])|(K[SY])|(LA)|(M[EDAINSOT])|(N[EVHJMYCD])|(O[HKR])|(PA)|(RI)|(S[CD])|(T[NX])|(UT)|(V[TA])|(W[AVIY]))\d{7}$', str(dl_item))):
                 drivinglicense.append(dl_item)
     return drivinglicense 
def extract_address(values):
    address=[]
    for item in values:
       addr=""
       count=0
       #print(item)
       x=usaddress.parse(str(item))
       for i in x:
           if i[1]=='AddressNumber':
               addr+=i[0] +" "
               count+=1
           if i[1]=='StreetNamePreDirectional':
               addr+=i[0]+" "
               count+=1
           if i[1]=='StreetName':
               addr+=i[0]+" "
               count+=1
           if i[1]=='OccupancyType':
               addr+=i[0]+" "
               count+=1
           if i[1]=='OccupancyIdentifier':
               addr+=i[0]+" "
               count+=1
           if i[1]=='PlaceName':
               addr+=i[0]+" "
               count+=1
           if i[1]=='StateName':
               addr+=i[0]+" "
               count+=1
           if i[1]=='ZipCode':
               addr+=i[0]+" "
               count+=1
           if i[0]=='CountryName':
               addr+=i[0]+" "
               count+=1
       if count>4:
           address.append(addr)
    return address
def extract_username(values):
     username=[]
     for item in values:
       word_item=str(item).split()
       for uname_item in word_item:
             if bool(re.match(r'^[a-z0-9]*$', str(uname_item))):
                 username.append(uname_item)
     return username
def extract_password(values):
    password = []
    for item in values:
        word_item = str(item).split()
        for passname_item in word_item:
                if bool(re.match(r'^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[@$!%*?&])[A-Za-z\d@$!%*?&]{8,16}$', str(passname_item))):
                                    password.append(passname_item)
    return password
def ie_preprocess(document):
    document = ' '.join([i for i in document.split() if i not in stop])
    return (ne_chunk(pos_tag(word_tokenize(document))))
def extract_names(document):
    names = []
    sentences = ie_preprocess(document)
    for tagged_sentence in sentences:
      #print(tagged_sentence)
      if 'PERSON' in str(tagged_sentence):
          tagged_sentence=str(tagged_sentence).replace('PERSON','')
          tagged_sentence=str(tagged_sentence).replace('/NNP','')
          tagged_sentence=str(tagged_sentence).replace('(','')
          tagged_sentence=str(tagged_sentence).replace(')','')
          names.append(tagged_sentence)
    return names
def pdf_ssn(values):
    ssn_list=[]
    val=values
    for i in range(len(values)):
        ssn=""
        if val[i] =='-':
            if val[i-1].isdigit() and len(val[i-1])==3 and len(val[i+1])==2  and val[i+1].isdigit() and val[i+2]=='-' and val[i+3].isdigit() and len(val[i+3])==4:
                ssn+=val[i-1]+val[i]+val[i+1]+val[i+2]+val[i+3]
                ssn_list.append(ssn)   
    return ssn_list
def pdf_creditcard(values):
    creditcard_list=[]
    val=values
    for i in range(len(values)):
        cc=""
        if val[i] =='-':
            if val[i-1].isdigit() and len(val[i-1])==4 and len(val[i+1])==4  and val[i+1].isdigit() and val[i+2]=='-' and val[i+3].isdigit() and len(val[i+3])==4 and val[i+4]=='-' and val[i+5].isdigit() and len(val[i+5])==4:
                cc+=val[i-1]+val[i]+val[i+1]+val[i+2]+val[i+3]+val[i+4]+val[i+5]
                creditcard_list.append(cc)   
    return creditcard_list
def extract_gender(values):
    gender=['male','female']
    gender_list=[]
    for item in values:
       word_item=str(item).split()
       for gender_item in word_item:
           if gender_item.lower() in gender:
               gender_list.append(gender_item)
    return gender_list
               
def main_xlsx(file):
    loc = (file) 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 
    values=[]     
    for i in range(sheet.nrows): 
      for j in range(sheet.ncols):
        if sheet.cell_value(i,j)!='':
           values.append(sheet.cell_value(i, j))
    return values

def main_docx(file):
    doctext = docxpy.process(file)
    doc_item = str(doctext).split()
    return doc_item

def main_pdf(file):
    pdfFileObj = open(file, 'rb') 
 
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj) 
    pdf=[] 
    pages=int(pdfReader.numPages)
    for page in range(pages):
# creating a page object 
     pageObj = pdfReader.getPage(page) 
     po=pageObj.extractText().split()
     for i in po:
      if i!='':
         i = re.sub(r"[\n\t]*", "", i)
         pdf.append(i)
    pdfFileObj.close()
    return(pdf)
def main_image(file):
    text = pytesseract.image_to_string(file)
    text=text.split()
    print(text)
    return text
def main_pdf2image(file):
    images = pdf2image.convert_from_path(file,200)
    image_list=[]
    permanent_image=[]
    for img in images:
       text = pytesseract.image_to_string(img)
       text=text.split()
       image_list.append(text)
    for i in image_list:
        for j in i:
            permanent_image.append(j)
    return permanent_image    
    
def _main(format_input):
   #print("Enter the type of file you want to input:")
   #format_input=input()
 
    if  'xlsx' in format_input:
       return main_xlsx(format_input)
    elif 'docx' in format_input:
       return main_docx(format_input)
    elif 'pdf' in format_input:
       return main_pdf(format_input)
    elif 'jpg' in format_input:
       return main_image(format_input)
      
    else:
       return "Not Supported File"
  
if __name__== "__main__":
       mypath=r"C:\Users\pk\Desktop\PII\input file\PII- test-2.pdf"
       ssn_flag_list=[]
       credit_card_flag_list=[]
       phonenumbers_flag_list=[]
       email_flag_list=[]
       dob_flag_list=[]
       passport_flag_list=[]
       drivinglicense_flag_list=[]
       address_flag_list=[]
       password_flag_list=[]
       names_flag_list=[]
       PI_flag_list=[]
   
  
       PI_count=0
       jsonoutput={}
       #print(f) 
       output=_main(mypath)
       values=output
       #print(values)
       csv_list=[]
       #print("SSN:-")
       ssn=extract_ssn(values)
       jsonoutput['SSN']=ssn
       #print(ssn)
       if ssn==[]:
           ssn=pdf_ssn(values)
           jsonoutput['SSN']=ssn
           #print(ssn)
       #print("CREDIT-CARD:")
       credit_card=extract_creditcard(values)
       jsonoutput['CREDITCARD']=credit_card
       #print(credit_card)
       if credit_card==[]:
          credit_card=pdf_creditcard(values)
          jsonoutput['CREDITCARD']=credit_card
          #print(credit_card)
       
       phonenumbers=extract_phonenumbers(values)
       jsonoutput['PHONENUMBERS']=phonenumbers
       #print("PHONE-NUMBERS:")
       #print(phonenumbers)
       #print("E-Mail:")
       email=extract_email(values)
       jsonoutput['E-MAIL']=email
       #print(email)
       #print("DoB:")
       dob=extract_dob(values)
       jsonoutput['DOB']=dob
       #print(dob)
       #print("PASSPORT:")
       passport=extract_passport(values)
       jsonoutput['PASSPORT']=passport
       #print(passport)
       #print("DRIVING LICENSE:")
       drivinglicense=extract_drivinglicense(values)
       jsonoutput['DRIVINGLICENSE']=drivinglicense
       #print(drivinglicense)
       #print("ADDRESS:")
       address=extract_address(values)
       jsonoutput['ADDRESS']=address
       #print(address)
       username = extract_username(values)
       #jsonoutput['USERNAME']=username
       #print(username)
       #print("Password:")
       password=extract_password(values)
       jsonoutput['PASSWORD']=password
       #print(password)
       #print("NAME:")
       new_values=[]
       for i in values:
           new_values.append(str(i))
       text=''.join(new_values)
       names=extract_names(text)
       #print(names)
       jsonoutput['NAME']=names
       #print("GENDER:")
       gender=extract_gender(values)
       jsonoutput['GENDER']=gender
       #print(gender)
       print("PII-Information:\n")
       print(jsonoutput)
       
       ssn_flag = 0
       credit_card_flag = 0
       phonenumbers_flag = 0
       email_flag = 0
       dob_flag = 0
       passport_flag = 0
       drivinglicense_flag = 0
       address_flag = 0
       password_flag = 0
       names_flag = 0
       PI_Flag='Yellow'
       if len(ssn) > 0: 
           ssn_flag = 1
           PI_Flag='Red'
           PI_count+=len(ssn)
       ssn_flag_list.append(ssn_flag)
       if len(credit_card) > 0:
           credit_card_flag = 1
           PI_Flag='Red'
           PI_count+=len(credit_card)
       credit_card_flag_list.append(credit_card_flag)
       if len(phonenumbers) > 0: 
           phonenumbers_flag = 1
           PI_count+=len(phonenumbers)
       phonenumbers_flag_list.append(phonenumbers_flag)
       if len(email) > 0: 
           email_flag = 1
           PI_count+=len(email)
       email_flag_list.append(email_flag)
       if len(dob) > 0: 
           dob_flag = 1
           PI_count+=len(dob)
       dob_flag_list.append(dob_flag)
       if len(passport) > 0: 
           passport_flag = 1
           PI_count+=len(passport)
       passport_flag_list.append(passport_flag)
       if len(drivinglicense) > 0: 
           drivinglicense_flag = 1
           PI_count+=len(drivinglicense)
       drivinglicense_flag_list.append(drivinglicense_flag)
       if len(address) > 0:
           address_flag = 1
           PI_count+=len(address)
       address_flag_list.append(address_flag)
       if len(password) > 0: 
           password_flag = 1
           PI_count+=len(password)
       password_flag_list.append(password_flag)
       if len(names) > 0: 
           names_flag = 1
           #PI_Flag='Red'
           PI_count+=len(names)
       names_flag_list.append(names_flag)
       if PI_count==0:
           PI_Flag='Green'
       PI_flag_list.append(PI_Flag)
       data={'SSN':ssn_flag_list,
             'CREDIT_CARD':credit_card_flag_list,
             'PHONENUMBERS':phonenumbers_flag_list,
             'EMAIL':email_flag_list,
             'DOB':dob_flag_list,
             'PASSPORT':passport_flag_list,
             'DRIVINGLICENSE':drivinglicense_flag_list,
             'ADDRESS':address_flag_list,
             'PASSWORD':password_flag_list,
              'NAMES':names_flag_list,
              'FLAGS':PI_flag_list}
       df = pd.DataFrame(data)
       #print(df)
       print("PII-Count:\n")
       print(PI_count)
       
       df.to_csv('out1.csv')
       
      
       ourMessage='Number of PIIs Identified :' +" "+str(PI_count)
       messageVar = tk.Label(main, text = ourMessage)
       messageVar.config(bg = 'cyan',font=('helvetica', 16),width = 50, height = 25)
       messageVar.pack()
       main.mainloop()
         
    

   
       
               
       