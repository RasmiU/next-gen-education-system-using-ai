from textblob import TextBlob
import nltk
import re
import string,sys,os
from rake_nltk import Rake
from nltk.tokenize import word_tokenize,sent_tokenize
import io
from io import StringIO
import sys
import random
from string import punctuation
from nltk.util import ngrams
from nltk.corpus import stopwords
import xlrd

loc = ("Final topics .xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
n_rows=sheet.nrows

filename='course_content.txt'
f=open(filename,'r')
outF = open('questions.txt',"a")


#questions=[]
#answers=[]


def find_key(sent,tagged_list,choices):
    flag=0
    for tag in tagged_list:
        if(tag[1]=='NNP'):
            key_word=tag[0]
            flag=1
            break
    if(flag==0):
        for tag in tagged_list:
            if(tag[1]=='NNPS'):
                key_word=tag[0]
                flag=1
                break
    if(flag==0):
        for tag in tagged_list:
            if(tag[1]=='NN'):
                key_word=tag[0]
                flag=1
                break
    if(flag==0):
        for tag in tagged_list:
            if(tag[1]=='NNS'):
                key_word=tag[0]
                flag=1
                break
    #tagged_list_copy=tagged_list
    if(flag==1):
        display(sent,key_word,choices)

def display(qtn,ans,choices):
    blank='________'
    qtn = re.sub(ans, blank, qtn, 1, flags=re.IGNORECASE)
    #qtn=str.replace(qtn,ans,blank)
    mc=[]
    disp_mc=[]

    len_ch=len(choices)
    temp_choices=['software','testing','code','agile']
    if len_ch < 4:
        len_diff = 4 - len_ch
        for a in range(0, len_diff-1):
            choices.append(temp_choices[a])

    mc=random.sample(choices, 3)
    mc.append(ans)
    disp_mc=random.sample(mc, 4)
    
    #questions.append(qtn)
    #answers.append(ans)
   
    outF.write("Q:")
    outF.write(qtn)
    outF.write("\n")
    outF.write("Options:")
    outF.write(str(disp_mc))
    outF.write("\nAns:")
    outF.write(ans)
    outF.write("\n\n")
    
    #outF.close()


#read summary into text and tokenize text into sentences in collection
#for i in range(1,n_rows+1):

for i in range(1,14):
    text=''
    str1=sheet.cell_value(i,0)
    try:
        str2=sheet.cell_value(i+1,0)
    except IndexError:
       str2=sheet.cell_value(i,3) 
    outF.write(str1)
    outF.write(":\n")

    line=f.readline()
    line=line.strip()
    h=0
    while line:
        if(line==str1):
            h=1
            line=f.readline()
            line=line.strip()
        if(line==str2):
            h=0
            break
        if(h==1):
            text=text+line
        line=f.readline()
        line=line.strip()

    collection=sent_tokenize(text)
    choices=[]
    #find noun keywords from text
    r = Rake(min_length=1, max_length=1) 
    r.extract_keywords_from_text(text)
    text_keys=r.get_ranked_phrases()
    text_keys_tagged=nltk.pos_tag(text_keys)
    for tag_key in text_keys_tagged:
        if(tag_key[1]=='NNP'):
            choices.append(tag_key[0])
        #if(tag_key[1]=='NNPS'):
        #   choices.append(tag_key[0])
        if(tag_key[1]=='NN'):
            choices.append(tag_key[0])
        #if(tag_key[1]=='NNS'):
        #    choices.append(tag_key[0])

    #find the relevant keywords from each sentence
    r = Rake(min_length=1, max_length=1) 
    stoplist = set(stopwords.words('english') + list(punctuation))

    for collec in collection:
        r.extract_keywords_from_text(collec)
        if(r.get_ranked_phrases()):
            phrase_list=r.get_ranked_phrases()
            for pl in phrase_list:
                if(pl in stoplist):
                    phrase_list.remove(pl)
            for pl in phrase_list:
                if(pl=='’'):
                    phrase_list.remove(pl)
            for pl in phrase_list:
                if(len(pl)==1):
                    phrase_list.remove(pl)
            #print(collec,"\t",phrase_list,"\n")
            tagged=nltk.pos_tag(phrase_list)
            #print(tagged,"\n")
            res = sum([i.strip(string.punctuation).isalpha() for i in collec.split()])
            if res<= 40 :
                find_key(collec,tagged,choices)
        
