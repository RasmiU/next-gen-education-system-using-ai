import fitz
import nltk
import os
import re
import math
import operator
import xlrd
from nltk.stem import WordNetLemmatizer
from nltk.corpus import stopwords
from nltk.tokenize import sent_tokenize,word_tokenize
nltk.download('averaged_perceptron_tagger')
Stopwords = set(stopwords.words('english'))
wordlemmatizer = WordNetLemmatizer()

fileout = open('course_content.txt',"a")
loc = ("Final topics .xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
n_rows=sheet.nrows
doc = fitz.open('copy.pdf') #copy.pdf is the textbook
#for i in range(1,n_rows+1):
for i in range(1,5):
    page_no=int(sheet.cell_value(i,2)) #page to be read
    page = doc.loadPage(page_no) 
    text = page.getText('text')

    search_string=sheet.cell_value(i,0) #heading to be searched for
    next_string=sheet.cell_value(i,3) #the heading that'll come after search_string
    outF = open('extract_page.txt',"w")
    if(text.find(search_string) == -1): #search_string not found in page_no
        print("couldn't find"+search_string+" !")
    else:
        start=text.find(search_string) 
        length=len(search_string)
        next=start+length
        if(text.find(next_string) != -1): #case where search_string and next_string are in the same page
            start_next=text.find(next_string)
            text_input=text[next:start_next] #extracting paragraph between search_string and next_string
            #writing the paragraph to a file
            outF.write(search_string)
            outF.write("\n")
            outF.write(text_input)
        else: #if the two headings aren't in the same page
            text_input=text[next:] #extract all text coming after search_string
            outF.write(search_string)
            outF.write("\n")
            outF.write(text_input)
            flag=0 #flag=0 => next_string not found yet
            while(flag==0): #will keep on iterating until next_string is found in some page
                page_no=page_no+1
                page = doc.loadPage(page_no) #reading next page of page_no
                text = page.getText('text')
                temp=text.find(next_string)
                if(temp != -1):
                    text_input=text[:temp]
                    flag=1
                else:
                    text_input=text
                outF.write(text_input)
    outF.close()
    #print(text_input)
    f = open('input_text.txt',"w")
    outF=open('extract_page.txt', 'r')
    line=outF.readline()
    while line:
        cleanedLine = line.strip()
        if cleanedLine: # is not empty
            if(cleanedLine[-1]=='-'): 
                cleanedLine=cleanedLine[:-1]
                f.write(cleanedLine)
            else:
                f.write(cleanedLine)
                f.write("   ")
        line=outF.readline()
    '''with open('extract_page.txt', 'r') as outf:
        for line in outf:
                cleanedLine = line.strip()
                if cleanedLine:
                    if(cleanedLine[-1]=='-'): # is not empty
                        cleanedline=cleanedLine[:-1]
                        f.write(cleanedLine)
                    else:
                        f.write(cleanedLine)
                        f.write("   ")'''
    outF.close()
    f.close()

    def lemmatize_words(words):
        lemmatized_words = []
        for word in words:
            lemmatized_words.append(wordlemmatizer.lemmatize(word))
        return lemmatized_words
    def stem_words(words):
        stemmed_words = []
        for word in words:
            stemmed_words.append(stemmer.stem(word))
        return stemmed_words
    def remove_special_characters(text):
        regex = r'[^a-zA-Z0-9\s]'
        text = re.sub(regex,'',text)
        return text
    def freq(words):
        words = [word.lower() for word in words]
        dict_freq = {}
        words_unique = []
        for word in words:
            if word not in words_unique:
                words_unique.append(word)
        for word in words_unique:
            dict_freq[word] = words.count(word)
        return dict_freq
    def pos_tagging(text):
        pos_tag = nltk.pos_tag(text.split())
        pos_tagged_noun_verb = []
        for word,tag in pos_tag:
            if tag == "NN" or tag == "NNP" or tag == "NNS" or tag == "VB" or tag == "VBD" or tag == "VBG" or tag == "VBN" or tag == "VBP" or tag == "VBZ":
                pos_tagged_noun_verb.append(word)
        return pos_tagged_noun_verb
    def tf_score(word,sentence):
        freq_sum = 0
        word_frequency_in_sentence = 0
        len_sentence = len(sentence)
        for word_in_sentence in sentence.split():
            if word == word_in_sentence:
                word_frequency_in_sentence = word_frequency_in_sentence + 1
        tf =  word_frequency_in_sentence/ len_sentence
        return tf
    def idf_score(no_of_sentences,word,sentences):
        no_of_sentence_containing_word = 0
        for sentence in sentences:
            sentence = remove_special_characters(str(sentence))
            sentence = re.sub(r'\d+', '', sentence)
            sentence = sentence.split()
            sentence = [word for word in sentence if word.lower() not in Stopwords and len(word)>1]
            sentence = [word.lower() for word in sentence]
            sentence = [wordlemmatizer.lemmatize(word) for word in sentence]
            if word in sentence:
                no_of_sentence_containing_word = no_of_sentence_containing_word + 1
        idf = math.log10(no_of_sentences/no_of_sentence_containing_word)
        return idf
    def tf_idf_score(tf,idf):
        return tf*idf
    def word_tfidf(dict_freq,word,sentences,sentence):
        word_tfidf = []
        tf = tf_score(word,sentence)
        idf = idf_score(len(sentences),word,sentences)
        tf_idf = tf_idf_score(tf,idf)
        return tf_idf
    def sentence_importance(sentence,dict_freq,sentences):
        sentence_score = 0
        sentence = remove_special_characters(str(sentence)) 
        sentence = re.sub(r'\d+', '', sentence)
        pos_tagged_sentence = [] 
        no_of_sentences = len(sentences)
        pos_tagged_sentence = pos_tagging(sentence)
        for word in pos_tagged_sentence:
            if word.lower() not in Stopwords and word not in Stopwords and len(word)>1: 
                word = word.lower()
                word = wordlemmatizer.lemmatize(word)
                sentence_score = sentence_score + word_tfidf(dict_freq,word,sentences,sentence)
        return sentence_score
    file = 'input_text.txt'
    file = open(file , 'r')
    text = file.read()
    tokenized_sentence = sent_tokenize(text)
    text = remove_special_characters(str(text))
    text = re.sub(r'\d+', '', text)
    tokenized_words_with_stopwords = word_tokenize(text)
    tokenized_words = [word for word in tokenized_words_with_stopwords if word not in Stopwords]
    tokenized_words = [word for word in tokenized_words if len(word) > 1]
    tokenized_words = [word.lower() for word in tokenized_words]
    tokenized_words = lemmatize_words(tokenized_words)
    word_freq = freq(tokenized_words)
    input_user = 50
    no_of_sentences = int((input_user * len(tokenized_sentence))/100)
    c = 1
    sentence_with_importance = {}
    for sent in tokenized_sentence:
        sentenceimp = sentence_importance(sent,word_freq,tokenized_sentence)
        sentence_with_importance[c] = sentenceimp
        c = c+1
    sentence_with_importance = sorted(sentence_with_importance.items(), key=operator.itemgetter(1),reverse=True)
    cnt = 0
    summary = []
    sentence_no = []
    for word_prob in sentence_with_importance:
        if cnt < no_of_sentences:
            sentence_no.append(word_prob[0])
            cnt = cnt+1
        else:
            break
    sentence_no.sort()
    cnt = 1
    for sentence in tokenized_sentence:
        if cnt in sentence_no:
            summary.append(sentence)
        cnt = cnt+1
    summary = " ".join(summary)
    #removing junk values from summary
    summary = re.sub("[0-9]+/[0-9]+/[0-9]+", "", summary)
    summary = re.sub("[0-9]+:[0-9]+ [A-Z]+", "", summary)
    summary = re.sub("Page [0-9]+", "", summary)
    summary = re.sub("pre[0-9]+_ch[0-9]+.qxd  ", "", summary)

    fileout.write(search_string)
    fileout.write("\n")
    fileout.write(summary)
    fileout.write("\n \n")
    


