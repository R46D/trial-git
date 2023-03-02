#Required libraries to download if not available

'''!pip3 install pandas
   !pip3 install selenium
   !pip3 install nltk
   !pip3 install webdriver_manager
   !pip3 install pyhyphen
   nltk.download('punkt')'''


#Importing the necessary Packages
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import nltk
from hyphen import Hyphenator
from nltk.tokenize import sent_tokenize
from nltk.tokenize import word_tokenize
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import re



#Using webdriver for chrome
driver = webdriver.Chrome(ChromeDriverManager().install())   

#reading input file
file=pd.read_excel('Input.xlsx')

POSITIVE_SCORE=[]
NEGATIVE_SCORE=[]
POLARITY_SCORE=[]
SUBJECTIVITY_SCORE=[]
AVG_SENTENCE_LENGTH=[]
PERCENTAGE_OF_COMPLEX_WORDS=[]
FOG_INDEX=[]
AVG_NUMBER_OF_WORDS_PER_SENTENCE=[]
COMPLEX_WORD_COUNT=[]
WORD_COUNT=[]
SYLLABLE_PER_WORD=[]
PERSONAL_PRONOUNS=[]
AVG_WORD_LENGTH=[]

#opening and reading the Master Dictionary and Stop Words text files
positivewords=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/MasterDictionary/positive-words.txt","r")
negativewords=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/MasterDictionary/negative-words.txt","r")
stopwordsa=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_Auditor.txt","r")
stopwordsb=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_Currencies.txt","r")
stopwordsc=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_DatesandNumbers.txt","r")
stopwordsd=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_Generic.txt","r")
stopwordse=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_GenericLong.txt","r")
stopwordsf=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_Geographic.txt","r")
stopwordsg=open("C:/Users/Randolph/Desktop/Work/Sentimental Analysis/StopWords/StopWords_Names.txt","r")

positive_words=positivewords.read()
negative_words=negativewords.read()
stopwords_a=stopwordsa.read()
stopwords_b=stopwordsb.read()
stopwords_c=stopwordsc.read()
stopwords_d=stopwordsd.read()
stopwords_e=stopwordse.read()
stopwords_f=stopwordsf.read()
stopwords_g=stopwordsg.read()

#Making all the words lowercase in order to correctly compare them
positive_words=positive_words.lower()
negative_words=negative_words.lower()
stopwords_a=stopwords_a.lower()
stopwords_b=stopwords_b.lower()
stopwords_c=stopwords_c.lower()
stopwords_d=stopwords_d.lower()
stopwords_e=stopwords_e.lower()
stopwords_f=stopwords_f.lower()
stopwords_g=stopwords_g.lower()

#Cleaning the text files to be usable

positive_words=positive_words.split("\n")
negative_words=negative_words.split('\n')
stopwords_a=stopwords_a.split("\n")
stopwords_b=stopwords_b.replace("|","\n")
stopwords_b=stopwords_b.split("\n",)
stopwords_c=stopwords_c.split("\n")
stopwords_d=stopwords_d.split("\n")
stopwords_e=stopwords_e.split("\n")
stopwords_f=stopwords_f.split("\n")
stopwords_g=stopwords_g.split("\n")

#removing trailing spaces in stopwords b
for i in range(len(stopwords_b)):
    stopwords_b[i]=stopwords_b[i].strip()

#removing unncessary text from stopwords c,f and g
for i in range(len(stopwords_c)):
    stopwords_c[i] = stopwords_c[i].split("|", 1)[0]
    stopwords_c[i] = stopwords_c[i].strip()

for i in range(len(stopwords_f)):
    stopwords_f[i] = stopwords_f[i].split("|", 1)[0]
    stopwords_f[i] = stopwords_f[i].strip()
    
for i in range(len(stopwords_g)):
    stopwords_g[i] = stopwords_g[i].split("|", 1)[0]
    stopwords_g[i] = stopwords_g[i].strip()

#merging the stopwords into one list

stopwords=stopwords_a+stopwords_b+stopwords_c+stopwords_d+stopwords_e+stopwords_f+stopwords_g


#Reading input file

for k in range(len(file['URL'])):
    url = file['URL'][k]

#opening the url
    driver.get(url)

#Extracting text from article
    try:
        texts = driver.find_element(By.CLASS_NAME,'td-post-content').text
        
#removing the unncessary last line from the extracted text
    
        texts=texts[:texts.rfind('\n')]
        
#counting the number of sentences 

        number_of_sentences = sent_tokenize(texts)

        Nsentences=len(number_of_sentences)

#Removing punctuations from extracted text 
        punctuations = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''
        cleaned_texts= ""

        for x in texts:
            if x not in punctuations:
                cleaned_texts += x

#Counting pronouns
        pronounRegex = re.compile(r'\b(I|we|my|ours|(?-i:us))\b',re.I)
        pronouns = pronounRegex.findall(cleaned_texts)
        TP=len(pronouns)

#seperating the cleaned_texts into words
    
        cleaned_texts=word_tokenize(cleaned_texts)



#Removing stopwords from the cleaned text
    
        filtered_texts=[]
        for i in range(len(cleaned_texts)):
            if cleaned_texts[i].casefold() not in stopwords:
                filtered_texts.append(cleaned_texts[i])

       
        
#Calculating new word count
        TWC=len(filtered_texts)

#Calculating Positive, Negative, Polarity and Subjectivity Score

        PS=0
        for i in range(len(filtered_texts)):
            if filtered_texts[i].casefold() in positive_words:
                PS=PS+1
        NS=0
        for i in range(len(filtered_texts)):
            if filtered_texts[i].casefold() in negative_words:
                NS=NS+1

        PLS = (PS-NS)/ ((PS+NS) + 0.000001)

        SS= (PS + NS)/ ((TWC) + 0.000001)

#calculating average sentence length 
        ASL= TWC/Nsentences

#calculating word per sentence  
        AWPS=TWC/Nsentences

#calculating syllable per word and complex words using pyhyphen package
        CP=0
        SPW=0
        h_en = Hyphenator('en_US')
        for i in range(len(filtered_texts)):
            g=h_en.syllables(filtered_texts[i])
            SPW=SPW+len(g)
            if len(g)>2:
                CP=CP+1
            del g

#calculating percentage of complex words
        PCW = CP/TWC

#Calculating Fog Index
        FI=0.4*(ASL+PCW)

#Calculating the number of characters
        cleaned_texts_characters=''
        for i in range(0,len(cleaned_texts)):
            cleaned_texts_characters = cleaned_texts_characters+cleaned_texts[i]

        NC = len(cleaned_texts_characters)

#Calculating average word length
        AWL=NC/TWC

#Assigning Values    

        POSITIVE_SCORE.append(PS)
        NEGATIVE_SCORE.append(NS)
        POLARITY_SCORE.append(PLS)
        SUBJECTIVITY_SCORE.append(SS)
        AVG_SENTENCE_LENGTH.append(ASL)
        PERCENTAGE_OF_COMPLEX_WORDS.append(PCW)
        FOG_INDEX.append(FI)
        AVG_NUMBER_OF_WORDS_PER_SENTENCE.append(AWPS)
        COMPLEX_WORD_COUNT.append(CP)
        WORD_COUNT.append(TWC)
        SYLLABLE_PER_WORD.append(SPW)
        PERSONAL_PRONOUNS.append(TP)
        AVG_WORD_LENGTH.append(AWL)
        
    except (Exception,):
        print('Error with',url)
        POSITIVE_SCORE.append(0)
        NEGATIVE_SCORE.append(0)
        POLARITY_SCORE.append(0)
        SUBJECTIVITY_SCORE.append(0)
        AVG_SENTENCE_LENGTH.append(0)
        PERCENTAGE_OF_COMPLEX_WORDS.append(0)
        FOG_INDEX.append(0)
        AVG_NUMBER_OF_WORDS_PER_SENTENCE.append(0)
        COMPLEX_WORD_COUNT.append(0)
        WORD_COUNT.append(0)
        SYLLABLE_PER_WORD.append(0)
        PERSONAL_PRONOUNS.append(0)
        AVG_WORD_LENGTH.append(0)
        
    
driver.quit()
file['POSITIVE SCORE']=   POSITIVE_SCORE
file['NEGATIVE SCORE'] =  NEGATIVE_SCORE
file['POLARITY SCORE'] =  POLARITY_SCORE
file['SUBJECTIVITY SCORE'] =  SUBJECTIVITY_SCORE
file['AVG SENTENCE LENGTH'] =  AVG_SENTENCE_LENGTH
file['PERCENTAGE OF COMPLEX WORDS'] =  PERCENTAGE_OF_COMPLEX_WORDS
file['FOG INDEX'] =  FOG_INDEX
file['AVG NUMBER OF WORDS PER SENTENCE'] =  AVG_NUMBER_OF_WORDS_PER_SENTENCE
file['COMPLEX WORD COUNT'] =  COMPLEX_WORD_COUNT
file['WORD COUNT'] =  WORD_COUNT
file['SYLLABLE PER WORD'] =  SYLLABLE_PER_WORD
file['PERSONAL PRONOUNS'] =  PERSONAL_PRONOUNS
file['AVG WORD LENGTH'] =  AVG_WORD_LENGTH

file.to_excel('Output.xlsx',index=False)
file.to_csv('Output.csv',index=False)