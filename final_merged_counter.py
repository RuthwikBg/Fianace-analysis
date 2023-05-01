import re
from nltk import tokenize
import pandas as pd
# nltk.download('punkt')
import glob
import os
import docxpy

# initialize lists to store values
final_list_files = []
final_list_sentences_count = []
final_list_date_count = []
final_list_word_count = []
numbers = []

# contains path of folder containing docx files
path = "add folder path here"


# counts dates
def date_counter(text):
    # regEx1 is the regular expression to detect dates
    regEx1 = '\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|(?:\d{1,2}[-/th|st|nd|rd\s]*)?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z\s,.]*(?:\d{1,2}[-/th|st|nd|rd)\s,]*)?(?:\d{2,4})|\d{4} to present|prior to \d{4}|in \d{4}|In \d{4}|until \d{4}|for \d{4}|of \d{4}|from the \d{4}|from \d{4}'
    dateList1 = re.findall(regEx1, text)
    final_list_date_count.append(len(dateList1))


# counts numbers(both words and digits)
def number_counter(text):
    # regex is regular expression to detect numbers in digits
    regex = r'\$\s*[\d+,\d+]+[\.[0-9]+]?|$[0-9]|\b(?:\d+\.\d+|\d+)\b'
    # numbers_in_words is regular expression to detect number in words
    numbers_in_words = r'(?i)\b(twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand)\s*(one|two|three|four|five|six|seven|eight|nine)?\b|\b(zero|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen)\b'
    nlist = re.findall(regex, text)  # findall creates a list of all values that matches regex
    wlist = []
    matches = re.findall(numbers_in_words, text)
    for i in matches:
        x = list(i)
        w = " ".join(x).strip()
        wlist.append(w)
    f = nlist + wlist
    numbers.append(len(f))  # adds total number count to numbers[]


# counts the number of words in text
def word_counter(text):
    l = text.split()
    final_list_word_count.append(len(l))  # adds total word count to final_list_word_count[]


# removes unnecessary period and return clean text
def cleaner(txt):
    txt = txt.replace('etc.', 'etc')
    txt = txt.replace('approx.', 'approx')
    txt = txt.replace('mins.', 'mins')
    return txt


# counts sentences containing 3 or more words
def sentence_counter(l):
    for i in l:
        if len(i.split(" ")) < 3:
            l.remove(i)
    final_list_sentences_count.append(len(l))  # adds total sentences count to final_list_sentences_count[]


# extracts text from docx files
def getText(file):
    text = docxpy.process(file)
    return text


try:
    for i, doc in enumerate(glob.iglob(path + "*.docx")):
        try:
            filename = doc.split('\\')[-1]
            path1 = os.path.abspath(doc)
            txt = getText(path1)
            ftxt = cleaner(txt)
            l = tokenize.sent_tokenize(ftxt)
            sentence_counter(l)
            date_counter(ftxt)
            word_counter(ftxt)
            number_counter(ftxt)
            final_list_files.append(filename)
        except:
            print("error "+ filename)
finally:
    pass

# stores all lists in dataframe
df = pd.DataFrame.from_dict({'Filename':final_list_files,'Sentence count':final_list_sentences_count,'Dates count':final_list_date_count,'Word count':final_list_word_count,'Number count':numbers})
# creates excel file and adds the data in data frame to the file
df.to_excel( path , header=True, index=False)
