from lxml import etree
import zipfile
import os
from textblob import TextBlob
import pandas as pd
from textblob.sentiments import NaiveBayesAnalyzer
from tkinter import filedialog

# Namespace
namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
# Full data frame
df = pd.DataFrame(columns=['Title', 'Author', 'Comment', 'Classification', 'Positivity', 'Negativity', 'Polarity', 'Subjectivity'])
# Boolean for short analysis (drops extra columns)
short = False
# These values have no meaning and also confuse the analysis
skip = ['_Marked as resolved_', '_Re-opened_']

# Right now the long way takes FOREVER. Trying to fix but it seems like it's the NaiveBayesAnalyzer
def long_comments(docxFileName):
    total_comments = 0
    docx = 1 # Debug iterator
    try:
        docxZip = zipfile.ZipFile(directory + "/" + docxFileName)
        commentsXML = docxZip.read('word/comments.xml')
        etr = etree.XML(commentsXML)
        comments = etr.xpath('//w:comment', namespaces=namespace)
        print("Read docx" + docx) # Debug output
        for c in comments:
            # attributes
            title = docxFileName
            author = str(c.xpath('@w:author', namespaces=namespace))
            commentstring = c.xpath('string(.)', namespaces=namespace)
            author = author.replace('[', '').replace(']', '').replace("'", '')  # Formatting
            if commentstring not in skip:
                print("Analyzing comment" + total_comments + 1) # Debug line
                # Polarity and subjectivity from default analyzer
                sentimentstring2 = TextBlob(commentstring)
                polarity = sentimentstring2.sentiment.polarity
                subjectivity = sentimentstring2.sentiment.subjectivity
                # Classification, positivity, and negativity from NaiveBayesAnalyzer
                sentimentstring = TextBlob(commentstring, analyzer=NaiveBayesAnalyzer())
                classification = sentimentstring.sentiment.classification
                positivity = sentimentstring.sentiment.p_pos
                negativity = sentimentstring.sentiment.p_neg
                print("Posting to dataframe") # Debug line
                new_row = [title, author, commentstring, classification, positivity, negativity, polarity, subjectivity]
            j = len(df)
            df.loc[j] = new_row
            print("No Comment found for file")
            total_comments += 1
    except KeyError:
        pass

def short_comments(docxFileName):
    total_comments = 0
    try:
        docxZip = zipfile.ZipFile(directory + "/" + docxFileName)
        commentsXML = docxZip.read('word/comments.xml')
        etr = etree.XML(commentsXML)
        comments = etr.xpath('//w:comment', namespaces=namespace)
        for c in comments:
            title = docxFileName
            author = str(c.xpath('@w:author', namespaces=namespace))
            commentstring = c.xpath('string(.)', namespaces=namespace)
            author = author.replace('[', '').replace(']', '').replace("'", '')  # Formatting
            if commentstring not in skip:
                sentimentstring = TextBlob(commentstring)
                polarity = sentimentstring.sentiment.polarity
                subjectivity = sentimentstring.sentiment.subjectivity
                # Posts null values to df to prevent errors
                classification = ""
                positivity = ""
                negativity = ""

                new_row = [title, author, commentstring, classification, positivity, negativity, polarity, subjectivity]
            a = len(df)
            df.loc[a] = new_row
            total_comments += 1
    except KeyError:
        print("No Comment found for file")

# Asks for directory
directory = filedialog.askdirectory()
# Prevents user error (on terminal until I can learn tkinter)
while directory=='':
    print("Close program? (Y/N)")
    userinput = input("Enter here: ")
    if userinput in ["Y", "y", "yes", "YES"]:
        exit()
    else:
        directory = filedialog.askdirectory()

# Asks if long or short (on terminal until I can learn tkinter)
print("Long analysis? (Y/N)")
userinput = input("Enter here: ")

if userinput in ["Y","y","yes","YES"]:
    for filename in os.listdir(directory):
        if filename.endswith(".docx"):
            long_comments(filename)
else:
    short = True
    for filename in os.listdir(directory):
        if filename.endswith(".docx"):
            short_comments(filename)

# Null value cleaner
if short:
    del df['Classification']
    del df['Positivity']
    del df['Negativity']

# Make this in tkinter
outputname = input("Enter file name: ")

# Creates csv with user-input file name
df.to_csv(outputname+'.csv')
print('Saved successfully!')