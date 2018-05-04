from xlrd import open_workbook
from textblob import TextBlob
import re
import datetime
import json

from sklearn import linear_model
reg = linear_model.BayesianRidge()

from sklearn.externals import joblib

from sklearn import tree

from sklearn.naive_bayes import GaussianNB
gnb = GaussianNB()


emoticons_str = r"""
    (?:
        [:=;] # Eyes
        [oO\-]? # Nose (optional)
        [D\)\]\(\]/\\OpP] # Mouth
    )"""
 
regex_str = [
    emoticons_str,
    r'<[^>]+>', # HTML tags
    r'(?:@[\w_]+)', # @-mentions
    r"(?:\#+[\w_]+[\w\'_\-]*[\w_]+)", # hash-tags
    r'http[s]?://(?:[a-z]|[0-9]|[$-_@.&amp;+]|[!*\(\),]|(?:%[0-9a-f][0-9a-f]))+', # URLs
 
    r'(?:(?:\d+,?)+(?:\.?\d+)?)', # numbers
    r"(?:[a-z][a-z'\-_]+[a-z])", # words with - and '
    r'(?:[\w_]+)', # other words
    r'(?:\S)' # anything else
]
    
tokens_re = re.compile(r'('+'|'.join(regex_str)+')', re.VERBOSE | re.IGNORECASE)
emoticon_re = re.compile(r'^'+emoticons_str+'$', re.VERBOSE | re.IGNORECASE)

 
def tokenize(s):
    return tokens_re.findall(s)

 
def preprocess(s, lowercase=False):
    tokens = tokenize(s)
    if lowercase:
        tokens = [token if emoticon_re.search(token) else token.lower() for token in tokens]
    return tokens

def fridaySentiment():
    wb = open_workbook('friday.xlsx')
    sentiment_dict = {}
    sentiment_avg = 0
    count=0

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        rows = []
        values = []

        for row in range(1, number_of_rows):
            value  = (sheet.cell(row,7).value)
            try:
                value = str(value)
            except ValueError:
                pass
            finally:
                values.append(value)

        for i in range(0,len(values)):
            blob = TextBlob(values[i])
            sentiment = blob.sentiment.polarity
            sentiment_avg += sentiment
            count += 1
        sentiment_avg /= count
    return sentiment_avg

def ThursdaySentiment():
    wb = open_workbook('thur.xlsx')
    sentiment_dict = {}
    sentiment_avg = 0
    count=0

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        rows = []
        values = []

        for row in range(1, number_of_rows):
            value  = (sheet.cell(row,7).value)
            try:
                value = str(value)
            except ValueError:
                pass
            finally:
                values.append(value)

        for i in range(0,len(values)):
            blob = TextBlob(values[i])
            sentiment = blob.sentiment.polarity
            sentiment_avg += sentiment
            count += 1
        sentiment_avg /= count
    return sentiment_avg

def main():
	# load the trained model
	clf = joblib.load('dmpa_lab_project.pkl')
	print '[0] - Stock price will go down; [1] - Stock price will go up'	

	thurs_sentiment = ThursdaySentiment()
	print 'Thursday sentiment value:', thurs_sentiment
	print 'Model prediction for Thursday', clf.predict([[thurs_sentiment]])

	friday_sentiment = fridaySentiment()
	print 'Friday sentiment value:', friday_sentiment
	print 'Model Prediction for Friday', clf.predict([[friday_sentiment]])


main() # call of the main function