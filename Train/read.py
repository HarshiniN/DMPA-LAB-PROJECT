from xlrd import open_workbook
from textblob import TextBlob
import re
import datetime
import json

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


def getSentiment():
    wb = open_workbook('sample.xlsx')
    sentiment_dict = {}

    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

        rows = []
        values = []

        for row in range(1, number_of_rows):
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(value)
                except ValueError:
                    pass
                finally:
                    values.append(value)

        date_dict = {}
        for i in range(2,number_of_rows*3-1,3):
            date = values[i-1]
        
            if date in date_dict:
                date_dict[date] += repr(values[i])
            else:
                date_dict[date] = repr(values[i])

        for date, tweets in date_dict.items():
            blob = TextBlob(tweets)
            sentiment = blob.sentiment.polarity
            year, month, day = (int(x) for x in date.split('-'))    
            ans = datetime.date(year, month, day)
            day_name = ans.strftime("%A")
            # print("Sentiment for date: " + date + " ( " + day_name + " ) is: " + str(sentiment))
            sentiment_dict[date] = {
                'sentiment': sentiment,
                'day_name': day_name
            }

    return sentiment_dict


def main():
    sentiment_dict = getSentiment()
    print("Sentiment analysis done... writing the sentiment dictionary into a file.")
    print("Sentiment Dictionary -")
    print json.dumps(sentiment_dict, sort_keys=True, indent=4)
    file = open("sentiment_dictionary.json","w")
    file.write(json.dumps(sentiment_dict, sort_keys=True, indent=4))


main()  # call of the main function