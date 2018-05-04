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
            list_tweets = tweets.split('.')
            sentiment_total = 0
            cnt = 0
            for each in list_tweets:
                cnt += 1
                blob = TextBlob(each)    
                sentiment_total += blob.sentiment.polarity
            
            sentiment = sentiment_total/cnt 
            if sentiment == 0.0:
                sentiment = 0.05

            year, month, day = (int(x) for x in date.split('-'))    
            ans = datetime.date(year, month, day)
            day_name = ans.strftime("%A")
            # print("Sentiment for date: " + date + " ( " + day_name + " ) is: " + str(sentiment))
            sentiment_dict[date] = {
                'sentiment': sentiment,
                'day_name': day_name
            }

    return sentiment_dict

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

def calDiff():
    wb = open_workbook('AMZN.xlsx')
    change_dict = {}
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
    print(sheet.cell(0,0).value)

    for row in range(1,number_of_rows):
        one = float(sheet.cell(row,1).value)
        two = float(sheet.cell(row,2).value)
        date = str((sheet.cell(row,0).value))
        dateoffset = 693594
        dateStr = datetime.date.fromordinal(dateoffset + int(float(date))).strftime('%Y-%m-%d')

        if((two-one)>0):
            change_dict[dateStr] = 1

        else:
            change_dict[dateStr] = 0
    return change_dict

def mapSentivalToStockval(sentiment_dict,change_dict):
    newMap_dict = {}
    count=0
    avgVal=0

    day = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    for dates in sentiment_dict:
        for datec in change_dict:
            if(dates==datec):
                if(sentiment_dict[dates]["day_name"]==day[5]):
                    flag=1
                elif(sentiment_dict[dates]["day_name"]==day[6]):
                    flag=1
                else:
                    flag=0
                if(flag==1):
                    avgVal+=sentiment_dict[dates]["sentiment"]
                    count+=1
                    continue
                if sentiment_dict[dates]["day_name"]==day[0]:
                    avgVal+=sentiment_dict[dates]["sentiment"]
                    newMap_dict[dates]={
                    'sentiment': (avgVal/(count+1)),
                    'stock': change_dict[datec]
                    }
                    avgVal=0
                    count=0
                else:
                    newMap_dict[dates]={
                    'sentiment': sentiment_dict[dates]["sentiment"],
                    'stock': change_dict[datec]
                    }
    return newMap_dict

def calTrends():
    wb = open_workbook('last3weeks.xlsx')
    trends = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols

    for row in range(1,number_of_rows):
        one = float(sheet.cell(row,1).value)
        two = float(sheet.cell(row,2).value)
        if((two-one)>0):
            trends.append(1)

        else:
            trends.append(0)

    return trends

def main():
    sentiment_dict = getSentiment()
    print("Sentiment analysis done... writing the sentiment dictionary into a file.")
    print("Sentiment Dictionary -")
    print(json.dumps(sentiment_dict, sort_keys=True, indent=4))
    file = open("sentiment_dictionary.json","w")
    file.write(json.dumps(sentiment_dict, sort_keys=True, indent=4))
    change_dict = calDiff()
    print("Stock value change :")
    print(json.dumps(change_dict, sort_keys=True, indent=4))
    newMap_dict = mapSentivalToStockval(sentiment_dict,change_dict)
    print("Mapped values :")
    print(json.dumps(newMap_dict, sort_keys=True, indent=4))

    # arrays to be given to Numpy for model training
    X_Sentiment = []
    Y_StockVal  = []
    for date in newMap_dict:
        train_data = newMap_dict[date]
        X_Sentiment.append([train_data["sentiment"]])
        Y_StockVal.append(train_data["stock"])

    clf = tree.DecisionTreeClassifier()
    clf = clf.fit(X_Sentiment, Y_StockVal)

    # Break point analysis
    # print clf.predict([[0.05614616]])
    # print clf.predict([[0.05614617]])

    count_tp = 0.0
    count_tn = 0.0
    count_fp = 0.0
    count_fn = 0.0

    t_value = 0.05614616

    # By taking the current trends into consideration
    trends = calTrends()
    for each in trends:
        Y_StockVal.append(each)
        if each == 1:
            X_Sentiment.append([0.075])
        else: 
            X_Sentiment.append([0.035])


    for date, value in newMap_dict.items():
        if value["sentiment"] > t_value:
            if value["stock"] == 1:
                count_tp += 1.0
            elif value["stock"] == 0:
                count_fn += 1.0
        elif value["sentiment"] <= t_value:
            if value["stock"] == 0:
                count_tn += 1.0
            elif value["stock"] == 1:
                count_fp += 1.0

    print 'count_tp:', count_tp
    print 'count_tn:', count_tn
    print 'count_fp:', count_fp
    print 'count_fn:', count_fn

    accuracy_with_trends = (count_tp+count_tn)/(count_tn+count_fp+count_fn+count_tp)
    print 'Accuracy of the model:', accuracy_with_trends

    print '[0] - Stock price will go down; [1] - Stock price will go up'

    print 'Break point for our DecisionTreeClassifier:', t_value

    # thurs_sentiment = ThursdaySentiment()
    # print 'Thursday sentiment value:', thurs_sentiment
    # print 'Model prediction for Thursday', clf.predict([[thurs_sentiment]])

    # friday_sentiment = fridaySentiment()
    # print 'Friday sentiment value:', friday_sentiment
    # print 'Model Prediction for Friday', clf.predict([[friday_sentiment]])
    
    # Model persistence
    joblib.dump(clf, 'dmpa_lab_project.pkl') 
    print 'Model saved.'


main() # call of the main function