#Association rule mining code

import xlrd
import re
import nltk
from nltk.tokenize import RegexpTokenizer
from nltk import bigrams, trigrams
import math
from apyori import apriori
from nltk.stem.lancaster import LancasterStemmer
from nltk.stem import WordNetLemmatizer
from xlrd import open_workbook
from xlutils.copy import copy

#Load data
rb = xlrd.open_workbook("C:/Users/yekang/Desktop/2016/data/raw_data.xls")
s = rb.sheet_by_index(0)
rb_write = open_workbook("C:/Users/yekang/Desktop/2016/data/association_rule.xls")
wb = copy(rb_write)
s_wb = wb.get_sheet(1)
num_row = 7184
i = 0

#Preprocess loaded data
stopwords = nltk.corpus.stopwords.words('english')
tokenizer = RegexpTokenizer("[\w’]+", flags=re.UNICODE)
lancaster_stemmer = LancasterStemmer()
wordnet_lemmatizer = WordNetLemmatizer()

#Compute the frequency for each term.
vocabulary = []
docs = {}
all_tips = []
item = ""
final_tokens = []
while i < num_row:
    final_tokens = []
    item_list = s.row_values(i, start_colx=0, end_colx=4)
    if item_list[3] == 2:
        item = " ".join(item_list[:3])
        tokens = nltk.regexp_tokenize(item, '\w+|\$[\d\.]+|\S+/+&')

        tokens = [token.lower() for token in tokens if len(token) > 2]
        tokens = [token for token in tokens if token not in stopwords]

        tokens = [wordnet_lemmatizer.lemmatize(token) for token in tokens]
        tags_en = nltk.pos_tag(tokens)
        tokens_pos = ['/'.join(t[:-1]) for t in nltk.pos_tag(tokens) if ((t[1] == 'NN') or (t[1] == 'VBP') or (t[1] == 'VBS') or (t[1] == 'VBZ') or (t[1] == 'VBG') or (t[1] == 'VB')or (t[1] == 'VBD') or (t[1] == 'VBN') or (t[1] == 'VBP') & (t[0] not in stopwords))]
        final_tokens.extend(tokens_pos)
    vocabulary.append(final_tokens)
    i += 1
    print(i)

#Association rule analysis
kwargs = {"min_support": 0.01, "min_confidence": 0.3, "min_lift": 5}
results = list(apriori(vocabulary, **kwargs))
print(results)

index = 0
while index < len(results):
    s_wb.write(index, 0, str(results[index].items))
    s_wb.write(index, 1, str(results[index].support))
    if len(results[index][2]) >= 1:
        s_wb.write(index, 2, str(results[index].ordered_statistics[0]))
    if len(results[index][2]) == 2:
        s_wb.write(index, 3, str(results[index].ordered_statistics[1]))
    index += 1

wb.save('C:/Users/yekang/Desktop/2016/data/association_rule.xls')
