#Wordcloud code
import xlrd
import re
import nltk
from nltk.tokenize import RegexpTokenizer
from nltk import bigrams, trigrams
import math
from wordcloud import WordCloud
from nltk.stem import WordNetLemmatizer

#Load data
rb = xlrd.open_workbook("C:/Users/yekang/Desktop/2016/data/raw_data.xls")
s = rb.sheet_by_index(0)
num_row = 1
i = 0

#Preprocessing
stopwords = nltk.corpus.stopwords.words('english')
tokenizer = RegexpTokenizer("[\w’]+", flags=re.UNICODE)

#Function definition
def freq(word, doc):
    return doc.count(word)


def word_count(doc):
    return len(doc)


def tf(word, doc):
    return (freq(word, doc) / float(word_count(doc)))


def num_docs_containing(word, list_of_docs):
    count = 0
    for document in list_of_docs:
        if freq(word, document) > 0:
            count += 1
    return count


def idf(word, list_of_docs):
    return math.log(len(list_of_docs) / float(num_docs_containing(word, list_of_docs)))


def tf_idf(word, doc, list_of_docs):
    return (tf(word, doc) * idf(word, list_of_docs))

wordnet_lemmatizer = WordNetLemmatizer()

#Compute the frequency for each term.
vocabulary = []
docs = {}
all_tips = []
item = ""
while i < num_row:
    item_list = s.row_values(i, start_colx=0, end_colx=11)
    if item_list[10] == 0:
        item = " ".join(item_list[:10])
        tokens = nltk.regexp_tokenize(item, '\w+|\$[\d\.]+|\S+/+&')
        tokens = [token.lower() for token in tokens if len(token) > 2]
        tokens = [token for token in tokens if token not in stopwords]
        tokens = [wordnet_lemmatizer.lemmatize(token) for token in tokens]

        tags_en = nltk.pos_tag(tokens)
        tokens_pos = ['/'.join(t[:-1]) for t in nltk.pos_tag(tokens) if ((t[1] == 'NN') & (t[0] not in stopwords))]

    #bi_tokens = bigrams(tokens)
    # tri_tokens = trigrams(tokens)


    #bi_tokens = [' '.join(token).lower() for token in bi_tokens]
    #bi_tokens = [token for token in bi_tokens if token not in stopwords]

    # tri_tokens = [' '.join(token).lower() for token in tri_tokens]
    # tri_tokens = [token for token in tri_tokens if token not in stopwords]
    final_tokens = []
    final_tokens.extend(tokens_pos)
    #final_tokens.extend(bi_tokens)
    # final_tokens.extend(tri_tokens)
    docs[i] = {'freq': {}, 'tf': {}, 'idf': {}, 'tf-idf': {}, 'tokens': []}

    for token in final_tokens:
        #The frequency computed for each tip
        docs[i]['freq'][token] = freq(token, final_tokens)
        #The term-frequency (Normalized Frequency)
        docs[i]['tf'][token] = tf(token, final_tokens)
        docs[i]['tokens'] = final_tokens

    vocabulary.append(final_tokens)
    print(i)
    i += 1

"""
for doc in docs:
    for token in docs[doc]['tf']:
        #The Inverse-Document-Frequency
        docs[doc]['idf'][token] = idf(token, vocabulary)
        #The tf-idf
        docs[doc]['tf-idf'][token] = tf_idf(token, docs[doc]['tokens'], vocabulary)

"""
#Now let's find out the most relevant words by tf-idf.
words = {}
for doc in docs:
    for token in docs[doc]['tf']:
        if token not in words:
            words[token] = docs[doc]['tf'][token]
        else:
            if docs[doc]['tf'][token] > words[token]:
                words[token] = docs[doc]['tf'][token]

"""
for token in docs[doc]['tf-idf']:
    print(token, docs[doc]['tf-idf'][token])


for item in sorted(words.items(), key=lambda x: x[1], reverse=True):
    print("%f <= %s" % (item[1], item[0]))


# tags = make_tags(words, maxsize=120)
# create_tag_image(tags, 'cloud_large.png', size=(900, 600), fontname='Lobster')

with open('C:/Users/Young Eun/Desktop/2016/data/words.csv', 'w', encoding='utf-8') as f:
    f.write('word,freq \n')
    writer = csv.writer(f)
    writer.writerows(words.items())
"""

#Font setting
wordcloud = WordCloud(font_path = r'C:\Windows\Fonts\Cambria.ttc')
#Wordcloud input
word_tf = zip(list(words.keys()),list(words.values()))
wordcloud.generate_from_frequencies(word_tf).to_file('WordCloud_Cluster1.png')