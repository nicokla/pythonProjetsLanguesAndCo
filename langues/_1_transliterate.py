# -*- coding: UTF-8 -*-

#https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes
#http://polyglot.readthedocs.io/en/latest/Transliteration.html

#from polyglot.downloader import downloader
#print(downloader.supported_languages_table("transliteration2"))
#downloader.download('embeddings2.en')
#downloader.download('transliteration2.fr')
#downloader.download('transliteration2.ar')
#downloader.download('transliteration2.ru')
#downloader.download('transliteration2.hi')
#downloader.download('transliteration2.zh')
#downloader.download('transliteration2.ja')
#downloader.download('transliteration2.he')
#downloader.download('transliteration2.yi')
#downloader.download('transliteration2.ko')
#downloader.download('transliteration2.th')
#downloader.download('transliteration2.am')
#downloader.download('transliteration2.el')

from polyglot.text import Text
from unidecode import unidecode

def transliteratePolyglot(blob):
    text = Text(blob)
    l=text.transliterate()
    t=l[0]
    for i in range(len(l)-1):
        t=t+' '+l[i+1]
    return t

blob='누구'
print transliteratePolyglot(blob)
#unidecode(blob)












