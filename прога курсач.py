import urllib.request
import html
import re
from urllib.parse import quote, urlsplit, urlunsplit

def iri_to_uri(iri):
    parts = urlsplit(iri)
    uri = urlunsplit((
        parts.scheme, 
        parts.netloc.encode('idna').decode('ascii'), 
        quote(parts.path),
        quote(parts.query, '='),
        quote(parts.fragment),
    ))
    return uri

def download_page(pageUrl):
    try:
        page = urllib.request.urlopen(pageUrl)
        text = page.read().decode('utf-8')
        return text
    except:
        return('Error')

commonUrl = 'https://fi.glosbe.com/ru/fi/'
regPostTranslation = re.compile('<div class="text-info"><strong class=" phr">(.*?)</strong>', flags= re.DOTALL)
regTag = re.compile('<.*?>', flags= re.DOTALL)
words = []
translations = []

with open('words.txt') as wordlist:
    for word in wordlist:
        words.append(word.strip('\n'))
    wordlist.close()

for word in words:
    pageUrl = iri_to_uri(commonUrl + word)
    page = download_page(pageUrl)
    try:
        translation = (regPostTranslation.search(page)).group()
        translation = regTag.sub('', translation)
        translation = str(translation.encode('utf-8'))
    except:
        translation = 'Нет перевода'
    translations.append(translation)

with open('translations.txt', 'w') as trlist:
    for i in range(len(words)):
        trlist.write(words[i] + ' ' + translations[i] + '\n')
    trlist.close()
