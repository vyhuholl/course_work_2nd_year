import urllib.request
import re
import xlsxwriter
import codecs

def download_page(pageUrl):
    page = urllib.request.urlopen(pageUrl)
    text = codecs.decode(page.read(), encoding = 'utf-8')
    return text

commonUrl = 'http://aranea.juls.savba.sk/guest/run.cgi/first?corpname=AranFinn_x&reload=&iquery=&queryselector=cqlrow&lemma=&lpos=&phrase=&word=&wpos=&char=&cql='
wordlist_sharp = ['ter%C3%A4v%C3%A4', 'jyrkk%C3%A4', 'vihlova', 'tarkka', 'korotettu', 'kirpe%C3%A4', 'kipakka', 'ter%C3%A4v%C3%A4kulmainen']
regTag = re.compile('<.*?>', flags = re.DOTALL)

questionnaire = open('questionnaire.xlsx', 'w', encoding = 'cp1251')

with open('nouns sharp.txt', encoding = 'cp1251') as wordlist:
    for line in wordlist:
        translation, word = line.split(' ')
        word = word[:-1]
        for sharp in wordlist_sharp:
            url = commonUrl + '%5Blemma+%3D+%22' + sharp + '%22%5D+%5Blemma+%3D+%22' + word + '%22%5D&default_attr=word'
            print(url)
            page = download_page(url)
            number = re.search('"(.*?) occurrences', page)
            print(number)
            try:
                number = number[1:]
                number = number[:-12]
                number = int(number)
                example = re.search('<td class="lc "(.*?)</td>', page) + sharp + re.search('<td class="rc "(.*?)</td>', page)
                example = re.sub('', regTag, example)
                if number >= 10:
                    questionnaire.write(sharp, 'финский', 'острый', translation, '', '', 'острый', '', example)
            except:
                continue

wordlist.close()
questionnaire.close()
 
