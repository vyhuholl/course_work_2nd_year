import urllib.request
import re
import xlsxwriter

def download_page(pageUrl):
    page = urllib.request.urlopen(pageUrl)
    text = page.read().decode('utf-8')
    return text

commonUrl = 'http://aranea.juls.savba.sk/guest/run.cgi/first?corpname=AranFinn_x&reload=&iquery=&queryselector=cqlrow&lemma=&lpos=&phrase=&word=&wpos=&char=&cql='
wordlist_sharp = ['ter%C3%A4v%C3%A4', 'jyrkk%C3%A4', 'vihlova', 'tarkka', 'korotettu', 'kirpe%C3%A4', 'kipakka', 'ter%C3%A4v%C3%A4kulmainen', 'yl%C3%A4vireinen', '%C3%A4t%C3%A4kk%C3%A4', 'pist%C3%A4v%C3%A4', 'k%C3%A4rjek%C3%A4s', 'fiksu', 't%C3%A4sm%C3%A4llinen', 'ankara', 'tuima', 'tyylik%C3%A4s']
wordlist_size = ['iso', 'aikuinen', 'suuri', 'mittava', 'pieni', 'nuori', 'snadi', 'pikku']
wordlist_smooth = ['sile%C3%A4', 'sujuva', 'kitkaton', 'juoheva', 'sutjakka', 'sulava', 'pehme%C3%A4', 'luonteva', 'luonnikas']
regTag = re.compile('<.*?>', flags = re.DOTALL)
regNumber = re.compile('\"(.*?) hits')
regLeftContext = re.compile('<td class="lc "(.*?)</td>', flags = re.DOTALL)
regRightContext = re.compile('<td class="rc "(.*?)</td>', flags = re.DOTALL)
questionnaire_sharp = xlsxwriter.Workbook('questionnaire_sharp.xlsx')
sharp_worksheet1 = questionnaire_sharp.add_worksheet('Финский')
sharp_worksheet2 = questionnaire_sharp.add_worksheet('Финский стандартная форма')
questionnaire_size = xlsxwriter.Workbook('questionnaire_size.xlsx')
size_worksheet1 = questionnaire_size.add_worksheet('Финский')
size_worksheet2 = questionnaire_size.add_worksheet('Финский стандартная форма')
questionnaire_smooth = xlsxwriter.Workbook('questionnaire_smooth.xlsx')
smooth_worksheet1 = questionnaire_smooth.add_worksheet('Финский')
smooth_worksheet2 = questionnaire_smooth.add_worksheet('Финский стандартная форма')
sharp_worksheet1.write('лексема', 'A1', 'bold')
sharp_worksheet1.write('язык', 'B1', 'bold')
sharp_worksheet1.write('микрофрейм', 'C1', 'bold')
sharp_worksheet1.write('контекст на языке', 'D1', 'bold')
sharp_worksheet1.write('фрейм', 'E1', 'bold')
sharp_worksheet1.write('такс. класс', 'F1', 'bold')
sharp_worksheet1.write('поле', 'G1', 'bold')
sharp_worksheet1.write('тип значения', 'H1', 'bold')
sharp_worksheet1.write('пример', 'I1', 'bold')
sharp_worksheet1.write('комментарий', 'J1', 'bold')
size_worksheet1.write('лексема', 'A1', 'bold')
size_worksheet1.write('язык', 'B1', 'bold')
size_worksheet1.write('микрофрейм', 'C1', 'bold')
size_worksheet1.write('контекст на языке', 'D1', 'bold')
size_worksheet1.write('фрейм', 'E1', 'bold')
size_worksheet1.write('такс. класс', 'F1', 'bold')
size_worksheet1.write('поле', 'G1', 'bold')
size_worksheet1.write('тип значения', 'H1', 'bold')
size_worksheet1.write('пример', 'I1', 'bold')
size_worksheet1.write('комментарий', 'J1', 'bold')
smooth_worksheet1.write('лексема', 'A1', 'bold')
smooth_worksheet1.write('язык', 'B1', 'bold')
smooth_worksheet1.write('микрофрейм', 'C1', 'bold')
smooth_worksheet1.write('контекст на языке', 'D1', 'bold')
smooth_worksheet1.write('фрейм', 'E1', 'bold')
smooth_worksheet1.write('такс. класс', 'F1', 'bold')
smooth_worksheet1.write('поле', 'G1', 'bold')
smooth_worksheet1.write('тип значения', 'H1', 'bold')
smooth_worksheet1.write('пример', 'I1', 'bold')
smooth_worksheet1.write('комментарий', 'J1', 'bold')
for i in range(len(wordlist_sharp)):
    sharp_worksheet2.write(0, i + 2, wordlist_sharp[i])
for i in range(len(wordlist_size)):
    size_worksheet2.write(0, i + 2, wordlist_size[i])
for i in range(len(wordlist_smooth)):
    smooth_worksheet2.write(0, i + 2, wordlist_smooth[i])

with open('nouns sharp.txt', encoding = 'utf-8') as wordlist_sharp_n:
    wordlist_sharp_n = list(wordlist_sharp_n)
    i_sharp = 1
    for i in range(len(wordlist_sharp_n)):
        translation, word = wordlist_sharp_n[i].split(' ')
        sharp_worksheet2.write(i + 1, 0, translation)
        sharp_worksheet2.write(i + 1, 1, word)
        for j in range(len(wordlist_sharp)):
            sharp = wordlist_sharp[j]
            url = commonUrl + '%5Blemma+%3D+%22' + sharp + '%22%5D+%5Blemma+%3D+%22' + word + '%22%5D&default_attr=word'
            page = download_page(url)
            if regNumber.search(page) != None:
                number = (regNumber.search(page)).group()
                number = (number.lstrip('"')).rstrip(' hits')
                number = int(number)
                LeftContext = (regLeftContext.search(page)).group()
                LeftContext = re.sub('\n', '', LeftContext)
                LeftContext = re.sub('  ', '', LeftContext)
                RightContext = (regRightContext.search(page)).group()
                RightContext = re.sub('\n', '', RightContext)
                RightContext = re.sub('  ', '', RightContext)
                example = re.sub(regTag, '', LeftContext + ' ' + sharp + ' ' + RightContext)
                print(example)
                if number >= 10:
                    sharp_worksheet1.write(i_sharp, 0, word)
                    sharp_worksheet1.write(i_sharp, 1, 'финский')
                    sharp_worksheet1.write(i_sharp, 2, 'острый' + translation)
                    sharp_worksheet1.write(i_sharp, 3, word)
                    sharp_worksheet1.write(i_sharp, 6, 'острый')
                    sharp_worksheet1.write(i_sharp, 8, ехаmple)
                    sharp_worksheet2.write(i + 1, j + 2, ' +')
                    i_sharp += 1

with open('nouns size.txt', encoding = 'utf-8') as wordlist_size_n:
    i_size = 1
    wordlist_size_n = list(wordlist_size_n)
    for i in range(len(wordlist_size_n)):
        translation, word = wordlist_size_n[i].split(' ')
        size_worksheet2.write(i + 1, 0, translation)
        size_worksheet2.write(i + 1, 1, word)
        for j in range(len(wordlist_size)):
            size = wordlist_size[j]
            url = commonUrl + '%5Blemma+%3D+%22' + size + '%22%5D+%5Blemma+%3D+%22' + word + '%22%5D&default_attr=word'
            page = download_page(url)
            if regNumber.search(page) != None:
                number = (regNumber.search(page)).group()
                number = (number.lstrip('"')).rstrip(' hits')
                number = int(number)
                LeftContext = (regLeftContext.search(page)).group()
                LeftContext = re.sub('\n', '', LeftContext)
                LeftContext = re.sub('  ', '', LeftContext)
                RightContext = (regRightContext.search(page)).group()
                RightContext = re.sub('\n', '', RightContext)
                RightContext = re.sub('  ', '', RightContext)
                example = re.sub(regTag, '', LeftContext + ' ' + sharp + ' ' + RightContext)
                print(example)
                if number >= 10:
                    size_worksheet1.write(i_size, 0, word)
                    size_worksheet1.write(i_size, 1, 'финский')
                    if j < 4:
                        size_worksheet1.write(i_size, 2, 'большой' + translation)
                    else:
                        size_worksheet1.write(i_size, 2, 'маленький' + translation)
                    size_worksheet1.write(i_size, 3, word)
                    size_worksheet1.write(i_size, 6, 'размер')
                    size_worksheet1.write(i_size, 8, ехаmple)
                    size_worksheet2.write(i + 1, j + 2, ' +')
                    i_size += 1

with open('nouns smooth.txt', encoding = 'utf-8') as wordlist_smooth_n:
    i_smooth = 1
    wordlist_smooth_n = list(wordlist_smooth_n)
    for i in range(len(wordlist_smooth_n)):
        translation, word = line.split(' ')
        size_worksheet2.write(i + 1, 0, translation)
        size_worksheet2.write(i + 1, 1, word)
        for j in range(len(wordlist_smooth)):
            smooth = wordlist_smooth[j]
            url = commonUrl + '%5Blemma+%3D+%22' + smooth + '%22%5D+%5Blemma+%3D+%22' + word + '%22%5D&default_attr=word'
            page = download_page(url)
            if regNumber.search(page) != None:
                number = (regNumber.search(page)).group()
                number = (number.lstrip('"')).rstrip(' hits')
                number = int(number)
                LeftContext = (regLeftContext.search(page)).group()
                LeftContext = re.sub('\n', '', LeftContext)
                LeftContext = re.sub('  ', '', LeftContext)
                RightContext = (regRightContext.search(page)).group()
                RightContext = re.sub('\n', '', RightContext)
                RightContext = re.sub('  ', '', RightContext)
                example = re.sub(regTag, '', LeftContext + ' ' + sharp + ' ' + RightContext)
                print(example)
                if number >= 10:
                    smooth_worksheet1.write(i_smooth, 0, word)
                    smooth_worksheet1.write(i_smooth, 1, 'финский')
                    smooth_worksheet1.write(i_sharp, 2, 'гладкий' + translation)
                    smooth_worksheet1.write(i_smooth, 3, word)
                    smooth_worksheet1.write(i_smooth, 6, 'острый')
                    smooth_worksheet1.write(i_smooth, 8, ехаmple)
                    smooth_worksheet2.write(i + 1, j + 2, ' +')
                    i_smooth += 1

wordlist_sharp_n.close()
wordlist_size_n.close()
wordlist_smooth_n.close()
questionnaire_sharp.close()
questionnaire_size.close()
questionnaire_smooth.close()

 
