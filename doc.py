
#use python to write excel
#given 4 doc(string),count the time that each word appear in docs
import xlwt

DOCNUM = 4      #the number of docs

Doc = [None]*DOCNUM
Doc[0] = "prediction of whole country sales"
Doc[1] = "country sales rise in July"
Doc[2] = "decrease in home sales in June"
Doc[3] = "July country sales rise prediction"

#get all the words in all the docs
def seek_for_words():
    words = set([])
    for i in range(DOCNUM):
        lis = Doc[i].split(' ')
        for word in lis:
            words.add(word)
    return words

#if the word in the doc?
def is_in_doc(word,doc):
    lis = doc.split(' ')
    for s in lis:
        if(word == s):
            return True
    return False

#for each given word,count the time it appear in docs
def freq_of_word(word):
    freqs = [None]*DOCNUM   #the time of the word appear in each doc
    Doc_freq = 0        #the total times
    for i in range(DOCNUM):
        lis = Doc[i].split(' ')
        num = 0
        for j in range(len(lis)):
            if(word == lis[j]):
                num += 1
        freqs[i] = num
        Doc_freq += num
    return(Doc_freq,freqs)


words = list(seek_for_words())
words.sort(key = lambda x:x.lower())
excel = xlwt.Workbook()
table = excel.add_sheet('Matrix',cell_overwrite_ok=True)
table2 = excel.add_sheet('info',cell_overwrite_ok=True)

#the titles of tables in excel
for i in range(len(words)):
    table.write(i+1,0,words[i])
    table2.write(i+1,0,words[i])
for j in range(DOCNUM):
    table.write(0,j+1,'Doc'+str(j+1))
table2.write(0,0,'term')
table2.write(0,1,'doc.freq')
table2.write(0,2,'=>')
table2.write(0,3,'posting list')

for i in range(len(words)):
    word = words[i]
    freq = freq_of_word(word)
    table2.write(i+1,1,freq[0])
    table2.write(i+1,2,'=>')
    s = ''
    for n in range(DOCNUM):
        if(freq[1][n]):
            s = s + str(n) + '[freq=' + str(freq[1][n]) + '], '
    s = s[0:-2]                                 #remove ', ' in the end
    table2.write(i+1,3,s)
    for j in range(DOCNUM):
        if(is_in_doc(word,Doc[j])):
            table.write(i+1,j+1,1)
        else:
            table.write(i+1,j+1,0)
excel.save('test.xls')
