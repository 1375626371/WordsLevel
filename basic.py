import xlwings as xw
import pickle
import os
wb=xw.Book(r'高中英语单词检索词汇总表(人教版)(必修1至选修8).xlsm')
sht=wb.sheets['Sheet1']
sht2=wb.sheets['Sheet2']
class WordsLevel(object):
    wordnum=0
    days=0
    data=sht2.range((1,1),(3000,3)).value
    def Bookselect(self):
        #wb=xw.Book(r'高中英语单词检索词汇总表(人教版)(必修1至选修8).xlsm')
        #sht=wb.sheets['Sheet1']
        #sht2=wb.sheets['Sheet2']
        print('可供选择的书籍有：')
        print(sht.range((1,6),(8,6)).value)
        bookname=input('请输入你要背诵的单词书,中间以空格作为分割。（例如：必修1 必修2）').split(' ')
        print(bookname)
    def setplan(self):
        while WordsLevel.data[WordsLevel.wordnum]!=[None,None,None]:
            WordsLevel.wordnum+=1
            #print(WordsLevel.wordnum)
        WordsLevel.days=input('大侠选择的单词共有%d,准备几天结果掉它们？' % WordsLevel.wordnum)
    def Saveschedule(self,duixiang):
        file= open('parameter.pickle','wb')
        pickle.dump(a,file)
        file.close()
    def Readschedule(self):
        with open('parameter.pickle', 'rb') as file:
            a=pickle.load(file)
        return a
hasschedule=0
a=WordsLevel()
for filename in os.listdir("."):
    if filename[-6:]=='pickle':
        hasschedule=1
        break
if hasschedule == 0:
    a.setplan()
    a.Saveschedule(a)
else:
    a=a.Readschedule()
print(a.days)
