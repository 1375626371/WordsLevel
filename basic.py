import xlwings as xw
import pickle
import os
import math
import copy
def clear():os.system('cls')
class WordsLevel(object):
    wordnum=0#总单词数
    daygoal=0#每天背诵单词数
    basicnum=3000#假定单词最大数
    data=[]#需要背诵的全部单词
    days=0#背诵天数
    day=0#处于第几天
    remeberenglishwords=[]#储存英文单词计划
    remeberchinesewords=[]#储存中文单词计划
    wrongenglishwords=[]
    wrongchinesewords=[]
    def Englishcompare(self,inputword,rightword):
        a=0
        if inputword.lower()==rightword.lower():
            a=1
        while rightword.find('(')!=-1:
            b=list(rightword)
            b.remove('(')
            b.remove(')')
            if b[0]==' ':
                b.pop(0)
            if b[-1]==' ':
                b.pop(-1)
            c=''.join(b)
            e=rightword.find('(')
            f=rightword.find(')')
            g=list(rightword[:e])+list(rightword[f+1:])
            g=''.join(g)
            if inputword.lower()==c.lower():
                a=1
            if inputword.lower()==g.lower():
                a=1
            rightword=c
        if rightword.find('=')!=-1:
            c=rightword.find('=')
            b=list(rightword)
            if inputword.lower()==''.lower().join(b[:c]):
                a=1
            if inputword.lower()==''.lower().join(b[c+1:]):
                a=1
        return a
    def Chinesecompare(self,inputword,rightword):
        a=list(inputword)
        b=list(rightword)
        c=0
        for i in range(len(a)):
            if a[i]in b:
                c+=1
        if c>=2:
            return 1
        else:return 0
    def __init__(self):
        hasschedule=0
        for filename in os.listdir("."):
            if filename[-6:]=='pickle':
                WordsLevel.Readschedule(WordsLevel)
                hasschedule=1
                break
        if hasschedule == 0 or WordsLevel.days==0 or WordsLevel.wordnum==0:
            wb=xw.Book(r'高中英语单词检索词汇总表(人教版)(必修1至选修8).xlsm')
            sht=wb.sheets['Sheet1']
            sht2=wb.sheets['Sheet2']
            WordsLevel.data=sht2.range((1,1),(WordsLevel.basicnum,5)).value
            while WordsLevel.data[WordsLevel.wordnum][0]!=None:
                WordsLevel.wordnum+=1
                #print(WordsLevel.wordnum)
            WordsLevel.wordnum -= 1
            for b in range(WordsLevel.basicnum-WordsLevel.wordnum-1):
                WordsLevel.data.pop()
            WordsLevel.data.pop(0)
            WordsLevel.setplan(WordsLevel)
        else:
            WordsLevel.Readschedule(WordsLevel)
    def Wordremember(self):
        print('少侠，欢迎来到练武场，这是%d次练习，距离大成还需%d次，首先我们练习英文大典之汉英转化,单词清单如下'%(WordsLevel.day+1,WordsLevel.days-WordsLevel.day))
        for i in range(WordsLevel.daygoal):
            print('%s'%WordsLevel.remeberenglishwords[WordsLevel.day][i][1],end='')
            print('\t\t\t\t%s'%WordsLevel.remeberenglishwords[WordsLevel.day][i][0])
        print('在中文后输入对应的英文，按回车键提交输入。')
        for a in range(WordsLevel.daygoal):
            b=input('%s'%WordsLevel.remeberenglishwords[WordsLevel.day][a][1]).strip()
            if  WordsLevel.Englishcompare(WordsLevel,b,WordsLevel.remeberenglishwords[WordsLevel.day][a][0])==0:
                WordsLevel.wrongenglishwords.append(WordsLevel.remeberenglishwords[WordsLevel.day][a])
                WordsLevel.wrongenglishwords[-1][3]=WordsLevel.day
                WordsLevel.wrongenglishwords[-1][4]+=2
                d=0
                while(d!=3):
                    c=input("招式错误，正确的是%s,再练习三遍，中间以空格相隔，按回车提交。"%WordsLevel.remeberenglishwords[WordsLevel.day][a][0]).strip()
                    c=c.split(' ')
                    if  len(c)<3:
                        print('要练三遍不准偷懒！！！')
                        continue
                    if  len(c)>3:
                        print('说好了三遍，你想走火入魔？？？！！！')
                        continue
                    for i in range(3):
                        if WordsLevel.Englishcompare(WordsLevel,c[i],WordsLevel.remeberenglishwords[WordsLevel.day][a][0])==1 :
                            d+=1
                        else:d=0
        
        clear()
        print('少侠，恭喜完成英文大典之汉英转化，接下来首先我们练习英文大典之汉英转化，请在英文后输入对应的中文，按回车键提交输入。')
        for a in range(WordsLevel.daygoal):
            b=input('%s'%WordsLevel.remeberchinesewords[WordsLevel.day][a][0]).strip()
            if  WordsLevel.Chinesecompare(WordsLevel,b,WordsLevel.remeberenglishwords[WordsLevel.day][a][1])==0:
                WordsLevel.wrongchinesewords.append(WordsLevel.remeberchinesewords[WordsLevel.day][a])
                WordsLevel.wrongchinesewords[-1][3]=WordsLevel.day
                WordsLevel.wrongchinesewords[-1][4]+=2
                d=0
                while(d!=3):
                    c=input("招式错误，正确的是%s,再练习三遍，中间以空格相隔，按回车提交。"%WordsLevel.remeberchinesewords[WordsLevel.day][a][1]).strip()
                    c=c.split(' ')
                    if  len(c)<3:
                        print('要练三遍不准偷懒！！！')
                        continue
                    if  len(c)>3:
                        print('说好了三遍，你想走火入魔？？？！！！')
                        continue
                    for i in range(3):
                        if WordsLevel.Chinesecompare(WordsLevel,c[i],WordsLevel.remeberchinesewords[WordsLevel.day][a][1])==1 :
                            d+=1
                        else:d=0
        c=0
        for m in range(len(WordsLevel.wrongenglishwords)):
            if WordsLevel.day-WordsLevel.wrongenglishwords[m][3]:c=1
        if c==1:
            print('接下来我们需要将昨天汉英转化复习一下')
            for m in range(len(WordsLevel.wrongenglishwords)):
                if WordsLevel.day-WordsLevel.wrongenglishwords[m][3]>=1:
                    b=input('%s'%WordsLevel.wrongenglishwords[m][1]).strip()
                    if  WordsLevel.Englishcompare(WordsLevel,b,WordsLevel.wrongenglishwords[m][0])==0:
                        WordsLevel.wrongenglishwords[m][3]=WordsLevel.day
                        WordsLevel.wrongenglishwords[m][4]+=1
                        d=0
                        while(d!=3):
                            c=input("招式错误，正确的是%s,再练习三遍，中间以空格相隔，按回车提交。"%WordsLevel.wrongenglishwords[m][0]).strip()
                            c=c.split(' ')
                            if  len(c)<3:
                                print('要练三遍不准偷懒！！！')
                                continue
                            if  len(c)>3:
                                print('说好了三遍，你想走火入魔？？？！！！')
                                continue
                            for i in range(3):
                                if WordsLevel.Englishcompare(WordsLevel,c[i],WordsLevel.wrongenglishwords[m][0])==1:
                                    d+=1
                                else:d=0
                    else:WordsLevel.wrongenglishwords[m][4]-=1
                    if WordsLevel.wrongenglishwords[m][4]==0:
                        WordsLevel.wrongenglishwords.pop(m)
        c=0
        for m in range(len(WordsLevel.wrongchinesewords)):
            if WordsLevel.day-WordsLevel.wrongchinesewords[m][3]>=1:c=1
        if c==1:
            print('接下来我们需要将昨天英汉转化复习一下')
            for i in range(len(WordsLevel.wrongchinesewords)):
                if WordsLevel.day-WordsLevel.wrongchinesewords[m][3]>=1:
                    b=input('%s'%WordsLevel.wrongchinesewords[i][0]).strip()
                    if  WordsLevel.Chinesecompare(WordsLevel,b,WordsLevel.wrongchinesewords[i][1])==0:
                        WordsLevel.wrongchinesewords[i][3]=WordsLevel.day
                        WordsLevel.wrongchinesewords[i][4]+=1
                        d=0
                        while(d!=3):
                            c=input("招式错误，正确的是%s,再练习三遍，中间以空格相隔，按回车提交。"%WordsLevel.wrongchinesewords[i][1]).strip()
                            c=c.split(' ')
                            if  len(c)<3:
                                print('要练三遍不准偷懒！！！')
                                continue
                            if  len(c)>3:
                                print('说好了三遍，你想走火入魔？？？！！！')
                                continue
                            for m in range(3):
                                if WordsLevel.Chinesecompare(WordsLevel,c[i],WordsLevel.wrongchinesewords[i][1])==1:
                                    d+=1
                                else:d=0
                    else:WordsLevel.wrongchinesewords[i][4]-=1
                    if WordsLevel.wrongchinesewords[i][4]==0:
                        WordsLevel.wrongchinesewords.pop(i)
        
        WordsLevel.day+=1
        print('恭喜少侠完成今天的试炼，距离高考大业又进了一步')
        WordsLevel.Saveschedule(WordsLevel)
        WordsLevel.Mainmenu(WordsLevel)
    def Mainmenu(self):
        '主菜单'
        print('1.背单词')
        #a=int(input('请输入你要去的地方的方位，按回车键结束！'))
        a=input('请输入你要去的地方的方位，按回车键结束！')
        while a and (a=='1'or a=='2'):
            a=int(a.strip())
        if a==1:
            WordsLevel.Wordremember(WordsLevel)
    def Setremember(self):
        '随机生成中英文背诵表'
        for a in range(WordsLevel.days):
            for b in range(WordsLevel.daygoal):
                if len(WordsLevel.data):
                    WordsLevel.remeberenglishwords[a][b]=WordsLevel.data.pop()
                    WordsLevel.remeberenglishwords[a][b][4]=0
                else:
                    break
        a=WordsLevel.daygoal
        while a>=0:
            if WordsLevel.remeberenglishwords[WordsLevel.days-1][a-1]==None:
                WordsLevel.remeberenglishwords[WordsLevel.days-1].pop(a-1)
            a-=1
        WordsLevel.remeberchinesewords=copy.deepcopy(WordsLevel.remeberenglishwords)  
        WordsLevel.remeberchinesewords.reverse()
    def Bookselect(self):
        '选择需要背诵的书籍'
        #wb=xw.Book(r'高中英语单词检索词汇总表(人教版)(必修1至选修8).xlsm')
        #sht=wb.sheets['Sheet1']
        #sht2=wb.sheets['Sheet2']
        print('可供选择的书籍有：')
        #print(WordsLevel.sht.range((1,6),(8,6)).value)
        bookname=input('请输入你要背诵的单词书,中间以空格作为分割。（例如：必修1 必修2）').split(' ')
        print(bookname)
    def setplan(self):
        '建立自己的背诵目标，只运行一次'
        a=input('请问少侠是否要挑战三角符号的单词（不推荐）(y or n)(按回车结束输入)')
        if a=='n':
            b=WordsLevel.wordnum-1
            while b>=0:
                if(WordsLevel.data[b][3]==1):
                    WordsLevel.data.pop(b)
                b-=1
        WordsLevel.wordnum=len(WordsLevel.data)
        WordsLevel.daygoal=int(input('少侠选择的单词共有%d,准备一天结果掉几个？' % WordsLevel.wordnum))
        WordsLevel.days=math.ceil(WordsLevel.wordnum/WordsLevel.daygoal)
        WordsLevel.remeberenglishwords=[[None for i in range(WordsLevel.daygoal)] for i in range(WordsLevel.days)]
        WordsLevel.Setremember(WordsLevel)
        WordsLevel.Saveschedule(WordsLevel) 
    def Saveschedule(self):
        '保存进度'
        file= open('parameter.pickle','wb')
        basicnum=[WordsLevel.wordnum,WordsLevel.daygoal,WordsLevel.days,WordsLevel.day]
        alldata=[basicnum,WordsLevel.data,WordsLevel.remeberenglishwords,WordsLevel.remeberchinesewords,WordsLevel.wrongchinesewords,WordsLevel.wrongenglishwords]
        #print(alldata)
        pickle.dump(alldata,file)
        file.close()
    def Readschedule(self):
        with open('parameter.pickle', 'rb') as file:
            a=pickle.load(file)
        WordsLevel.wordnum,WordsLevel.daygoal,WordsLevel.days,WordsLevel.day=a[0]
        WordsLevel.data,WordsLevel.remeberenglishwords,WordsLevel.remeberchinesewords,WordsLevel.wrongchinesewords,WordsLevel.wrongenglishwords=a[1],a[2],a[3],a[4],a[5]
a=WordsLevel()
a.Mainmenu()