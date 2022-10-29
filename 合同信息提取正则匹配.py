import xlwings as xw
import re

app = xw.App(visible=False,add_book=False)

source_data = app.books.open(r'F:\02_DATABASE\订单与拍摄\套系DB.xlsx')

Source_sht = source_data.sheets(1)
source_last_low = Source_sht.range("A65536").end('up').row
CN_NUM = {
    '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '零': 0,
    '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9, '貮': 2, '两': 2,}

def Photo_NO():
    Photo_NO1, Photo_NO2 = r"精修\d+|精调\d+|精选\d+", r"\d+"
    for i in range(2,source_last_low+1):
        res = re.search(Photo_NO1,Source_sht.cells(i,3).value)
        if res != None:
            res2 = re.search(Photo_NO2,res.group())
            Source_sht.cells(i,23).value = res2.group()
            print(res2.group())
    source_data.save()

def photos():
    p1,p2 = r'\d+张',r'\d+'
    count = 0
    for i in range(2,source_last_low+1):
        res = re.search(p1,Source_sht.cells(i,3).value)
        if res != None:
            count = count + 1
            res2 = re.search(p2,res.group())
            Source_sht.cells(i,24).value = res2.group()
            print(res.group(),res2.group(),count)
    source_data.save()

def hotel():
    h1,h2 = r'精品|蜜月|星级',r'酒店.+|[晚|夜]'
    count = 0
    for i in range(2,source_last_low+1):
        if re.search(h2,Source_sht.cells(i,3).value) != None:
            res = re.search(h1,Source_sht.cells(i,3).value)
            if res != None:
                count = count + 1
                Source_sht.cells(i,25).value = res.group()
                print(res.group(),count)
        else:
            Source_sht.cells(i, 25).value = '无酒店'
    source_data.save()

def hotel_days():
    d1,d2= r'\d[晚|夜]|[一|二|三|四|五|六|七|八|九|十|两][晚|夜]',r'\d|[一|二|三|四|五|六|七|八|九|十|两]'
    count = 0
    for i in range(2,source_last_low+1):
        res = re.search(d1,Source_sht.cells(i,3).value)
        if res != None:
            count = count + 1
            res2 = re.search(d2,res.group())
            if re.search(r'\d',res2.group()) == None:
                res2 = CN_NUM[res2.group()]
                Source_sht.cells(i, 26).value = res2
            else:
                Source_sht.cells(i,26).value = res2.group()
                print(res.group(), res2, count)
        else:
            Source_sht.cells(i, 26).value = 0
    source_data.save()

def Tec_team():
    t1,t2 = r'[首席|样片|资深|集团]\D{3}总监','首席|样片|资深|集团'
    count = 0
    for i in range(2, source_last_low + 1):
        res = re.search(t1, Source_sht.cells(i, 3).value)
        if res != None:
            count = count + 1
            res2 = re.search(t2, res.group())
            Source_sht.cells(i, 27).value = res2.group()
            print(res2.group(), count)
    source_data.save()

# Photo_NO()
# photos()
# hotel()
# hotel_days()
Tec_team()


source_data.close()
app.quit()
