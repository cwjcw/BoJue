import xlwings as xw
import re


app = xw.App(visible=False,add_book=False)
source = app.books.open(r'F:\07_数据分析\B拍补拍\B拍\202210\补拍申请单-Grid分析 1.1-10.13.xls') # 表格存放路径
source_sht = source.sheets(1)
source_last_low = source_sht.range("A65536").end('up').row
key = app.books.open(r'F:\07_数据分析\B拍补拍\关键字.xlsx')
key_sht = key.sheets(1)
for i in range(2,source_last_low+1):
    reason = source_sht.cells(i,5).value

    huizong = ''
    liebiao = []
    for x in range(1,34):
        # print(type(key_sht.cells(x,1).value))
        result = re.search(key_sht.cells(x,1).value,reason)

        if result != None:
             huizong = huizong  + (result.group()+',')
             liebiao.append(result.group())
    print(i)
    source_sht.cells(i,6).value = huizong
    # print(source_sht.cells(i,6).value)
source.save()
source.close()
key.close()
app.quit()