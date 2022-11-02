import requests
import json
import time
import xlwings as xw

app = xw.App(visible=False,add_book=False)
wb = app.books.open(r'C:\Users\Administrator\Desktop\出勤表.xlsx')
sht = wb.sheets[0]

emp_id = ['U17330','U09928','U07338','U15303','U02835','U00364','U10670','U00011','U02073','U14228']

headers = {
   'Open-Authorization': '4VzqOE93a8ZeLvFcrATsRf2YuZThrVJempx1xTNY',
   'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36 QIHU 360SE',
   'Content-Type': 'application/json'
}

Title = ['userid', 'groupname', 'checkin_type' , 'exception_type','checkin_time', 'location_title', 'location_detail',
         'wifiname', 'notes', 'wifimac', 'mediaids', 'lat', 'lng', 'deviceid', 'sch_checkin_time', 'groupid',
         'schedule_id', 'timeline_id']

def get_token():
    s = 'secret'
    url1 = f'https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=ID&corpsecret={s}'
    response = requests.request("POST", url1, headers=headers)
    # print(response)
    x = json.loads(response.text)
    ACCESS_TOKEN = x['access_token']
    url = f"https://qyapi.weixin.qq.com/cgi-bin/checkin/getcheckindata?access_token={ACCESS_TOKEN}"
    return url

def get_data():
    for i in emp_id:
        # print(i)
        payload = json.dumps({
        "opencheckindatatype": 3,
        "starttime": 1664553600, #  Unix时间戳
        "endtime": 1664899200, #  Unix时间戳
        "useridlist": i
        })
        response = requests.request("POST", get_token(), headers=headers, data=payload)
        # print(response.text)
        x = json.loads(response.text) # 把Response格式转为字典格式（文本转字典）
        # print(x)
        checkindata = x['checkindata'] # 数据储存在key(checkindata)对的Value里面,这是一个列表
        # print(checkindata)
        # print(checkindata)
        data = []
        for n in checkindata:
            # print(n)
            # print(n)
            day_data = []
            for t in Title:
                # print(n[t])
                day_data.append(str((n[t])))
                # print(data)
                # print(list(data))

            data.append(day_data)
        # print(data)
        for d in data:
            # print(d)
            row = sht.range('A65536').end('up').row + 1
            # print(d[4])
            d[4] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(int(d[4])))
            d[-4] = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(int(d[-4])))
            # print(d[4])
            sht.range(row, 1).value = d
        # print(data)


# print(get_token())
get_data()
print('输入完成')
wb.save()
wb.close()
app.quit()
