from pykeyboard import *
from pymouse import *
import time
import win32api,win32con,win32clipboard as w
import xlwings as xw

app = xw.App(visible=False,add_book=False)
wb = app.books.open(r'F:\01_经营分析\排挡\排档自动化\微信群相关\1111.xls')
sht = wb.sheets(1)
last_low = sht.range("B65536").end('up').row


k = PyKeyboard()
m = PyMouse()
for i in range(2,4):
    def get_text():
        w.OpenClipboard()
        d=w.GetClipboardData(win32con.CF_TEXT)
        w.CloseClipboard()
        return d
    # 复制内容到剪切板
    def set_text(astring):
        w.OpenClipboard()
        w.EmptyClipboard()
        d=w.SetClipboardData(win32con.CF_UNICODETEXT,astring)
        w.CloseClipboard()
    # 定义了一些key值
    vk_code={'ctrl':0x11,'enter':0x0D,'a':0x41,'v':0x56,'x':0x58}
    # 键盘按下
    def key_down(keyname):
        win32api.keybd_event(vk_code[keyname],0,0,0)
    # 键盘抬起
    def key_up(key_name):
        win32api.keybd_event(vk_code[key_name],0,win32con.KEYEVENTF_KEYUP,0)
    #按键组合操作
    def simulate_key(firstkey,sencondkey):
        key_down(firstkey)
        key_down(sencondkey)
        key_up(sencondkey)
        key_up(firstkey)
        time.sleep(2)
        key_down('enter') #按下回车
        key_up('enter') # 抬起回车
        # print('simulate_key执行完成！')

    def set_group():
        m.click(26,250,1,1) # 点击工作台
        time.sleep(1)
        m.click(377,176,1,1) # 点击客户群
        time.sleep(1)
        m.click(536,269, 1, 1)  # 点击创建客户群
        time.sleep(4)
        m.click(1893, 46, 1, 1)  # 进入拉人界面
        people = [sht.cells(i,89).value,sht.cells(i,5).value]
        for p in range(len(people)):
            time.sleep(1)
            m.click(707, 303, 1, 1)  # 进入查询界面
            set_text(people[p])  # 查找建群人员
            time.sleep(1)
            simulate_key('ctrl', 'v')
            time.sleep(1)
            m.click(674, 350, 1, 1)  # 选择拉进群的对象
            time.sleep(1)
        m.click(1127, 776, 1, 1)  # 点击“确认”建群
        time.sleep(1)

    def get_QRcode():
            save_path = r"\\192.168.12.2\酒店\群二维码\{}".format(sht.cells(i,5).value)
            set_text(save_path)
            m.click(1853,43,1,1) # 进入群设置选项
            m.click(1720,170,1,1) # 进入二维码界面
            time.sleep(1)
            m.click(1481, 761, 1, 1)  # 点击保存二维码
            time.sleep(2)
            m.click(1583, 178, 1, 1)  # 点击修改二维码保存路径
            time.sleep(2)
            simulate_key('ctrl', 'v')  # 粘贴内容到获得焦点的输入框

            m.click(1748, 643, 1, 1)  # 点击保存
            time.sleep(2)
            m.click(1574, 152, 1, 1)  # 关闭二维码
            print('群创建成功！')

    def set_name_announce():

        set_group()
        time.sleep(1)
        group_name = '{}【{}&{} 铂爵旅拍VIP沟通群】'.format(str(sht.cells(i, 6).value)[:10], sht.cells(i, 9).value, sht.cells(i, 11).value)
        set_text(group_name.replace("None&",'').replace('&None',''))
        m.click(471,39,1,1) # 点击群名
        # time.sleep(1)
        simulate_key('ctrl','v')  # 粘贴内容到获得焦点的输入框

        # 输入群公告文本
        announce = '\
以下是为您订制的拍摄行程安排：\
\n \
第1天：{arr_day}{arr_time} | 到达中航紫金广场门店办理手续（约90分钟），办理完手续后购买第二天上岛船票 \n \
购买方式：微信搜索“厦门轮渡有限公司”\n \
船票时间：{arr_day}早上08: 10班次（现在就可以买了哦）\n \
购买地点：邮轮中心厦鼓码头→三丘田码头（湖里区）\n \
乘船方式：无需取票刷身份证即可登船 / 回程时间不限 / 如需要帮助请联系：0592 - 2968571，选完服装自行到酒店出示身份证即可办理入住\n \
\n \
第2天：{photo_day}\n \
拍摄 | 早上6：40在酒店前台领取早餐，酒店大厅集合候车7：00分发车前往厦鼓码头，下船后有工作人员接前往鼓浪屿门店\n \
\n \
第3天：{leave_day}{select_time} | 中航紫金广场门店选片（约90分钟）\n \
入住酒店 ：{hotel},{hotel_addr},（确认预定后，7天内酒店是无法更改或取消，临时取消属自动放弃）\n \
{checkin_day}下午14：00过后可办理入住，{checkout_day}中午12：00前退房\n \
铂爵旅拍门店地址：厦门市思明区环岛东路1813号中航紫金广场B塔铂爵旅拍1 - 5层\n \
\n \
需自备品：\n \
双方必备：身份证\
女士必备：蕾丝胸贴、乳（晕）头贴、浅色安全裤一条、平底鞋一双走路穿、修腋毛、防晒霜，防晒伞。\n \
男士必准备：理发、刮胡子、黑白纯色长袜、船袜各一双。\n \
\n \
为拍摄效果更好建议准备：\n \
女士：不要头发两截色、建议染栗色、亚麻色头发、做浅色水晶美甲、带美瞳，隐形眼镜\n \
门店鞋参考码：女士35 - 39码,男士40 - 44码,可自带或不带。'.format(arr_day=sht.cells(i,84).value, arr_time=sht.cells(i,43).value, photo_day=str(sht.cells(i,43).value)[:10],
                                                         leave_day=sht.cells(i,19).value, select_time=sht.cells(i,20).value, checkin_day=str(sht.cells(i,85).value)[:10],
                                                         checkout_day=str(sht.cells(i,86).value)[:10], hotel=sht.cells(i,87).value, hotel_addr=sht.cells(i,88).value
                                                             )
        set_text(announce)
        m.click(1853,43,1,1) # 进入群设置选项
        time.sleep(1)
        m.click(1751,224,1,1) # 点击群公告
        m.click(930,410,1,1) # 点击群公告编辑框
        time.sleep(1)
        simulate_key('ctrl','v')  #粘贴内容到获得焦点的输入框

        m.click(918,676,1,1) # 点击发布群公告
        time.sleep(2)
        m.click(996, 574, 1, 1)  # 确认公告发给所有人
        time.sleep(2)
        get_QRcode()


    set_name_announce()


wb.close()
app.quit()

