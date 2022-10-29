import time
import pymssql as sql
import pandas as pd

# 就目前而言，get_data 和 re_data是两个互相独立的文件，需要结合起来，一键生成
t1 = time.time()

def get_data():
    # 连接数据库
    conn = sql.connect(server='192.168.30.61',user='bojueinner',
                       password='BoJue*!inner.2022', database='PLL_ERP_Co_01'
                       )

    # 定义游标
    cursor = conn.cursor()

    cursor.execute(
    '''
    SELECT
    t2.CustomId,--客户编号
    t1.CreateTime,-- 创建时间
    t2.InfoSaleId,-- 创建人ID
    t4.PersonName,--创建人姓名
    t5.DepartName, --所属部门
    t3.ShopName,--所属门店
    t6.CustSource,--来源渠道
    CONVERT(nvarchar(50),t1.OldCustName) as OldCustName,--老客户姓名
    CONVERT(nvarchar(50),t1.CustName) as CustName,--新客户姓名
    t1.Mob,--新客户电话
    t1.CustStatus,-- 是否急量 ，需要转换，0为非急量，1为急量
    CONVERT(nvarchar(50),t1.Remark) as Remark, -- 备注
    t1.R1Creator,-- 过滤组第一次沟通人
    t1.CustSynTime1,-- 第一次沟通时间
    CONVERT(nvarchar(50),t1.Remark1) as Remark1,--第一次沟通结果
    t1.CustSynTime2,-- 第二次沟通时间
    CONVERT(nvarchar(50),t1.Remark2) as Remark2,--第二次沟通结果
    t1.CustSynTime3,--第三次沟通时间
    CONVERT(nvarchar(50),t1.Remark3) as Remark3, -- 第三次沟通结果
    t2.InfoStatus,-- 是否有效 需要转换，0有效，1被合并，2无效，4主动放弃，9流失，3待定
    t8.PersonName,-- 销售姓名
    t2.CCommContent, -- 销售沟通内容
    T7.BillStatus, -- 订单状态 需转换 0未核销，1生效，2退单
    T7.ModifyTime -- 订单状态更新时间
    FROM bjFJSRecord t1
    LEFT JOIN bjCustomM t2 ON t1.Mob=t2.Mob
    LEFT JOIN dbo.bjshopsM T3 ON T1.ShopId=T3.ShopId
    LEFT JOIN comPerson T4 ON T2.InfoSaleId=T4.PersonId
    LEFT JOIN bjDepartNewExe T5 ON T4.NewDepartId=T5.ERPDepartId
    LEFT JOIN bjCustSource T6 ON T2.FromSourceID=T6.TypeId
    LEFT JOIN bjSaleOrderM T7 ON t2.CustomId = T7.CustomId
    LEFT JOIN comPerson T8 ON T2.SalerId = T8.PersonId
    WHERE
    t1.CreateTime > '2022-06-19'
    order by
    CreateTime
    '''
    )

    # 获取字段名
    title = []
    for i in range(len(cursor.description)):
        title.append(cursor.description[i][0])
    # print(len(title))


    # 获取内容
    data = []
    for row in cursor:
        data.append(list(row))
    conn.close()
    # pd.DataFrame(data).to_csv('test.csv',index=False,header=False,encoding='ANSI')
    return pd.DataFrame(data,columns=title)

# print(get_data())

def re_data():
    # path1 = r'F:\04_数字化\无代码\整体上线\ERP同步数据\SQL代码\转介绍状态查询.xlsx'
    path2 = r'C:\Users\Administrator\Desktop\伙伴云上传.csv'

    wb = get_data()

    data1 = wb.loc[:, 'CustStatus']
    data1.replace({'0': '非急量', '1': "急量"}, inplace=True)
    print('data1_done')

    data2 = wb.loc[:, 'InfoStatus']
    data2.replace({0: '有效', 1: "被合并", 2: "无效", 4: "主动放弃", 3: "待定", 9: "流失"}, inplace=True)
    print('data2_done')

    data3 = wb.loc[:, 'BillStatus']
    data3.replace({0: '未核销', 1: "生效", 2: "退单"}, inplace=True)
    print('data3_done')

    # print(wb.loc[:,['CustStatus','InfoStatus','BillStatus']])

    wb.columns = ['客户编号', '创建时间', '创建人ID', '创建人姓名', '所属部门', '所属门店', '来源渠道', '老客户姓名', '新客户姓名', '新客户电话',
                  '是否急量', '备注', '过滤组沟通人', '第一次沟通时间', '第一次沟通结果', '第二次沟通时间', '第二次沟通结果', '第三次沟通时间',
                  '第三次沟通结果', '是否有效', '销售姓名', '销售沟通内容', '订单状态', '订单状态更新时间']

    # 最后生成一个CSV文件，Excel有问题伙伴云无法识别
    wb.to_csv(path2, index=False, encoding='UTF-8')

re_data()
t2 = time.time()

print((t2-t1))