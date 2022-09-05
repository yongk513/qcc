# coding:utf-8

import datetime
import requests
import xlrd
from bs4 import BeautifulSoup
from xlutils.copy import copy
import json
import urllib3
import threading
import time

urllib3.disable_warnings()

def get_config():
    with open('qccConfig.txt', 'r') as f:
        config = f.read()
    return config


# 线程锁
lock = threading.Lock()

proxy = []

# 创建一个空子典
proxies = {}

config = get_config()
jsonCon = eval(config)
fileName = jsonCon['FileName']
proxyUrl = jsonCon['dailiurl']
start = jsonCon['start']
end = jsonCon['end']
threadNumber = jsonCon['threadNumber']
proxyNumber = jsonCon['proxyNumber']
proxyErrNumber = jsonCon['proxyErrNumber']
errorNumber = jsonCon['errorNumber']


# IP代理池调度
def get_ip_pool():
    global proxy
    # 加锁
    lock.acquire(blocking=True, timeout=5)
    if len(proxy) <= 5:
        get_xdlIp(proxyUrl)
    if proxy[0]['count'] >= int(proxyErrNumber):
        proxy.pop(0)
    # 解锁
    lock.release()
    global proxies
    proxies = proxy[0]


# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='gbk')


# 读取excel
def read_excel(file_name):
    """
    Reads an excel file and returns a list of lists
    :param file_name:
    :return:
    """
    workbook = xlrd.open_workbook(file_name)
    worksheet = workbook.sheet_by_index(0)
    data = []
    for row in range(worksheet.nrows):
        data.append(worksheet.row_values(row))

    return data


def log(msg):
    # 追加日志文件
    # sys.stdout.flush()
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' ' + msg + '\n')
    with open('log.txt', 'a') as f:
        f.write(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ' ' + msg + '\n')



def get_xdlIp(url):
    bools = True
    while bools:
        html = requests.get(url)
        log("获取代理IP：" + html.text)
        html.encoding = 'utf-8'
        html = html.text
        html = json.loads(html)
        errorCode = html['ERRORCODE']
        if errorCode == '0':
            bools = False
            for i in html['RESULT']:
                proxyHost = i['ip']
                proxyPort = i['port']
                proxyMeta = "http://%(host)s:%(port)s" % {
                    "host": proxyHost,
                    "port": proxyPort,
                }
                proxies = {
                    "http": proxyMeta,
                    "https": proxyMeta,
                    "count": 0,
                    "use": 0
                }
                proxy.append(proxies)
        else:
            time.sleep(1)

    return proxies


# def get_ip(url):
#     html = requests.get(url)
#     log("获取代理IP：" + html.text)
#     html.encoding = 'utf-8'
#     html = html.text
#     html = json.loads(html)
#     proxyHost = html['data'][0]['ip']
#     proxyPort = html['data'][0]['port']
#     proxyMeta = "http://%(host)s:%(port)s" % {
#         "host": proxyHost,
#         "port": proxyPort,
#     }
#     proxies = {
#         "http": proxyMeta,
#         "https": proxyMeta
#     }
#     return proxies


def GetMiddleStr(content, startStr, endStr):
    startIndex = content.index(startStr)
    if startIndex >= 0:
        startIndex += len(startStr)
        endIndex = content[startIndex:].index(endStr)
        if endIndex >= 0:
            return content[startIndex:startIndex + endIndex]
        else:
            return ''

def get_html(url, proxies1):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:104.0) Gecko/20100101 Firefox/104.0',
        'cookie': 'qcc_did=e0a42459-48ba-4c68-8eae-6614464f98f2',
    }
    # log("代理IP：" + str(proxies1))
    html = requests.get(url, headers=headers, proxies=proxies1)
    html.encoding = 'utf-8'
    return html.text


def get_data(html):
    list = []
    soup = BeautifulSoup(html, 'lxml')
    title = soup.select("table")
    tr = title[0].select("tr")
    # print(len(tr))
    for i in range(len(tr)):
        json1 = {
            "OperName": "",
            "Name": "",
            "CreditCode": "",
            "ZCZB": "",
            "CLSJ": "",
        }
        name = tr[i].select("td")[2].select("div")[0].select("span")[0].select("a")[0].text
        try:
            shuju = tr[i].select("td")[2].select("div")[0].select("div")[4].select("div")[0].text.replace("\n",
                                                                                                          "|").replace(
                " ", "")
        except Exception as e:
            # print(e)
            continue
        arr = shuju.split("|")
        json1["Name"] = name
        json1["OperName"] = arr[1]
        try:
            json1["CreditCode"] = arr[4].replace("复制", "")
        except:
            json1["CreditCode"] = "无数据"
        json1["ZCZB"] = arr[2]
        json1["CLSJ"] = arr[3]
        list.append(json1)
    return list


def diff_list(excelData, list):
    gsName = excelData['Name'].strip()
    gsCreditCode = excelData['CreditCode'].strip()
    gsfr = excelData['excelfr'].strip()

    rData = {
        "统一社会信用代码": '',
        "结果": '',
        "企查查-公司名称": '',
        "企查查-法人": ''
    }
    bool = False
    for i, data in enumerate(list):
        # log("爬到的数据:" + str(i) + "、" + str(data))
        gsmc = data['Name'].strip()
        # 截取内容到括号前
        if gsmc.find("（") != -1:
            gsmc_l = gsmc[:gsmc.find("（")].strip()
            # print("gsmc_l")
            # print(gsmc_l)
        fr = data['OperName'].replace("负责人：", "").replace("法定代表人：", "").strip()
        code = data['CreditCode'].replace("统一社会信用代码：", "").strip()
        # print(data)
        if gsName == gsmc and gsCreditCode == code and gsfr == fr:
            list[i]['结果'] = "一级"
            list[i]['sort'] = 1
        elif gsName == gsmc and gsCreditCode == code and gsfr != fr:
            list[i]['结果'] = "二级"
            list[i]['sort'] = 2
        elif gsmc.find(gsName) >= 0 and gsCreditCode == code and gsfr == fr:
            list[i]['结果'] = "三级"
            list[i]['sort'] = 3
        elif gsmc.find(gsName) >= 0 and gsCreditCode == code and gsfr != fr:
            list[i]['结果'] = "四级"
            list[i]['sort'] = 4
        elif gsName == gsmc and gsCreditCode != code and gsfr == fr:
            list[i]['结果'] = "五级"
            list[i]['sort'] = 5
        elif gsName == gsmc and gsCreditCode != code and gsfr != fr:
            list[i]['结果'] = "六级"
            list[i]['sort'] = 6
        elif gsmc.find(gsName) >= 0 and gsCreditCode != code and gsfr == fr:
            list[i]['结果'] = "八级"
            list[i]['sort'] = 8
        elif gsmc.find(gsName) >= 0 and gsCreditCode != code and gsfr != fr:
            if gsmc.find("（") != -1:
                # print("gsmc_l222")
                # print(gsmc_l,gsmc)
                if gsName == gsmc_l and gsCreditCode != code and gsfr != fr:
                    list[i]['结果'] = "七级"
                    list[i]['sort'] = 7
                else:
                    list[i]['结果'] = "九级"
                    list[i]['sort'] = 9
            else:
                list[i]['结果'] = "九级"
                list[i]['sort'] = 9
        else:
            list[i]['结果'] = "十级"
            list[i]['sort'] = 10
    # 排序
    list.sort(key=lambda x: x['sort'])

    rData['统一社会信用代码'] = list[0]['CreditCode'].replace("统一社会信用代码：", "")
    rData['企查查-公司名称'] = list[0]['Name']
    rData['企查查-法人'] = list[0]['OperName'].replace("负责人：", "").replace("法定代表人：", "").strip()
    rData['结果'] = list[0]['结果']
    return rData


def insert_excel(data, row, col, wb):
    ws = wb.get_sheet(0)
    ws.write(row, col, data)


# 创建一个线程任务
def thread_task(i, data, wb):
    log('第' + str(i) + '条数据:' + data[5] + '----' + data[8])
    Name = data[5]
    newName = data[5]
    CreditCode = data[8]
    excelfr = data[14]

    excelData = {
        "Name": Name,
        "CreditCode": CreditCode,
        "excelfr": excelfr,
    }

    Name = Name.encode('gbk').decode('gbk')
    bool = True
    count = 0

    while bool:
        get_ip_pool()
        try:
            proxy[0]['use'] += 1
            str1 = get_html("https://www.qcc.com/web/search?key=" + Name + "&isTable=true", proxies)
            resultList = get_data(str1)
            rdata = diff_list(excelData, resultList)

            log("结果：" + str(rdata))

            insert_excel(rdata['统一社会信用代码'], i, 9, wb)
            insert_excel(rdata['企查查-公司名称'], i, 10, wb)
            insert_excel(rdata['企查查-法人'], i, 11, wb)
            insert_excel(rdata['结果'], i, 12, wb)
            insert_excel("https://www.qcc.com/web/search?key=" + newName + "&isTable=true", i, 13, wb)

            log("代理池IP数量：" + str(len(proxy)) + " ---- 正在使用：" + proxy[0]['http'] + " ---- 使用次数" + str(
                proxy[0]['use']) + " ---- 错误次数：" + str(proxy[0]['count']))
            proxy[0]['count'] = 0

            bool = False
        except Exception as e:
            proxy[0]['count'] += 1
            log("代理池IP数量：" + str(len(proxy)) + " ---- 正在使用：" + proxy[0]['http'] + " ---- 使用次数" + str(
                proxy[0]['use']) + " ---- 错误次数：" + str(proxy[0]['count']))
            # print("失败次数：" + str(count))
            count += 1
            bool = True
            if count >= int(errorNumber):
                bool = False


proxy_i = 0
if __name__ == '__main__':

    log("初始化配置文件")
    log("文件名称：" + fileName)
    log("代理地址：" + proxyUrl)
    log("开始行数：" + start)
    log("结束行数：" + end)
    log("设置代理数量：" + proxyNumber)
    log("设置线程数量：" + threadNumber)
    log("错误跳过次数：" + errorNumber)
    log("IP错误换IP次数" + proxyErrNumber)

    datas = read_excel(fileName)
    rb = xlrd.open_workbook(fileName)
    wb = copy(rb)
    # 获取sheet的行数
    row = rb.sheets()[0].nrows - 1
    tlist = list()
    for i, data in enumerate(datas):
        if int(end) > 0:
            if i > int(end):
                continue
        if i < int(start):
            continue

        if i % 100 == 0:
            wb.save(fileName)
            log("第" + start + "-" + str(i) + "条数据保存成功")
        #     proxy.pop(0)
        #
        # if len(proxy) <= 5:
        #     get_xdlIp(proxyUrl)

        t = threading.Thread(target=thread_task, args=(i, data, wb))
        tlist.append(t)

        if len(tlist) == int(threadNumber):
            for t in tlist:
                t.start()
            for t in tlist:
                t.join()
            tlist.clear()


        # log('第' + str(i) + '条数据:' + data[5] + '----' + data[8])
        # Name = data[5]
        # newName = data[5]
        # CreditCode = data[8]
        # excelfr = data[11]
        #
        # excelData = {
        #     "Name": Name,
        #     "CreditCode": CreditCode,
        #     "excelfr": excelfr,
        # }
        #
        # Name = Name.encode('gbk').decode('gbk')
        # bool = True
        # while bool:
        #     # str1 = get_html("https://www.qcc.com/web/search?key=" + Name + "&isTable=true", proxies)
        #     # resultList = get_data(str1)
        #     # rdata = diff_list(excelData, resultList)
        #     # print("----------------------------------------------------")
        #     # print(rdata)
        #     # print(i, rdata['统一社会信用代码'], rdata['结果'])
        #     # print("----------------------------------------------------")
        #     # insert_excel(rdata['统一社会信用代码'], i, 9, wb)
        #     # insert_excel(rdata['结果'], i, 10, wb)
        #     # wb.save(fileName)
        #     # bool = False
        #     try:
        #         str1 = get_html("https://www.qcc.com/web/search?key=" + Name + "&isTable=true", proxies)
        #         resultList = get_data(str1)
        #         rdata = diff_list(excelData, resultList)
        #         # print(rdata)
        #         log("结果：" + str(rdata))
        #         insert_excel(rdata['统一社会信用代码'], i, 9, wb)
        #         insert_excel(rdata['企查查-公司名称'], i, 10, wb)
        #         insert_excel(rdata['企查查-法人'], i, 11, wb)
        #         insert_excel(rdata['结果'], i, 12, wb)
        #         insert_excel("https://www.qcc.com/web/search?key=" + newName + "&isTable=true", i, 13, wb)
        #         wb.save(fileName)
        #         bool = False
        #     except Exception as e:
        #         print(e)
        #         proxies = get_ip(proxyUrl)
        #         bool = True
        #         # if str(e) == 'substring not found':
        #         #     bool = False
        # print('\n')
