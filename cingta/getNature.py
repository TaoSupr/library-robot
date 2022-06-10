import json
import traceback
from selenium import webdriver
import xlsxwriter
import requests

def getJsonLoad(pageNo, pageSize=30):
    postLoad = {}
    postLoad["entityTag"] = {
        "entity":["A1574"],
        "tag":[]
    }
    postLoad["buriedPointParam"] = {
        "entityTag":{
            "label":"单位名称",
            "value":["华中科技大学"]
        },
        "year":{
            "label":"年度区间",
            "value":["1950","2022"]
        }
    }
    postLoad["year"] = ["1950","2022"]
    postLoad["type"] = []
    postLoad["indexid"]=17
    postLoad["onlySchool"] = False
    postLoad["pageNo"] = pageNo
    postLoad["pageSize"] = pageSize
    postLoad["orderby"] = ""
    return json.dumps(postLoad)

def get_cookies():
    # driver = webdriver.Chrome()
    # driver.get("https://beta.cingta.com/#/login")

    # driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[3]/div/div[2]/div[1]/div/input').clear()
    # driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[3]/div/div[2]/div[1]/div/input').send_keys("fangji@hust.edu.cn") # 输入自己的ESI帐号
    # driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[3]/div/div[2]/div[2]/div/input').clear()
    # driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[3]/div/div[2]/div[2]/div/input').send_keys("cingta2022") # 输入自己的ESI密码
    # driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[3]/div/div[2]/button').click()

    # # 等待窗口响应
    # time.sleep(3)
    # # 获取登录后的cookies
    # cookies = driver.get_cookies()
    # # 关闭浏览器
    # driver.close()
    # # print(cookies)
    cookie = {}
    cookie["absms_crm2_userName"]="fangji%40hust.edu.cn"    
    cookie["absms_crm2_password"]="cingta2022"
    cookie["authorization"] = "eyJhbGciOiJIUzUxMiJ9.eyJ0aW1lc3RhbXAiOjE2NTQ3NjE2MjksInVzZXJuYW1lIjoiZmFuZ2ppQGh1c3QuZWR1LmNuIn0.sVDqLpsNdq13P-NW3vaqWDq7s0MTyr9x6rBJYr8EVOA2u6_NcnPu28lQ4njPBVfpaXgF_sjfwZC4pOMcBLSfDg"
    cookie["sessionid"] = "g3vlxc5lsd2sy2ymfkilhxpo8bxxidx2"
    # 将cookies保存成字典格式
    # for items in cookies:
    #     cookie[items.get("name")] = items.get("value")
    return cookie


#init workbook
def __initWorkBook():
    workBook = xlsxwriter.Workbook('./国家自然科学基金.xlsx')
    print('[info]: workbook initialized...')
    return workBook

def __closeWorkBook(workBook):
    workBook.close()
    print('[info]: workbook closed...')


def getInfo():
    cookies = get_cookies()
    Workbook = __initWorkBook()
    sheet = Workbook.add_worksheet()
    curline = 1
    pageNo = 1
    while True:   
        try:
            search_res = requests.post('https://ud.cingta.com/basedata/get_maindata/', data=getJsonLoad(pageNo=pageNo), cookies=cookies)
            data = search_res.json().get("data")
            if pageNo == 1 and curline == 1:
                for i in range(len(data.get("columnData"))):
                    sheet.write(0, i, data.get("columnData")[i].get("label"))
                print('[info]: sheet initialized...')
            list = data.get("table")
            if len(list) == 0:
                break
            for item in list:
                linedata = item.get("linedata")
                for i in range(len(linedata)):
                    sheet.write(curline, i, linedata[i])
                curline = curline + 1
            print(f'[info]: {curline-1} lines processed...')
            pageNo = pageNo + 1
        except Exception as ex:
            traceback.print_exc()
            break
    __closeWorkBook(workBook=Workbook)
    return 

if __name__ == '__main__':
    getInfo()
