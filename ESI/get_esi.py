from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import requests
import os

ROOT = './22SubjectsInfoESI/'


# 需要爬取的所有机构资料信息
institutionDict = {
    "华中科技大学": "HUAZHONG UNIVERSITY OF SCIENCE & TECHNOLOGY",
    "中国科学技术大学":"UNIVERSITY OF SCIENCE & TECHNOLOGY OF CHINA, CAS",
    "浙江大学":"ZHEJIANG UNIVERSITY",
    "西安交通大学":"XI'AN JIAOTONG UNIVERSITY",
    "上海交通大学":"SHANGHAI JIAO TONG UNIVERSITY",
    "清华大学":	"TSINGHUA UNIVERSITY",
    "南京大学":	"NANJING UNIVERSITY",
    "哈尔滨工业大学": "HARBIN INSTITUTE OF TECHNOLOGY",
    "复旦大学":	"FUDAN UNIVERSITY",
    "北京大学":	"PEKING UNIVERSITY",
    "重庆大学":	"CHONGQING UNIVERSITY",
    "中山大学":	"SUN YAT SEN UNIVERSITY",
    "山东大学": "SHANDONG UNIVERSITY",
    "南开大学": "NANKAI UNIVERSITY",
    "大连理工大学":"DALIAN UNIVERSITY OF TECHNOLOGY",
    "厦门大学":"XIAMEN UNIVERSITY",
    "东南大学":"SOUTHEAST UNIVERSITY - CHINA",
    "北京师范大学":"BEIJING NORMAL UNIVERSITY",
    "华中师范大学":"E CHINA NORMAL UNIV",
    "郑州大学":"ZHENGZHOU UNIVERSITY",
    "中国海洋大学":"OCEAN UNIVERSITY OF CHINA",
    "西北工业大学":"NORTHWESTERN POLYTECHNICAL UNIVERSITY",
    "西北农林科技大学":"NORTHWEST A&F UNIVERSITY - CHINA",
    "东北大学（中国）":"NORTHEASTERN UNIVERSITY - CHINA",
    "国防科技大学":"NATIONAL UNIVERSITY OF DEFENSE TECHNOLOGY - CHINA",
    "云南大学":"YUNNAN UNIVERSITY",
    "中国人民大学":"RENMIN UNIVERSITY OF CHINA",
    "新疆大学":"XINJIANG UNIVERSITY",
    "中央名族大学":"MINZU UNIVERSITY OF CHINA",
    "中南大学":	"CENTRAL SOUTH UNIVERSITY",
    "中国农业大学":	"CHINA AGRICULTURAL UNIVERSITY",
    "中国科学院大学":	"UNIVERSITY OF CHINESE ACADEMY OF SCIENCES, CAS",
    "中国地质大学":	"CHINA UNIVERSITY OF GEOSCIENCES",
    "西安电子科技大学":	"XIDIAN UNIVERSITY",
    "武汉理工大学":	"WUHAN UNIVERSITY OF TECHNOLOGY",
    "武汉大学":	"WUHAN UNIVERSITY",
    "同济大学":	"TONGJI UNIVERSITY",
    "天津大学":	"TIANJIN UNIVERSITY",
    "四川大学":	"SICHUAN UNIVERSITY",
    "兰州大学":	"LANZHOU UNIVERSITY",
    "吉林大学":	"JILIN UNIVERSITY",
    "华中农业大学":	"HUAZHONG AGRICULTURAL UNIVERSITY",
    "华南理工大学":	"SOUTH CHINA UNIVERSITY OF TECHNOLOGY",
    "湖南大学":	"HUNAN UNIVERSITY",
    "电子科技大学":	"UNIVERSITY OF ELECTRONIC SCIENCE & TECHNOLOGY OF CHINA",
    "北京邮电大学":	"BEIJING UNIVERSITY OF POSTS & TELECOMMUNICATIONS",
    "北京理工大学":	"BEIJING INSTITUTE OF TECHNOLOGY",
    "北京航空航天大学":	"BEIHANG UNIVERSITY",
    "牛津大学":	"UNIVERSITY OF OXFORD",
    "剑桥大学":	"UNIVERSITY OF CAMBRIDGE",
    "帝国理工学院":	"IMPERIAL COLLEGE LONDON",
    "新加坡国立大学":	"NATIONAL UNIVERSITY OF SINGAPORE",
    "南洋理工大学":	"NANYANG TECHNOLOGICAL UNIVERSITY",
    "约翰霍普金斯":	"JOHNS HOPKINS UNIVERSITY",
    "斯坦福大学":	"STANFORD UNIVERSITY",
    "麻省理工学院":	"MASSACHUSETTS INSTITUTE OF TECHNOLOGY (MIT)",
    "加州大学伯克利分校":	"UNIVERSITY OF CALIFORNIA BERKELEY",
    "哈佛大学":	"HARVARD UNIVERSITY",
    "多伦多大学":	"UNIVERSITY OF TORONTO",
    "香港科技大学 中国香港":	"HONG KONG UNIVERSITY OF SCIENCE & TECHNOLOGY",
    "香港大学 中国香港":	"UNIVERSITY OF HONG KONG",
    "台湾清华大学":	"NATIONAL TSING HUA UNIVERSITY",
    "台湾大学":	"NATIONAL TAIWAN UNIVERSITY",
    "中国科学院":	"CHINESE ACADEMY OF SCIENCES",
    "印度理工学院(BOMBAY)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - BOMBAY",
    "印度理工学院(DELHI)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - DELHI",
    "印度理工学院(DHARWAD)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - DHARWAD",
    "印度理工学院(GANDHINAGAR)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - GANDHINAGAR",
    "印度理工学院(GUWAHATI)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - GUWAHATI",
    "印度理工学院(HYDERABAD)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - HYDERABAD",
    "印度理工学院(INDORE)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - INDORE",
    "印度理工学院(JODHPUR)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - JODHPUR",
    "印度理工学院(KANPUR)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - KANPUR", 
    "印度理工学院(KHARAGPUR)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - KHARAGPUR",
    "印度理工学院(MADRAS)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - MADRAS",
    "印度理工学院(TIRUPATI)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - TIRUPATI",
    "印度理工学院(ROPAR)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - ROPAR",
    "印度理工学院(ROORKEE)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - ROORKEE",
    "印度理工学院(PATNA)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - PATNA",
    "印度理工学院(PALAKKAD)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - PALAKKAD",
    "印度理工学院(MANDI)":	"INDIAN INSTITUTE OF TECHNOLOGY (IIT) - MANDI",
    "马德里理工大学（UPM）":	"UNIVERSIDAD POLITECNICA DE MADRID",
    "米兰理工大学":	"POLYTECHNIC UNIVERSITY OF MILAN",
    "都灵理工大学":	"POLYTECHNIC UNIVERSITY OF TURIN",
    "新加坡科技研究局":	"AGENCY FOR SCIENCE TECHNOLOGY & RESEARCH (ASTAR)",
    "阿卜杜勒阿齐兹国王大学":	"KING ABDULAZIZ UNIVERSITY",
    "苏黎世联邦理工学院 （ETH）":	"ETH ZURICH",
    "瑞士保罗谢勒研究所（PSI）":	"PAUL SCHERRER INSTITUTE",
    "欧洲核子研究组织（CERN）":	"EUROPEAN ORGANIZATION FOR NUCLEAR RESEARCH (CERN)",
    "洛桑联邦理工学院（EPFL）":	"ECOLE POLYTECHNIQUE FEDERALE DE LAUSANNE",
    "日本国立研究与开发法人理化学研究所":	"RIKEN",
    "京都大学":	"KYOTO UNIVERSITY",
    "东京工业大学":	"TOKYO INSTITUTE OF TECHNOLOGY",
    "东北大学（日本）":	"TOHOKU UNIVERSITY",
    "大阪大学":	"OSAKA UNIVERSITY",
    "东京大学 曰本":	"UNIVERSITY OF TOKYO",
    "普林斯顿大学":	"PRINCETON UNIVERSITY",
    "美国农业部":	"UNITED STATES DEPARTMENT OF AGRICULTURE (USDA)",
    "美国能源部":	"UNITED STATES DEPARTMENT OF ENERGY (DOE)",
    "美国环境保护署":	"UNITED STATES ENVIRONMENTAL PROTECTION AGENCY",
    "美国国立卫生研究院":	"NATIONAL INSTITUTES OF HEALTH (NIH) - USA",
    "NSF":	"NATIONAL SCIENCE FOUNDATION (NSF)",
    "NASA":	"NATIONAL AERONAUTICS & SPACE ADMINISTRATION (NASA)",
    "代尔夫特理工大学（TUD）":	"DELFT UNIVERSITY OF TECHNOLOGY",
    "首尔大学 韩国":	"SEOUL NATIONAL UNIVERSITY (SNU)",
    "韩国科学技术研究院（Korea Institute of Science and Technology，简称KIST":	"KOREA INSTITUTE OF SCIENCE & TECHNOLOGY (KIST)",
    "法国国家科学研究中心":	"CENTRE NATIONAL DE LA RECHERCHE SCIENTIFIQUE (CNRS)",
    "亚琛工业大学（RWTH）":	"RWTH AACHEN UNIVERSITY",
    "慕尼黑工业大学（TUM）":	"TECHNICAL UNIVERSITY OF MUNICH",
    "马普学会":	"MAX PLANCK SOCIETY",
    "亥姆霍兹联合会":	"HELMHOLTZ ASSOCIATION",
    "柏林工业大学（TUB）":	"TECHNICAL UNIVERSITY OF BERLIN",
    "奥尔堡大学":	"AALBORG UNIVERSITY",
    "鲁汶大学":	"KU LEUVEN"
}

subjectNameList = [
    'CLINICAL MEDICINE',
    'MULTIDISCIPLINARY',
    'CHEMISTRY',
    'MATERIALS SCIENCE',
    'ENGINEERING',
    'BIOLOGY & BIOCHEMISTRY',
    'PHYSICS',
    'MOLECULAR BIOLOGY & GENETICS',
    'NEUROSCIENCE & BEHAVIOR',
    'ENVIRONMENT_ECOLOGY',
    'SOCIAL SCIENCES, GENERAL',
    'PLANT & ANIMAL SCIENCE',
    'GEOSCIENCES',
    'PSYCHIATRY_PSYCHOLOGY',
    'PHARMACOLOGY & TOXICOLOGY',
    'IMMUNOLOGY',
    'AGRICULTURAL SCIENCES',
    'COMPUTER SCIENCE',
    'MICROBIOLOGY',
    'ECONOMICS & BUSINESS',
    'SPACE SCIENCE',
    'MATHEMATICS'
]

def get_cookies():
    driver = webdriver.Chrome()
    driver.get("https://access.clarivate.com/login?app=esi")

    driver.find_element_by_xpath('//*[@id="mat-input-0"]').clear()
    driver.find_element_by_xpath('//*[@id="mat-input-0"]').send_keys("m202173712@hust.edu.cn") # 输入自己的ESI帐号
    driver.find_element_by_xpath('//*[@id="mat-input-1"]').clear()
    driver.find_element_by_xpath('//*[@id="mat-input-1"]').send_keys("FC7RQq2W8W9_q&V") # 输入自己的ESI密码
    driver.find_element_by_xpath('//*[@id="signIn-btn"]').click()

    # 等待窗口响应
    time.sleep(5)
    # 获取登录后的cookies
    cookies = driver.get_cookies()
    # 关闭浏览器
    driver.close()
    # print(cookies)
    cookie = {}
    # 将cookies保存成字典格式
    for items in cookies:
        cookie[items.get("name")] = items.get("value")
    return cookie


def makeRootFile():
    if not os.path.exists(ROOT):
        os.makedirs(ROOT)
        if os.path.exists(ROOT):
            print('[info]: 文件创建成功')
        else:
            print('[error]: 文件创建失败')

def str2html(s):
    return s.replace(' ', '%20').replace('&', '%26').replace('_', '%2F').replace(',', '%2C')


def recordingFrontName():
    cookies = get_cookies()
    for institutionC, institutionE in institutionDict.items():
        dictLi = []
        page = 0
        institutionInURL = str2html(institutionE)
        while True:
            start = page * 50
            page = page + 1
            requestUrl = "https://esi.clarivate.com/IndicatorsDataAction.action?&type=documents&author=&researchField="\
                + "&institution=" + institutionInURL \
                + "&journal=&territory=&article_UT=&researchFront=&articleTitle=&docType=Top&year="\
                + "&page=" + str(page) + "&start=" + str(start)\
                + "&limit=50&sort=%5B%7B%22property%22%3A%22citations%22%2C%22direction%22%3A%22DESC%22%7D%5D"     
            try:
                search_res = requests.get(requestUrl, cookies=cookies)
            except:
                cookies = get_cookies()
                continue
            if search_res.text == 'No results found':
                break
            re_text = search_res.json()
            list = re_text.get("data")
            print(institutionC + str(page * 50) +"processed")
            # 保存这一页的数据到字典indDict中
            for item in list:
                indDict = {}
                for key, value in item.items():
                    indDict[key] = value
                dictLi.append(indDict)
        # dict转换成DataFrame并保存到excel里
        print(institutionC + "has finished")
        df = pd.DataFrame(dictLi)
        df.to_excel(ROOT + institutionC + '.xlsx', index=False)


def recording20Subjects():
    cookies = get_cookies()
    dictLi = []
    for subject in subjectNameList:
        page = 0
        while True:
            start = page * 50
            page = page + 1
            requestUrl = "https://esi.clarivate.com/IndicatorsDataAction.action?&type=documents&author=" \
                + "&researchField=" + str2html(subject) \
                + "&institution=&journal=&territory=&article_UT=&researchFront=&articleTitle=&docType=Top&year="\
                + "&page=" + str(page) + "&start=" + str(start)\
                + "&limit=50&sort=%5B%7B%22property%22%3A%22citations%22%2C%22direction%22%3A%22DESC%22%7D%5D"     
            try:
                search_res = requests.get(requestUrl, cookies=cookies)
            except:
                cookies = get_cookies()
                continue
            if search_res.text == 'No results found':
                break
            re_text = search_res.json()
            list = re_text.get("data")
            print(subject + " " + str(page * 50) +" processed")
            # 保存这一页的数据到字典indDict中
            for item in list:
                indDict = {}
                for key, value in item.items():
                    indDict[key] = value
                if "researchFrontName" not in indDict.keys():
                    indDict["researchFrontName"] = ""
                if "hotpaper" not in indDict.keys():
                    indDict["hotpaper"] = ""
                del indDict["rowSeq"]
                dictLi.append(indDict)
        # dict转换成DataFrame并保存到excel里
        print(subject + " has finished")
        df = pd.DataFrame(dictLi)
        df.to_excel(ROOT + subject + '.xlsx', index=False)




def recordingTopPapers():
    cookies = get_cookies()
    for institutionC, institutionE in institutionDict.items():
        dictLi = []
        for subject in subjectNameList:
            page = 0
            while True:
                start = page * 50
                page = page + 1
                requestUrl = "https://esi.clarivate.com/IndicatorsDataAction.action?&type=documents&author=" \
                    + "&researchField=" + str2html(subject)\
                    + "&institution=" + str2html(institutionE) \
                    + "&journal=&territory=&article_UT=&researchFront=&articleTitle=&docType=Top&year="\
                    + "&page=" + str(page) + "&start=" + str(start)\
                    + "&limit=50&sort=%5B%7B%22property%22%3A%22citations%22%2C%22direction%22%3A%22DESC%22%7D%5D"     
                try:
                    search_res = requests.get(requestUrl, cookies=cookies)
                except:
                    cookies = get_cookies()
                    continue
                if search_res.text == 'No results found':
                    break
                re_text = search_res.json()
                list = re_text.get("data")
                print(institutionC + " " + subject + " " + str(page * 50) +" processed")
                # 保存这一页的数据到字典indDict中
                for item in list:
                    indDict = {}
                    for key, value in item.items():
                        indDict[key] = value
                    if "researchFrontName" not in indDict.keys():
                        indDict["researchFrontName"] = ""
                    if "hotpaper" not in indDict.keys():
                        indDict["hotpaper"] = ""
                    del indDict["rowSeq"]
                    dictLi.append(indDict)
        # dict转换成DataFrame并保存到excel里
        print(institutionC + " has finished")
        df = pd.DataFrame(dictLi)
        df.to_excel(ROOT + institutionC + '.xlsx', index=False)



if __name__ == '__main__':
    makeRootFile()
    recording20Subjects()

        






    