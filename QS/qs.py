#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common import exceptions
import pandas as pd
import time
import xlsxwriter
import os
import numpy as np


QS_YEAR = 2022
QS_1_SUBJECT_ROOT = f'./qs_{QS_YEAR}/一级学科'
QS_2_SUBJECT_ROOT = f'./qs_{QS_YEAR}/二级学科'
QS_1_RESULT = f'./qs_{QS_YEAR}/qs_{QS_YEAR}_一级学科_总表.xlsx'
QS_2_RESULT = f'./qs_{QS_YEAR}/qs_{QS_YEAR}_二级学科_总表.xlsx'

#init driver
chrome_options = Options()
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
driver = webdriver.Chrome(options=chrome_options)
curl = 'https://www.topuniversities.com/university-rankings/university-subject-rankings/2022/arts-humanities'
driver.get(curl)
time.sleep(10)

#init workbook
def initSecClassWorkBook(subjectName):
    if not os.path.exists(QS_2_SUBJECT_ROOT):
        os.makedirs(QS_2_SUBJECT_ROOT)
        if os.path.exists(QS_2_SUBJECT_ROOT):
            print('[info]: 文件创建成功')
        else:
            print('[error]: 文件创建失败')
    return xlsxwriter.Workbook(f'{QS_2_SUBJECT_ROOT}/qs_{QS_YEAR}_{subjectName}.xlsx')


def initFirstClassWorkBook(subjectName):
    if not os.path.exists(QS_1_SUBJECT_ROOT):
        os.makedirs(QS_1_SUBJECT_ROOT)
        if os.path.exists(QS_1_SUBJECT_ROOT):
            print('[info]: 文件创建成功')
        else:
            print('[error]: 文件创建失败')
    return xlsxwriter.Workbook(f'{QS_1_SUBJECT_ROOT}/qs_{QS_YEAR}_{subjectName}.xlsx')

# init sheet
def initSecClassSheet(workBook):
    Sheet = workBook.add_worksheet()
    #initial the table head
    Sheet.write(0, 0, 'Rank')
    Sheet.write(0, 1, 'University')
    Sheet.write(0, 2, 'Location')
    Sheet.write(0, 3, 'Overall Score')
    Sheet.write(0, 4, 'H-index Citations')
    Sheet.write(0, 5, 'Citations per Paper')
    Sheet.write(0, 6, 'Academic Reputation')
    Sheet.write(0, 7, 'Employer Reputation')
    Sheet.write(0, 8, 'Subject')
    return Sheet

def initFirstClassSheet(workBook):
    Sheet = workBook.add_worksheet()
    #initial the table head
    Sheet.write(0, 0, 'Rank')
    Sheet.write(0, 1, 'University')
    Sheet.write(0, 2, 'Location')
    Sheet.write(0, 3, 'Overall Score')
    Sheet.write(0, 4, 'International Research Network')
    Sheet.write(0, 5, 'H-index Citations')
    Sheet.write(0, 6, 'Citations per Paper')
    Sheet.write(0, 7, 'Academic Reputation')
    Sheet.write(0, 8, 'Employer Reputation')
    Sheet.write(0, 9, 'Subject')
    return Sheet

# switch to rankings indicators
def switchRankingsIndicators():
    switcher = driver.find_element(By.XPATH, '//*[@id="block-tu-d8-content"]/div/article/div/div/div[3]/div/div[1]/div/div[1]/div/div/ul/li[2]/a')
    driver.execute_script('arguments[0].click();', switcher)
    time.sleep(1)

# refresh subject
def switchSubject(subjectCnt):
    subjectSwitcher = driver.find_element(By.XPATH, '//*[@id="ranking-fillters"]/div[7]/div/div')
    driver.execute_script('arguments[0].click();', subjectSwitcher)
    menu = subjectSwitcher.find_elements_by_class_name('item')
    driver.execute_script('arguments[0].click();', menu[subjectCnt])
    time.sleep(10)

# switch to 100 results per pages
def changeResultsPerPageto100():
    resultPerPageSwitcher = driver.find_element(By.XPATH, '//*[@id="block-tu-d8-content"]/div/article/div/div/div[3]/div/div[1]/div/div[3]/div[4]/div[1]/div[2]/i')
    driver.execute_script('arguments[0].click();', resultPerPageSwitcher)
    oneHundredResultsPerPage = driver.find_element(By.XPATH, '//*[@id="block-tu-d8-content"]/div/article/div/div/div[3]/div/div[1]/div/div[3]/div[4]/div[1]/div[2]/div[2]/div[4]')
    driver.execute_script('arguments[0].click();', oneHundredResultsPerPage)
    time.sleep(5)

# get rank info of this second class subject
def recordingSecClassSubject():
    curline = 1
    SubjectSel = driver.find_element(By.XPATH, '//*[@id="ranking-data-load_ind"]')
    subject = driver.find_element(By.XPATH, '//*[@id="ranking-fillters"]/div[7]/div/div/div[1]').text
    workBook = initSecClassWorkBook(subject)
    Sheet = initSecClassSheet(workBook)
    # get ranking total number
    totalRankNum = (int)(driver.find_element(By.XPATH, '//*[@id="_totalcountresults"]').text)
    print("[totalRankNum of {}: {}]".format(subject, totalRankNum))
    # switch to 100 results per pages
    changeResultsPerPageto100()
    switchRankingsIndicators()
    # get rank info of this table
    finished = 0
    while True:
        try:
            tables = SubjectSel.find_elements_by_class_name("row.ind-row")
            for item in tables:
                rank = item.find_element_by_class_name("_univ-rank ").text
                if rank[0] == '=':
                    rank = rank[1:]
                Sheet.write(curline, 0, rank)
                university = item.find_element_by_class_name("uni-link").text
                Sheet.write(curline, 1, university)
                location = item.find_element_by_class_name("location ").text
                Sheet.write(curline, 2, location)
                subTables = item.find_elements_by_class_name("td-wrap-in")
                for i in range(1, len(subTables)):
                    Sheet.write(curline, i + 2, subTables[i].text)
                Sheet.write(curline, 8, subject)
                curline = curline + 1
            nextButt = driver.find_element_by_class_name('page-link.next')
            driver.execute_script('arguments[0].click();', nextButt)
            time.sleep(1)
            finished = finished + 100
            print('[processing: {} finished...]'.format(finished))
        except exceptions.NoSuchElementException:
            print("[学科{}已经搜索完毕]".format(subject))
            workBook.close()
            break


# get rank info of this first class subject
def recordingFirstClassSubject():
    curline = 1
    SubjectSel = driver.find_element(By.XPATH, '//*[@id="ranking-data-load_ind"]')
    subject = driver.find_element(By.XPATH, '//*[@id="ranking-fillters"]/div[7]/div/div/div[1]').text
    workBook = initFirstClassWorkBook(subject)
    Sheet = initFirstClassSheet(workBook)
    # get ranking total number
    totalRankNum = (int)(driver.find_element(By.XPATH, '//*[@id="_totalcountresults"]').text)
    print("[totalRankNum of {}: {}]".format(subject, totalRankNum))
    # switch to 100 results per pages
    changeResultsPerPageto100()
    switchRankingsIndicators()
    # get rank info of this table
    finished = 0
    while True:
        try:
            tables = SubjectSel.find_elements_by_class_name("row.ind-row")
            for item in tables:
                try:
                    rank = item.find_element_by_class_name("_univ-rank ").text
                    if rank[0] == '=':
                        rank = rank[1:]
                    Sheet.write(curline, 0, rank)
                    university = item.find_element_by_class_name("uni-link").text
                    Sheet.write(curline, 1, university)
                    location = item.find_element_by_class_name("location ").text
                    Sheet.write(curline, 2, location)
                    subTables = item.find_elements_by_class_name("td-wrap-in")
                    for i in range(1, len(subTables)):
                        Sheet.write(curline, i + 2, subTables[i].text)
                    Sheet.write(curline, 9, subject)
                    curline = curline + 1
                except BaseException as e:
                    print('[error]: {}排名{}数据存在问题:{},做跳过处理\n'.format(university, rank, e), list(map(lambda x:x.text, subTables)))
            nextButt = driver.find_element_by_class_name('page-link.next')
            driver.execute_script('arguments[0].click();', nextButt)
            finished = finished + 100
            print('[processing: {} finished...]'.format(finished))
        except exceptions.NoSuchElementException:
            print("[学科{}已经搜索完毕]".format(subject))
            #workBook.close()
            break


# get all subjects
def startRecording():
    firstClassSubjects = 5
    secondClassSubjects = 51
    try:
        for subjectCnt in range(firstClassSubjects):
            switchSubject(subjectCnt)
            recordingFirstClassSubject()
            print(subjectCnt)
        mergeAllParts(1)
        for subjectCnt in range(firstClassSubjects, secondClassSubjects):
            switchSubject(subjectCnt)
            recordingSecClassSubject()
            print(subjectCnt)
        mergeAllParts(2)
    except BaseException:
        print('[error, terminated]')


def mergeAllParts(classType):
    if classType == 1:
        root = QS_1_SUBJECT_ROOT
        result = QS_1_RESULT
    elif classType == 2:
        root = QS_2_SUBJECT_ROOT
        result = QS_2_RESULT
    else:
        return
    excel_names = []
    for excel_name in os.listdir(root):
        excel_names.append(excel_name)
    df_list = []
    for excel_name in excel_names:
        # 读取每个excel到df
        excel_path = f"{root}/{excel_name}"
        df_split = pd.read_excel(excel_path)
        df_list.append(df_split)
    df_merged = pd.concat(df_list)
    df_merged.to_excel(result, index=False)



