#encoding=utf-8
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common import exceptions
import time
import xlsxwriter

#init driver
chrome_options = Options()
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
driver = webdriver.Chrome(options=chrome_options)
curl = 'https://www.topuniversities.com/university-rankings/university-subject-rankings/2022/arts-humanities'
driver.get(curl)
time.sleep(10)

#init workbook
def initWorkBook(subjectName):
    return xlsxwriter.Workbook('qs_2022_' + subjectName + '.xlsx')

# init sheet
def initSheet(workBook):
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

# get rank info of this subject
def recordingCurSubject():
    curline = 1
    SubjectSel = driver.find_element(By.XPATH, '//*[@id="ranking-data-load_ind"]')
    subject = driver.find_element(By.XPATH, '//*[@id="ranking-fillters"]/div[7]/div/div/div[1]').text
    workBook = initWorkBook(subject)
    Sheet = initSheet(workBook)
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


# get all subjects
def startRecording():
    allSubjects = 56
    try:
        for subjectCnt in range(55, allSubjects):
            switchSubject(subjectCnt)
            recordingCurSubject()
            print(subjectCnt)
    except BaseException:
        print('[error, terminated]')

    

# close the data flow
startRecording()
driver.close()

