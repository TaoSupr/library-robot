import traceback
import requests
import xlsxwriter
from pyquery import PyQuery as pq

#init workBook
def __initWorkBook():
    workBook =  xlsxwriter.Workbook('./QSUniversities.xlsx')
    print("[info]: workBook initialized")
    return workBook

# init sheet
def __initSheet(workBook):
    Sheet = workBook.add_worksheet()
    #initial the table head
    Sheet.write(0, 0, 'Rank')
    Sheet.write(0, 1, 'University')
    Sheet.write(0, 2, 'Location')
    Sheet.write(0, 3, 'Overall Score')
    Sheet.write(0, 4, 'Academic Reputation')
    Sheet.write(0, 5, 'Employer Reputation')
    Sheet.write(0, 6, 'Citations per Faculty')
    Sheet.write(0, 7, 'Faculty Student Ratio')
    Sheet.write(0, 8, 'International Students Ratio')
    Sheet.write(0, 9, 'International Faculty Ratio')
    Sheet.write(0, 10, 'International Research Network')
    Sheet.write(0, 11, 'Employment Outcomes')
    print("[info]: sheet initialized")
    return Sheet


def __closeWorkBook(workBook):
    workBook.close()
    print("[info]: WorkBook closed")

def record():
    workBook = __initWorkBook()
    sheet = __initSheet(workBook=workBook)
    queryList = ["overall_rank_dis", "uni", "city+location", "overall", "ind_76", "ind_77", "ind_73", 
                 "ind_36", "ind_14", "ind_18", "ind_15", "ind_3819456"]
    try:
        ret = requests.get("https://www.topuniversities.com/sites/default/files/qs-rankings-data/en/3816281_indicators.txt?rd6fcf")
        data = ret.json().get("data")
        curLine = 1
        for item in data:
            for i in range(len(queryList)):
                if i == 2:
                    qs = queryList[2].split("+")
                    city = item.get(qs[0])
                    location = item.get(qs[1])
                    if city != "":
                        city = pq(city).text()
                    if location != "":
                        location = pq(location).text()
                    sheet.write(curLine, 2, city + ", " + location)
                elif item.get(queryList[i]) != "":
                    sheet.write(curLine, i, pq(item.get(queryList[i])).text())
                else:
                    sheet.write(curLine, i, "")
            print(f"[info]: {curLine} processed...")
            curLine = curLine + 1
    except:
        traceback.print_exc()
    __closeWorkBook(workBook=workBook)


if __name__ == "__main__":
    record()