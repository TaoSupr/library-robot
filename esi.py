import numpy as np
import pandas as pd

DATA_SOURCE = 'data.xlsx'
GENERATE_TABLES_RESULT = 'generate_tables.xlsx'
TABLE_CMP_RESULT = 'table_cmp_target.txt'

bestSubjectsOfHust = 0
best20Schools = 0
writer = 0

# get esi ranking top 20 schools
def getTop20():
    data_sheet = pd.read_excel(DATA_SOURCE, sheet_name=0, nrows=1000)
    data_of_school = data_sheet[data_sheet['学科'] == '全领域']
    ranking = data_of_school.sort_values(by=['ESI排名位置百分比'])
    return ranking[0:20]


# read table from file
def readFromFile(filePath):
    retDict = {}
    with open(filePath, mode='r') as f:
        curSubject = ''
        for line in f.readlines():
            flag = 0
            for word in line.split():
                if flag == 0:
                    retDict[word] = []
                    curSubject = word
                elif flag & 1:
                    retDict[curSubject].append(float(word))
                flag = flag + 1
    return retDict


# compare two tables
def cmptables(oldTablePath, newTablePath):
    oldTable = readFromFile(oldTablePath)
    newTable = readFromFile(newTablePath)
    with open(TABLE_CMP_RESULT, mode='w') as f:
        curline = ''
        for key, values in newTable.items():
            if not key in oldTable.keys():
                curline = key
                for i in range(len(values)):
                    curline = curline + '\t-'
            else:
                oldValues = oldTable[key]
                curline = key
                for i in range(min(len(values), len(oldValues))):
                    diff = values[i] - oldValues[i]
                    if diff % 1 == 0:
                        curline = curline + '\t' + '%.0f'%(diff)
                    else:
                        curline = curline + '\t' + '%.2f'%(diff)
            print(curline)
            f.write(curline + '\n')


def getBestSubjectsOfHust():
    data_sheet = pd.read_excel(DATA_SOURCE, sheet_name=0)
    data_sheet_of_hust = data_sheet[data_sheet['中文名称']=='华中科技大学'].sort_values(by='ESI排名位置百分比')
    data_sheet_of_hust.drop(columns=['中文名称'], inplace=True)
    allFiled = data_sheet_of_hust.loc[data_sheet_of_hust['学科']=='全领域'].copy()
    data_sheet_of_hust.drop(allFiled.index, inplace=True)
    data_sheet_of_hust = pd.concat([allFiled, data_sheet_of_hust], ignore_index=True)
    data_sheet_of_hust['ESI排名位置百分比'] = data_sheet_of_hust.apply(lambda x : '%.2f' % (float(x['ESI排名位置百分比'])*100) + '%', axis=1)
    return data_sheet_of_hust


def generateTable2():
    data_sheet_of_hust = bestSubjectsOfHust[['学科', '排名', 'ESI排名位置百分比', '大陆高校排名']]
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table2')


def generateTable3():
    data_sheet_of_hust = bestSubjectsOfHust[['学科', '排名', 'Web of Science Documents', 'Cites', 'Cites/Paper']]		
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table3')


def generateTable4():
    data_sheet_of_hust = bestSubjectsOfHust[['学科', 'Top Papers', 'Highly Cited Papers', 'Hot Papers']]		
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table4')


def generateTable5():
    data_sheet_of_hust = best20Schools[['中文名称', '排名', '优势学科数量', '前千分之一学科数', '前万分之一学科数']]		
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table5')


def generateTable6():
    data_sheet_of_hust = best20Schools[['中文名称', 'Web of Science Documents', 'Cites', 'Cites/Paper']]		
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table6')


def generateTable7():
    data_sheet_of_hust = best20Schools[['中文名称', 'Top Papers', 'Highly Cited Papers', 'Hot Papers']]		
    data_sheet_of_hust.index = np.arange(1, len(data_sheet_of_hust) + 1)
    data_sheet_of_hust.to_excel(writer, sheet_name='table7')


def generateTable8():
    data_sheet = pd.read_excel(DATA_SOURCE, sheet_name=0, usecols="A,P,Q,U")
    dict = {}
    curinfo = ''
    for school in best20Schools['中文名称'].tolist():
        data_sheet_sorted = data_sheet[data_sheet['中文名称'] == school].sort_values(by=['ESI排名位置百分比'])
        data_sheet_sorted.drop(data_sheet_sorted[data_sheet_sorted['学科'] == '全领域'].index.tolist(), axis=0, inplace=True)
        for index, row in data_sheet_sorted.iterrows():
            curinfo = curinfo + ('{}({}, {}%) '.format(row[1], row[0], '%.2f' % (row[3]*100)))
        dict[school] = curinfo
    df = pd.DataFrame.from_dict(dict, orient='index', columns=['info'])
    df.to_excel(writer, sheet_name='table8')


def generateAllTables():
    global best20Schools
    global bestSubjectsOfHust
    global writer
    best20Schools = getTop20()
    bestSubjectsOfHust = getBestSubjectsOfHust()
    writer = pd.ExcelWriter(GENERATE_TABLES_RESULT)
    generateTable2()
    generateTable3()
    generateTable4()
    generateTable5()
    generateTable6()
    generateTable7()
    generateTable8()
    writer.close()

#cmptables('oldtable.txt', 'newtable.txt')

generateAllTables()