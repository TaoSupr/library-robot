from numpy import NaN
import pandas as pd
import os

DATA_SOURCE = 'analysis_esi_changes/'

filelist = ['201901.xlsx','201903.xlsx','201905.xlsx','201909.xlsx',
            '201911.xlsx','202001.xlsx','202005.xlsx','202007.xlsx',
            '202009.xlsx','202011.xlsx','202103.xlsx','202105.xlsx',
            '202107.xlsx','202109.xlsx','202203.xlsx','202205.xlsx']

def showDiffs():
    df = pd.read_excel(DATA_SOURCE + 'result.xlsx')
    for file in filelist:
        print("processing..." + file)
        dfnext = pd.read_excel(DATA_SOURCE + file)
        newData = dfnext[~dfnext['Full title'].isin(df['Full title'])]
        df = pd.concat([df, newData.loc[:, ['Full title', 'ISSN', 'EISSN']]], ignore_index=True)
        newCategoryName = dfnext[df['Full title'].isin(dfnext['Full title'])]
        newCategoryName = newCategoryName[['Full title', 'Category name']]
        categoryName = file[:6] + 'Category name'
        newCategoryName.rename(columns={'Category name': categoryName}, inplace=True)
        print("newCategory\n")
        print(newCategoryName.head())
        for index, row in newCategoryName.iterrows():
            df.loc[df['Full title']==row['Full title'], categoryName] = row[categoryName]
        print("df\n")
        print(df.head())
    df.to_excel(DATA_SOURCE + 'result.xlsx', index=False)


def unchanged():
    df = pd.read_excel(DATA_SOURCE + 'result.xlsx')
    df = df[~pd.isnull(df['201901Category name'])]
    df = df[~pd.isnull(df['202205Category name'])]
    df = df[df['201901Category name']==df['202205Category name']]
    df.to_excel(DATA_SOURCE + 'unchanged.xlsx', index=False)


def changed():
    df = pd.read_excel(DATA_SOURCE + 'result.xlsx')
    df = df[~pd.isnull(df['201901Category name'])]
    df = df[~pd.isnull(df['202205Category name'])]
    df = df[df['201901Category name']!=df['202205Category name']]
    df.to_excel(DATA_SOURCE + 'changed.xlsx', index=False)

def new():
    df = pd.read_excel(DATA_SOURCE + 'result.xlsx')
    df = df[pd.isnull(df['201901Category name'])]
    df = df[~pd.isnull(df['202205Category name'])]
    df.to_excel(DATA_SOURCE + 'new.xlsx', index=False)


def old():
    df = pd.read_excel(DATA_SOURCE + 'result.xlsx')
    df = df[~pd.isnull(df['201901Category name'])]
    df = df[pd.isnull(df['202205Category name'])]
    df.to_excel(DATA_SOURCE + 'old.xlsx', index=False)


if __name__ == "__main__":
    #print(os.path.abspath('./'))
    #unchanged()
    #changed()
    new()
    old()