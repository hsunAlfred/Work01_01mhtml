import os
from pprint import pprint
import shutil
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
import time

def folders():
    try:
        #path = 'C:\\Users\\servi\\Desktop\\MZiQ\\03觀光' 
        path = 'C:\\Users\\servi\\Desktop\\MZiQ\\05如時電子客戶同業' 

        res = os.listdir(path)
        folder = []
        '''
        for r in res:
            if 'xlsx' in r:
                if not( os.path.exists(path + "\\" + r[:-5]) ):
                    os.mkdir(path + "\\" + r[:-5])
                    folder.append(path + "\\" + r[:-5])
        '''
        for r in res:
            for f in folder:
                if r[:-5] in f:
                    shutil.move(path + '\\' + r, f + '\\' + r)
    except Exception as e:
        print(e)

def sta():
    path = 'C:\\Users\\servi\\Desktop\\MZiQ'
    res = os.listdir(path)
    result = {}
    for r in res:
        if r in ['01如時客戶', '02KY', '03觀光', '04零售百貨', '05如時電子客戶同業']:
            resFolder = os.listdir(path + '\\' + r)
            result[r] = {}
            for fo in resFolder:
                if 'xlsx' in fo:continue
                resFile = os.listdir(path + '\\' + r + '\\' + fo)
                result[r][fo] = 0
                for fi in resFile:
                    if 'png' in fi:
                        try:
                            temp = int( fi.split('-')[1])
                        except:
                            temp = int( fi.split('-')[1][:-4])
                        result[r][fo] = temp if temp > result[r][fo] else result[r][fo]
                    if 'mhtml' in fi: 
                        try:
                            temp = int( fi.split('-')[1])
                        except:
                            temp = int( fi.split('-')[1][:-6])
                        result[r][fo] = temp if temp > result[r][fo] else result[r][fo]
    
    with pd.ExcelWriter('sta.xlsx') as writer:
        for k, v in result.items():
            ki, vi = [], []
            for kk, vv in v.items():
                ki.append(kk)
                vi.append(vv)
            df = pd.DataFrame({'co_id':ki, 'count':vi})
            df.to_excel(writer, sheet_name = k, index = False)

    return result

def tea0901():
    try:
        path = 'C:\\Users\\servi\\Desktop\\MZiQ\\茶0901'
        allF = os.listdir(path)
        folders = [f for f in allF if '.mhtml'not in f]
        files = [f for f in allF if '.mhtml' in f]
        
        for fi in files:
            for fo in folders:
                if fi.split('-')[0] == fo.split(' ')[0]:
                    shutil.move(path + '\\' + fi, path + '\\' + fo + '\\' + fi)
    except Exception as e:
        return e

def makeFolders():
    path = "C:\\Users\\User\\Desktop\\下載資料"
    res = os.listdir(path)
    need = ['01如時客戶', '02KY', '03觀光', '04零售', '05如時電子客戶同業']
    for n in need:
        folders = os.listdir( path + '\\' + n )
        for folder in folders:
            if '.' in folder:
                os.mkdir(path + '\\' + n + '\\' + folder.split('.')[0])
            shutil.move(path + '\\' + n + '\\' + folder, path + '\\' + n + '\\' + folder.split('.')[0])

def comnbineData():
    path = "C:\\Users\\servi\\Desktop\\MZiQ"
    res = os.listdir(path)
    need = ['01如時客戶', '02KY', '03觀光', '04零售百貨', '05如時電子客戶同業']
    #need = ['02KY']
    df = pd.DataFrame()
    investedCompany = []
    for n in need:
        folders = os.listdir( path + '\\' + n )
        for folder in folders:
            if '.xlsx' not in folder: 
                files = os.listdir( path + '\\' + n + '\\' + folder )
                for f in files:
                    if ".xlsx" in f:
                        fileLoc = path + '\\' + n + '\\' + folder + '\\' + f
                        df_read = pd.read_excel(fileLoc, sheet_name="All ShareHolders", skiprows=1)

                        df_temp = df_read[ ['Name', 'Type', 'Style', 'T/O', 'Assets Under Management ($MM)', 'Location', 'Shareholder type', 'Report Date'] ]
                        fn = f.split('.')[0].split(' ')[0]
                        investedCompany.extend( [fn]*df_temp.shape[0] )
                        df = pd.concat( [df, df_temp] )
    
        df['Invested Company'] = investedCompany
        #print(df)
        df.to_excel('temp_' + n + '.xlsx', index = False)

def setWebdriver():
    '''建立webdeiver物件，並禁止彈出式視窗'''
    chrome_options=webdriver.ChromeOptions()
    prefs={"profile.default_content_setting_values.notifications":2}
    chrome_options.add_experimental_option("prefs",prefs)
    '''背景執行'''
    chrome_options.add_argument("--headless")
    driver=webdriver.Chrome(options=chrome_options)
    return driver

def contactInfo(target):
    driver = setWebdriver()
    path = "C:\\Users\\servi\\Desktop\\MZiQ\\" + target
    folders = os.listdir(path)
    companyContactInfo = {}
    for folder in folders:
        if '.xlsx' in folder: continue
        files = os.listdir(path + '\\' + folder)
        for fi in files:
            if '.mhtml' in fi:
                url = path + '\\' + folder + '\\' + fi
                '''載入網站'''
                driver.get(url)
                driver.implicitly_wait(20)

                '''取得網頁原始碼'''
                soup=BeautifulSoup(driver.page_source,'html.parser')
                sp_title = soup.select('.title')[0].text
                #input(sp_title)
                if sp_title not in companyContactInfo.keys(): 
                    companyContactInfo[sp_title] = {'Company':[], 
                        'Contact Name':[], 'Title':[], 'Email':[]}

                sp_info = soup.select('.contact__info')
                n = 0
                temp = []
                for info in sp_info:
                    temp.append(info.text)
                    if '@' in str(info) or '--' in str(info):
                        if len(temp)!=3: continue
                        if temp[2] not in companyContactInfo[sp_title]['Email']:
                            companyContactInfo[sp_title]['Company'].append(sp_title)
                            companyContactInfo[sp_title]['Contact Name'].append(temp[0])
                            companyContactInfo[sp_title]['Title'].append(temp[1])
                            companyContactInfo[sp_title]['Email'].append(temp[2])
                        temp = []
                    n += 1
                #print(fi, n, n%3)
    driver.quit()

    n, companyCompare = 0, {"No":[], "Company":[]}
    with pd.ExcelWriter('Contact Not Choice_{}.xlsx'.format(target)) as writer:
        for k, v in companyContactInfo.items():
            n += 1
            companyCompare["No"].append(n)
            companyCompare["Company"].append(k)
            df = pd.DataFrame(v)
            df.to_excel(writer, sheet_name = str(n), index = False)
        df = pd.DataFrame(companyCompare)
        df.to_excel(writer, sheet_name = "對照表", index = False)
    
    contactChoice = { 'Company':[], 'Contact Name':[], 'Title':[], 'Email':[] }
    for k, v in companyContactInfo.items():
        contactChoice['Company'].append(v['Company'][0])
        contactChoice['Contact Name'].append(v['Contact Name'][0])
        contactChoice['Title'].append(v['Title'][0])
        contactChoice['Email'].append(v['Email'][0])

        #if v['Title'][1]
        p2 = v['Title'][1].lower() if len(v['Title'])>=2 else ""
        p4 = v['Title'][3].lower() if len(v['Title'])>=4 else ""
        p, a, vp = 'Portfolio'.lower(), 'Analyst'.lower(), 'Vice President'.lower()
        player2, player4 = [p in p2, a in p2, vp in p2], [p in p4, a in p4, vp in p4]

        if player2[0] and not player4[0]:
            if p2 != "":
                contactChoice['Company'].append(v['Company'][1])
                contactChoice['Contact Name'].append(v['Contact Name'][1])
                contactChoice['Title'].append(v['Title'][1])
                contactChoice['Email'].append(v['Email'][1])
        elif not player2[0] and player4[0]:
            if p4 != "": 
                contactChoice['Company'].append(v['Company'][3])
                contactChoice['Contact Name'].append(v['Contact Name'][3])
                contactChoice['Title'].append(v['Title'][3])
                contactChoice['Email'].append(v['Email'][3])
        elif player2[0] and player4[0]:
            if p2 != "":
                contactChoice['Company'].append(v['Company'][1])
                contactChoice['Contact Name'].append(v['Contact Name'][1])
                contactChoice['Title'].append(v['Title'][1])
                contactChoice['Email'].append(v['Email'][1])
        else:
            if player2[1] and not player4[1]:
                if p2 != "":
                    contactChoice['Company'].append(v['Company'][1])
                    contactChoice['Contact Name'].append(v['Contact Name'][1])
                    contactChoice['Title'].append(v['Title'][1])
                    contactChoice['Email'].append(v['Email'][1])
            elif not player2[1] and player4[1]:
                if p4 != "":
                    contactChoice['Company'].append(v['Company'][3])
                    contactChoice['Contact Name'].append(v['Contact Name'][3])
                    contactChoice['Title'].append(v['Title'][3])
                    contactChoice['Email'].append(v['Email'][3])
            elif player2[1] and player4[1]:
                if p2 != "":
                    contactChoice['Company'].append(v['Company'][1])
                    contactChoice['Contact Name'].append(v['Contact Name'][1])
                    contactChoice['Title'].append(v['Title'][1])
                    contactChoice['Email'].append(v['Email'][1])
            else:
                if player2[1] and not player4[1]:
                    if p2 != "":
                        contactChoice['Company'].append(v['Company'][1])
                        contactChoice['Contact Name'].append(v['Contact Name'][1])
                        contactChoice['Title'].append(v['Title'][1])
                        contactChoice['Email'].append(v['Email'][1])
                elif not player2[1] and player4[1]:
                    if p4 != "":
                        contactChoice['Company'].append(v['Company'][3])
                        contactChoice['Contact Name'].append(v['Contact Name'][3])
                        contactChoice['Title'].append(v['Title'][3])
                        contactChoice['Email'].append(v['Email'][3])
                elif player2[1] and player4[1]:
                    if p2 != "":
                        contactChoice['Company'].append(v['Company'][1])
                        contactChoice['Contact Name'].append(v['Contact Name'][1])
                        contactChoice['Title'].append(v['Title'][1])
                        contactChoice['Email'].append(v['Email'][1])
                else:
                    if p4 != "":
                        contactChoice['Company'].append(v['Company'][3])
                        contactChoice['Contact Name'].append(v['Contact Name'][3])
                        contactChoice['Title'].append(v['Title'][3])
                        contactChoice['Email'].append(v['Email'][3])
                    elif p2 != "":
                        contactChoice['Company'].append(v['Company'][1])
                        contactChoice['Contact Name'].append(v['Contact Name'][1])
                        contactChoice['Title'].append(v['Title'][1])
                        contactChoice['Email'].append(v['Email'][1])
    df = pd.DataFrame(contactChoice)
    df.to_excel('{}聯絡資訊.xlsx'.format(target), index = False)

def mhtmlTest(target):
    driver = setWebdriver()
    path = "C:\\Users\\servi\\Desktop\\MZiQ\\" + target
    folders = os.listdir(path)
    mhtml = {}
    for folder in folders:
        if '.xlsx' in folder: continue
        files = os.listdir(path + '\\' + folder)
        for fi in files:
            if '.mhtml' in fi:
                url = path + '\\' + folder + '\\' + fi
                '''載入網站'''
                driver.get(url)
                driver.implicitly_wait(20)

                '''取得網頁原始碼'''
                soup=BeautifulSoup(driver.page_source,'html.parser')
                sp_title = soup.select('.title')[0].text
                
                if sp_title not in mhtml.keys(): mhtml[sp_title] = ""
                mhtml[sp_title] += str(fi)[:-6] + ","
    df = pd.Series(mhtml)
    df.to_excel('mhtmlInfo_' + target + '.xlsx')
    #print(df)
    
    driver.quit()


def lastCombine():
    ne = ['01如時客戶', '02KY', '03觀光', '04零售百貨', '05如時電子客戶同業']
    sheetName = ['法人一覽表', '異常', 'Sheet1']

    title_summary, title_error, title_sheet = [], [], []
    summaryDict, errorDict, s1Dict = {}, {}, {}
    for n in ne:
        fName = "20200902 Potential investors list_{}.xlsx".format(n)
        path = "C://Users//servi//Desktop//MZiQ//整理0910//{}".format(fName)
        
        #if list(df_summary.columns)[0] not in summaryDict.keys():
            #summaryDict[list(df_summary.columns)[0]] = list(df_summary.columns)[1:]

        df_summary = pd.read_excel(path, sheet_name=sheetName[0])
        if title_summary == []: title_summary = list(df_summary.columns)[1:]
        rs, cs = df_summary.shape
        for r in range(rs):
            k = df_summary.iloc[r, 0]
            v = list(df_summary.iloc[r, 1:])
            if k not in summaryDict.keys():
                summaryDict[k] = v
            else:
                temp_IC = summaryDict[k][5].split(', ')
                temp_IN = v[5].split(', ')
                for temp in temp_IN:
                    if temp not in temp_IC: temp_IC.append(temp)
                summaryDict[k][5] = ', '.join(temp_IC)
        
        df_error = pd.read_excel(path, sheet_name=sheetName[1])
        if title_error == []: title_error = list(df_error.columns)[1:]
        rs, cs = df_error.shape
        for r in range(rs):
            k = df_error.iloc[r, 0]
            v = list(df_error.iloc[r, 1:])
            if k not in errorDict.keys(): errorDict[k] = v
        
        df_sheet = pd.read_excel(path, sheet_name=sheetName[2])
        if title_sheet == []: title_sheet = list(df_sheet.columns)[1:]
        rs, cs = df_sheet.shape
        for r in range(rs):
            k = df_sheet.iloc[r, 0]
            v = list(df_sheet.iloc[r, 1:])
            if k not in s1Dict.keys():
                s1Dict[k] = v
            else:
                temp_IC = s1Dict[k][0].split(', ')
                temp_IN = v[0].split(', ')
                for temp in temp_IN:
                    if temp not in temp_IC: temp_IC.append(temp)
                s1Dict[k][0] = ', '.join(temp_IC)

    with pd.ExcelWriter('MZiQ Info(Complete).xlsx') as writer:
        df = pd.DataFrame(summaryDict, index = title_summary).T
        df.to_excel(writer, sheet_name="MZiQ彙總表")
    
        df = pd.DataFrame(errorDict, index = title_error).T
        df.to_excel(writer, sheet_name="異常")

        df = pd.DataFrame(s1Dict, index = title_sheet).T
        df.to_excel(writer, sheet_name="對照表")

#測試用
def seleniumTest():
    driver = setWebdriver()
    path = 'C:\\Users\\User\\Desktop\\MZiQ Code\\04零售百貨'
    folders = os.listdir(path)
    companyContactInfo = {}
    for folder in folders:
        files = os.listdir(path + '\\' + folder)
        for f in files:
            if '.mhtml' in f:
                url = path + '\\' + folder + '\\' + f
                '''載入網站'''
                driver.get(url)
                driver.implicitly_wait(20)

                '''取得網頁原始碼'''
                soup=BeautifulSoup(driver.page_source,'html.parser')
                sp_title = soup.select('.title')[0].text
                #input(sp_title)
                if sp_title not in companyContactInfo.keys(): 
                    companyContactInfo[sp_title] = {'Company':[], 
                        'Contact Name':[], 'Title':[], 'Email':[]}

                sp_info = soup.select('.contact__info')
                n = 0
                temp = []
                for info in sp_info:
                    temp.append(info.text)
                    if '@' in str(info) or '--' in str(info):
                        companyContactInfo[sp_title]['Company'].append(sp_title)
                        companyContactInfo[sp_title]['Contact Name'].append(temp[0])
                        companyContactInfo[sp_title]['Title'].append(temp[1])
                        companyContactInfo[sp_title]['Email'].append(temp[2])
                        temp = []
                    n += 1
                print(n, n%3)
    driver.quit()

    n, companyCompare = 0, {"No":[], "Company":[]}
    with pd.ExcelWriter('Contact Not Choice.xlsx') as writer:
        for k, v in companyContactInfo.items():
            n += 1
            companyCompare["No"].append(n)
            companyCompare["Company"].append(k)
            df = pd.DataFrame(v)
            df.to_excel(writer, sheet_name = str(n), index = False)
        df = pd.DataFrame(companyCompare)
        df.to_excel(writer, sheet_name = "對照表", index = False)



def main():
    #sta()
    #tea0901()
    #comnbineData()
    start = time.time()
    ne = ['01如時客戶', '02KY', '03觀光', '04零售百貨', '05如時電子客戶同業']
    sheetName = ['法人一覽表', '異常', 'Sheet1']
    #ne = ['05如時電子客戶同業']

    '''
    for n in ne:
        print(n)
        contactInfo(n)
    end = time.time()
    print(end-start)
    
    for n in ne:
        print(n)
        mhtmlTest(n)
    end = time.time()
    print(end-start)
    '''

    lastCombine()
    
    #seleniumTest()
    print('done')

main()
