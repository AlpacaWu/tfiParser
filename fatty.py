import requests as rq
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import numpy as np
import openpyxl
import pandas as pd
if __name__ == "__main__":
    url='https://www.tfi.org.tw/BoxOfficeBulletin/weekly'
    driver=webdriver.Chrome(ChromeDriverManager().install())
    driver.get(url)
    result = driver.find_elements_by_xpath('//tr[@data-id]/td[5]/a')
    all_command=[r.get_attribute('onclick') for r in result]
    all_file=[]
    for i,a in enumerate(all_command):
        a_list=a.partition('getData(')
        tar=a_list[2][:-1]
        tar_list=tar.split(",")
        num=int(tar_list[0])
        if num <=310 and num >=65:
            fileurl=tar_list[1][1:-1]
            completed_filename='https://www.tfi.org.tw'+fileurl
            filename=fileurl.split("/")[-1]
            all_file.append(filename)
            #driver.get(completed_filename)
    #time.sleep(60)
    for i,f in enumerate(all_file):
        excel = openpyxl.load_workbook(r'.\fatty\{}'.format(f)) #這裡的橘色部分改成你下載檔案的路徑
        sheet=excel.worksheets[0]
        rows=sheet.rows
        content=[]
        for j,r in enumerate(rows):
            line=[column.value for k,column in enumerate(r)]
            content.append(line)
        try:
            content=np.array(content)[1:,[1,2,3,6,10,11]]
        except:
            content=np.array(content)[1:,[1,2,3,6,9,10]]
        if i==0:
            all_content=content
        else:
            all_content=np.concatenate((all_content,content),axis=0)
    names=[]
    for n in all_content[:,1]:
        names.append(str(n))
    unique,index,c=np.unique(names,return_index=True,return_counts=True)
    final=all_content[index,:]
    df=pd.DataFrame(final)
    df.to_csv('Result.csv')
       
        
        
