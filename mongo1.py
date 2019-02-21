import requests
import xlrd
from openpyxl import load_workbook
import pymongo
import pandas as pd
import time
#myclient= pymongo.MongoClient('mongodb://ankit:123454321@cluster0-shard-00-00-pafqz.azure.mongodb.net:27017,cluster0-shard-00-01-pafqz.azure.mongodb.net:27017,cluster0-shard-00-02-pafqz.azure.mongodb.net:27017/test?ssl=true&replicaSet=Cluster0-shard-0&authSource=admin&retryWrites=true')

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["mca"]
mycol = mydb["company Master Data"]
wb=load_workbook('sample.xlsx')
sht=wb.active
url="http://www.mca.gov.in/mcafoportal/viewCompanyMasterData.do"
headers = {
        'User-Agent': 'Mozilla/5.0'
    }
p=1
i=input('Enter number: ')
while 1:
    print i
    time.sleep(2)
    Name=sht['A'+str(i)].value
    if Name =='None':
        break
    print "Scraping "+Name +" Data"
    i=i+1
    #print Name
    s=requests.Session()
    #r=s.get(url,headers=headers)
    data={'altScheme':'CIN',
          'companyID':Name}
    try:
        r=s.post('http://www.mca.gov.in/mcafoportal/exportCompanyMasterData.do',headers=headers,params=data)
    except:
        time.sleep(10)
        try:
            r=s.post('http://www.mca.gov.in/mcafoportal/exportCompanyMasterData.do',headers=headers,params=data)
        except:
            print "error in scraping "+Name+" Data"
            continue
    try:
        df=pd.read_excel(xlrd.open_workbook(file_contents=r.content), engine='xlrd')
    except:
        time.sleep(5)
        try:
            r=s.post('http://www.mca.gov.in/mcafoportal/exportCompanyMasterData.do',headers=headers,params=data)
        except:
            print "error in scraping "+Name+" Data"
            continue
        try:
            df=pd.read_excel(xlrd.open_workbook(file_contents=r.content), engine='xlrd')
        except:
            print "error in scraping "+Name+" Data" 
            continue
            
        
    r=df.shape[0]
    values=[]
    #dict = {'CIN':df.loc[0,2],'Company/ LLP Name','ROC Code','Registration Number','Company Category','Company SubCategory','Class of Company ','Authorised Capital(Rs)','Paid up Capital(Rs)','Number of Members(Applicable in case of company without Share Capital)','Date of Incorporation','Registered Address','Email Id','Whether Listed or not','Date of last AGM','Date of Balance Sheet','Company Status(for efiling)'}
    dins=[]
    charges=[]
    #print df
    for m in range(0,17):
        #print df.iloc[m,2]
        values.append(df.iloc[m,2])

    for m in range(17,r):
        if df.iloc[m,0] == 'Charges':
            #print 'yes'
            c=m
        if df.iloc[m,0] == 'Directors/Signatory Details':
            d=m
    dic={}
    for m in range(c+2,d-1):
        dic={'Assests under charge':df.iloc[m,0],
        'Charge Amount':df.iloc[m,1],
        'Date of Creation':df.iloc[m,2],
        'Date of Modification':df.iloc[m,3],
        'Status':df.iloc[m,4]}
        charges.append(dic)
    din1=[]
    dic1={''}
    for m in range(d+2,r):
        dic1={'DIN/PAN':df.iloc[m,0],
        'Name':df.iloc[m,1],
        'Begin date':df.iloc[m,2],
        'End date':df.iloc[m,3]}
        din1.append(dic1)

    dict5 = {
    'CIN':values[0],
    'Company/ LLP Name':values[1],
    'ROC Code':values[2],
    'Registration Number':values[3],
    'Company Category':values[4],
    'Company SubCategory':values[5],
    'Class of Company ':values[6],
    'Authorised Capital(Rs)':values[7],
    'Paid up Capital(Rs)':values[8],
    'Number of Members(Applicable in case of company without Share Capital)':values[9],
    'Date of Incorporation':values[10],
    'Registered Address':values[11],
    'Email Id':values[12],
    'Whether Listed or not':values[13],
    'Date of last AGM':values[14],
    'Date of Balance Sheet':values[15],
    'Company Status(for efiling)':values[16],
    'Charges':charges,
    'Directors/Signatory Details':din1}

    try:
        x = mycol.insert_one(dict5)
        print Name +" Data Scraped successfuly"
    except:
        print "error in scraping "+Name+" Data"
        pass


#print list of the _id values of the inserted documents:
#print(x.inserted_ids)
