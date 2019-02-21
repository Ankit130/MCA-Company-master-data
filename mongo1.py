import requests
#import xlrd
from openpyxl import load_workbook
#import pandas as pd
import time
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
    i=i+1
    #print Name
    s=requests.Session()
    #r=s.get(url,headers=headers)
    data={'altScheme':'CIN',
          'companyID':Name}
    try:
        r=s.post('http://www.mca.gov.in/mcafoportal/exportCompanyMasterData.do',headers=headers,params=data)
    except:
        continue
    with open(Name+'.xls', 'wb') as output:
        output.write(r.content)
            
   #loc[0,2],'Company/ LLP Name','ROC Code','Registration Number','Company Category','Company SubCategory','Class of Company ','Authorised Capital(Rs)','Paid up Capital(Rs)','Number of Members(Applicable in case of company without Share Capital)','Date of Incorporation','Registered Address','Email Id','Whether Listed or not','Date of last AGM','Date of Balance Sheet','Company Status(for efiling)'}



#print list of the _id values of the inserted documents:
#print(x.inserted_ids)
