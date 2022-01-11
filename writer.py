import requests
from requests.exceptions import MissingSchema
import time
import reader 
import urllib.request,re
import xlrd
import xlwt
from xlwt import Workbook 
from xlutils.copy import copy
# Hardcode the address of the excel file here
file_location=r'C:\Users\Vaibhav\Desktop\PBL2'
wb = Workbook() 
sheet1 = wb.add_sheet('Sheet 1') 
#workbook=xlwt.open_workbook(file_location)
#writeworkbook=copy(workbook)
#sheet=writeworkbook.get_sheet(0)
url1=reader.read()
#print(url1)
sheet1.write(0,0,'Name')
sheet1.write(0,1,'Email')
sheet1.write(0,2,'Number')



for i  in range (len(url1)):
    start=time.time()
    Number=[]
    Email=[]
    a=url1[i]
    #print(a)
   # url1=(sheet.read(i,0))
    #print(url1)
    url=a
    
    try:
        f = urllib.request.urlopen(url)
        s = f.read().decode('utf-8')
                #REGEX for Email and Mobile Number
        Number=(re.findall(r"([+][9][1]|[9][1]|[0]){0,1}([7-9]{1})([0-9]{9})",s)[0:2])
        Email=(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,6}",s)[0:2])
    except:
        pass
            
    if Number==[] and Email==[]:
        try:
            # COUNTACT-US
            f = urllib.request.urlopen(url+'/Contact-us')
            s = f.read().decode('utf-8')
            Number=(re.findall(r"([+][9][1]|[9][1]|[0]){0,1}([7-9]{1})([0-9]{9})",s)[0:2])
            Email=(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,6}",s)[0:2])

            
        except:
            pass
    if Number==[] and Email==[]:
        try:
            # CONTACT
            f = urllib.request.urlopen(url+'/Contact')
            s = f.read().decode('utf-8')
            Number=(re.findall(r"([+][9][1]|[9][1]|[0]){0,1}([7-9]{1})([0-9]{9})",s)[0:2])
            Email=(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,6}",s)[0:2])

            
        except:
            pass
    if Number==[] and Email==[]:
        try:
            # Contact Details
            f = urllib.request.urlopen(url+'/Contact Details')
            s = f.read().decode('utf-8')
            Number=(re.findall(r"([+][9][1]|[9][1]|[0]){0,1}([7-9]{1})([0-9]{9})",s)[0:2])
            Email=(re.findall(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,6}",s)[0:2])

            
        except:
            pass
    if Number==[]:
        Number=['Not available']  
    if Email==[]:
        Email=['Not available']  
           
        
    print(Number,Email)
    # Write the data into the Excel file

    sheet1.write(i+1,0,url)
  
    sheet1.write(i+1,2,Email)
    for j in range(len(Number)):
        sheet1.write(i+1,3+j,Number[j])
    # Hardcode the address of the excel file  here
    wb.save(r'C:\Users\Vaibhav\Desktop\PBL2.xls')
