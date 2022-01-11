
# The script intends to read data from the excel file 
def read():
    import requests
    from requests.exceptions import MissingSchema
    import xlrd
    #stores the name of the sites from the excel as a list
    sites=[]
    # Hardcode the location of the excel file here
    file_location=r'C:\Users\Vaibhav\Desktop\PBL2.xls'
    # Opens the excel file
    workbook=xlrd.open_workbook(file_location)
    # Opens the FIRST sheet 
    sheet=workbook.sheet_by_index(0)
    # Counts the Number of row in the sheet
    n=sheet.nrows
    for i  in range (1,n):
        #Converts the name of the sites into an URL 
        # Stores the URL of the sites into the list
        site='http://www.'+sheet.cell_value(i,0)
        try:
            request = requests.get(site)
            if request.status_code == 200:
                sites.append(site)
        except:
            pass
    return(sites)
x=read()
print(x)




   
