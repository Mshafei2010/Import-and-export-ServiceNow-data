import requests
import pandas as pd
from dotenv import load_dotenv
import os

# URL to Read data from ServiceNow
def buildUrl (instanceName, api, table, query, limit):
    url = 'https://'+ instanceName + '.service-now.com' + api + table + "?" + "sysparm_query=" + query + "&sysparm_limit=" + limit
    return url

#Read data from ServiceNow
def readServiceNowData(url,username,password):
    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}
   
    # Do the HTTP request
    response = requests.get(url, auth=(username, password), headers=headers , verify=False)
    # Check for HTTP codes other than 200
    if response.status_code != 200:
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        exit()
    # Decode the JSON response into a dictionary and use the data
    responseJSON = response.json()['result']
    return responseJSON

#Upload data to ServiceNow
def postRecordToServiceNow(instanceName, table, record, username, password, verify=False):
    url = f"https://{instanceName}.service-now.com/api/now/table/{table}"
    headers = {"Content-Type":"application/json","Accept":"application/json"}
    try:
        resp = requests.post(url, auth=(username, password), headers=headers, json=record, verify=verify, timeout=15)
    except requests.exceptions.RequestException as e:
        print(f"POST error: {e}")
        return False, str(e)
    if resp.status_code in (200,201):
        return True, resp.json()
    return False, f"HTTP {resp.status_code}: {resp.text}"

def read_workbook(path):
    xls = pd.ExcelFile(path)
    sheets = {}
    for sheet_name in xls.sheet_names[:3]:
        df = xls.parse(sheet_name)
        sheets[sheet_name] = df.fillna('').to_dict(orient='records')
    return sheets


if __name__ == "__main__":
    ExcelFile = r"C:\Users\mshafei\Desktop\code\Book1.xlsx"
    sheets = read_workbook (ExcelFile)

    load_dotenv(r"C:\Users\mshafei\Desktop\code\variables.env")

    # Set the request parameters
    instanceName = os.getenv("INSTANCE_NAME") # "dev000305"
    api = os.getenv("API_PATH") #"/api/now/table/"
    table = os.getenv("SN_TABLE") #"sys_user"

    # Eg. User name="admin", Password="admin" for this code sample.
    user = os.getenv("SN_USER") 
    pwd = os.getenv("SN_PASS")

    #query = 'active%3Dtrue%5Estate%3D2'
    query = os.getenv("QUERY") #active=true
    limit = os.getenv("LIMIT") #'100'
    
    
    url = buildUrl(instanceName,api,table,query,limit)

    # Read data from ServiceNow
    responseJSON = readServiceNowData(url,user,pwd)
    for item in responseJSON:
        print(str(item['user_name']) + " " + str(item['phone']) + " , " + str(item['email']) )

    #Upload Data to ServiceNow
    for sheet_name, records in sheets.items():
        print(f"Sheet: {sheet_name}")
        for record in records:
            responseJSON = postRecordToServiceNow(instanceName, table, {'user_name': str(record['email']), 'phone': str(record['phone']) , 'email':str(record['email']), 'first_name':str(record['name']).split(" ")[0],'last_name':str(record['name']).split(" ")[1]}, user, pwd)

    # Read Updated data from ServiceNow
    responseJSON = readServiceNowData(url,user,pwd)
    for item in responseJSON:

        print(str(item['user_name']) + " " + str(item['phone']) + " , " + str(item['email']) )
