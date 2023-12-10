# from http import client
import os
import time
import pyodbc
import base64
import pathlib
import requests
import datetime
from time import sleep 
from turtle import update
from ssl import create_default_context
from xml.dom.minidom import Document 
import urllib.request
from xml.dom.minidom import parse
import xml.etree.ElementTree as ET
from datetime import date, timedelta

# Define constants
senderId = "LBMC"
senderPassword = "2t2fPXW!!&lt;9y"
amt = 0
TIMEOUT = 90
ENDPOINT_URL = "https://api.intacct.com/ia/xml/xmlgw.phtml"
DATABASE_DRIVER = 'ODBC Driver 17 for SQL Server'
INVOICE_SUB_STR = 'Benefits Invoice - September 2023.xlsx'
YEAR_MONTH = 'September 2023'
DB_CONFIG = {
    'server': 'lbmcbenefits.database.windows.net',
    'database': 'LBMCbenefits',
    'username': 'LBMC@lbmcbenefits',
    'password': '3fP3Z4AE69tgyOBoa3sF',
}

def establish_connection():
    connection_string = (
        f'DRIVER={{{DATABASE_DRIVER}}};'
        f'SERVER={DB_CONFIG["server"]};'
        f'DATABASE={DB_CONFIG["database"]};'
        f'UID={DB_CONFIG["username"]};'
        f'PWD={DB_CONFIG["password"]};'
    )

    # Establish a connection
    conn = pyodbc.connect(connection_string)
    print("Connection to SQL Server successful.")
    return conn

class XMLRequestClient:
    @staticmethod
    def post(request):
        header = {'Content-type': 'application/xml'}
        conn = urllib.request.Request(ENDPOINT_URL, headers = header, method='POST')

        result = urllib.request.urlopen(conn, request.toxml(encoding="ascii"), TIMEOUT)
        
        if result.getcode() != 200:
            raise Exception("Received HTTP status code, " + result.getcode())

        response = parse(result)
        return response

def get_session():
    controlid = str(time.time()).replace('.', '')
    companyid = "LBMC"
    userid = "DWReader"
    userpassword = "$KgWYS168TB"
    
    url = "https://api.intacct.com/ia/xml/xmlgw.phtml"
    header = {'Content-type': 'application/xml'}
    payload = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<request>\r\n  <control>\r\n    <senderid>"+senderId+"</senderid>\r\n    <password>"+senderPassword+"</password>\r\n    <controlid>"+controlid+"</controlid>\r\n    <uniqueid>false</uniqueid>\r\n    <dtdversion>3.0</dtdversion>\r\n    <includewhitespace>false</includewhitespace>\r\n  </control>\r\n  <operation>\r\n    <authentication>\r\n      <login>\r\n        <userid>"+userid+"</userid>\r\n        <companyid>"+companyid+"</companyid>\r\n        <password>"+userpassword+"</password>\r\n  <locationid>101</locationid>\r\n    </login>\r\n    </authentication>\r\n    <content>\r\n      <function controlid=\"1ee01cfe-aa00-4931-9731-f8591a0e54d2\">\r\n        <getAPISession />\r\n      </function>\r\n    </content>\r\n  </operation>\r\n</request>"     
    response = requests.request("POST", url, data=payload, headers=header)
    #print(response.text)
    root = ET.fromstring(response.content)

    for child in root.iter('sessionid'):
        session = (child.text)
    return session

def get_clients(cursor):
    sql = """
        set nocount on 
        select distinct ClientCode from dbo.[BILLING_STEP_3]
        order by 1 desc
    """
    cursor.execute(sql)
    row = cursor.fetchall() 
    return  row 

def get_detail(c, cursor):
    sql = """
    select premium, lastname + ', ' + firstname, period,
    case when [plan] like '%bcbs%dental%' then 'EP-BCBS-DENTAL'
        when [plan] like '%bcbs%(%)%' then 'EP-BCBS-HEALTH'
        when [plan] like '%bcbs%vision%' then 'EP-BCBS-VISION'
        when [plan] like '%cigna%p%/%' then 'EP-CIGNA-HEALTH'
        when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
        when [plan] like '%cigna%vision%' then 'EP-CIGNA-VISION'
        when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
        when [Provider Name] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
        when [Provider Name] like '%colonial%' and [plan] like '%critical%' then 'EP-COLONIAL-CRITICAL'
        when [Provider Name] like '%colonial%' and [plan] like '%accid%' then 'EP-COLONIAL-ACCIDENT'
        when [plan] like '%STANDARD LIFE%' AND [Provider Name] LIKE '%LINCOLN%' then 'EP-LINCOLN-STD'
        when [Provider Name] like '%lincoln%' and [plan] like '%long term disability%' then 'EP-LINCOLN-LTD'
        when [Provider Name] like '%lincoln%' and [plan] like '%vol%short%term disability%' then 'EP-LINCOLN-STD-VOL'
        when [Provider Name] like '%lincoln%' and [plan] like '%short%term disability%' then 'EP-LINCOLN-STD'
        when [Provider Name] like '%lincoln%' and [plan] like '%supplemental life ins%' then 'EP-LINCOLN-LIFE'
        END itemid
    , replace(replace(clientname,'.',''),'/','') clientname, clientcode, Customer_ID, project_ID
    from dbo.[BILLING_STEP_3]
    where clientcode = ? and 
     case when [plan] like '%bcbs%dental%' then 'EP-BCBS-DENTAL'
        when [plan] like '%bcbs%(%)%' then 'EP-BCBS-HEALTH'
        when [plan] like '%bcbs%vision%' then 'EP-BCBS-VISION'
        when [plan] like '%cigna%p%/%' then 'EP-CIGNA-HEALTH'
        when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
        when [plan] like '%cigna%vision%' then 'EP-CIGNA-VISION'
        when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
        when [Provider Name] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
		when [plan] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
        when [Provider Name] like '%colonial%' and [plan] like '%critical%' then 'EP-COLONIAL-CRITICAL'
		when [plan] like '%Colonial Critical Illness%' then 'EP-COLONIAL-CRITICAL'
		when [plan] like '%Colonial Life Group Critical Care%'  then 'EP-COLONIAL-CRITICAL'
        when [Provider Name] like '%colonial%' and [plan] like '%accid%' then 'EP-COLONIAL-ACCIDENT'
		when [plan] like '%Colonial Accident Plan%' or [plan] like '%Colonial Life Group Accident%' then 'EP-COLONIAL-ACCIDENT'
		when [plan] like '%Colonial Accident%' then 'EP-COLONIAL-ACCIDENT'
        when [plan] like '%STANDARD LIFE%' AND [Provider Name] LIKE '%LINCOLN%' then 'EP-LINCOLN-STD'
        when [Provider Name] like '%lincoln%' and [plan] like '%long term disability%' then 'EP-LINCOLN-LTD'
        when [Provider Name] like '%lincoln%' and [plan] like '%vol%short%term disability%' then 'EP-LINCOLN-STD-VOL'
        when [Provider Name] like '%lincoln%' and [plan] like '%short%term disability%' then 'EP-LINCOLN-STD'
        when [Provider Name] like '%lincoln%' and [plan] like '%supplemental life ins%' then 'EP-LINCOLN-LIFE'
        END IS NOT NULL
    """
    cursor.execute(sql,c)
    row = cursor.fetchall() 
    return  row 
 
def post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems,customer, clientID):
    newdoc = Document();
    request = newdoc.createElement('request')
    newdoc.appendChild(request)
    control = newdoc.createElement('control')
    request.appendChild(control)
    senderid = newdoc.createElement('senderid')
    control.appendChild(senderid).appendChild(newdoc.createTextNode(senderId))
    senderpassword = newdoc.createElement('password')
    control.appendChild(senderpassword).appendChild(newdoc.createTextNode("2t2fPXW!!<9y"))
    controlid = newdoc.createElement('controlid')
    control.appendChild(controlid).appendChild(newdoc.createTextNode("testRequestId"))
    uniqueid = newdoc.createElement('uniqueid')
    control.appendChild(uniqueid).appendChild(newdoc.createTextNode("false"))
    dtdversion = newdoc.createElement('dtdversion')
    control.appendChild(dtdversion).appendChild(newdoc.createTextNode("3.0"))
    
    operation = newdoc.createElement('operation')
    request.appendChild(operation) 
    authentication = newdoc.createElement('authentication')
    operation.appendChild(authentication) 
    sessionid = newdoc.createElement('sessionid')
    authentication.appendChild(sessionid).appendChild(newdoc.createTextNode(sessionId))

    content = newdoc.createElement('content')
    operation.appendChild(content)
    function = newdoc.createElement('function')
    content.appendChild(function).setAttributeNode(newdoc.createAttribute('controlid'))
    function.attributes["controlid"].value = "testFunctionId"
 
    createX = newdoc.createElement('create_supdoc')
    function.appendChild(createX)

    docidx = newdoc.createElement('supdocid')
    createX.appendChild(docidx).appendChild(newdoc.createTextNode(clientID+'Ben092023'))

    folderX = newdoc.createElement('supdocfoldername')
    createX.appendChild(folderX).appendChild(newdoc.createTextNode('EP Benefits Billing 092023'))

    folderDescriptionX = newdoc.createElement('supdocdescription')
    createX.appendChild(folderDescriptionX).appendChild(newdoc.createTextNode('Description of folder'))

    attachmentsX = newdoc.createElement('attachments')
    createX.appendChild(attachmentsX) 

    attachmentX = newdoc.createElement('attachment')
    attachmentsX.appendChild(attachmentX)

    attachmentNameX = newdoc.createElement('attachmentname')
    attachmentX.appendChild(attachmentNameX).appendChild(newdoc.createTextNode(customer + ' ' +clientID+' - Benefits Invoice - September 2023.xlsx'))
    
    attachmentTypeX = newdoc.createElement('attachmenttype')
    attachmentX.appendChild(attachmentTypeX).appendChild(newdoc.createTextNode('xlsx'))

    file_path = os.path.join("Andreas", f'{customer} {clientID} - {"Benefits Invoice - September 2023.xlsx"}')
    with open(file_path, 'rb') as file:
        data = file.read()
    encoded = base64.b64encode(data).decode('UTF-8')
 
    attachmentDataX = newdoc.createElement('attachmentdata')
    attachmentX.appendChild(attachmentDataX).appendChild(newdoc.createTextNode(str(encoded)))
     
    print(request.toprettyxml()) 
    result = XMLRequestClient.post(request) 
    xmlData = result.toprettyxml() 
    print(xmlData)
    print('Done')  
    try:
        cursor.execute("""
        insert into BillingLog (logmessage,clientid,projectid,amt,updatedate,billstage) 
        values (?,?,?,?,getdate(),'Post attachment to Intacct')
        """,xmlData,customerID, projectID, amt)
        conn.commit()
        print("Inserting, ", projectID)
    except Exception as e:
            print(e)
#Begin

def main():
    todays_date = date.today()
    createDate = todays_date
    cdMonth = todays_date.month
    cdYear = todays_date.year
    cdDay = todays_date.day 

    conn = establish_connection()
    cursor = conn.cursor()

    sessionId = get_session() 

    clients = get_clients(cursor)

    for c in clients:
        invoiceItems = get_detail(c, cursor)
        projectID = invoiceItems[0][7]
        customerID = invoiceItems[0][6]
        customer = invoiceItems[0][4]
        clientID = invoiceItems[0][5]
        post_data(conn, cursor,sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems, customer,clientID)

if __name__ == "__main__":
    main()