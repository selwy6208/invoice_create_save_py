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
companyId = "LBMC"
userId = "DWReader"
userPassword = "$KgWYS168TB"
    
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

def establish_connection():
    connection_string = (
        f'DRIVER={{{DATABASE_DRIVER}}};'
        f'SERVER={DB_CONFIG["server"]};'
        f'DATABASE={DB_CONFIG["database"]};'
        f'UID={DB_CONFIG["username"]};'
        f'PWD={DB_CONFIG["password"]};'
    )

    try:
        with pyodbc.connect(connection_string) as conn:
            print("Connection to SQL Server successful.")
            return conn
    except Exception as e:
        print(f"Error establishing connection: {e}")
        raise

def get_session():
    controlId = str(time.time()).replace('.', '')

    header = {'Content-type': 'application/xml'}
    payload = f"""<?xml version="1.0" encoding="UTF-8"?>
        <request>
          <control>
            <senderid>{senderId}</senderid>
            <password>{senderPassword}</password>
            <controlid>{controlId}</controlid>
            <uniqueid>false</uniqueid>
            <dtdversion>3.0</dtdversion>
            <includewhitespace>false</includewhitespace>
          </control>
          <operation>
            <authentication>
              <login>
                <userid>{userId}</userid>
                <companyid>{companyId}</companyid>
                <password>{userPassword}</password>
                <locationid>101</locationid>
              </login>
            </authentication>
            <content>
              <function controlid="1ee01cfe-aa00-4931-9731-f8591a0e54d2">
                <getAPISession />
              </function>
            </content>
          </operation>
        </request>"""
    
    response = requests.request("POST", ENDPOINT_URL, data=payload, headers=header)
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

def get_detail(clientCode, cursor):
    sql = """
    SET NOCOUNT ON
    SELECT 
        premium, 
        lastname + ', ' + firstname as fullName,
        period,
        clientName, 
        ClientCode, 
        Customer_ID, 
        project_ID,
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
    from dbo.[BILLING_STEP_3]
    where clientcode = ? and 
        case when [plan] like '%bcbs%dental%' then 'EP-BCBS-DENTAL'
            when [plan] like '%bcbs%(%)%' then 'EP-BCBS-HEALTH'
            when [plan] like '%bcbs%vision%' then 'EP-BCBS-VISION'
            WHEN [Provider Name] LIKE '%BLUE CROSS BLUE%' AND [Description] LIKE '%LOAD PLAN%' THEN 'EP-BCBS-HEALTH'
            WHEN [Provider Name] LIKE '%BLUE CROSS BLUE%' AND [Description] LIKE '%VISION%' THEN 'EP-BCBS-VISION'
            WHEN [Provider Name] LIKE '%BLUE CROSS BLUE%' AND [Description] LIKE '%DENTAL%' THEN 'EP-BCBS-DENTAL'
            when [plan] like '%cigna%p%/%' then 'EP-CIGNA-HEALTH'
            WHEN [Provider Name] LIKE '%CIGNA%' AND [Description] LIKE '%VISION%' THEN 'EP-CIGNA-VISION'
            WHEN [Provider Name] LIKE '%CIGNA%' AND [Description] LIKE '%OAP%' THEN 'EP-CIGNA-HEALTH'
            when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
            when [plan] like '%cigna%vision%' then 'EP-CIGNA-VISION'
            when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
            when [Description] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
            WHEN [Description] LIKE '%cigna%health%' then 'EP-CIGNA-HEALTH'
            WHEN [Description] LIKE 'Cigna Heath%' THEN 'EP-CIGNA-HEALTH'
            when [Provider Name] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
            when [plan] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
            when [Provider Name] like '%colonial%' and [plan] like '%critical%' then 'EP-COLONIAL-CRITICAL'
            when [plan] like '%Colonial Critical Illness%' then 'EP-COLONIAL-CRITICAL'
            when [Provider Name] like '%colonial life%' and [Description] like '%critical illness%' then 'EP-COLONIAL-CRITICAL'
            when [plan] like '%Colonial Life Group Critical Care%'  then 'EP-COLONIAL-CRITICAL'
            when [Provider Name] like '%colonial%' and [plan] like '%accid%' then 'EP-COLONIAL-ACCIDENT'
            when [plan] like '%Colonial Accident Plan%' or [plan] like '%Colonial Life Group Accident%' then 'EP-COLONIAL-ACCIDENT'
            when [plan] like '%Colonial Accident%' then 'EP-COLONIAL-ACCIDENT'
            when [plan] like '%STANDARD LIFE%' AND [Provider Name] LIKE '%LINCOLN%' then 'EP-LINCOLN-STD'
            when [Provider Name] like '%lincoln%' and [plan] like '%long term disability%' then 'EP-LINCOLN-LTD'
            when [Provider Name] like '%lincoln%' and [plan] like '%vol%short%term disability%' then 'EP-LINCOLN-STD-VOL'
            when [Provider Name] like '%lincoln%' and [plan] like '%short%term disability%' then 'EP-LINCOLN-STD'
            when [Provider Name] like '%lincoln%' and [plan] like '%supplemental life ins%' then 'EP-LINCOLN-LIFE'
            when [Provider Name] like '%lincoln%' and [Description] like '%Supplemental Life Insurance and AD&D%' then 'EP-LINCOLN-LIFE'
            END IS NOT NULL
    """
    cursor.execute(sql, clientCode)
    row = cursor.fetchall() 
    return  row 
 
def post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, customer, clientID):
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
    attachmentX.appendChild(attachmentNameX).appendChild(
        newdoc.createTextNode(customer + ' ' +clientID+' - ' + INVOICE_SUB_STR)
    )
    
    attachmentTypeX = newdoc.createElement('attachmenttype')
    attachmentX.appendChild(attachmentTypeX).appendChild(newdoc.createTextNode('xlsx'))

    file_path = os.path.join("Andreas", f'{customer} {clientID} - {INVOICE_SUB_STR}')
    print(file_path, "file path testing")
    if (file_path):
        with open(file_path, 'rb') as file:
            data = file.read()
    encoded = base64.b64encode(data).decode('UTF-8')
 
    attachmentDataX = newdoc.createElement('attachmentdata')
    attachmentX.appendChild(attachmentDataX).appendChild(newdoc.createTextNode(str(encoded)))
     
    # print(request.toprettyxml()) 
    result = XMLRequestClient.post(request) 
    xmlData = result.toprettyxml() 
    # print(xmlData)
    # print('Done')  
    try:
        query = """
        INSERT INTO BillingLog (logmessage, clientid, projectid, amt, updatedate, billstage) 
        VALUES (?, ?, ?, ?, GETDATE(), 'Post attachment to Intacct')
        """
        cursor.execute(query, (xmlData, customerID, projectID, amt))
        conn.commit()
        # print("Inserting, ", projectID)
    except Exception as e:
        print(f"Error inserting into BillingLog: {e}")

def main():
    todaysDate = date.today()
    createDate = todaysDate
    cdMonth = todaysDate.month
    cdYear = todaysDate.year
    cdDay = todaysDate.day 

    conn = establish_connection()
    cursor = conn.cursor()

    sessionId = get_session() 

    clients = get_clients(cursor)

    for clientCode in clients:
        invoiceItems = get_detail(clientCode, cursor)
        projectID = invoiceItems[0][6]
        customerID = invoiceItems[0][5]
        customer = invoiceItems[0][3]
        clientID = invoiceItems[0][4]
        post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, customer, clientID)

if __name__ == "__main__":
    main()