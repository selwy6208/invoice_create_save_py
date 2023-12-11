from http import client
import os
import time
import pyodbc 
import requests
import datetime
from time import sleep 
from turtle import update
import urllib.request
from xml.dom.minidom import Document 
from xml.dom.minidom import parse
from ssl import create_default_context
#from matplotlib import projections
import xml.etree.ElementTree as ET
from datetime import date, timedelta

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
    IIF(premium IS NULL, 0, premium) as premium
	, lastname + ', ' + firstname as fullName
	, period
	, Customer_ID
	, project_ID
    , ClientName
	, ClientCode 
    ,case when [plan] like '%bcbs%dental%' then 'EP-BCBS-DENTAL'
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
        END  itemid
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
    cursor.execute(sql, clientCode)
    row = cursor.fetchall() 
    return  row 
 
def post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems,customer,clientID,contacts):
    print(projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems,customer,clientID,contacts, "all data test")
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
    includewhitespace = newdoc.createElement('includewhitespace')
    control.appendChild(includewhitespace).appendChild(newdoc.createTextNode("false")) 
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

    createX = newdoc.createElement('create_sotransaction')
    function.appendChild(createX)
    ttype = newdoc.createElement('transactiontype')
    createX.appendChild(ttype).appendChild(newdoc.createTextNode('EP - BEN - Benefits Invoices'))

    dateCreatedX = newdoc.createElement('datecreated')
    createX.appendChild(dateCreatedX)
    
    createDateYear = newdoc.createElement('year')
    createDateMonth = newdoc.createElement('month')
    createDateDay = newdoc.createElement('day')

    dateCreatedX.appendChild(createDateYear).appendChild(newdoc.createTextNode(str(2023))) 
    dateCreatedX.appendChild(createDateMonth).appendChild(newdoc.createTextNode(str(8)))        #Month
    dateCreatedX.appendChild(createDateDay).appendChild(newdoc.createTextNode(str(22)))         #Day

    cust = newdoc.createElement('customerid')
    createX.appendChild(cust).appendChild(newdoc.createTextNode(str(customerID)))
    
    termDateX = newdoc.createElement('termname')
    createX.appendChild(termDateX).appendChild(newdoc.createTextNode('EFT on Due Date'))
    
    dateDueX = newdoc.createElement('datedue')
    createX.appendChild(dateDueX)
    
    dueDateYear = newdoc.createElement('year')
    dueDateMonth = newdoc.createElement('month')
    dueDateDay = newdoc.createElement('day')

    dateDueX.appendChild(dueDateYear).appendChild(newdoc.createTextNode(str(2023))) 
    dateDueX.appendChild(dueDateMonth).appendChild(newdoc.createTextNode(str(9)))               #Month
    dateDueX.appendChild(dueDateDay).appendChild(newdoc.createTextNode(str(4)))                 #Day
    
    messageX = newdoc.createElement('message')
    createX.appendChild(messageX).appendChild(newdoc.createTextNode('09/01/2023 - 09/30/2023'))
    if contacts == 0:
        shipToX = newdoc.createElement('shipto')
        createX.appendChild(shipToX)
        
        contactShipX = newdoc.createElement('contactname')  
        shipToX.appendChild(contactShipX).appendChild(newdoc.createTextNode(str(invoiceItems[0][3]))) 

        billToX = newdoc.createElement('billto')
        createX.appendChild(billToX)

        contactX = newdoc.createElement('contactname')
        billToX.appendChild(contactX).appendChild(newdoc.createTextNode(str(invoiceItems[0][3]))) 
 
    attachmentX = newdoc.createElement('supdocid')
    createX.appendChild(attachmentX).appendChild(newdoc.createTextNode( clientID+'Ben092023'))
 
    stateX = newdoc.createElement('state')
    createX.appendChild(stateX).appendChild(newdoc.createTextNode('Pending'))
 
    proj = newdoc.createElement('projectid')
    createX.appendChild(proj).appendChild(newdoc.createTextNode(str(projectID)))

    invoiceItemsX = newdoc.createElement('sotransitems')
    createX.appendChild(invoiceItemsX)

    for a in invoiceItems:
        lineItemsX = newdoc.createElement('sotransitem')
        invoiceItemsX.appendChild(lineItemsX)
       
        itemX = newdoc.createElement('itemid')
        lineItemsX.appendChild(itemX).appendChild(newdoc.createTextNode(str(a[5])))
        qtyX = newdoc.createElement('quantity')
        lineItemsX.appendChild(qtyX).appendChild(newdoc.createTextNode(str('1')))
        unitX = newdoc.createElement('unit')
        lineItemsX.appendChild(unitX).appendChild(newdoc.createTextNode(str('Each')))
        amtX = newdoc.createElement('price')
        lineItemsX.appendChild(amtX).appendChild(newdoc.createTextNode(str(a[0])))
        memoX = newdoc.createElement('memo')
        if a[1] is None:
            lineItemsX.appendChild(memoX).appendChild(newdoc.createTextNode(str(a[2])))  
        if a[2] is None:
            lineItemsX.appendChild(memoX).appendChild(newdoc.createTextNode(str(a[1])))  
        if a[2] is None and a[1] is None:
            lineItemsX.appendChild(memoX).appendChild(newdoc.createTextNode(str('')))  
        else:
            lineItemsX.appendChild(memoX).appendChild(newdoc.createTextNode(str(a[1] + ' ' + a[2])))   
        customFieldsX = newdoc.createElement('customfields')
        lineItemsX.appendChild(customFieldsX)
        customFieldX = newdoc.createElement('customfield')
        customFieldsX.appendChild(customFieldX)
        cfX = newdoc.createElement('customfieldname')
        customFieldX.appendChild(cfX).appendChild(newdoc.createTextNode('DATA_ID'))
        cfvX = newdoc.createElement('customfieldvalue')
        customFieldX.appendChild(cfvX).appendChild(newdoc.createTextNode(a[1]))
 
    print(request.toprettyxml()) 
    result = XMLRequestClient.post(request) 
    xmlData = result.toprettyxml() 
    print(xmlData)
    print('Done') 

    try:
        cursor.execute("""
        insert into BillingLog (logmessage,clientid,projectid,amt,updatedate,billstage) 
        values (?,?,?,?,getdate(),'Post invoice to Intacct')
        """,xmlData,customerID, projectID, a[0])
        conn.commit()
        print("Inserting, ", projectID) 
    except Exception as e:
            print(e)
    
    if """<status>failure</status>""" in xmlData and contacts == 0:
      print("Failure, trying again without project contacts")
    #   post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems, customer, clientID, 1)
      
    if """<status>failure</status>""" in xmlData and contacts == 1:
      print("Failure, moving on. Something else is failing.") 
    
#Begin
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
        projectID = invoiceItems[0][4]
        customerID = invoiceItems[0][3]
        customer = invoiceItems[0][5]
        clientID = invoiceItems[0][6]

        post_data(conn, cursor, sessionId, projectID, customerID, amt, createDate, cdYear, cdMonth, cdDay, invoiceItems, customer,clientID, 0)

if __name__ == "__main__":
    main()
 