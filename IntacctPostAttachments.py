import os
import time
import pyodbc
import base64
import requests
from time import sleep 
from turtle import update
from ssl import create_default_context
from xml.dom.minidom import Document 
import urllib.request
from xml.dom.minidom import parse
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from constants import *

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
            <senderid>{SENDER_ID}</senderid>
            <password>{SNEDER_PASSWORD}</password>
            <controlid>{controlId}</controlid>
            <uniqueid>false</uniqueid>
            <dtdversion>3.0</dtdversion>
            <includewhitespace>false</includewhitespace>
          </control>
          <operation>
            <authentication>
              <login>
                <userid>{USER_ID}</userid>
                <companyid>{COMPANY_ID}</companyid>
                <password>{USER_PASSWORD}</password>
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

    session = None  # Provide a default value

    for child in root.iter('sessionid'):
        session = child.text

    if session is None:
        raise Exception("Session not found in the XML response")

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
        [Premium], 
        lastname + ', ' + firstname as fullName,
        [Period],
        [ClientName], 
        [ClientCode], 
        [Customer_ID], 
        [project_ID]
    from dbo.[BILLING_STEP_3]
    where ClientCode = ?
    """
    cursor.execute(sql, clientCode)
    row = cursor.fetchall() 
    return  row 
 
def post_data(conn, cursor, sessionId, projectID, customerID, amt, customer, clientID):
    newdoc = Document();
    request = newdoc.createElement('request')
    newdoc.appendChild(request)
    control = newdoc.createElement('control')
    request.appendChild(control)
    senderid = newdoc.createElement('senderid')
    control.appendChild(senderid).appendChild(newdoc.createTextNode(SENDER_ID))
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
        print("Inserting, ", projectID)
    except Exception as e:
        print(f"Error inserting into BillingLog: {e}")

def main():
    conn = establish_connection()
    cursor = conn.cursor()

    session_id = get_session()

    clients = get_clients(cursor)

    for client_code in clients:
        invoice_items = get_detail(client_code, cursor)
        project_id = invoice_items[0][6]
        customer_id = invoice_items[0][5]
        customer = invoice_items[0][3]
        client_id = invoice_items[0][4]
        post_data(conn, cursor, session_id, project_id, customer_id, AMT, customer, client_id)

if __name__ == "__main__":
    main()