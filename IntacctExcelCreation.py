import os
from pydoc import cli
from sqlite3 import Cursor
from statistics import pvariance
from traceback import format_exc
from numpy import true_divide
import pyodbc
import pandas as pd 

server = 'lbmcbenefits.database.windows.net'
database = 'LBMCbenefits'
username = 'LBMC@lbmcbenefits'
password = '3fP3Z4AE69tgyOBoa3sF'

connection_string = (
    f'DRIVER={{ODBC Driver 17 for SQL Server}};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
)

# Establish a connection
conn = pyodbc.connect(connection_string)
print("Connection to SQL Server successful.")

# Create a cursor for executing SQL queries
cursor = conn.cursor()

def get_clients():
    sql = """
        set nocount on 
        select distinct clientcode from dbo.[BILLING_STEP_3]
        order by 1 desc
    """
    cursor.execute(sql)
    row = cursor.fetchall() 
    return  row  

def working(client):
    sql = pd.read_sql_query( '''
        set nocount on 
        SELECT 
            lastname + ', ' + firstname fullName
            , replace(replace(clientName,'.',''),'/','') clientName
            , [clientCode] 
            , [Period]
            , [Plan]
            , [Scenario]
            , [Provider Name]
            , [premium]
        FROM dbo.[BILLING_STEP_3]
        where clientcode = ? and
            case 
                when [plan] like '%bcbs%dental%' then 'EP-BCBS-DENTAL'
                when [plan] like '%bcbs%(%)%' then 'EP-BCBS-HEALTH'
                when [plan] like '%bcbs%vision%' then 'EP-BCBS-VISION'
                WHEN [Provider Name] LIKE '%BLUE CROSS BLUE%' THEN 'EP-BCBS-HEALTH'
                when [plan] like '%cigna%p%/%' then 'EP-CIGNA-HEALTH'
                WHEN [Provider Name] LIKE '%CIGNA%' THEN 'EP-CIGNA-VISION'
                when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
                when [plan] like '%cigna%vision%' then 'EP-CIGNA-VISION'
                WHEN [PLAN] LIKE '%DENTAL%' THEN 'EP-CIGNA-DENTAL' 
                when [plan] like '%cigna%dental%' then 'EP-CIGNA-DENTAL'
                when [Provider Name] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
                when [plan] like '%symet%' then  'EP-SYMETRA-INDEMNITY'
                when [Provider Name] like '%colonial%' and [plan] like '%critical%' then 'EP-COLONIAL-CRITICAL'
                when [plan] like '%Colonial Critical Illness%' then 'EP-COLONIAL-CRITICAL'
                when [Provider Name] like '%colonial life%' then 'EP-COLONIAL-CRITICAL'
                when [plan] like '%Colonial Life Group Critical Care%'  then 'EP-COLONIAL-CRITICAL'
                when [Provider Name] like '%colonial%' and [plan] like '%accid%' then 'EP-COLONIAL-ACCIDENT'
                when [plan] like '%Colonial Accident Plan%' or [plan] like '%Colonial Life Group Accident%' then 'EP-COLONIAL-ACCIDENT'
                when [plan] like '%Colonial Accident%' then 'EP-COLONIAL-ACCIDENT'
                when [plan] like '%STANDARD LIFE%' AND [Provider Name] LIKE '%LINCOLN%' then 'EP-LINCOLN-STD'
                when [Provider Name] like '%lincoln%' and [plan] like '%long term disability%' then 'EP-LINCOLN-LTD'
                when [Provider Name] like '%lincoln%' and [plan] like '%vol%short%term disability%' then 'EP-LINCOLN-STD-VOL'
                when [Provider Name] like '%lincoln%' and [plan] like '%short%term disability%' then 'EP-LINCOLN-STD'
                when [Provider Name] like '%lincoln%' and [plan] like '%supplemental life ins%' then 'EP-LINCOLN-LIFE'
                when [Provider Name] like '%lincoln%' then 'EP-LINCOLN-LIFE'
                END IS NOT NULL
            '''  , con=conn, params=client
        )

    df = pd.DataFrame(sql)

    df = df.sort_values(by=['fullName', 'Provider Name'])
    
    rows = df[df.columns[0]].count()
    client = df["clientName"].max()
    clientID = df["clientCode"].max()
    employees = df['fullName'].count() 

    detail = df[['fullName', 'Period', 'premium', 'Provider Name', 'Plan']]
    detail.rename(columns={'fullName':'EE','premium':'Premium'},inplace=True)

    sumByPlan = df.groupby(['Plan'],as_index=True).agg({'premium':'sum', 'fullName':'count'})
    sumByPlan.reset_index(inplace=True)
    sumByPlan.rename(columns={'fullName':'Employees','premium':'Amount'}, inplace=True)

    gb = df.groupby(['fullName']).sum()
    gb.reset_index(inplace=True) 
    piv = gb.pivot(index = 'fullName', columns='Plan', values='premium')
    piv.reset_index(inplace=True)
    piv.rename(columns={'fullName':'EE'}, inplace=True)

    loc = os.path.join("address", f'{client} {clientID} - {"Benefits Invoice - September 2023.xlsx"}')

    writer = pd.ExcelWriter(loc,  engine='xlsxwriter')

    sumByPlan.to_excel(writer, 'Summary', index=False, startrow= 7)
    format_Summary(sumByPlan=sumByPlan,client=client,clientID=clientID,writer=writer)

    piv.to_excel(writer,'Summary Detail', index=False, startrow= 7)
    format_SummaryDetail(writer=writer,piv=piv, client=client,clientID=clientID)

    detail.to_excel(writer, 'Detail', index=False, startrow=7 )
    format_Detail(writer=writer,detail=detail,client=client,clientID=clientID)

    writer.close() ## changed from writer.save, research: .save is deprecated with this version of python

def format_Detail(writer,detail, client, clientID): 
    workbook = writer.book
    worksheet = writer.sheets['Detail'] 
    money_fmt = workbook.add_format({'num_format': '$ #,##0.00', 'align': 'right'})
    worksheet.set_column('A:A', 35)
    worksheet.set_column('E:E', 12, money_fmt)

    chEnd = { 1:'A', 2:'B', 3:'C', 4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J', 11:'K',12:'L',13:'M',14:'N' }
    for cw in detail:
        column_width = max(detail[cw].astype(str).map(len).max(), len(cw))
        col_idx = detail.columns.get_loc(cw)
        writer.sheets['Detail'].set_column(col_idx, col_idx, column_width+5)

    tablerange = 'A8:'+ str(chEnd[detail.shape[1]])+ str(detail[detail.columns[0]].count()+9)
    column_settings = [{'header':column} if column =="fullName"   else {'header':column,  'total_function':'sum'} for column in detail.columns]
    worksheet.add_table(tablerange, { 
        'columns':column_settings,
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })
    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path)    
    
    bold = workbook.add_format({'bold': True,'font':15})
    green = workbook.add_format({'bold': True,'font':15, 'color':'7da53d'})
    worksheet.write('D2', client,bold)
    worksheet.write('D3', "Monthly Client Summary",green)
    worksheet.write('D4', "September 2023",green)
    worksheet.write('A6', client + ' (' + clientID + ')', bold)

def format_Summary(writer,sumByPlan, client, clientID): 
    workbook = writer.book
    worksheet = writer.sheets['Summary'] 
    moneyfmt = workbook.add_format({'num_format': 44, 'align': 'right'})
    nmbfmt = workbook.add_format({'num_format': '#,##0', 'align': 'right'}) 
    worksheet.set_column('B:B', 12, moneyfmt)
    worksheet.set_column('C:C', 12, nmbfmt)

    chEnd = {1:'A', 2:'B', 3:'C', 4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J', 11:'K',12:'L',13:'M',14:'N'}

    for cw in sumByPlan:
        column_width = max(sumByPlan[cw].astype(str).map(len).max(), len(cw))
        col_idx = sumByPlan.columns.get_loc(cw)
        writer.sheets['Summary'].set_column(col_idx, col_idx, column_width+5)

    tablerange = 'A8:'+ str(chEnd[sumByPlan.shape[1]])+ str(sumByPlan[sumByPlan.columns[0]].count()+9)
    column_settings = [{'header':column} if column =="Description"   else {'header':column,  'total_function':'sum'} for column in sumByPlan.columns]
    worksheet.add_table(tablerange, { 
        'columns':column_settings, 
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })
    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path) 
    
    bold   = workbook.add_format({'bold': True,'font':15})
    green  = workbook.add_format({'bold': True,'font':15, 'color':'7da53d'})
    worksheet.write('D2', "LBMC Employment Partners, LLC",bold)
    worksheet.write('D3', "Monthly Client Summary",green)
    worksheet.write('D4', "September 2023",green)
    worksheet.write('A6', client + ' (' + clientID + ')', bold)

def format_SummaryDetail(writer,piv, client, clientID): 
    workbook = writer.book
    worksheet = writer.sheets['Summary Detail'] 
    moneyfmt = workbook.add_format({'num_format': '$ #,##0.00', 'align': 'right'})
    nmbfmt = workbook.add_format({'num_format': '#,##0', 'align': 'right'}) 
    worksheet.set_column('B:J', 12, moneyfmt) 
    chEnd = {1:'A', 2:'B', 3:'C', 4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J', 11:'K',12:'L',13:'M',14:'N',15:'O',16:'P',17:'Q',18:'R',19:'S',20:'T',21:'U',22:'V',23:'W',24:'X',25:'Y',26:'Z'}

    for cw in piv:
        column_width = max(piv[cw].astype(str).map(len).max(), len(cw))
        col_idx = piv.columns.get_loc(cw)
        writer.sheets['Summary Detail'].set_column(col_idx, col_idx, column_width+5)
        
    tablerange = 'A8:'+ str(chEnd[piv.shape[1]])+ str(piv[piv.columns[0]].count()+9)
      
    column_settings = [{'header':column} if column =="EE"   else {'header':column,  'total_function':'sum'} for column in piv.columns]
    worksheet.add_table(tablerange, { 
        'columns':column_settings,
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })
    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path) 
    
    bold   = workbook.add_format({'bold': True,'font':15})
    green  = workbook.add_format({'bold': True,'font':15, 'color':'7da53d'})
    worksheet.write('D2', "LBMC Employment Partners, LLC",bold)
    worksheet.write('D3', "Monthly Client Summary",green)
    worksheet.write('D4', "September 2023",green)
    worksheet.write('A6', client + ' (' + clientID + ')', bold)
 
clients = get_clients()
for c in clients:
    working(c)