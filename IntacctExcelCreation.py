import os
import pyodbc
import sqlalchemy as sa
import pandas as pd 
from constants import *

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

def get_clients(cursor):
    sql = """
        set nocount on 
        select distinct ClientCode from dbo.[BILLING_STEP_3]
        order by 1 desc
    """
    cursor.execute(sql)
    row = cursor.fetchall()
    return row

def working(conn, client):
    sql = pd.read_sql_query("""
        set nocount on 
        SELECT 
            [Description]
            , [Plan]
			, [EmployeeId]
            , lastname + ', ' + firstname fullName
            , [EE]
            , [Period]
            , [ClientCode] 
            , [ClientName]
            , [Scenario]
            , [Provider Name]
			, [Amounts]
            , [Premium]
            , [Coverage]
        FROM dbo.[BILLING_STEP_3]
        where ClientCode = ?
        """  , con=conn, params=client
    )

    df = pd.DataFrame(sql)

    df = df.sort_values(by=['fullName', 'Provider Name'])
    
    rows = df[df.columns[0]].count()
    client = df["ClientName"].max()
    clientID = df["ClientCode"].max()
    employees = df['fullName'].count() 

    detail = df[['EE', 'Period', 'Premium', 'Provider Name', 'Plan', 'Coverage']]

    sumByPlan = df.groupby(['Description'],as_index=True).agg({'Amounts':'sum', 'fullName':'count'})
    sumByPlan.reset_index(inplace=True)
    sumByPlan.rename(columns={'fullName':'Employees','Amounts':'Amount'}, inplace=True)

    gb = df.groupby(['fullName', 'Description']).sum()
    gb.reset_index(inplace=True) 

    piv = gb.pivot(index = 'EE', columns='Description', values='Amounts')
    piv.reset_index(inplace=True)

    loc = os.path.join("Andreas", f'{client} {clientID} - {INVOICE_SUB_STR}')

    with pd.ExcelWriter(loc, engine='xlsxwriter') as writer:
        sumByPlan.to_excel(writer, 'Summary', index=False, startrow=7)
        format_Summary(sumByPlan=sumByPlan, client=client, clientID=clientID, writer=writer)

        piv.to_excel(writer, 'Summary Detail', index=False, startrow=7)
        format_SummaryDetail(writer=writer, piv=piv, client=client, clientID=clientID)

        detail.to_excel(writer, 'Detail', index=False, startrow=7)
        format_Detail(writer=writer, detail=detail, client=client, clientID=clientID)

    # conn.close()

def format_Detail(writer, detail, client, clientID): 
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

    tableRange = 'A8:' + str(chEnd[detail.shape[1]]) + str(detail[detail.columns[0]].count() + 9)
    column_settings = [{'header': column} if column == "fullName" else {'header': column, 'total_function': 'sum'}
                       for column in detail.columns]
    
    worksheet.add_table(tableRange, { 
        'columns': column_settings,
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })

    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path)    
    
    bold = workbook.add_format({'bold': True, 'font': 15})
    green = workbook.add_format({'bold': True, 'font': 15, 'color': '7da53d'})
    worksheet.write('D2', client, bold)
    worksheet.write('D3', "Monthly Client Summary", green)
    worksheet.write('D4', "September 2023", green)
    worksheet.write('A6', f"{client} ({clientID})", bold)

def format_Summary(writer, sumByPlan, client, clientID): 
    workbook = writer.book
    worksheet = writer.sheets['Summary'] 
    money_fmt = workbook.add_format({'num_format': 44, 'align': 'right'})
    nmb_fmt = workbook.add_format({'num_format': '#,##0', 'align': 'right'}) 
    worksheet.set_column('B:B', 12, money_fmt)
    worksheet.set_column('C:C', 12, nmb_fmt)

    chEnd = {
        1: 'A', 2: 'B', 3: 'C', 4: 'D',5: 'E', 6: 'F', 7: 'G', 
        8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M', 14: 'N'
    }

    for cw in sumByPlan:
        column_width = max(sumByPlan[cw].astype(str).map(len).max(), len(cw))
        col_idx = sumByPlan.columns.get_loc(cw)
        writer.sheets['Summary'].set_column(col_idx, col_idx, column_width + 5)

    tableRange = 'A8:' + str(chEnd[sumByPlan.shape[1]]) + str(sumByPlan[sumByPlan.columns[0]].count() + 9)

    column_settings = [{'header': column} if column == "Description" else {'header': column, 'total_function': 'sum'}
                       for column in sumByPlan.columns]
    
    worksheet.add_table(tableRange, { 
        'columns':column_settings, 
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })

    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path) 
    
    bold = workbook.add_format({'bold': True, 'font': 15})
    green = workbook.add_format({'bold': True, 'font': 15, 'color': '7da53d'})

    worksheet.write('D2', "LBMC Employment Partners, LLC", bold)
    worksheet.write('D3', "Monthly Client Summary", green)
    worksheet.write('D4', YEAR_MONTH, green)
    worksheet.write('A6', f"{client} ({clientID})", bold)

def format_SummaryDetail(writer,piv, client, clientID): 
    workbook = writer.book
    worksheet = writer.sheets['Summary Detail'] 
    money_fmt = workbook.add_format({'num_format': '$ #,##0.00', 'align': 'right'})
    nmbfmt = workbook.add_format({'num_format': '#,##0', 'align': 'right'}) 
    worksheet.set_column('B:J', 12, money_fmt) 

    chEnd = {
        1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 
        11: 'K', 12: 'L', 13: 'M', 14:'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 
        20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'
    }

    for cw in piv:
        column_width = max(piv[cw].astype(str).map(len).max(), len(cw))
        col_idx = piv.columns.get_loc(cw)
        writer.sheets['Summary Detail'].set_column(col_idx, col_idx, column_width + 5)
        
    tableRange = 'A8:' + str(chEnd[piv.shape[1]]) + str(piv[piv.columns[0]].count() + 9)
      
    column_settings = [{'header': column} if column == "EE" else {'header': column, 'total_function': 'sum'}
                       for column in piv.columns]
    
    worksheet.add_table(tableRange, { 
        'columns':column_settings,
        'autofilter': True,
        'total_row': True,
        'style': 'Table Style Medium 4'
    })

    script_directory = os.path.dirname(os.path.abspath(__file__))
    image_path = os.path.join(script_directory, 'assets', 'LBMC-EmpPartners-logo.png')

    worksheet.insert_image('A1', image_path) 
    
    bold = workbook.add_format({'bold': True, 'font': 15})
    green = workbook.add_format({'bold': True, 'font': 15, 'color': '7da53d'})
    worksheet.write('D2', "LBMC Employment Partners, LLC", bold)
    worksheet.write('D3', "Monthly Client Summary", green)
    worksheet.write('D4', "September 2023", green)
    worksheet.write('A6', f"{client} ({clientID})", bold)
 
def main():
    conn = establish_connection()
    cursor = conn.cursor()

    clients = get_clients(cursor)

    for client in clients:
        working(conn, client)

    # Close the database connection
    conn.close()

if __name__ == "__main__":
    main()