# Importing the necessary packages
import pandas as pd
import datetime
import win32com.client as win32
import mysql.connector

#
# Creating outlook connection
#
olApp = win32.Dispatch("Outlook.Application")
#
# Reading roster base file
#
roster = r'D:\Users\skathir\Desktop\PR Roster\Roster1.xlsx'
data = pd.read_excel(roster)
#
# Creating mysql connection
#
con = mysql.connector.connect(
    host="cmtlabs-db-pr.aka.amazon.com",
    user="pr_reader",
    password="************",
    database="price_rejects",
)
#
# Creating new DF
#
new = pd.DataFrame()
#
# Getting the present HC
#
column_headers = data.columns
for column_header in column_headers:
    try:
        if column_header.date() == datetime.date.today():
            result = data[column_header]
    except AttributeError:
        pass

df = pd.DataFrame(
    data=data['ID'],
)
df1 = pd.DataFrame(data=result)
df_total = [df, df1]
full = pd.concat(
    df_total,
    axis=1,
)
full.columns = ['Auditor_Id', 'Shift']
df3 = pd.DataFrame(
    data=full,
)
condition = df3[df3['Shift'] == 'GS']
present_HC = condition['Auditor_Id']
present_HC = present_HC.astype(str).str.cat(sep=', ')
#
# Getting the current pending listings from mysql database.
#
res = con.cursor()
Total_pending = "select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,category_code,competitor_url,competitor_price,competitorshipping_price,uploaded_date,status,super_status from price_rejects.price_reject_new where status='allocated';"
res.execute(Total_pending)
pending_1 = res.fetchall()
pending = len(pending_1)
print(pending)

inflow = 'select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,competitor_url,competitor_price,uploaded_date,mapper,mapped_date,new_url,error_type,issue_group,issue_desc,rootcause,resolution,comments,auditor,audited_date,status from price_rejects.price_reject_new where date(Uploaded_Date) between DATE_FORMAT(NOW(),"%Y-%m-01") AND NOW();'
res.execute(inflow)
inflow_pending = res.fetchall()
#
#
inflow_pending_df = pd.DataFrame(
    inflow_pending,
    columns=[
        'region',
        'merchant_id',
        'asin',
        'competitor',
        'reject_code',
        'recommended_price',
        'reject_date',
        'buyer_login',
        'gl',
        'competitor_url',
        'competitor_price',
        'uploaded_date',
        'mapper',
        'mapped_date',
        'new_url',
        'error_type',
        'issue_group',
        'issue_desc',
        'rootcause',
        'resolution',
        'comments',
        'auditor',
        'audited_date',
        'status',
    ],
)
#
# Spliting the uploaded date from datetime.
#
inflow_pending_df['Date'] = inflow_pending_df['uploaded_date'].dt.date
#
# Sorting the values based on date.
#
u_date = inflow_pending_df['Date'].value_counts().sort_index()
#
# Getting the Audited listings in current month from mysql database.
#
audited = "select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,competitor_url,competitor_price,uploaded_date,mapper,mapped_date,new_url,error_type,issue_group,issue_desc,rootcause,resolution,comments,auditor,audited_date,status from price_rejects.price_reject_new where auditor!='audi' and date(audited_Date) between DATE_FORMAT(NOW(),'%Y-%m-01') AND NOW();"
res.execute(audited)
audited_counts = res.fetchall()
audited_counts_df = pd.DataFrame(
    audited_counts,
    columns=[
        'region',
        'merchant_id',
        'asin',
        'competitor',
        'reject_code',
        'recommended_price',
        'reject_date',
        'buyer_login',
        'gl',
        'competitor_url',
        'competitor_price',
        'uploaded_date',
        'mapper',
        'mapped_date',
        'new_url',
        'error_type',
        'issue_group',
        'issue_desc',
        'rootcause',
        'resolution',
        'comments',
        'auditor',
        'audited_date',
        'status',
    ],
)
#
# Spliting the audited date from datetime.
#
audited_counts_df['audited_date1'] = audited_counts_df['audited_date'].dt.date
#
# Sorting the values based on date.
#
a_date = audited_counts_df['audited_date1'].value_counts().sort_index()
#
audi_result = 'select region,merchant_id,asin,competitor,reject_code,recommended_price,reject_date,buyer_login,gl,competitor_url,competitor_price,uploaded_date,mapper,mapped_date,new_url,error_type,issue_group,issue_desc,rootcause,resolution,comments,auditor,audited_date,status from price_rejects.price_reject_new where auditor="audi" and date(audited_Date) between DATE_FORMAT(NOW(),"%Y-%m-01") AND NOW();'
res.execute(audi_result)
audi_result = res.fetchall()
audi_result_df = pd.DataFrame(
    audi_result,
    columns=[
        'region',
        'merchant_id',
        'asin',
        'competitor',
        'reject_code',
        'recommended_price',
        'reject_date',
        'buyer_login',
        'gl',
        'competitor_url',
        'competitor_price',
        'uploaded_date',
        'mapper',
        'mapped_date',
        'new_url',
        'error_type',
        'issue_group',
        'issue_desc',
        'rootcause',
        'resolution',
        'comments',
        'auditor',
        'audited_date',
        'status',
    ],
)
#
# Getting the AUDI Tool listings.
#
audi_result_df['audited_date2'] = audi_result_df['audited_date'].dt.date
audi = audi_result_df['audited_date2'].value_counts().sort_index()
#
# Seeting the ID column as Index from the roster.
#
data1 = data.set_index('ID', inplace=True)
data1 = data.transpose()
data1.index = pd.to_datetime(data1.index, format='%m/%d/%Y')
#
# Getting the Current HC column from the roster.
#
data1['Rostered_HC'] = (data1[['singhecy', 'hemalj', 'hgoyal']] == 'GS').sum(axis=1)
#
# Creating the capacity based on HC's.
#
data1['Current_Capacity'] = data1['Rostered_HC'] * 90
#
# Creat the new column from the DF & Adding the values.
#
new['Inflow_No'] = u_date
new['Audited_No'] = a_date
new['AUDI_No'] = audi
new['Pending_for_Day'] = new['Inflow_No'] - (new['Audited_No'] + new['AUDI_No'])
new['Pending_for_Day'] = new['Pending_for_Day'].astype(int)
new['Coverage_%'] = ((new['Audited_No'] / new['Inflow_No']) * 100).map(
    "{:,.2f}%".format
)
#
new_list = []
for i in new['Pending_for_Day']:
    c = c + i
    new_list.append(c)
new['Pending_for_Month'] = new_list

data1.index.names = ['Date']
df3 = pd.concat([new, data1], axis=1).reindex(new.index)
df4 = pd.DataFrame(
    data=df3,
    columns=[
        'Inflow_No',
        'AUDI_No',
        'Audited_No',
        'Pending_for_Day',
        'Pending_for_Month',
        'Rostered_HC',
        'Current_Capacity',
        'Coverage_%',
    ],
)
#
# Convert DF to HTML format.
#
table_html = df4.to_html(table_id='Summary')
#
# Read the relevant text file for the outlook body.
#
with open(
    r'D:\Users\skathir\PycharmProjects\PR_Inflow_Tracker\table style.txt'
) as file:
    table_style = file.read()

with open(
    r'D:\Users\skathir\PycharmProjects\PR_Inflow_Tracker\before table html.txt'
) as body_file:
    body_html = body_file.read()

with open(
    r'D:\Users\skathir\PycharmProjects\PR_Inflow_Tracker\after_table.txt'
) as body_file:
    last_body_html = body_file.read()
#
# Getting the current date.
#
df2 = pd.to_datetime('today').strftime('%m-%d-%Y')
#
# Outlook Creation & Send.
#
mail_item = olApp.CreateItem(0)
mail_item.To = 'priamp@amazon.com'
mail_item.CC = 'pricerejects-team@amazon.com'
mail_item.Subject = 'PR Summary on' + ' ' + df2
#
final_mail_body = body_html.format(pending, [present_HC])

mail_item.HTMLBody = (
    final_mail_body + '<br/> PR Summary <br/>' + table_html + last_body_html
)
mail_item.send
