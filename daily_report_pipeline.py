import warnings
warnings.filterwarnings('ignore')
# Import imaplib, email, os, configparser, glob, numpy, pandas, pyodbc, shutil 
# and os packages
import imaplib, email, os
import configparser
import datetime
import glob
import numpy as np
import pandas as pd
import pyodbc
import shutil

# Check the mail
# Read password configuration file
config = configparser.ConfigParser()
config.read(r"E:\cepps\Web_Report\Credit_Karma\etc\config.txt")
user_config = config.get("configuration","user")
password_config = config.get("configuration","password")
imap_url_config = config.get("configuration","imap_url")
smtp_url_config = config.get("configuration","smtp_url")

# Store login info
user = user_config
password = password_config
imap_url = imap_url_config

# Load directory to save attachements
# Where you want your attachments to be saved (ensure this directory exists) 
attachment_dir = r'E:\cepps\Web_Report\Credit_Karma\attachments'

# Create date variable formatted as DD-MM-YYYY
date = (datetime.date.today() - datetime.timedelta(0)).strftime("%d-%b-%Y")

# Create auth function to pass authorization info
# sets up the auth
def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con

# Create get_body function to retrieve the specified email
# extracts the body from the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)

# Create get_attachments function to retrieve the attachments from the 
# specified email
# allows you to download attachments
def get_attachments(msg):
    for part in msg.walk():
        if part.get_content_maintype()=='multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()

        if bool(fileName):
            filePath = os.path.join(attachment_dir, fileName)
            with open(filePath,'wb') as f:
                f.write(part.get_payload(decode=True))

# Create search function in case program should search for specific email (this 
# has no use in current iteration, but may be important later)
# search for a particular email
def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data

# Create get_emails function to retrieve all specified emails
# extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

# Connect to email imap server and select the requisite mailbox
con = auth(user,password,imap_url)
con.select('INBOX/CK_Reports')

# Fetch email from bay sent since the day the program runs. This should be the 
# most recent email in the mailbox. Specify email format as RFC822. Store email 
# as byte data. Extract attachments from raw email data and store in previously 
# designated directory.
result, data = con.search(None, '(SENTSINCE {0})'.format(date))
for ids in data[0].split():
    result, data = con.fetch(ids,'(RFC822)')
    raw = email.message_from_bytes(data[0][1])
    get_attachments(raw)

# Clear mailbox, close connection, and log out.
#typ, data = con.search(None, 'ALL')
#
#for num in data[0].split():
#   con.store(num, '+FLAGS', '\\Deleted')
#
#con.expunge()
#con.close()
#con.logout()

# Use glob to find all attachments in a specified pathname. Retrieve the latest 
# file and print the results
# * means all if need specific format then *.csv
list_of_files = glob.glob(r'E:\cepps\Web_Report\Credit_Karma\attachments\*') 
latest_file_1 = max(list_of_files, key=os.path.getctime)
latest_file_2 = min(list_of_files, key=os.path.getctime)
print(latest_file_1)
print(latest_file_2)

# Read in the data from the most recent Daily CK Report. Strip the column 
# names, convert them to lower case, and replace spaces with underscores. 
leads_1 = pd.read_csv(latest_file_1, encoding="ISO-8859-1", 
                      error_bad_lines=False)
leads_2 = pd.read_csv(latest_file_2, encoding="ISO-8859-1", 
                      error_bad_lines=False)
leads = leads_1.append(leads_2)
leads.columns = leads.columns.str.strip().str.lower().str.replace(' ', '_') 

# Filter leads data set for Credit Karma leads
leads = leads[(leads.irmpname == 'CreditKarma')]

# Store applicant_ssn in ssn as string type. Format ssn to 9 characters 
# including leading zeros. Drop duplicates
leads['ssn'] = leads['applicant_ssn'].astype(str) 
leads['ssn'] = leads['ssn'].apply(lambda x: '{0:0>9}'.format(x)) 
leads = leads.drop_duplicates(subset=['ssn'], keep="first")

# Format string types
leads['irpid'] = leads['irpid'].astype(str)
leads['application_number'] = leads['application_number'].astype(str)
leads['applicant_ssn'] = leads['applicant_ssn'].astype(str)
leads['applicant_address_zip'] = leads['applicant_address_zip'].astype(str)
leads['app._cell_phone'] = leads['app._cell_phone'].astype(str)
leads['app._home_phone'] = leads['app._home_phone'].astype(str)
leads['app._work_phone'] = leads['app._work_phone'].astype(str)
leads['PQ_DECISION'] = leads['decision_status'].astype(str)
leads['LEAD_TYPE'] = leads['loan_type'].astype(str)
leads['LEAD_STATE'] = leads['applicant_address_state'].astype(str)
leads['CK_TRACKING_ID'] = leads['sub_id'].astype(str)

# Format date types
leads['application_date'] = pd.to_datetime(leads['application_date'],
     infer_datetime_format=True)
leads['decision_date/time'] = pd.to_datetime(leads['decision_date/time'],
     infer_datetime_format=True)
leads['LEADDATE'] = pd.to_datetime(leads['application_date'],
     infer_datetime_format=True)

# Format integers
leads['LEAD_SCORE'] = leads['applicant_credit_score'].fillna(0).astype(int)
leads['AMT_APPLIED'] = leads['amt._fin.'].astype(int)

# Create lead identifier
leads['lead_app'] = 'lead'

# Store the SSNs as a list
leads_ssn = leads['ssn'].tolist()

# Move attachments to archive folder
source = 'E:\\cepps\\Web_Report\\Credit_Karma\\attachments\\'
dest1 = 'E:\\cepps\\Web_Report\\Credit_Karma\\archive\\'
files = os.listdir(source)
for f in files:
        shutil.move(source+f, dest1)

# Connect to RMCDW. Select SSNo and Cifno for SSNs in the leads_ssn list
# Parameters
server = 'server-DW'
db = 'RMCDW'

# Create the connection
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + db + ';Trusted_Connection=yes')

# query db
sql = 'SELECT SSNo, Cifno FROM vw_BORROWER t1 WHERE t1.SSNo in %s' % str(tuple(leads_ssn))
ssn = pd.io.sql.read_sql(sql, conn)

# Connect to NLS_Prod to retrieve application data. Select statement provided 
# by Brian Killen
# Parameters
server = 'NLS-PROD-SQL03'
db = 'NLS_Prod'

# Create the connection
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + db + ';Trusted_Connection=yes')

# query db
sql = """

select distinct 
    c.firstname1 as firstname,
    c.lastname1 as lastname,
    ta.task_refno,
    td.userdef100 as LoanNumber	,
    st.statflags as StatusCodes,
    TS.status_Code as TaskStatus,
    case when td.userdef100 is null then null else ls1.status_code end as LoanStatus,

    COALESCE(( select distinct 'Yes' from task_routing_history as trh
              left join workflow_action as wa (nolock) on trh.workflow_action_id = wa.workflow_action_id
              left join dbo.workflow_result as wr (nolock) on wr.workflow_result_id = trh.workflow_result_id
                 where task_refno = td.task_refno  
                 and action = 'PULL CREDIT REPORT'
                 and result = 'yes'
                ),'No')  as WasCreditReportPulled,

    ( select top 1 cast(created_date as date) from task_routing_history as trh
    left join workflow_action as wa (nolock) on trh.workflow_action_id = wa.workflow_action_id
                 left join dbo.workflow_result as wr (nolock) on wr.workflow_result_id = trh.workflow_result_id
    where task_refno = td.task_refno  
    and	wa.action = 'BOOK LOAN'
    and wr.result = 'yes'
    order by created_date desc) as FundedDate,
    
    CAST(CASE WHEN (h.userdef05 IS NULL OR LEN(h.userdef05) < 13 OR h.userdef05 like '% [a-z]%') THEN '' ELSE RIGHT(LEFT(h.userdef05,CHARINDEX('.',h.userdef05)-1),4) END AS INT) AS CreditScore,

    coalesce(atb.netbalance,
    case when td.userdef100 is null then 0 else
    cast(CASE WHEN (CASE WHEN lsp.interest_method in ('7S', '78', 'FA') THEN 'P' WHEN lsp.interest_method = 'SI' THEN 'I' ELSE '0' END 
                    = 'I') THEN LEFT((l.original_note_amount+COALESCE(l.addonint_total,'0.00')),9) else
                        COALESCE((REPLACE(ld.userdef11,',','')),LEFT((l.original_note_amount+COALESCE(l.addonint_total,'0.00')),9)) END as numeric(18,2)) end)  AS FinalLoanAmount,

        ( select top 1 cast(created_date as date) as ApplicationDate from task_routing_history as trh
                 left join workflow_action as wa (nolock) on trh.workflow_action_id = wa.workflow_action_id
                 where task_refno = td.task_refno  
                 and action = 'ENTER APPLICATION DATA'
                 order by created_date) as ApplicationEnterDate,
                 cast(td.userdef54 as numeric(18,0)) as FinalTerm,
        
        td.userdef55 as InterestRate,
        td.userdef57 as FinalAPR,
        ActionsTaken.action as LastAction,
        ActionsTaken.result as LastResult,
        cast(td.userdef34 as numeric(18,2)) as AmountRequested,
        cast(td.userdef10 as numeric(18,2)) as SmallApprovedAmount,
        cast(td.userdef19 as numeric(18,2)) as LargeApprovedAmount,
        td.userdef14  as SmallAPR,
        td.userdef21  as LargeAPR,
        td.userdef12 as SmallTerm,
        td.userdef20 as LargeTerm,
        c.cifno as Cifno,
        td.userdef11 as ReasonDeclined,
        cast(td.userdef07 as numeric(18,2)) as FreeIncome,
        cast(td.userdef04 as numeric(18,2)) as MonthlyGrossIncome,
        cast(td.userdef05 as numeric(18,2)) as MonthlyNetIncome,
        CASE trh2.workflow_result_id 
            when 2 then 'FB'
            when 3 then 'NB'
            when 28 then 'PB'
            when 32 then 'PB'
        END as  CustomerType
        
    
  from cif as c
    INNER JOIN NLS_Prod.dbo.tmr AS t ON tmr_code_id = 1 and c.cifno = t.child_refno
    LEFT OUTER JOIN NLS_Prod.dbo.cif_detail (NOLOCK) h ON c.cifno = h.cifno
    INNER JOIN task as ta (nolock) on t.parent_refno = ta.task_refno
    LEFT JOIN NLS_Prod.dbo.task_detail AS td ON ta.task_refno = td.task_refno and ta.task_template_no = 1
    LEFT JOIN loanacct as l (nolock) on tmr_code_id = 1 and COALESCE(l.loan_number,'') = COALESCE(td.userdef100,'')
    LEFT JOIN NLS_Prod.dbo.loanacct_detail AS ld ON ld.acctrefno = l.acctrefno 
    outer apply (select top 1 wa.action,wr.result
                from  task_routing_history as trh (nolock) 
                left join workflow_action as wa (nolock) on trh.workflow_action_id = wa.workflow_action_id  
                left join dbo.workflow_result as wr (nolock) on wr.workflow_result_id = trh.workflow_result_id
                where trh.task_refno = ta.task_refno
                order by 1 desc) as ActionsTaken
    left JOIN NLS_Prod.dbo.Loanacct_setup (NOLOCK) lsp
                ON l.acctrefno = lsp.acctrefno 

    Left join task_routing_history trh2 (NOLOCK) on trh2.task_refno = ta.task_refno and trh2.workflow_action_id = 2
    Left join loan_status_codes as ls1 (nolock) on ls1.status_code_no = l.status_code_no
    JOIN NLS_Prod.dbo.task_status_codes TS (NOLOCK) ON TS.status_code_id = Ta.status_code_id
    outer apply (select top 1 netbalance from rmc.[dbo].[vw_AgedTrialBalance] (nolock) where loannumber = l.loan_number order by rundate) as atb
    left join (	SELECT DISTINCT acctrefno, 
                            STUFF((SELECT ','+ status_code 
                            FROM (SELECT     lsh.acctrefno, lsc.status_code
                                    FROM         NLS_Prod.dbo.loanacct AS l WITH (NOLOCK) 
                                                INNER JOIN NLS_Prod.dbo.loanacct_statuses AS lsh WITH (NOLOCK) ON l.acctrefno = lsh.acctrefno 
                                                            INNER JOIN NLS_Prod.dbo.loan_status_codes AS lsc WITH (NOLOCK) ON lsc.status_code_no = lsh.status_code_no) s1
                                            WHERE s1.acctrefno = s2.acctrefno 
                                        AND status_code <> ''  FOR XML PATH('')),1,1,'') AS statflags
                            FROM (SELECT     lsh.acctrefno, lsc.status_code
                    FROM         NLS_Prod.dbo.loanacct AS l WITH (NOLOCK) 
                                INNER JOIN NLS_Prod.dbo.loanacct_statuses AS lsh WITH (NOLOCK) ON l.acctrefno = lsh.acctrefno 
                                            INNER JOIN NLS_Prod.dbo.loan_status_codes AS lsc WITH (NOLOCK) ON lsc.status_code_no = lsh.status_code_no) s2
                            WHERE status_code <> ''
                            ) as st on st.acctrefno = l.acctrefno
    where 
    COALESCE(td.userdef100,'') <> 'CLOSE BRANCH'
    and COALESCE(l.loan_group_no,1) <> 6
    and ( select top 1 cast(created_date as date) as ApplicationDate from task_routing_history as trh
                 left join workflow_action as wa (nolock) on trh.workflow_action_id = wa.workflow_action_id
                 where task_refno = td.task_refno  
                 and action = 'ENTER APPLICATION DATA'
                 order by created_date) is not null
            --and ta.task_refno = 327484
    ORDER BY 2

"""
app_table = pd.io.sql.read_sql(sql, conn)

# Format Cifno as string
app_table['Cifno'] = app_table['Cifno'].astype(int)
app_table['Cifno'] = app_table['Cifno'].astype(str)
ssn['Cifno'] = ssn['Cifno'].astype(str)

# Merge app_table and ssn
ck_app_data = pd.merge(app_table, ssn[["SSNo", "Cifno"]], on = 'Cifno', 
                       how = 'inner')

# Drop duplicates
ck_app_data = ck_app_data.drop_duplicates()

# Format String types
ck_app_data['firstname'] = ck_app_data['firstname'].astype(str)
ck_app_data['lastname'] = ck_app_data['lastname'].astype(str)
ck_app_data['task_refno'] = ck_app_data['task_refno'].astype(int)
ck_app_data['task_refno'] = ck_app_data['task_refno'].astype(str)
ck_app_data['LoanNumber'] = ck_app_data['LoanNumber'].astype(str)
ck_app_data['StatusCodes'] = ck_app_data['StatusCodes'].astype(str)
ck_app_data['TaskStatus'] = ck_app_data['TaskStatus'].astype(str)
ck_app_data['LoanStatus'] = ck_app_data['LoanStatus'].astype(str)
ck_app_data['WasCreditReportPulled'] = ck_app_data['WasCreditReportPulled'].astype(str)
ck_app_data['ssn'] = ck_app_data['SSNo'].astype(str)
ck_app_data['ssn'] = ck_app_data['ssn'].apply(lambda x: '{0:0>9}'.format(x)) 

# Format date types
ck_app_data['FundedDate'] = pd.to_datetime(ck_app_data['FundedDate'])
ck_app_data['ApplicationEnterDate'] = pd.to_datetime(ck_app_data['ApplicationEnterDate'])

# Format numerics
ck_app_data['InterestRate'] = ck_app_data['InterestRate'].astype(float)
ck_app_data['FinalAPR'] = ck_app_data['FinalAPR'].astype(float)
ck_app_data['SmallAPR'] = ck_app_data['SmallAPR'].astype(float)
ck_app_data['LargeAPR'] = ck_app_data['LargeAPR'].astype(float)
ck_app_data['SmallTerm'] = ck_app_data['SmallTerm'].astype(float)
ck_app_data['LargeTerm'] = ck_app_data['LargeTerm'].astype(float)
ck_app_data['SSNo'] = ck_app_data['SSNo'].astype(int)
ck_app_data['SSNo'] = ck_app_data['SSNo'].astype(int)

# Create app table identifier
ck_app_data['lead_app'] = 'app'

# Add current date to ck_app_data file formatted as YYYY_MM_DD
# Use pandas to adress the issue
from datetime import datetime
os.chdir('E:\\cepps\\Web_Report\\Credit_Karma\\ck_app_data')
datestring = datetime.strftime(datetime.now(), ' %Y_%m_%d')
#Fill in your path
ck_app_data.to_excel(excel_writer=r"E:\cepps\Web_Report\Credit_Karma\ck_app_data\{0}".format('ck_app_data_' + datestring + '.xls'))

# Drop duplicates again
leads = leads.drop_duplicates(subset=['ssn'], keep="first")

# Sort by application date, de-dup, keep latest applications
ck_app_data = ck_app_data.sort_values(['ApplicationEnterDate'], 
                                      ascending=[False])
apps = ck_app_data.drop_duplicates(["ssn"], keep='first')

# merge leads and apps
merge_1 = pd.merge(leads, apps, on = 'ssn', how = 'left')

# Format numerics, set amount approved
merge_1['AmountRequested'] = merge_1['AmountRequested'].fillna(0).astype(int)
merge_1["AMT_APPROVED"] = merge_1[["SmallApprovedAmount", 
       "LargeApprovedAmount"]].max(axis=1)
merge_1['AMT_APPROVED'] = merge_1['AMT_APPROVED'].fillna(0).astype(float)

# Flag apps with app date older than lead date. Set AMT_APPROVED, 
# TERM_APPROVED, RATE_APPROVED, AMOUNT_FUNDED, TERM_FUNDED and RATE_FUNDED, to 
# 0 for apps older than lead date.
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'APP_BEFORE_LEAD'] = 1
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'AMT_APPROVED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'TERM_APPROVED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'RATE_APPROVED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'AMOUNT_FUNDED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'TERM_FUNDED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'RATE_FUNDED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'FundedDate'] = np.nan
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 
            'ApplicationEnterDate'] = np.nan

# Apply funnel. Order is very important
merge_1['ISPQAPP'] = 1
merge_1.loc[merge_1.ApplicationEnterDate.isnull(), 'IS_PREQUAL'] = 0
merge_1.loc[(merge_1.PQ_DECISION == 'Auto Approved') & (merge_1.lead_app_x == 'lead'), 'IS_PREQUAL'] = 1
merge_1.loc[merge_1.ApplicationEnterDate.notnull(), 'IS_APPL'] = 1
merge_1.loc[merge_1.AmountRequested > 0, 'IS_APPL'] = 1
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 'IS_APPL'] = 0
merge_1.loc[merge_1.ApplicationEnterDate.isnull(), 'IS_APPL'] = 0
merge_1['IS_APPROVE'] = 0
merge_1.loc[(merge_1.TaskStatus == 'FUNDED') | (merge_1.TaskStatus == 'ELIGIBLE'), 'IS_APPROVE'] = 1
merge_1.loc[(merge_1.AMT_APPROVED > 0) & (merge_1.ReasonDeclined.isnull()), 'IS_APPROVE'] = 1
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 'IS_APPROVE'] = 0
merge_1.loc[merge_1.ApplicationEnterDate.isnull(), 'IS_APPROVE'] = 0
merge_1.loc[merge_1.IS_APPROVE == 1, 'IS_APPL'] = 1
merge_1['IS_APPL'] = merge_1['IS_APPL'].fillna(0).astype(int)
merge_1.loc[merge_1.IS_APPROVE == 1, 'IS_PREQUAL'] = 1
merge_1.loc[merge_1.IS_APPL == 1, 'IS_PREQUAL'] = 1
merge_1['IS_PREQUAL'] = merge_1['IS_PREQUAL'].fillna(0).astype(int)
merge_1['IS_FUNDED'] = 0
merge_1.loc[merge_1.TaskStatus == 'FUNDED', 'IS_FUNDED'] = 1
merge_1['LoanNumber'] = merge_1['LoanNumber'].replace('None', np.NaN)
merge_1.loc[merge_1.LoanNumber.notnull(), 'IS_FUNDED'] = 1
merge_1.loc[merge_1.ApplicationEnterDate < merge_1.LEADDATE, 'IS_FUNDED'] = 0
merge_1.loc[merge_1.ApplicationEnterDate.isnull(), 'IS_FUNDED'] = 0

#Set term, rate, amount funded, and cost
merge_1['TERM_FUNDED'] = 0
merge_1['RATE_FUNDED'] = 0
merge_1['APP_BEFORE_LEAD'] = 0
merge_1['COST'] = 0
merge_1.loc[merge_1.LargeTerm > 0, 'TERM_APPROVED'] = merge_1['LargeTerm']
merge_1.loc[merge_1.LargeTerm <= 0, 'TERM_APPROVED'] = merge_1['SmallTerm']
merge_1.loc[merge_1.LargeAPR > 0, 'RATE_APPROVED'] = merge_1['LargeAPR']
merge_1.loc[merge_1.LargeAPR <= 0, 'RATE_APPROVED'] = merge_1['SmallAPR']
merge_1['AMOUNT_FUNDED'] = merge_1['FinalLoanAmount']
merge_1.loc[merge_1.AMOUNT_FUNDED <= 2500, 'COST'] = 125
merge_1.loc[merge_1.AMOUNT_FUNDED == 0, 'COST'] = 0
merge_1.loc[merge_1.AMOUNT_FUNDED > 2500, 'COST'] = 200
merge_1.loc[merge_1.IS_FUNDED == 1, 'TERM_FUNDED'] = merge_1['FinalTerm']
merge_1.loc[merge_1.IS_FUNDED == 1, 'RATE_FUNDED'] = merge_1['FinalAPR']

# Sort and drop duplicates
merge_1 = merge_1.sort_values(['APP_BEFORE_LEAD', 'IS_FUNDED', 'IS_APPROVE'], 
                              ascending=[False, False, False])
merge_2 = merge_1.drop_duplicates(["ssn"], keep='last')

# Format final dataset
last = merge_2[['CK_TRACKING_ID', 'LEADDATE', 'ApplicationEnterDate', 
                'FundedDate', 'ISPQAPP', 'IS_PREQUAL', 'IS_APPL', 'IS_APPROVE', 
                'IS_FUNDED', 'AMT_APPLIED', 'AMT_APPROVED', 'TERM_APPROVED', 
                'RATE_APPROVED', 'AMOUNT_FUNDED', 'TERM_FUNDED', 
                'RATE_FUNDED']]
last = last.sort_values(['IS_FUNDED'], ascending = [False])

# Show frenquency distribution of funnel
last.groupby(["ISPQAPP", "IS_PREQUAL", "IS_APPL", "IS_APPROVE", 
              "IS_FUNDED"]).size().reset_index(name="Funnel")

# Output CK file with formatted date in filename
datestring = datetime.strftime(datetime.now(), '%m%d%Y')
last.to_csv(r"E:\cepps\Web_Report\Credit_Karma\output\{0}".format('CreditKarma_Regional_'+datestring+'.csv'), 
                encoding='utf-8', 
                float_format='%.2f', index=False)

# * means all if need specific format then *.csv
list_of_files = glob.glob(r'E:\cepps\Web_Report\Credit_Karma\output\*') 
latest = max(list_of_files, key=os.path.getctime)
print(latest)

# specify port, if required, using a colon and port number following the 
# hostname
user = user_config
password = password_config
imap_url = imap_url_config
smtp_url = smtp_url_config
default_address = [user_config] 
host = smtp_url 
send_to = user_config
replyto = user_config # unless you want a different reply-to
subject_text = 'TEST - Daily CK Report - TEST'
# text with appropriate HTML tags
body = """TEST - Daily CK Report - TEST""" 

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from os.path import basename

def send_mail(send_from: str, subject: str, text: str, 
send_to: list, files= None):

    send_to= default_address if not send_to else send_to

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = ', '.join(send_to)  
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    for f in files or []:
        with open(f, "rb") as fil: 
            ext = f.split('.')[-1:]
            attachedfile = MIMEApplication(fil.read(), _subtype = ext)
            attachedfile.add_header(
                'content-disposition', 'attachment', filename=basename(f) )
        msg.attach(attachedfile)


    smtp = smtplib.SMTP(host=smtp_url, port= 587) 
    smtp.starttls()
    smtp.login(user,password)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

send_mail(send_from = user, 
          subject=subject_text,
          text=body,
          send_to= None,
          files=[latest])