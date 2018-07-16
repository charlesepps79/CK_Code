# -*- coding: utf-8 -*-
"""
Created on Wed Jul 11 09:46:17 2018

@author: cepps
"""

import imaplib, email, os
import configparser

config = configparser.ConfigParser()
config.read(r"E:\cepps\Web_Report\Credit_Karma\etc\config.txt")
password_config = config.get("configuration","password")

user = 'cepps@regionalmanagement.com'
password = password_config
imap_url = 'imap-mail.outlook.com'

#Where you want your attachments to be saved (ensure this directory exists) 
attachment_dir = r'E:\cepps\Web_Report\Credit_Karma\attachments'

# sets up the auth
def auth(user,password,imap_url):
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user,password)
    return con

# extracts the body from the email
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)
    
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
                
#search for a particular email
def search(key,value,con):
    result, data  = con.search(None,key,'"{}"'.format(value))
    return data

#extracts emails from byte array
def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = con.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs

con = auth(user,password,imap_url)
con.select('INBOX/CK_Reports')

result, data = con.fetch(b'1','(RFC822)')
raw = email.message_from_bytes(data[0][1])
get_attachments(raw)

# Clear mailbox and log out
typ, data = con.search(None, 'ALL')
for num in data[0].split():
   con.store(num, '+FLAGS', '\\Deleted')
con.expunge()
con.close()
con.logout()

import glob

list_of_files = glob.glob(r'E:\cepps\CK_Reports_Attachment_Test\*') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print (latest_file)

import numpy as np
import pandas as pd
import pyodbc

leads = pd.read_csv(latest_file, encoding="ISO-8859-1", error_bad_lines=False)
leads.columns = leads.columns.str.strip().str.lower().str.replace(' ', '_') 
leads['ssn'] = leads['applicant_ssn'].astype(str) 
leads['ssn'] = leads['ssn'].apply(lambda x: '{0:0>9}'.format(x)) 

leads = leads[(leads.irmpname == 'CreditKarma')]

leads_ssn = leads['ssn'].tolist()

leads.describe(include = 'all')

# Parameters
server = 'server-DW'
db = 'RMCDW'

# Create the connection
conn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + db + ';Trusted_Connection=yes')

# query db
sql = 'SELECT SSNo, Cifno FROM vw_BORROWER t1 WHERE t1.SSNo in %s' % str(tuple(leads_ssn))
ssn = pd.io.sql.read_sql(sql, conn)

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

app_table.head()

ck_app_data = pd.merge(app_table, ssn[["SSNo", "Cifno"]], on = 'Cifno', how = 'inner')

ck_app_data = ck_app_data.drop_duplicates()

ck_app_data['firstname'] = ck_app_data['firstname'].astype(str)
ck_app_data['lastname'] = ck_app_data['lastname'].astype(str)
ck_app_data['task_refno'] = ck_app_data['task_refno'].astype(int)
ck_app_data['task_refno'] = ck_app_data['task_refno'].astype(str)
ck_app_data['LoanNumber'] = ck_app_data['LoanNumber'].astype(str)
ck_app_data['StatusCodes'] = ck_app_data['StatusCodes'].astype(str)
ck_app_data['TaskStatus'] = ck_app_data['TaskStatus'].astype(str)
ck_app_data['LoanStatus'] = ck_app_data['LoanStatus'].astype(str)
ck_app_data['WasCreditReportPulled'] = ck_app_data['WasCreditReportPulled'].astype(str)
ck_app_data['FundedDate'] = pd.to_datetime(ck_app_data['FundedDate'])
ck_app_data['ApplicationEnterDate'] = pd.to_datetime(ck_app_data['ApplicationEnterDate'])
ck_app_data['InterestRate'] = ck_app_data['InterestRate'].astype(float)
ck_app_data['FinalAPR'] = ck_app_data['FinalAPR'].astype(float)
ck_app_data['SmallAPR'] = ck_app_data['SmallAPR'].astype(float)
ck_app_data['LargeAPR'] = ck_app_data['LargeAPR'].astype(float)
ck_app_data['SmallTerm'] = ck_app_data['SmallTerm'].astype(float)
ck_app_data['LargeTerm'] = ck_app_data['LargeTerm'].astype(float)
ck_app_data['SSNo'] = ck_app_data['SSNo'].astype(int)

from datetime import datetime

#Use pandas to adress the issue

os.chdir('E:\\cepps\\Web_Report\\Credit_Karma\\ck_app_data')
datestring = datetime.strftime(datetime.now(), ' %Y_%m_%d')

ck_app_data.to_excel(excel_writer=r"E:\cepps\Web_Report\Credit_Karma\ck_app_data\{0}".format('ck_app_data_' + datestring + '.xls'))#Fill in your path