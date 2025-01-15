from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
import numpy as np
import pyodbc 
import warnings
import win32com.client
import openpyxl
warnings.filterwarnings('ignore')

#SQL Connection 
#Using pyodbc library I was able to connect to SQL Server using a connection string
#region
server ='testazsql01'
database = 'Test'
driver = 'SQL Server'
trusted_connection = 'Yes'
conn_str = f'DRIVER={driver};SERVER={server};DATABASE={database};Trusted_Connection={trusted_connection}'
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()
#endregion

#Time Variables
#these variables will be used in the dataframe as well as dynamically naming our files each month
#region
today = datetime.now().date()
today_form = today.strftime("%m-%Y")
today_full = today.strftime("%m/%d/%Y")
prev_date = prev = today - relativedelta(months=1)
prev_form = prev.strftime("%m-%Y")
title = f'Case Record Review {today_form}'

#endregion

#SQL Query
#This code has been redacted with sensitive information but demonstrates the methodology used
#A CTE was created to pull in the Max Review Date and connected to a larger database to pull active client information
#The table was filtered for specific criteria for qualifications that meet eligiblity for record review (detailed below)
query = """WITH Max_Review_Date AS (
    SELECT 
        [CASE_NO],
        CONVERT(date, MAX(REVIEW_END_DT)) AS Last_DT
    FROM 
        Test.[dbo].[Table1] AS sbc
    GROUP BY 
        [CASE_NO]),

CAL AS (
SELECT wfr.*,
     DATEDIFF(MONTH, CONVERT(date, wfr.APD), GETDATE()) AS [Months in Care],
	 Max_Review_Date.LAST_DT,
    DATEDIFF(MONTH, Max_Review_Date.Last_DT, GETDATE()) AS [Last Review]
From Test.[dbo].Table2(GETDATE(),GETDATE()) as clients
Left Join Test.dbo.Table3 as wfr on clients.CLT_NBR = wfr.CLT_NBR
Left Join Max_Review_Date on Max_Review_Date.CASE_NO = wfr.CaseNo )

Select DISTINCT
CLT_LNAME + ', ' + CLT_FNAME as [Child Name],
CIN,
CAL.CLT_NBR,
Convert(date,APD) as [APD],
Convert (date,CLT_DOB) as [DOB],
CLT_SEX as [Sex],
Convert(int,CaseNo) as [Case_NO],
CaseName as [CASE_NAME],
CASE 
	WHEN Clt_Status IN (0, 1, 5, 6, 12) THEN 'ACTIVE' 
	WHEN Clt_Status = 13 THEN 'EXTERNAL TRANSFER' 
	ELSE CLT_STATUS_DESC 
END AS Status_Desc, 
CAL.GOAL,
rtrim(ltrim(DimStaff.StaffName)) as [Worker], 
ltrim(rtrim(DimStaff.SupervisorName)) as [Supervisor], 
ltrim(rtrim(DimStaff.DirectorName)) as [Director], 
DimStaff.SITE, 
DimStaff.UNIT,
[Months in Care],
[Last Review]


From CAL
LEFT OUTER JOIN Test..Table4 as DimStaff ON CAL.STF_NBR = DimStaff.Wrkr_ID
LEFT JOIN Test.dbo.Table5 as S ON S.CLT_NBR = CAL.CLT_NBR
							AND S.EFFECT_DT = CAL.CLT_STATUS_DT
							AND S.STATUS = CAL.CLT_STATUS
Where ([Months in Care] > 3 AND 
    (
        ([Last Review] > 12) OR 
        ([Last Review] IS NULL)
    )) 


and NOT EXISTS(
Select DISTINCT
CLT_LNAME + ', ' + CLT_FNAME as [Child Name],
CIN,
CAL.CLT_NBR,
Convert(date,APD) as [APD],
Convert (date,CLT_DOB) as [DOB],
CLT_SEX as [Sex],
Convert(int,CaseNo) as [Case_NO],
CaseName as [CASE_NAME],
CASE 
	WHEN Clt_Status IN (0, 1, 5, 6, 12) THEN 'ACTIVE' 
	WHEN Clt_Status = 13 THEN 'EXTERNAL TRANSFER' 
	ELSE CLT_STATUS_DESC 
END AS Status_Desc, 
CAL.GOAL,
rtrim(ltrim(DimStaff.StaffName)) as [Worker], 
ltrim(rtrim(DimStaff.SupervisorName)) as [Supervisor], 
ltrim(rtrim(DimStaff.DirectorName)) as [Director], 
DimStaff.SITE, 
DimStaff.UNIT,
[Months in Care],
[Last Review]
From Test.dbo.Table6 as stat
Where ([Months in Care] > 3 AND 
    (
        ([Last Review] > 12) OR 
        ([Last Review] IS NULL)
    )) 
 and
CAL.CLT_NBR = stat.CLT_NBR and stat.STATUS_DESC IN ('END OF CASE','TRANSFER'))
Order by CASE_NO, APD desc
"""

#SQL to dataframe
#The data was converted to a pandas dataframe for further data cleaning and transformation, including adding two empty columns which will contain the assigned reviewer's name and the date of assignment.
#I also created a variable that contains the unique units to be used later. This is from the original dataframe to ensure a full, unique list.
#The length of the units list will also be used later in the method for sampling.
#region
data = pd.read_sql(query, conn)
df = pd.DataFrame(data)
df['Assigned Reviewer']='NULL'
df['Date Assigned'] = 'NULL'
units = df['UNIT'].unique().tolist()
units_num = len(units)
#endregion

#Load Previous Reviews and data cleaning
#We will also exclude cases that were selected during the previous month's review cycle even if not reviewed yet. One full month must pass before cases become eligible for reselection for review.
#I created a deep copy to protect the original dataframe's structure.
#region
data2 = pd.read_excel(rf'C:\Case Record Reviews\Case Record Review {prev_form}.xlsx',header=0, sheet_name=f'Case Record Review {prev_form}')
prev_df = pd.DataFrame(data2)
df.drop_duplicates(subset=['Case_NO'],inplace=True)
df = df[~df['Case_NO'].isin(prev_df['Case_NO'])]
df2 = df.copy()
#endregion


#Data Variables and Set Seed for reproducbility
#A seed is set to ensure the same results each time the code is run. This was helpful for testing the dataframe several times to ensure no duplicates were found.
#region
seed = datetime.now().toordinal()
set_seed = np.random.seed(seed)
#endregion


#Create Reviewer class with four instances and one method for sampling cases
#Parameters are passed in for the reviewer's name, number of cases to be assigned, and HEX code for formatting rows in Excel.
#An empty list is also created for each reviewer to append assigned cases and obtain a count to compare to the unit_num variable created above.
class Reviewer:
    def __init__(self, name, number, color):
        self.name = name
        self.number = number 
        self.color = color
        self.assignedcases = []
#This method samples one case per unit for Reviewer Cassie and Reviewer Steve. The selected cases have the reviewer name filled in the "Assigned Reviewer" column and today's date in the "Date "Assigned column.
#A for loop is used to sample each unit once. Each sample row is then dropped from the dataframe and added to the an all cases list that will contain all assigned cases.
#A new deep copy is created at the end of the loop to ensure the changes are saved in the new iteration of the loop.
    def cases(self):
        global units
        all_cases = []
        sample_num = self.number - len(self.assignedcases)
        if self.name == 'Cassie Jones' or self.name == 'Steve Smith':
            for unit in units:
                unit_rows = pd.DataFrame(df2[df2['UNIT'] == unit]).copy()
                unit_rows['Assigned Reviewer'] = self.name
                unit_rows['Date Assigned'] = today_full
                sample = unit_rows.sample(n=1, replace=False)
                all_cases.append(sample)
                self.assignedcases.append(sample['UNIT'])
                df2.drop(sample.index,inplace=True)
                df2.reset_index(drop=True, inplace=True) 
                unit_rows=df2.copy()
                #when the unit limit is reached, the sampling changes to select the remaining cases irrespecitve of unit.
                #The process is repeated with sampling, dropping, and appending.
                if len(self.assignedcases) >= len(units):
                    sample_num = self.number - len(self.assignedcases)
                    for _ in range(sample_num):
                        unit_rows = df2.copy()
                        unit_rows['Assigned Reviewer'] = self.name
                        unit_rows['Date Assigned'] = today_full
                        sample = unit_rows.sample(n=1, replace=False)
                        all_cases.append(sample)
                        self.assignedcases.append(sample['UNIT'])
                        df2.drop(sample.index,inplace=True)
                        df2.reset_index(drop=True, inplace=True)
                        unit_rows=df2.copy()
                        #When the reviewer number is the same as the length of the assigned cases for the reviewer, Python will exit the loop.
                        if len(self.assignedcases) == self.number:
                            break
        #For all other reviewers, the sample will include their assigned number of cases.
        #Cases are also sampled, dropped, and appended.
        else:
            unit_rows = df2.copy()
            unit_rows['Assigned Reviewer'] = self.name
            unit_rows['Date Assigned'] = today_full
            sample = unit_rows.sample(n=self.number, replace=False)
            all_cases.append(sample)
            self.assignedcases.append(sample['UNIT'])
            df2.drop(sample.index,inplace=True)
            df2.reset_index(drop=True, inplace=True)
        #finally all the cases are appended together 
        return pd.concat(all_cases)  

#Instance variables 
#The instance variables are created here, including name, number, and HEX code.
#region
cassie = Reviewer('Cassie Jones', 15, 'F9F907')
olivia = Reviewer('Olivia Roberts', 2, '92D050')
cassandra = Reviewer('Cassandra Burks', 2, 'BDD7EE')
steve = Reviewer('Steve Smith', 15, 'FFC000')
cases = pd.concat([cassie.cases(),olivia.cases(),cassandra.cases(),steve.cases()])
#endregion

#Export to Excel
#The file is exported to an excel file to share with others.
#The path is also saved for use with win32com library for automated email sending.
#region
cases.to_excel(rf'C:\Case Record Reviews\{title}.xlsx', index=False, sheet_name=title)
file_path = rf'C:\Case Record Reviews\{title}.xlsx'
#endregion


#Automated Emailing
#Python will attach and send the report each time it is run.
#Stakeholder are assigned to variables and later referenced to send the email.
#region
to = 'ShawM@myemail.org'
bcc = 'charlesi@myemail.org'
cc = ['RichardsG@myemail.org', 'HillP@myemail.org']
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.Subject = title
mail.HTMLBody = f"Hi Maddie, here is the Case Record Review for {today_form}:"
mail.To = to
mail.CC = '; '.join(cc)
mail.BCC = bcc
mail.Attachments.Add(file_path)
mail.Send()
#endregion




