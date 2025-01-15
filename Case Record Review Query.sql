WITH Max_Review_Date AS (
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
CAL.CLT_NBR = stat.CLT_NBR and stat.STATUS_DESC IN ('FINAL DISCHARGE','ADOPTION','COURT ORDERED DISCHARGE','INTER-AGENCY TRANSFER','DECEASED'))
Order by CASE_NO, APD desc
