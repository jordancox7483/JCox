USE [ULTIPRO_SAMPLECO]
GO
/****** Object:  StoredProcedure [dbo].[sp_U_PN_SAMPLE_TComp]    Script Date: 06/03/2010 14:15:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/**********************************************************************************************************
$CLIENT:          SAMPLE
$DATE:            11/17/2009
$CREATED BY:      ProNet Information Systems/jc 
$DESC:                
$DEPENDENCIES:    
$ULTIPRO VER.     10.3.0
$Last Revised:    05/25/2010
$Revision No.:    1
$Revised By:      
$Rev.Desc.:       11/23/2009 - added left outer join to account for PT employees not having deductions...
				  05/25/2010 - added stock purchase and employee address fields.
***********************************************************************************************************/

ALTER PROCEDURE [dbo].[sp_U_PN_SAMPLE_TComp]
--Variables passed to the SP
@AscFileName        VARCHAR(100),
@COID               CHAR(5),
@COIDAllCompanies CHAR(1),
@EndPerControl    CHAR(9),
@FormatCode         CHAR(10),
@SELECTByField    CHAR(25),
@SELECTByList       VARCHAR(512),
@StartPerControl  CHAR(9),
@SystemID           CHAR(12),
@TaxCalcGroupID   CHAR(5),
@COIDList           VARCHAR(512),
@OutputRecords    INT OUTPUT  

AS

Set NoCount On
Set Concat_Null_Yields_Null Off 

SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

BEGIN

SET NOCOUNT ON

--Variable created for export
DECLARE
@ExportCommand    VARCHAR(4000),
@ExportTable      VARCHAR(100),
@TempTable        VARCHAR(100),
@LineOut          VARCHAR(500),
@Path             VARCHAR(100), 
@File             VARCHAR(100),
@COIDTemp         VARCHAR(512),
@COIDFormatted    VARCHAR(512),
@COIDCount        VARCHAR(2),
@Time_Key         CHAR(12)


--Set path and Tables used
SET               @TempTable = 'U_PN_SAMPLE_TComp'                                      --Sets work table
SET               @ExportTable = 'U_PN_SAMPLE_TComp_Export'                             --Sets export table
SET               @Path = '\\mi1sfserv01\10556\Exports\TComp\'                       --Save path, company specific


--Set Auto-gen Number Variable
EXEC master..xp_usg_gettimedkey_VAL @Time_Key OUTPUT

--Check SysObjects for table and drop if exists
IF EXISTS (SELECT NAME FROM dbo.SysObjects WHERE dbo.SysObjects.Name = @TempTable
      AND dbo.SysObjects.xtype = 'U')
DROP TABLE dbo.U_PN_SAMPLE_TComp  

IF EXISTS (SELECT NAME FROM dbo.SysObjects WHERE dbo.SysObjects.Name = @ExportTable
      AND dbo.SysObjects.xtype = 'U')
DROP TABLE dbo.U_PN_SAMPLE_TComp_Export 

--Start Export                                                                      
--Export Select Statements
CREATE TABLE DBO.U_PN_SAMPLE_TComp ([Export] VARCHAR(1000),[EEID] VARCHAR(12))
SELECT EecEmpNo, EepSSN, 
'"' + LTRIM(RTRIM(EepNameLast)) + ', ' + LTRIM(RTRIM(EepNameFirst)) + '"' AS 'Name',
CONVERT(Varchar(10), EecDateOfLastHire, 110) as 'LastHireDate',
EecOrgLvl1 AS 'Org Level 1',
EecOrgLvl2 AS 'Org Level 2',
EecOrgLvl3 AS 'Org Level 3',
SUM(CASE WHEN PthTaxCode LIKE '%SUIER%' THEN PthCurTaxAmt ELSE 0 END) AS 'SUI ER',
SUM(CASE WHEN PthTaxCode LIKE '%FUTA%' THEN PthCurTaxAmt ELSE 0 END) AS 'FUTA ER',
SUM(CASE WHEN PthTaxCode IN ('USMEDER','USSOCER') THEN PthCurTaxAmt ELSE 0 END) AS 'FICA ER',
SUM(CASE WHEN PthTaxCode IN ('USMEDEE','USSOCEE') THEN PthCurTaxAmt ELSE 0 END) AS 'FICA EE',
'"' + COALESCE(LTRIM(RTRIM(EepAddressLine1)),'') +  '"' AS 'Address1',
'"' + COALESCE(LTRIM(RTRIM(EepAddressLine2)),'') + '"' AS 'Address2',
RTRIM(EepAddressCity) AS 'City',
RTRIM(EepAddressState) AS 'State',
RTRIM(EepAddressZipCode) AS 'Zip'
INTO #TmpTax
FROM PTaxHist, EmpComp, EmpPers
WHERE PthEEID = EecEEID AND PthCoID = EecCoID
	AND PthEEID = EepEEID
	AND PthCurTaxAmt <> 0
	AND PthPerControl BETWEEN @StartPerControl AND @EndPerControl
GROUP BY EecEmpNo, EepSSN, EepNameLast, EepNameFirst, EecDateOfLastHire, EecOrglvl1, EecOrgLvl2, EecOrgLvl3, EepAddressLine1, EepAddressLine2, EepAddressCity, EepAddressState, EepAddressZipCode
ORDER BY EecEmpNo,EepSSN,EecOrglvl1, EecOrgLvl2, EecOrgLvl3


SELECT EecEmpNo, 
'"' + LTRIM(RTRIM(EepNameLast)) + ', ' + LTRIM(RTRIM(EepNameFirst)) + '"' AS 'Name',
EecOrgLvl1 AS 'Org Level 1',
EecOrgLvl2 AS 'Org Level 2',
EecOrgLvl3 AS 'Org Level 3',
EecEmplStatus,
EecAnnSalary,

COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode = 'REG' AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Regular',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('CAOT','OT') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'OverTime',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('BDLT') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'DoubleTime',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('SICK','HOLID','VAC','VACPD','ACRVC','BEREV') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'PTO Pay',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('ONCAL') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'OnCall Pay',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('SPSBN') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'SPSBonus',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('SDCBN') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'SDCBonus',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('NCAR','AUTO') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'AutoAllowance',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('ARINC') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'AR WH Incentive',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('FTNES') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Workout Bonus',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('WTLSS') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Fitness Challenge',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('PAYEX') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Pay Exact',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('MVING') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Moving Expense',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('TUIT') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Tuition Reimbursement',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode = 'MAYBN' AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'MayBonus',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode = 'DECBN' AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'DecBonus',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode = 'COMM' AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'Commision',
COALESCE((SELECT SUM(PehCurAmt) FROM PEarHist WHERE PehEEID = EecEEID AND PehCoID = PehCoID AND PehEarnCode IN ('COMM','DECBN','MAYBN') AND PehPerControl BETWEEN @StartPerControl AND @EndPerControl GROUP BY EecEEID),'0') AS 'DirectCompTotal'
INTO #TmpEarn
FROM EmpComp, EmpPers
WHERE EecEEID = EepEEID


SELECT EecEmpNo, 
'"' + LTRIM(RTRIM(EepNameLast)) + ', ' + LTRIM(RTRIM(EepNameFirst)) + '"' AS 'Name',
EecOrgLvl1 AS 'Org Level 1',
EecOrgLvl2 AS 'Org Level 2',
EecOrgLvl3 AS 'Org Level 3',
SUM(CASE WHEN PdhDedCode IN ('MEDA','MEDB') THEN PdhEECurAmt ELSE 0 END) AS 'Med EE',
SUM(CASE WHEN PdhDedCode IN ('MEDA','MEDB') THEN PdhERCurAmt ELSE 0 END) AS 'Med ER',
SUM(CASE WHEN PdhDedCode IN ('DEN') THEN PdhEECurAmt ELSE 0 END) AS 'Den EE',
SUM(CASE WHEN PdhDedCode IN ('DEN') THEN PdhERCurAmt ELSE 0 END) AS 'Den ER',
SUM(CASE WHEN PdhDedCode IN ('BLFEE') THEN PdhEECurAmt ELSE 0 END) AS 'BasicLifeEE',
SUM(CASE WHEN PdhDedCode IN ('BLFEE') THEN PdhERCurAmt ELSE 0 END) AS 'BasicLifeER',
SUM(CASE WHEN PdhDedCode IN ('LTDS','LTDB') THEN PdhEECurAmt ELSE 0 END) AS 'LTD EE',
SUM(CASE WHEN PdhDedCode IN ('LTDS','LTDB') THEN PdhERCurAmt ELSE 0 END) AS 'LTD ER',
SUM(CASE WHEN PdhDedCode IN ('LTCEE') THEN PdhEECurAmt ELSE 0 END) AS 'LTC EE',
SUM(CASE WHEN PdhDedCode IN ('LTCSP') THEN PdhEECurAmt ELSE 0 END) AS 'LTC SP',
SUM(CASE WHEN PdhDedCode IN ('LTCER') THEN PdhERCurAmt ELSE 0 END) AS 'LTC ER',
SUM(CASE WHEN PdhDedCode IN ('PRSHR') THEN PdhERCurAmt ELSE 0 END) AS 'ProfitShare ER',
SUM(CASE WHEN PdhDedCode IN ('401K') THEN PdhEECurAmt ELSE 0 END) AS '401K EE',
SUM(CASE WHEN PdhDedCode IN ('401K') THEN PdhERCurAmt ELSE 0 END) AS '401K ER',
SUM(CASE WHEN PdhDedCode IN ('401KC') THEN PdhEECurAmt ELSE 0 END) AS '401KC EE',
SUM(CASE WHEN PdhDedCode IN ('401KC') THEN PdhERCurAmt ELSE 0 END) AS '401KC ER',
SUM(CASE WHEN PdhDedCode IN ('401KR') THEN PdhEECurAmt ELSE 0 END) AS '401KR EE',
SUM(CASE WHEN PdhDedCode IN ('401KR') THEN PdhERCurAmt ELSE 0 END) AS '401KR ER',
SUM(CASE WHEN PdhDedCode IN ('401K','401KC','401KR') THEN PdhEECurAmt ELSE 0 END) AS 'Total 401K EE',
SUM(CASE WHEN PdhDedCode IN ('401K','401KC','401KR') THEN PdhERCurAmt ELSE 0 END) AS 'Total 401K ER',

SUM(CASE WHEN PdhDedCode IN ('STOCK') THEN PdhEECurAmt ELSE 0 END) AS 'Stock'

INTO #TmpDed
FROM EmpComp, EmpPers, PDedHist
WHERE EecEEID = EepEEID AND PdhEEID = EecEEID AND PdhCoID = PdhCoID AND PdhPerControl BETWEEN @StartPerControl AND @EndPerControl
GROUP BY EecEEID, EecEmpNo, EepNameLast, EepNameFirst, EecOrgLvl1, EecOrgLvl2, EecOrgLvl3


INSERT INTO  DBO.U_PN_SAMPLE_TComp ([Export],[EEID])
SELECT 'EmpNo' + ',' + 'SSN' + ',' + 'Name'  + ',' + 'Address Line 1' + ',' + 'Address Line 2' + ',' + 'City' + ',' + 'State' + ',' + 'ZipCode' + ',' + 'Last Hire Date' 
		+ ',' + 'Org Level 1' + ',' + 'Org Level 2' + ',' + 'OrgLevel 3' + ',' + 'Med EE' + ',' + 'MedER' + ',' + 'Den EE' + ',' + 'DenER' + ',' + 'BasicLifeEE' + ',' + 'BasicLifeER' + ',' + 'LTD EE' + ',' + 'LTD ER' + ',' + 'LTC EE' + ',' + 'LTC SP' + ',' + 'LTC ER' + ',' + 'ProfitShare ER' 
		+ ',' + '401K EE' + ',' + '401K ER' + ',' + '401KC EE' + ',' + '401KC ER' + ',' + '401KR EE' + ',' + '401KR ER' + ',' + 'Total 401K EE' + ',' + 'Total 401K ER' + ',' + 'Stock Purchase'
		+ ',' + 'SUI ER' + ',' + 'FUTA ER' + ',' + 'FICA ER' + ',' + 'FICA EE' + ',' + 'Annual Salary' 
		+ ',' + 'Regular Pay' + ',' + 'Overtime' + ',' + 'DoubleTime' + ',' + 'PTO Pay' + ',' + 'On Call Pay' + ',' + 'SPS Bonus' + ',' + 'SDC Bonus' + ',' + 'Auto Allowance' + ',' + 'AR WH Incentive'
		+ ',' + 'Workout Bonus' + ',' + 'Fitness Challenge' + ',' + 'Pay Exact' + ',' + 'Moving Expense' + ',' + 'Tuition Reimbursement' 
		+ ',' + 'MayBonus' + ',' + 'DecBonus' + ',' + 'Commission' + ',' + 'DirectComp Total' AS 'Export', '00000000001A' AS EEID

INSERT INTO  DBO.U_PN_SAMPLE_TComp ([Export],[EEID])
SELECT COALESCE(#TmpTax.[EecEmpNo],'0') + ',' + COALESCE(RTRIM(#TmpTax.[EepSSN]),'0') + ',' + COALESCE(#TmpTax.[Name],'0') 	
	+ ',' + COALESCE([Address1],'') + ',' + COALESCE([Address2],'') + ',' + COALESCE([City],'') + ',' + COALESCE([State],'') + ',' + COALESCE([Zip],'')
	+ ',' + COALESCE(#TmpTax.[LastHireDate],'0') + ',' + COALESCE(#TmpTax.[Org Level 1],'0') + ',' + COALESCE(#TmpTax.[Org Level 2],'0') + ',' + COALESCE(#TmpTax.[Org Level 3],'0') + ',' + STR(COALESCE([Med EE],'0')) + ',' + STR(COALESCE([Med ER],'0')) + ',' + STR(COALESCE([Den EE],'0')) + ',' + STR(COALESCE([Den ER],'0')) + ',' + STR(COALESCE([BasicLifeEE],'0')) + ',' + STR(COALESCE([BasicLifeER],'0')) + ',' + STR(COALESCE([LTD EE],'0')) + ',' + STR(COALESCE([LTD ER],'0')) + ',' + STR(COALESCE([LTC EE],'0')) + ',' + STR(COALESCE([LTC SP],'0')) + ',' + STR(COALESCE([LTC ER],'0')) + ',' + STR(COALESCE([ProfitShare ER],'0'))
	+ ',' + STR(COALESCE([401K EE],'0')) + ',' + STR(COALESCE([401K ER],'0')) + ',' + STR(COALESCE([401KC EE],'0')) + ',' + STR(COALESCE([401KC ER],'0')) + ',' + STR(COALESCE([401KR EE],'0')) + ',' + STR(COALESCE([401KR ER],'0')) + ',' + STR(COALESCE([Total 401K EE],'0')) + ',' + STR(COALESCE([Total 401K ER],'0')) + ',' + STR(COALESCE([Stock],'0'))
	+ ',' + STR(COALESCE([SUI ER],'0')) + ',' + STR(COALESCE([FUTA ER],'0')) + ',' + STR(COALESCE([FICA ER],'0')) + ',' + STR(COALESCE([FICA EE],'0')) + ',' + STR(COALESCE([EecAnnSalary],'0')) 
	+ ',' + STR(COALESCE([Regular],'0'))+ ',' + STR(COALESCE([OverTime],'0'))+ ',' + STR(COALESCE([DoubleTime],'0'))+ ',' + STR(COALESCE([PTO Pay],'0'))+ ',' + STR(COALESCE([Oncall Pay],'0'))+ ',' + STR(COALESCE([SPSBonus],'0'))+ ',' + STR(COALESCE([SDCBonus],'0'))+ ',' + STR(COALESCE([AutoAllowance],'0'))+ ',' + STR(COALESCE([AR WH Incentive],'0'))
	+ ',' +STR(COALESCE([Workout Bonus],'0')) + ',' + STR(COALESCE([Fitness Challenge],'0')) + ',' + STR(COALESCE([Pay Exact],'0')) + ',' + STR(COALESCE([Moving Expense],'0')) + ',' + STR(COALESCE([Tuition Reimbursement],'0'))
	+ ',' + STR(COALESCE([MayBonus],'0')) + ',' + STR(COALESCE([DecBonus],'0')) + ',' + STR(COALESCE([Commision],'0')) + ',' +  STR(COALESCE([Regular],'0')+COALESCE([OverTime],'0')+COALESCE([DoubleTime],'0')+COALESCE([PTO Pay],'0')+COALESCE([Oncall Pay],'0')+COALESCE([SPSBonus],'0')+COALESCE([SDCBonus],'0')+COALESCE([AutoAllowance],'0')+COALESCE([AR WH Incentive],'0')+COALESCE([MayBonus],'0')+COALESCE([DecBonus],'0')+COALESCE([Commision],'0'))
AS 'Export',
	(RIGHT(REPLICATE('0',12) + COALESCE(#TmpTax.[EecEmpNo],'0'),12)) AS 'EEID'

FROM #TmpTax
	LEFT OUTER JOIN #TmpEarn ON #TmpTax.EecEmpNo = #TmpEarn.EecEmpNo
	LEFT OUTER JOIN #TmpDed ON #TmpTax.EecEmpNo = #TmpDed.EecEmpNo

--End Export







--Set export command
SET               @ExportCommand = '      SELECT
                                                [Export] AS LineOut 
                                                INTO dbo.' + @ExportTable +
                                                ' FROM ' + @TempTable +
                                                ' ORDER BY [EEID]  ' 
--Print @ExportCommand
exec (@ExportCommand)

--Set export path and grab file name from interface                                                                                                                                                                                                              
SELECT            @File =     'JMS_TComp_'+ (REPLACE(REPLACE(REPLACE(CONVERT(Varchar(20),GetDate(),100),'  ',' '),' ','_'),':','_')) + '.csv'


SET         @LineOut = 'bcp ' + db_name() + '.dbo.' + @ExportTable + ' out ' + 
                              @Path + @File + ' -o ' + 
                              @Path + 'ErrorLog.OUT -T -S ' + @@servername + ' -c -t'

print @LineOut
exec        master..xp_cmdshell @LineOut,NO_OUTPUT

SELECT            @OutputRecords =  count(*) - 1 From DBO.U_PN_SAMPLE_TComp (nolock)

--Output Variables for troubleshooting

IF EXISTS (SELECT NAME FROM dbo.SysObjects WHERE dbo.SysObjects.Name = 'U_VarValues'
      AND dbo.SysObjects.xtype = 'U')
INSERT INTO dbo.U_VarValues ([@AscFileName],[@COID],[@COIDAllCompanies],[@EndPerControl],[@FormatCode],[@SELECTByField],
                                    [@SELECTByList],[@StartPerControl],[@SystemID],[@TaxCalcGroupID],[@COIDList],[@COIDTemp],
                                    [@COIDCount],[@COIDFormatted],[@OutputRecords],[@ExportCommand],[@File],[@Path],[@LineOut],[Date])

SELECT
@AscFileName            AS '@AscFileName',
@COID                   AS '@COID',
@COIDAllCompanies AS '@COIDAllCompanies',
@EndPerControl          AS '@EndPerControl'  ,
@FormatCode             AS '@FormatCode'   ,
@SELECTByField          AS '@SELECTByField'   ,
@SELECTByList           AS '@SELECTByList'        ,
@StartPerControl  AS '@StartPerControl' ,
@SystemID               AS '@SystemID'      ,
@TaxCalcGroupID         AS '@TaxCalcGroupID'  ,
@COIDList               AS '@COIDList'      ,
@COIDTemp               AS '@COIDTemp',
@COIDCount              AS '@COIDCount',
@COIDFormatted          AS '@COIDFormatted',
@OutputRecords          AS '@OutputRecords',
@ExportCommand          AS '@ExportCommand',
@File                   AS '@File',
@Path                   AS '@Path',
@LineOut                AS '@LineOut',
GETDATE()               AS 'Date'

IF NOT EXISTS (SELECT NAME FROM dbo.SysObjects WHERE dbo.SysObjects.Name = 'U_VarValues'
      AND dbo.SysObjects.xtype = 'U')
SELECT
@AscFileName            AS '@AscFileName',
@COID                   AS '@COID',
@COIDAllCompanies AS '@COIDAllCompanies',
@EndPerControl          AS '@EndPerControl'  ,
@FormatCode             AS '@FormatCode'   ,
@SELECTByField          AS '@SELECTByField'   ,
@SELECTByList           AS '@SELECTByList'        ,
@StartPerControl  AS '@StartPerControl' ,
@SystemID               AS '@SystemID'      ,
@TaxCalcGroupID         AS '@TaxCalcGroupID'  ,
@COIDList               AS '@COIDList'      ,
@COIDTemp               AS '@COIDTemp',
@COIDCount              AS '@COIDCount',
@COIDFormatted          AS '@COIDFormatted',
@OutputRecords          AS '@OutputRecords',
@ExportCommand          AS '@ExportCommand',
@File                   AS '@File',
@Path                   AS '@Path',
@LineOut                AS '@LineOut',
GETDATE()               AS 'Date'
INTO
dbo.U_VarValues

SET NOCOUNT OFF   
END
