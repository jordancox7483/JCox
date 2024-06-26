USE [ULTIPRO_SAMPLE]
GO
/****** Object:  StoredProcedure [dbo].[sp_PN_vacat]    Script Date: 12/11/2009 11:59:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
--Client: SAMPLE
--ProNet Information Systems/Jordan Cox
--Calculates Salary Vacation Accruals
--11/12/2009
-- =============================================
ALTER PROCEDURE [dbo].[sp_PN_vacat] 
	-- Add the parameters for the stored procedure here
	@EEID AS VARCHAR(12)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

DECLARE @Ret AS VARCHAR(8)
DECLARE @Process AS VARCHAR(1)
DECLARE @VacYears AS INT
DECLARE @VacDate AS DATETIME
DECLARE @PayDate AS DATETIME
DECLARE @UWayDays AS INT

SET @Ret = NULL
SET @Process = NULL
SET @VacYears = NULL
SET @VacDate = NULL
SET @PayDate = NULL
SET @UWayDays = NULL

--DROP TABLE Tmp

--Determine if UD Field is populated, if not use Seniority
SELECT @VacDate =
	CASE 
		WHEN EecUDField07 IS NULL THEN EecDateOfSeniority
		ELSE EecUDField07 
	END
FROM EmpComp 
WHERE EecEEID = @EEID

--Load Paydate based on passed information
SELECT @PayDate = 
MbtPayDate
FROM M_Batch 
WHERE MbtEEID = @EEID
	AND MbtPeriodType = 'R'
	AND MbtPayGroup IN ('SALARY','SALNEX')

--Load years of service for vacat calc
SELECT @VacYears =
	ROUND(DATEDIFF(YEAR,@VacDate,@PayDate),0)

--Determine if pay date is in 01 or 06 to decide if EE should get accrual or not
SELECT @Process =
	CASE WHEN MONTH(@PayDate) IN ('11','6') THEN 'Y' ELSE 'N' END

--Load United Way Days from UD Field
SELECT @UWayDays = 
	CASE
		WHEN @Process = 'Y' AND MONTH(@PayDate) = '1' AND EecUDFIELD21 = '4' THEN 4
		WHEN @Process = 'Y' AND EecUDFIELD21 = '8' THEN 4
	ELSE 0
	END
FROM EmpComp 
WHERE EecEEID = @EEID

--Calculate accrual
SELECT
	CASE
		WHEN @VacYears < 1 THEN 
							CASE
								WHEN DAY(@VacDate) > 14 THEN (MONTH(@PayDate) - MONTH(@VacDate))
								ELSE (MONTH(@PayDate) - MONTH(@VacDate) + 1)
							END
		WHEN @Process = 'Y' AND @VacYears < 5 THEN 96/2
		WHEN @Process = 'Y' AND @VacYears < 10 THEN 120/2
		WHEN @Process = 'Y' AND @VacYears < 15 THEN 144/2
		WHEN @Process = 'Y' AND @VacYears < 20 THEN 160/2
		WHEN @Process = 'Y' AND @VacYears < 30 THEN 184/2
		WHEN @Process = 'Y' AND @VacYears >= 30 THEN 200/2
	ELSE 0
	END

+ @UWayDays

AS 'VACAT', @EEID AS 'EEID'
--INTO Tmp

END

