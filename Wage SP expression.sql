USE [ULTIPRO_BGF]
GO
/****** Object:  StoredProcedure [dbo].[sp_PN_vacat_wage]    Script Date: 01/13/2010 09:05:22 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
--Client: BGF Industries
--ProNet Information Systems/Jordan Cox
--Calculates Wage Vacation Accruals
--11/12/2009
-- =============================================
ALTER PROCEDURE [dbo].[sp_PN_vacat_wage] 
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
DECLARE @PaySeq AS INT
DECLARE @UWayDays AS INT

SET @Ret = NULL
SET @Process = NULL
SET @VacYears = NULL
SET @VacDate = NULL
SET @PayDate = NULL
SET @UWayDays = NULL

--DROP TABLE Tmp

--Determine if UD Field is populated, if not use Seniority (Wage policy doesn't use the UD Field)
SELECT @VacDate =
	CASE 
		WHEN EecUDField07 IS NULL THEN EecDateOfSeniority
		ELSE EecDateOfSeniority 
	END
FROM EmpComp 
WHERE EecEEID = @EEID

--Load Paydate based on passed information
SELECT @PayDate = 
MbtPayDate
FROM M_Batch, PgPayPer
WHERE MbtEEID = @EEID
	AND MbtPeriodType = 'R'
	AND MbtPayGroup NOT IN ('SALARY','SALNEX')
	AND MbtPayGroup = PgpPayGroup
	AND MbtPayDate = PgpPayDate

--Load Pay Sequence based on passed information
SELECT @PaySeq = 
PgpMonthlyPayPeriodNumber
FROM M_Batch, PgPayPer
WHERE MbtEEID = @EEID
	AND MbtPeriodType = 'R'
	AND MbtPayGroup NOT IN ('SALARY','SALNEX')
	AND MbtPayGroup = PgpPayGroup
	AND MbtPayDate = PgpPayDate

--Load years of service for vacat calc
SELECT @VacYears =
	ROUND(DATEDIFF(YEAR,@VacDate,@PayDate),0)

--Determine if pay date is in 01 or 06 to decide if EE should get accrual or not
SELECT @Process =
	CASE WHEN MONTH(@PayDate) IN ('1','6') AND @PaySeq = 1 THEN 'Y' ELSE 'N' END

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
		WHEN @VacYears <= 1 THEN 
							CASE
								WHEN MONTH(@VacDate) = MONTH(@PayDate) AND YEAR(@VacDate) <> YEAR(@PayDate) AND @PaySeq = 1 THEN 40
								ELSE 0
							END
		WHEN @Process = 'Y' AND @VacYears < 2 THEN 40/2
		WHEN @Process = 'Y' AND @VacYears < 10 THEN 80/2
		WHEN @Process = 'Y' AND @VacYears < 20 THEN 120/2
		WHEN @Process = 'Y' AND @VacYears >= 20 THEN 160/2
	ELSE 0
	END

+ @UWayDays

AS 'VACAT', @EEID AS 'EEID'
--INTO Tmp

END

