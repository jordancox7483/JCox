/*This will add a record to PAccHist where an employee accrual was missed for some reason*/

--SELECT * FROM PAccHist WHERE PahPerControl = 200912312 AND PahAccrOption = 'UNWVAC'

/*

SELECT * INTO U_EmpAccr_Backup FROM EmpAccr
SELECT * INTO U_PAccHist_Backup FROM PAccHist


TRUNCATE TABLE iPAccHist
INSERT INTO PaccHist
SELECT * FROM U_PAccHist_Backup

TRUNCATE TABLE EmpAccr
SET IDENTITY_INSERT EmpAccr ON
INSERT INTO EmpAccr (AuditAction ,AuditKey,EacAccrAllowedCurBal,EacAccrAllowedYTD,EacAccrAllowedYTDLOS,EacAccrCode,EacAccrOption,EacAccrPendingCurBal,EacAccrTakenCurBal,EacAccrTakenYTD,EacAccrTakenYTDLOS,EacAmtCarriedOver,EacAmtNotCarriedOver,EacCoID,EacDateAccruedThru,EacDateOfRollover,EacDatePendingMoved,EacEEID,EacLastAccrRate,EacSystemID)
SELECT * FROM U_EmpAccr_Backup
SET IDENTITY_INSERT EmpAccr OFF

*/



--BEGIN TRAN
INSERT INTO PAccHist 
(PahAccrAllowedCurAmt,
PahAccrAllowedCurBal,
PahAccrCalcExpKey,
PahAccrCalcRule,
PahAccrCode,
PahAccrCustAfterCalcPrg,
PahAccrCustCalcPrg,
PahAccrFixedDate,
PahAccrHoursOrDollars,
PahAccrInclAmt,
PahAccrInclHrs,
PahAccrOption,
PahAccrPendingCurAmt,
PahAccrPendingCurBal,
PahAccrPendingFixedDate,
PahAccrPendingNotMoved,
PahAccrSource,
PahAccrTakenCurAmt,
PahAccrTakenCurBal,
PahAccrualRate,
PahAmtCarriedOver,
PahAmtNotCarriedOver,
PahCoID,
PahDateAccruedThru,
PahDateOfRollover,
PahDatePendingMoved,
PahEEID,
PahEmpNo,
PahGenNumber,
PahGLBaseAcct,
PahGroupCreationID,
PahIncludeCurrentAmt,
PahIsRollOverRecord,
PahIsVacationPlan,
PahIsVoided,
PahIsVoidingRecord,
PahMaxAccrPerCheck,
PahMaxAccruedAllowed,
PahMaxAccruedAvailable,
PahMaxAccruedPending,
PahMaxCarryOver,
PahNoOfUnits,
PahOBType,
PahPayGroup,
PahPayoutOnRollover,
PahPendingRule,
PahPendingRuleDays,
PahPendingRuleListOfDays,
PahPerControl,
PahPostDateAccruedThru,
PahPostDateOfRollover,
PahPostDatePendingMoved,
PahRecID,
PahRolloverFixedDate,
PahRolloverNoOfUnits,
PahRolloverPer,
PahRolloverSequence,
PahTaxCalcGroupID,
PahUseCustAfterCalcPrg)

SELECT
PahAccrAllowedCurAmt = 
-1*( PahAccrAllowedCurAmt - 
	CASE
	WHEN 	
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN PahAccrAllowedCurBal > 40 THEN 40 
			ELSE COALESCE(PahAccrAllowedCurBal,0) 
		END
> 40 THEN 40 
	ELSE
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN PahAccrAllowedCurBal > 40 THEN 40 
			ELSE COALESCE(PahAccrAllowedCurBal,0) 
		END
END),
PahAccrAllowedCurBal = 	
	CASE
	WHEN 	
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
> 40 THEN 40 
	ELSE
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
END, 
PahAccrCalcExpKey = '', 
PahAccrCalcRule = AccAccrCalcRule,
PahAccrCode = EacAccrCode, 
PahAccrCustAfterCalcPrg = NULL, 
PahAccrCustCalcPrg = '', 
PahAccrFixedDate = AccAccrFixedDate,
PahAccrHoursOrDollars = AccAccrHoursOrDollars, 
PahAccrInclAmt = 0,
PahAccrInclHrs = 0,
PahAccrOption = EacAccrOption,
PahAccrPendingCurAmt = 0,
PahAccrPendingCurBal = 0,
PahAccrPendingFixedDate = NULL,
PahAccrPendingNotMoved = 0,
PahAccrSource = 'C',
PahAccrTakenCurAmt = 0,
PahAccrTakenCurBal = 0,
PahAccrualRate = 0,
PahAmtCarriedOver =
	CASE
	WHEN 	
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
> 40 THEN 40 
	ELSE
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
END, 

PahAmtNotCarriedOver = CASE WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal < 40 THEN 0 ELSE
EacAccrAllowedCurBal - EacAccrTakenCurBal -
	CASE
	WHEN 	
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
< 40 THEN 0 
	ELSE
		CASE							--Regular Accrual
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 5 THEN 3.07692
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 10 THEN 3.84615
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 15 THEN 4.61538
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 20 THEN 5.38462
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 30 THEN 6.15384
			WHEN (2009 - YEAR(EecDateOfSeniority)) < 99 THEN 7.6923
			ELSE 0
		END
	+
		CASE							--Rollover Amount 
			WHEN EacAccrAllowedCurBal - EacAccrTakenCurBal > 40 THEN 40 
			ELSE COALESCE(EacAccrAllowedCurBal - EacAccrTakenCurBal,0) 
		END
END
END, 

PahCoID = EacCoID,
PahDateAccruedThru = '2009-12-31 00:00:00.000',
PahDateOfRollover = NULL,
PahDatePendingMoved = NULL,
PahEEID = EacEEID,
PahEmpNo = EecEmpNo,
PahGenNumber = PrgGenNumber,
PahGLBaseAcct = NULL,
PahGroupCreationID = NULL,
PahIncludeCurrentAmt = 'N',
PahIsRollOverRecord = 'Y',
PahIsVacationPlan = 'N',
PahIsVoided = 'N',
PahIsVoidingRecord = 'N',
PahMaxAccrPerCheck = NULL,
PahMaxAccruedAllowed = NULL,
PahMaxAccruedAvailable = NULL,
PahMaxAccruedPending = NULL,
PahMaxCarryOver = NULL,
PahNoOfUnits = NULL,
PahOBType = NULL,
PahPayGroup = EecPayGroup,
PahPayoutOnRollover = AccPayoutOnRollover,
PahPendingRule = AccPendingRule,
PahPendingRuleDays = AccPendingRuleDays,
PahPendingRuleListOfDays = AccPendingRuleListOfDays,
PahPerControl = PrgPerControl,
PahPostDateAccruedThru = '2009-12-31 00:00:00.000',
PahPostDateOfRollover = NULL,
PahPostDatePendingMoved = NULL,
PahRecID = PahRecID + 250,
PahRolloverFixedDate = AccRolloverFixedDate,
PahRolloverNoOfUnits = AccRolloverNoOfUnits,
PahRolloverPer = AccRolloverPer,
PahRolloverSequence = AccRolloverSequence,
PahTaxCalcGroupID = PrgTaxCalcGroupID,
PahUseCustAfterCalcPrg = AccUseCustAfterCalcPrg

FROM AccrOpts, EmpAccr, PayReg, EmpComp, PAccHist
WHERE AccAccrCode = EacAccrCode
	AND AccAccrOption = EacAccrOption
	AND EacAccrCode = 'VACAT'
	AND EecEEID = EacEEID
	AND EecCoID = EacCoID
	AND PahCoID = PrgCoID
	AND PrgEEID = EacEEID
	AND PahEEID = EacEEID
	AND	PahCoID = EacCoID
	AND PrgPayGroup = 'BIWEEK'
	--AND PrgTransactionType IN ('D','C')
	AND PahAccrOption = 'UNWVAC'
	AND PrgPerControl = (SELECT MAX(PrgPerControl) FROM PayReg WHERE PrgPayGroup = 'BIWEEK' AND PrgEEID = EecEEID AND PrgCoID = EecCoID GROUP BY PrgEEID) --'200912312'
	AND PahAccrCode = 'VACAT' 
	AND PahAccrOption = 'UNWVAC'
	AND PahPerControl = (SELECT MAX(PahPerControl) FROM PAccHist WHERE PahEEID = EacEEID AND PahAccrCode = 'VACAT' AND PahAccrOption = 'UNWVAC' )
















--BEGIN TRAN
UPDATE EmpAccr
SET 
--SELECT
EacAccrTakenCurBal = 0,
EacAccrTakenYTD = 0,
EacAccrTakenYTDLOS = 0,
EacAccrAllowedCurBal = PahAccrAllowedCurBal,
EacAccrAllowedYTD = PahAccrAllowedCurBal,
EacAccrAllowedYTDLOS = 	PahAccrAllowedCurBal,
EacAmtCarriedOver = PahAmtCarriedOver,
EacAmtNotCarriedOver = PahAmtNotCarriedOver

FROM AccrOpts, PayReg, EmpComp, PAccHist--, EmpAccr
WHERE AccAccrCode = EacAccrCode
	AND AccAccrOption = EacAccrOption
	AND EacAccrCode = 'VACAT'
	AND EecEEID = EacEEID
	AND EecCoID = EacCoID
	AND PahCoID = PrgCoID
	AND PrgEEID = EacEEID
	AND PahEEID = EacEEID
	AND	PahCoID = EacCoID
	AND PrgPayGroup = 'BIWEEK'
	--AND PrgTransactionType IN ('D','C')
	AND PahAccrOption = 'UNWVAC'
	AND PrgPerControl = (SELECT MAX(PrgPerControl) FROM PayReg WHERE PrgPayGroup = 'BIWEEK' AND PrgEEID = EecEEID AND PrgCoID = EecCoID GROUP BY PrgEEID) --'200912312'
	AND PahAccrCode = 'VACAT' 
	AND PahAccrOption = 'UNWVAC'
	AND PahPerControl = (SELECT MAX(PahPerControl) FROM PAccHist WHERE PahEEID = EacEEID AND PahAccrCode = 'VACAT' AND PahAccrOption = 'UNWVAC' )
	AND PahAccrCalcExpKey = ''
--ROLLBACK
