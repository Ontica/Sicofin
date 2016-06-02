<% 
	Option Explicit	
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
%>

<!--#INCLUDE virtual="/empiria/bin/ms_scripts/rs.asp"-->

<% RSDispatch %>

<SCRIPT RUNAT=SERVER Language=javascript>

function IServerScripts()	{	
	this.AddSubsidiaryAccount = Function('sSubsAcct', 'nAccountId', 'sSectorKey', 'return AddSubsidiaryAccount(sSubsAcct, nAccountId, sSectorKey)');  
	this.AssignStdAccount = Function('nGralLedgerId', 'nStdAccountId', 'return AssignStdAccount(nGralLedgerId, nStdAccountId)');		
	this.AssignSubsidiaryAccount = Function('sSubsAcct', 'nAccountId', 'sSectorKey', 'return AssignSubsidiaryAccount(sSubsAcct, nAccountId, sSectorKey)');
	this.AssignSubsidiaryLedger = Function('nSubsLedgerId', 'nAccountId', 'return AssignSubsidiaryLedger(nSubsLedgerId, nAccountId)');		
	this.CalculateBaseAmount = Function('nExchangeRate', 'nAmount', 'return CalculateBaseAmount(nExchangeRate, nAmount)');
	this.CalculateWithLastExchange = Function('nFromCurrencyId', 'nToCurrencyId', 'nAmount', 'return CalculateWithLastExchange(nFromCurrencyId, nToCurrencyId, nAmount)');
	this.CanDeleteGL = Function('nGralLedgerId','return CanDeleteGL(nGralLedgerId)');
	this.CboAccountCurrencies = Function('nAccountId', 'nSelCurrencyId', 'return CboAccountCurrencies(nAccountId, nSelCurrencyId)');
	this.CboAccountSubsidiaryLedgers = Function('nAccountId', 'nSectorId', 'nSubsLedgerId', 'return CboAccountSubsidiaryLedgers(nAccountId, nSectorId, nSubsLedgerId)');
	this.CboGeneralLedgerCategories = Function('nGralLedgerClipId', 'nSelCategoryId', 'nFilterOperation', 'return CboGeneralLedgerCategories(nGralLedgerClipId, nSelCategoryId, nFilterOperation)');
	this.CboGeneralLedgerFilledCategories = Function('nGralLedgerClipId', 'nSelCategoryId', 'nFilterOperation', 'return CboGeneralLedgerFilledCategories(nGralLedgerClipId, nSelCategoryId, nFilterOperation)');
	this.CboGLGroupsForRule = Function('nRuleId', 'nGroupClip', 'nSelItem', 'return CboGLGroupsForRule(nRuleId, nGroupClip, nSelItem)');
	this.CboGLInCategory = Function('nCategoryId', 'return CboGLInCategory(nCategoryId)');
	this.CboGLSources = Function('nGralLedgerId', 'nSourceId', 'return CboGLSources(nGralLedgerId, nSourceId)');
	this.CboGralLedgersInGroup = Function('nGroupId', 'nSelectedItemId', 'return CboGralLedgersInGroup(nGroupId, nSelectedItemId)');
	this.CboGralLedgersInGroup2 = Function('nGroupId', 'return CboGralLedgersInGroup2(nGroupId)');
	this.CboOpenPeriodsDates = Function('nGralLedgerId', 'return CboOpenPeriodsDates(nGralLedgerId)');
	this.CboRuleChilds = Function('nRuleId', 'return CboRuleChilds(nRuleId)');
	this.CboRuleReports = Function('nRuleId', 'return CboRuleReports(nRuleId)');
	this.CboSectorsInAccount = Function('nAccountId', 'nSectorId', 'return CboSectorsInAccount(nAccountId, nSectorId)');	  
	this.CboSubsidiaryAccounts = Function('nSubsLedgerId', 'nSubsAccountId', 'return CboSubsidiaryAccounts(nSubsLedgerId, nSubsAccountId)');
	this.CheckAmounts = Function('nAmount', 'nExchangeRate', 'nBaseAmount', 'return CheckAmounts(nAmount, nExchangeRate, nBaseAmount)');
	this.CheckAvailableExchangeRates = Function('dFromDate', 'dToDate', 'nExchRateType', 'dExchRateDate', 'return CheckAvailableExchangeRates(dFromDate, dToDate, nExchRateType, dExchRateDate)');
	this.CheckPostingValues = Function('dVoucherDate', 'nAccountId', 'sSector', 'sSubsAcct', 'sBudgetKey', 'sResponsibilityArea', 'return CheckPostingValues(dVoucherDate, nAccountId, sSector, sSubsAcct, sBudgetKey, sResponsibilityArea)');
	this.CompareDates = Function('sDate1', 'sDate2', 'return CompareDates(sDate1, sDate2)');
	this.CurrencyName = Function('nCurrencyId', 'return CurrencyName(nCurrencyId)');
	this.CurrentExchangeRate = Function('nFromCurrencyId', 'nToCurrencyId', 'return CurrentExchangeRate(nFromCurrencyId, nToCurrencyId)');
	this.DeleteRule = Function('nRuleId', 'return DeleteRule(nRuleId)');
	this.DeleteRuleGroup = Function('nRuleGroupId', 'return DeleteRuleGroup(nRuleGroupId)');
	this.ExchangeRate = Function('nFromCurrencyId', 'nToCurrencyId', 'dDate', 'return ExchangeRate(nFromCurrencyId, nToCurrencyId , dDate)');
	this.ExistsCurrencyKey = Function('sCurrencyKey', 'return ExistsCurrencyKey(sCurrencyKey)');
	this.ExistsSubsAcctNumber = Function('nSubsLedgerId', 'sSubsAcct', 'return ExistsSubsAcctNumber(nSubsLedgerId, sSubsAcct)');
	this.FormatCurrency = Function('sAmount', 'nDecimals', 'return FormatCurrency(sAmount, nDecimals)');
	this.FormatDate = Function('dDate', 'sFormat', 'return FormatDate(dDate, sFormat)');	  
	this.FormatStdAccountNumber = Function('nStdAccountTypeId', 'sAccount', 'return FormatStdAccountNumber(nStdAccountTypeId, sAccount)');	  
	this.FormatStdAccountWithGLId = Function('nGralLedgerId', 'sStdAccount', 'return FormatStdAccountWithGLId(nGralLedgerId, sStdAccount)');	  
	this.FormatSubsAccount = Function('sSubsAcct', 'return FormatSubsAccount(sSubsAcct)');	  
	this.FormatSubsidiaryAccount = Function('sGralLedgerNumber', 'sSubsAcct', 'return FormatSubsidiaryAccount(sGralLedgerNumber, sSubsAcct)');
	this.FormatSubsidiaryAccountWithSLId = Function('nSubsLedgerId', 'sSubsAcct', 'return FormatSubsidiaryAccountWithSLId(nSubsLedgerId, sSubsAcct)');
	this.FormatWildCharsList = Function('sList', 'return FormatWildCharsList(sList)');
	this.GetAccountId = Function('nGralLedgerId', 'sAccount', 'return GetAccountId(nGralLedgerId, sAccount)');
	this.GetAccountProps = Function('dVoucherDate', 'nAccountId', 'nSectorId', 'return GetAccountProps(dVoucherDate, nAccountId, nSectorId)');
	this.GetAreasList = Function('sAreasList', 'bIncludeCheckBoxes', 'return GetAreasList(sAreasList, bIncludeCheckBoxes)');
	this.GetNextSubsAccountNumber = Function('nSubsLedgerId', 'return GetNextSubsAccountNumber(nSubsLedgerId)');	  
	this.GetSectorId = Function('sSectorKey', 'return GetSectorId(sSectorKey)');
	this.GetStdAccountId = Function('nStdAccountTypeId', 'sStdAccountNumber', 'return GetStdAccountId(nStdAccountTypeId, sStdAccountNumber)');
	this.GetStdAccountParentNumber = Function('nStdAccountTypeId', 'sStdAccountNumber', 'return GetStdAccountParentNumber(nStdAccountTypeId, sStdAccountNumber)');
	this.GetStdAccountParentRole = Function('nStdAccountTypeId', 'sStdAccountNumber', 'return GetStdAccountParentRole(nStdAccountTypeId, sStdAccountNumber)');
	this.GetGLSubsidiaryLedgerPrefix = Function('nGralLedgerId', 'return GetGLSubsidiaryLedgerPrefix(nGralLedgerId)');
	this.GLAccountRole = Function('nGralLegerId', 'sAccountNumber', 'return GLAccountRole(nGralLegerId, sAccountNumber)');	  
	this.GLGroupIsRoot = Function('nGLGroupId', 'return GLGroupIsRoot(nGLGroupId)');	  
	this.GLGroupStdAccountId = Function('nGLGroupId', 'return GLGroupStdAccountId(nGLGroupId)');	  
	this.IsDate = Function('sDate', 'return IsDate_(sDate)');
	this.IsDateInGLPeriod = Function('nGralLedgerId', 'dDate', 'return IsDateInGLPeriod(nGralLedgerId, dDate)');
	this.IsNumeric = Function('sNumber', 'nDecimals', 'return IsNumeric_(sNumber, nDecimals)');
	this.IsStdAccountNumberValid = Function('nStdAccountTypeId', 'sStdAccountNumber', 'return IsStdAccountNumberValid(nStdAccountTypeId, sStdAccountNumber)');
	this.IsSubsAccountNumberValid = Function('sSubsAccountNumber', 'return IsSubsAccountNumberValid(sSubsAccountNumber)');	  
	this.PendingPostingsReferences = Function('nGralLedgerId', 'nSelPosting', 'return PendingPostingsReferences(nGralLedgerId, nSelPosting)');
	this.PostingsTable = Function('nTransactionId', 'bAnalize', 'return PostingsTable(nTransactionId, bAnalize)');
	this.RuleChildsCount = Function('nRuleId', 'return RuleChildsCount(nRuleId)');		    
	this.RuleChildsType = Function('nRuleId', 'return RuleChildsType(nRuleId)');  		
	this.RuleItemsSection = Function('nRuleDefId', 'nRuleId', 'return RuleItemsSection(nRuleDefId, nRuleId)');		
	this.RuleLabel = Function('nRuleId', 'return RuleLabel(nRuleId)');
	this.RuleType = Function('nRuleId', 'return RuleType(nRuleId)');	  
	this.StandardAccountRole = Function('nStdAccountId', 'return StandardAccountRole(nStdAccountId)');
	this.StdAccountId = Function('nStdAccountTypeId', 'sStdAccountNumber', 'return StdAccountId(nStdAccountTypeId, sStdAccountNumber)');
	this.SubsAccountExtendedAttrs = Function('nSubsLedgerId', 'nSubsAccountId', 'return SubsAccountExtendedAttrs(nSubsLedgerId, nSubsAccountId)');
	this.SubsidiaryAccountName = Function('nGralLedgerId', 'sSubsAcctNumber', 'return SubsidiaryAccountName(nGralLedgerId, sSubsAcctNumber)');
	this.SubsidiaryAccountNumber = Function('nSubsAccountId', 'bShowComplete', 'return SubsidiaryAccountNumber(nSubsAccountId, bShowComplete)');
	this.TblSubsidiaryAccounts = Function('nSubsLedgerId', 'nSubsAccountId', 'sOrderBy', 'return TblSubsidiaryAccounts(nSubsLedgerId, nSubsAccountId, sOrderBy)');
	this.TransactionStatus = Function('nTransactionId', 'return TransactionStatus(nTransactionId)');
	this.TransactionType = Function('nVoucherId', 'return TransactionType(nVoucherId)');
	this.ValidateAreas = Function('sAreasList', 'return ValidateAreas(sAreasList)');
	this.ValidateTransaction = Function('nTransactionId', 'return ValidateTransaction(nTransactionId)');
	this.VoucherId = Function('nGralLedgerId', 'sVoucherNumber', 'return VoucherId(nGralLedgerId, sVoucherNumber)');	  
}

public_description = new IServerScripts();  

</SCRIPT>

<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">

Function AddSubsidiaryAccount(sSubsAcct, nAccountId, sSectorKey)
	Dim oGralLedger
	'*************************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CServer")
	AddSubsidiaryAccount = oGralLedger.AddSubsidiaryAccount(Session("sAppServer"), CStr(sSubsAcct), CLng(nAccountId), CStr(sSectorKey))
	Set oGralLedger = Nothing
End Function

Function AssignStdAccount(nGralLedgerId, nStdAccountId)
	Dim oGralLedger
	'****************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CGralLedger")
	AssignStdAccount = oGralLedger.AssignStdAccount(Session("sAppServer"), CLng(nGralLedgerId), CLng(nStdAccountId))
	Set oGralLedger = Nothing
End Function

Function AssignSubsidiaryAccount(sSubsAcct, nAccountId, sSectorKey)
  Dim oGralLedger
  '****************************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CServer")
	AddSubsidiaryAccount = oGralLedger.AssignSubsidiaryAccount(Session("sAppServer"), CStr(sSubsAcct), CLng(nAccountId), CStr(sSectorKey))
	Set oGralLedger = Nothing
End Function

Function AssignSubsidiaryLedger(nSubsLedgerId, nAccountId)
  Dim oGralLedger
  '*******************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
	oGralLedger.AddMapping Session("sAppServer"), CLng(nSubsLedgerId), CLng(nAccountId)
	Set oGralLedger = Nothing
End Function

Function CalculateBaseAmount(nExchangeRate, nAmount)
  Dim oGLVoucherUS
  '*************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	CalculateBaseAmount = oGLVoucherUS.CalculateBaseAmount(CStr(nExchangeRate), CStr(nAmount))
	Set oGLVoucherUS = Nothing
End Function
 
Function CalculateWithLastExchange(nFromCurrencyId, nToCurrencyId, nAmount)
  Dim oCurrency
  '************************************************************************
  On Error Resume Next
	Set oCurrency = Server.CreateObject("CurrencyMgr.CManager")
  CalculateWithLastExchange = CCur(nAmount) * oCurrency.LastExchange(Session("sAppServer"), nFromCurrencyId, nToCurrencyId)
  Set oCurrency = Nothing
End Function

Function CanDeleteGL(nGralLedgerId)
  Dim oGralLedger, sTemp	
	'********************************
	On Error Resume Next
	Set oGralLedger = Server.CreateObject("AOGralLedger.CGralLedger")
  CanDeleteGL = oGralLedger.CanDelete(Session("sAppServer"), CLng(nGralLedgerId))
  Set oGralLedger = Nothing
End Function

Function CboAccountCurrencies(nAccountId, nSelCurrencyId)
  Dim oGLVoucherUS, sTemp	
	'******************************************************
	On Error Resume Next
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboAccountCurrencies(Session("sAppServer"), CLng(nAccountId), CLng(nSelCurrencyId))
  Set oGLVoucherUS = Nothing    
	sTemp = "<SELECT name=cboCurrencies style='width:180' LANGUAGE=javascript onchange='return cboCurrencies_onchange()'>" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboAccountCurrencies = sTemp
End Function

Function CboAccountSubsidiaryLedgers(nAccountId, nSectorId, nSubsLedgerId)
  Dim oGLVoucherUS, sTemp	
	'***********************************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboSubsidaryLedgers(Session("sAppServer"), CLng(nAccountId), CLng(nSectorId), CLng(nSubsLedgerId))
  Set oGLVoucherUS = Nothing  
	sTemp = "<SELECT name=cboSubsidiaryLedgers style=""WIDTH: 100%"" onchange=""return updateCombo('cboSubsidiaryAccounts')"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf	
	CboAccountSubsidiaryLedgers = sTemp
End Function

Function CboGeneralLedgerCategories(nGralLedgerClipId, nSelCategoryId, nFilterOperation)
  Dim oGLVoucherUS, sTemp	
	'*************************************************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), CLng(Session("uid")), _
																									CLng(nGralLedgerClipId), CLng(nSelCategoryId), _
																									CLng(nFilterOperation))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGLCategories style=""WIDTH: 520px"" onchange=""return updateCboGralLedgers()"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboGeneralLedgerCategories = sTemp
End Function

Function CboGeneralLedgerFilledCategories(nGralLedgerClipId, nSelCategoryId, nFilterOperation)
  Dim oGLVoucherUS, sTemp	
	'*******************************************************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), CLng(Session("uid")), _
																									CLng(nGralLedgerClipId), CLng(nSelCategoryId), _
																									CLng(nFilterOperation))
  Set oGLVoucherUS = Nothing
 	sTemp = "<SELECT name=cboGLCategories style=""WIDTH: 520px"" onchange=""return updateCboGralLedgers()"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboGeneralLedgerFilledCategories = sTemp
End Function

Function CboGLGroupsForRule(nRuleId, nGroupClip, nSelItem) 
	Dim oReportsEngine, sTemp
	'*******************************************************
	Set oReportsEngine = Server.CreateObject("AOReportsEngine.CEngine")
	sTemp = oReportsEngine.CboGLGroupsForRule(Session("sAppServer"), CLng(nRuleId), CLng(Session("uid")), _
																						CLng(nGroupClip), CLng(nSelItem))
  Set oReportsEngine = Nothing
  
	sTemp = "<SELECT name=cboGLGroups style='width:100%' onchange='return updateGralLedgers();'>" & vbCrLf & _
					sTemp & _
					"</SELECT>"	
	CboGLGroupsForRule = sTemp	
End Function

Function CboGLInCategory(nCategoryId)
  Dim oGLVoucherUS, sTemp	
	'**********************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGralLedgers(Session("sAppServer"), CLng(Session("uid")), CLng(nCategoryId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGralLedgers style=""WIDTH: 520px"" onchange=""return updateInfo()"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboGLInCategory = sTemp
End Function

Function CboGLSources(nGralLedgerId, nSourceId)
  Dim oGLVoucherUS, sTemp	
	'********************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     	
  sTemp = oGLVoucherUS.CboGLSources(Session("sAppServer"), CLng(nGralLedgerId), CLng(nSourceId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGLSources style=""WIDTH: 520px"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboGLSources = sTemp
End Function

Function CboGralLedgersInGroup(nGroupId, nSelectedItemId)
  Dim oGLVoucherUS, sTemp	
	'*************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGralLedgers(Session("sAppServer"), CLng(Session("uid")), CLng(nGroupId), CLng(nSelectedItemId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGralLedgers style='width: 100%'>" & VbCrLf & _
					"<OPTION SELECTED value=0>-- Todas las contabilidades en el grupo seleccionado--</OPTION>" & VbCrLf & _
					sTemp & _
			    "</SELECT>" & VbCrLf
	CboGralLedgersInGroup = sTemp
End Function

Function CboGralLedgersInGroup2(nGroupId)
  Dim oGLVoucherUS, sTemp	
	'**************************************
	On Error Resume Next
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGralLedgers(Session("sAppServer"), CLng(Session("uid")), CLng(nGroupId))
  Set oGLVoucherUS = Nothing  
  If Len(sTemp) = 0 Then
		sTemp = "<OPTION SELECTED value=0>(El grupo seleccionado no tiene contabilidades)</OPTION>" & VbCrLf
	End If
	sTemp = "<SELECT name=cboGralLedgers style='WIDTH: 100%'>" & VbCrLf & _					
					sTemp & _
			    "</SELECT>" & VbCrLf
	CboGralLedgersInGroup2 = sTemp
End Function

Function CboOpenPeriodsDates(nGralLedgerId)
  Dim oGLVoucherUS, sTemp
	'****************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboOpenPeriodsDates(Session("sAppServer"), CLng(nGralLedgerId))
  Set oGLVoucherUS = Nothing
  sTemp = "<SELECT name='cboApplicationDates' style='WIDTH: 130px'>" & VbCrLf & _
					 sTemp & _
				  "</SELECT>" & VbCrLf
  CboOpenPeriodsDates = sTemp  
End Function

Function CboRuleChilds(nRuleId)
  Dim oRule
  '****************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  CboRuleChilds = oRule.CboRuleChilds(Session("sAppServer"), CLng(nRuleId))
  Set oRule = Nothing
End Function

Function CboRuleReports(nRuleId)
	Dim oReportsEngine, sTemp
	'*****************************
	Set oReportsEngine = Server.CreateObject("AOReportsEngine.CEngine")	
	sTemp = oReportsEngine.CboRuleReports(Session("sAppServer"), CLng(nRuleId), Session("uid"))
	Set oReportsEngine = Nothing
	
	sTemp = "<SELECT name=cboRuleReports style='width:240'>" & vbCrLf & _
					sTemp & _ 
					"</SELECT>" & vbCrLf
	CboRuleReports = sTemp
End Function

Function CboSectorsInAccount(nAccountId, nSectorId)
  Dim oGLVoucherUS, sTemp	
	'************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboAccountSectors(Session("sAppServer"), CLng(nAccountId),CLng(nSectorId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboSectors style=""WIDTH: 100%"" onchange=""return updateFromSectors()"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboSectorsInAccount = sTemp
End Function

Function CboSubsidiaryAccounts(nSubsLedgerId, nSubsAccountId)
  Dim oGLVoucherUS, sTemp	
	'**********************************************************
	On Error Resume Next
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboSubsidiaryAccounts(Session("sAppServer"), CLng(nSubsLedgerId), CLng(nSubsAccountId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboSubsidiaryAccounts style=""WIDTH: 100%"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboSubsidiaryAccounts = sTemp
End Function

Function CheckAmounts(nAmount, nExchangeRate, nBaseAmount)
  Dim oGLVoucherUS
  '*******************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	CheckAmounts = oGLVoucherUS.CheckAmounts(CStr(nAmount), CStr(nExchangeRate), CStr(nBaseAmount))
	Set oGLVoucherUS = Nothing
End Function

Function CheckAvailableExchangeRates(dFromDate, dToDate, nExchRateType, dExchRateDate)
  Dim oGLReports
	'***********************************************************************************
	On Error Resume Next	
	Set oGLReports = Server.CreateObject("SCFFixedReports.CReports")
  CheckAvailableExchangeRates = oGLReports.CheckAvailableExchangeRates(Session("sAppServer"), _
																																			 CDate(dFromDate), CDate(dToDate), _
																																			 CLng(nExchRateType), _
																																			 CDate(dExchRateDate))
  Set oGLReports = Nothing
End Function

Function CheckPostingValues(dVoucherDate, nAccountId, sSector, sSubsAcct, sBudgetKey, sResponsibilityArea)
  Dim oVoucherBS
	'*****************************************************************************************
	On Error Resume Next		
	Set oVoucherBS = Server.CreateObject("AOGLVoucher.CServer")		
	CheckPostingValues = oVoucherBS.CheckPostingValues(Session("sAppServer"), CDate(dVoucherDate), CLng(nAccountId), _
																										 CStr(sSector), CStr(sSubsAcct), CStr(sBudgetKey), _
																										 CStr(sResponsibilityArea))
  Set oVoucherBS = Nothing
End Function

Function CompareDates(sDate1, sDate2)
	If (CDate(sDate1) <= CDate(sDate2)) Then
		CompareDates = True
	Else
		CompareDates = False
	End If
End Function

Function CurrencyName(nCurrencyId)
  Dim oCurrenciesUS
	'*******************************
	On Error Resume Next	
	Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")		
  CurrencyName = oCurrenciesUS.CurrencyName(Session("sAppServer"), CLng(nCurrencyId), True)
  Set oCurrenciesUS = Nothing
End Function

Function CurrentExchangeRate(nFromCurrencyId, nToCurrencyId)
	Dim oCurrency
  '*********************************************************
  On Error Resume Next	
	Set oCurrency = Server.CreateObject("AOGLVoucherUS.CVoucher")
  CurrentExchangeRate = oCurrency.CurrentExchangeRate(Session("sAppServer"), CLng(nFromCurrencyId), CLng(nToCurrencyId))
  Set oCurrency = Nothing	
End Function

Function DeleteRule(nRuleId)
  Dim oRule
  '*************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  oRule.DeleteRule Session("sAppServer"), CLng(nRuleId)
  Set oRule = Nothing
End Function

Function DeleteRuleGroup(nRuleGroupId)
  Dim oRule 
  '***********************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  oRule.DeleteRuleGroup Session("sAppServer"), CLng(nRuleGroupId)
  Set oRule = Nothing
End Function    

Function ExchangeRate(nFromCurrencyId, nToCurrencyId, dDate)
  Dim oGralLedgerUS
  '*********************************************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
	ExchangeRate = oGralLedgerUS.ExchangeRate(Session("sAppServer"), CLng(nFromCurrencyId), CLng(nToCurrencyId), CDate(dDate))
	Set oGralLedgerUS = Nothing
End Function

Function ExistsCurrencyKey(sCurrencyKey)
  Dim oCurrencyUS
  '*************************************
	On Error Resume Next	
	Set oCurrencyUS = Server.CreateObject("AOCurrenciesUS.CServer")
	ExistsCurrencyKey = oCurrencyUS.ExistsKey(CStr(Session("sAppServer")), CStr(sCurrencyKey))
	Set oCurrencyUS = Nothing
End Function

Function ExistsSubsAcctNumber(nSubsLedgerId, sSubsAcct)
  Dim oGralLedger
  '****************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
	ExistsSubsAcctNumber = oGralLedger.ExistsSubsidiaryAccount(Session("sAppServer"), _
																														 CLng(nSubsLedgerId), CStr(sSubsAcct))
	Set oGralLedger = Nothing
End Function

Function FormatCurrency(sAmount, nDecimals)
  Dim oGLVoucherUS
  '****************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	FormatCurrency = oGLVoucherUS.FormatCurrency(CStr(sAmount), CLng(nDecimals))
	Set oGLVoucherUS = Nothing
End Function

Function FormatDate(dDate, sFormat)
	Dim oVoucherUS
	'********************************
	On Error Resume Next	
	Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	FormatDate = oVoucherUS.FormatDate(CDate(dDate), CStr(sFormat))
	Set oVoucherUS = Nothing	
End Function

Function FormatStdAccountNumber(nStdAccountTypeId, sAccount)
  Dim oStdAccount, sChar
  '*********************************************************
	On Error Resume Next		
	sChar = Right(sAccount, 1)
	If (sChar = "%") OR (sChar = "*") Then
		sAccount = Left(sAccount, Len(sAccount) - 1)
	Else 
		sChar = ""
	End If
	Set oStdAccount = Server.CreateObject("AOGralLedger.CStandardAccount")
	FormatStdAccountNumber = oStdAccount.FormatAccountNumber(Session("sAppServer"), CLng(nStdAccountTypeId), CStr(sAccount)) & sChar
	Set oStdAccount = Nothing
End Function

Function FormatStdAccountWithGLId(nGralLedgerId, sStdAccount)
  Dim oStdAccount
  '**********************************************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("AOGralLedger.CStandardAccount")
	FormatStdAccountWithGLId = oStdAccount.FormatAccountNumberWithGLId(Session("sAppServer"), CLng(nGralLedgerId), CStr(sStdAccount))
	Set oStdAccount = Nothing
End Function

Function FormatSubsAccount(sSubsAcct)
	Dim sTemp
	'**********************************	
	If (Left(sSubsAcct, 1) = "%") Or (Left(sSubsAcct, 1) = "*") Then
		FormatSubsAccount = sSubsAcct
		Exit Function
	End If
  sTemp = Right(String(16, "0") & sSubsAcct, 16)
  FormatSubsAccount = sTemp
End Function

Function FormatSubsidiaryAccount(sGralLedgerNumber, sSubsAcct)
  Dim oGLVoucherUS
  '***********************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	FormatSubsidiaryAccount = oGLVoucherUS.FormatSubsidiaryAccount(CStr(sGralLedgerNumber), CStr(sSubsAcct))
	Set oGLVoucherUS = Nothing		
End Function

Function FormatSubsidiaryAccountWithSLId(nSubsLedgerId, sSubsAcct)
  Dim oGLVoucherUS, sTemp
  '***************************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	sTemp = oGLVoucherUS.FormatSubsidiaryAccountWithSLId(Session("sAppServer"), CLng(nSubsLedgerId), CStr(sSubsAcct))
	FormatSubsidiaryAccountWithSLId = Mid(sTemp, Len(sTemp) - 16 + 1)
	Set oGLVoucherUS = Nothing		
End Function

Function FormatWildCharsList(sList)
  Dim oStdAccount
  '********************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
	FormatWildCharsList = oStdAccount.FormatWildCharsList(CStr(sList))
	Set oStdAccount = Nothing
End Function

Function GetAccountId(nGralLedgerId, sAccount)
  Dim oGLVoucherUS
  '*******************************************
	On Error Resume Next
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	GetAccountId = oGLVoucherUS.GetAccountId(Session("sAppServer"), CLng(nGralLedgerId), CStr(sAccount))
	Set oGLVoucherUS = Nothing
End Function

Function GetAccountProps(dVoucherDate, nAccountId, nSectorId)
  Dim oGralLedgerUS
  '********************************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGLVoucher.CServer")
	GetAccountProps = oGralLedgerUS.GetAccountProps(Session("sAppServer"), CDate(dVoucherDate), CLng(nAccountId), CLng(nSectorId))
	Set oGralLedgerUS = Nothing
End Function

Function GetAreasList(sAreasList, bIncludeCheckBoxes)
  Dim oStdAccount, sTemp
  '**************************************************
	On Error Resume Next	
	sTemp = "<TABLE border=0 cellPadding=1 cellSpacing=0 width=100%>" & vbCrLf
	If Len(sAreasList) <> 0 Then
		Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		sTemp = sTemp & oStdAccount.ResponsibilityAreasListWithArray(CStr(Session("sAppServer")), CStr(sAreasList), CBool(bIncludeCheckBoxes))
		Set oStdAccount = Nothing
	Else
		sTemp = sTemp & "<TR><TD><FONT color=maroon>No hay áreas seleccionadas.</FONT></TD></TR>"
	End If
	sTemp = sTemp & "</TABLE>" & vbCrLf
  GetAreasList = sTemp
End Function

Function GetNextSubsAccountNumber(nSubsLedgerId) 
  Dim oSubsLedger, sTemp
  '*********************************************
	On Error Resume Next	
	Set oSubsLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
	sTemp = oSubsLedger.GetNextSubsidiaryAccountNumber(Session("sAppServer"), CLng(nSubsLedgerId))
	If (Len(sTemp) <> 0) Then
		GetNextSubsAccountNumber = Mid(sTemp, Len(sTemp) - 16 + 1)
	Else
		GetNextSubsAccountNumber = sTemp
	End If	
	Set oSubsLedger = Nothing
End Function 

Function GetSectorId(sSectorKey)
  Dim oGLVoucherUS
  '*****************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	GetSectorId = oGLVoucherUS.SectorId(Session("sAppServer"), CStr(sSectorKey))
	Set oGLVoucherUS = Nothing
End Function

Function GetStdAccountId(nStdAccountTypeId, sStdAccountNumber)
  Dim oStdAccount
  '***********************************************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
	GetStdAccountId = oStdAccount.GetStdAccountId(Session("sAppServer"), _
																				        CLng(nStdAccountTypeId), CStr(sStdAccountNumber))
	Set oStdAccount = Nothing
End Function

Function GetStdAccountParentNumber(nStdAccountTypeId, sStdAccountNumber)
  Dim oStdAccount, oRecordset
  '*********************************************************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
	Set oRecordset = oStdAccount.GetParent(Session("sAppServer"), _
																				 CLng(nStdAccountTypeId), CStr(sStdAccountNumber))
	Set oStdAccount = Nothing
	
	If Not (oRecordset Is Nothing) Then
		GetStdAccountParentNumber = oRecordset.Fields("numero_cuenta_estandar")
	Else
		GetStdAccountParentNumber = ""
	End If
	oRecordset.Close
	Set oRecordset = Nothing
End Function

Function GetStdAccountParentRole(nStdAccountTypeId, sStdAccountNumber)
  Dim oStdAccount, oRecordset
  '*******************************************************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
	Set oRecordset = oStdAccount.GetParent(Session("sAppServer"), _
																				 CLng(nStdAccountTypeId), CStr(sStdAccountNumber))
	Set oStdAccount = Nothing
	
	If Not (oRecordset Is Nothing) Then
		GetStdAccountParentRole = oRecordset.Fields("rol_cuenta")
	Else
		GetStdAccountParentRole = ""
	End If
	oRecordset.Close
	Set oRecordset = Nothing
End Function

Function GetGLSubsidiaryLedgerPrefix(nGralLedgerId)
  Dim oGLVoucherUS
  '************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	GetGLSubsidiaryLedgerPrefix = oGLVoucherUS.GetGLSubsidiaryLedgerPrefix(Session("sAppServer"), CLng(nGralLedgerId))
	Set oGLVoucherUS = Nothing
End Function

Function GLAccountRole(nGralLegerId, sAccountNumber)
  Dim oGLVoucherUS
  '*************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	GLAccountRole = oGLVoucherUS.GLAccountRole(Session("sAppServer"), CLng(nGralLegerId), CStr(sAccountNumber))
	Set oGLVoucherUS = Nothing
End Function

Function GLGroupIsRoot(nGLGroupId)
  Dim oGLGroups
  '*******************************
	On Error Resume Next	
	Set oGLGroups = Server.CreateObject("AOGralLedger.CGralLedgerGroups")
	GLGroupIsRoot = oGLGroups.IsRoot(Session("sAppServer"), CLng(nGLGroupId))
	Set oGLGroups = Nothing
End Function

Function GLGroupStdAccountId(nGLGroupId)
  Dim oGLGroups
  '*************************************
	On Error Resume Next	
	Set oGLGroups = Server.CreateObject("AOGralLedger.CGralLedgerGroups")
	GLGroupStdAccountId = oGLGroups.GetGroup(Session("sAppServer"), CLng(nGLGroupId)).Fields("id_tipo_cuentas_std")
	Set oGLGroups = Nothing
End Function

Function IsDate_(sDate)
	IsDate_ = IsDate(sDate)
End Function

Function IsDateInGLPeriod(nGralLedgerId, dDate)
  Dim oGLVoucherUS
  '********************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	IsDateInGLPeriod = oGLVoucherUS.IsDateInGLPeriod(Session("sAppServer"), CLng(nGralLedgerId), CDate(dDate))
	Set oGLVoucherUS = Nothing
End Function

Function IsNumeric_(sNumber, nDecimals)
  Dim oGLVoucherUS
  '************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	IsNumeric_ = oGLVoucherUS.IsNumericOK(CStr(sNumber), CLng(nDecimals))
	Set oGLVoucherUS = Nothing
End Function

Function IsStdAccountNumberValid(nStdAccountTypeId, sStdAccountNumber)
  Dim oStdAccount
  '*******************************************************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("AOGralLedger.CStandardAccount")
	IsStdAccountNumberValid = oStdAccount.IsAccountNumberValid(Session("sAppServer"), CLng(nStdAccountTypeId), CStr(sStdAccountNumber))
	Set oStdAccount = Nothing
End Function

Function IsSubsAccountNumberValid(sSubsAccountNumber)
  Dim oGLVoucherUS
  '**************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	IsSubsAccountNumberValid = oGLVoucherUS.IsSubsAccountNumberValid(CStr(sSubsAccountNumber))
	Set oGLVoucherUS = Nothing
End Function

Function PendingPostingsReferences(nGralLedgerId, nSelPosting)
  Dim oGLVoucherUS
  '***********************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	PendingPostingsReferences = oGLVoucherUS.PendingPostingsReferences(Session("sAppServer"), CLng(nGralLedgerId), CLng(nSelPosting))
	Set oGLVoucherUS = Nothing
End Function

Function PostingsTable(nTransactionId, bAnalize)
	Dim oVoucher, sTemp
	'*********************************************
	On Error Resume Next	
	Set oVoucher = Server.CreateObject("AOGLVoucherUS.CVoucher")
	
	sTemp = oVoucher.GetPostings(Session("sAppServer"), CLng(nTransactionId), True, CBool(bAnalize))
	Set oVoucher = Nothing
	
	sTemp = "<TABLE class=applicationTable><TR class=applicationTableHeader>" & _
					"<TD nowrap width=120><b>Núm. de cuenta</b></TD><TD><b>Sec</b></TD><TD width=40%><b>Descripción</b></TD>" & _
					"<TD><b>Verif</b></TD><TD><b>Area</b></TD><TD align=center><b>Moneda</b></TD>" & _
					"<TD align=center nowrap><b>T. de cambio</b></TD>" & _
					"<TD colspan=3 align=center width=30%><b>Importes</b></TD></TR>" & VbCrlf & _
					"<TR class=applicationTableHeader><TD><b><i>Auxiliar</i></b></TD>" & _
					"<TD>&nbsp;</TD><TD><b><i>Concepto</i></b></TD><TD colspan=3>&nbsp;</TD>" & _
					"<TD align=center>&nbsp;</TD><TD align=center><b>Parcial</b></TD>" & _
					"<TD align=center><b>Debe</b></TD><TD align=center><b>Haber</b></TD></TR>" & _
					sTemp & _					
					"</TABLE>"
	PostingsTable = sTemp	
End Function

Function RuleChildsCount(nRuleId)
  Dim oRule
  '******************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  RuleChildsCount = oRule.RuleChildsCount(Session("sAppServer"), CLng(nRuleId))
  Set oRule = Nothing
End Function

Function RuleChildsType(nRuleId)
  Dim oRule
  '*****************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  RuleChildsType = oRule.RuleChildsType(Session("sAppServer"), CLng(nRuleId)) 
  Set oRule = Nothing
End Function

Function RuleItemsSection(nRuleDefId, nRuleId)
  Dim oRule 
  '*******************************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  RuleItemsSection = oRule.LeafItemsAsRows(Session("sAppServer"), CLng(nRuleDefId), CLng(nRuleId))
  Set oRule = Nothing
End Function

Function RuleLabel(nRuleId)
  Dim oRule 
  '************************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  RuleLabel = oRule.RuleLabel(Session("sAppServer"), CLng(nRuleId)) 
  Set oRule = Nothing
End Function

Function RuleType(nRuleId)
  Dim oRule
  '***********************
  Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
  RuleType = oRule.RuleType(Session("sAppServer"), CLng(nRuleId)) 
  Set oRule = Nothing
End Function

Function StandardAccountRole(nStdAccountId)
  Dim oGralLedger
  '****************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CStandardAccount")
	StandardAccountRole = oGralLedger.Role(Session("sAppServer"), CLng(nStdAccountId))
	Set oGralLedger = Nothing	
End Function

Function StdAccountId(nStdAccountTypeId, sStdAccountNumber)
  Dim oGralLedger
  '********************************************************
	On Error Resume Next	
	Set oGralLedger = Server.CreateObject("AOGralLedger.CStandardAccount")
	StdAccountId = oGralLedger.StdAccountId(Session("sAppServer"), _
																				  CLng(nStdAccountTypeId), CStr(sStdAccountNumber))
	Set oGralLedger = Nothing
End Function

Function SubsAccountExtendedAttributes(nSubsLedgerId, nSubsAccountId)
	Dim oGralLedger, nSubsidiaryLedgerType, oRecordset
	'******************************************************************
	Set oGralLedger = Server.CreateObject("AOGralLedgerUS.CServer")
	Set oRecordset  = oGralLedger.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(nSubsLedgerId))
	nSubsidiaryLedgerType = oRecordset("id_tipo_mayor_auxiliar")
	oRecordset.Close
	Set oRecordset = Nothing
	SubAccountExtendedAttributes = oGralLedger.SubsidiaryAccountExtendedAttrs(Session("sAppServer"), _
																																						CLng(nSubsidiaryLedgerType), CLng(nSubsAccountId))
	Set oGralLedger = Nothing
End Function

Function SubsidiaryAccountName(nGralLedgerId, sSubsAcctNumber)
  Dim oGralLedgerUS, sTemp	
	'***********************************************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
  SubsidiaryAccountName = oGralLedgerUS.SubsidiaryAccountName(Session("sAppServer"), CLng(nGralLedgerId), CStr(sSubsAcctNumber))
  Set oGralLedgerUS = Nothing
End Function

Function SubsidiaryAccountNumber(nSubsAccountId, bShowComplete)
  Dim oGralLedgerUS, sTemp	
	'************************************************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
  SubsidiaryAccountNumber = oGralLedgerUS.SubsidiaryAccountNumber(Session("sAppServer"), CLng(nSubsAccountId), CBool(bShowComplete))
  Set oGralLedgerUS = Nothing
End Function

Function TblSubsidiaryAccounts(nSubsLedgerId, nSubsAccountId, sOrderBy)
  Dim oGralLedgerUS, sHeader, sTemp	
	'********************************************************************
	On Error Resume Next
	sHeader = "<TABLE class=applicationTable height=100px>" & VbCrLf & _  
						"<TR class=applicationTableHeader>" & VbCrLf & _
					  "<TD nowrap><A href='' onclick=""return updateTable('numero_cuenta_auxiliar');"">Auxiliar</A></TD>" & VbCrLf & _
					  "<TD><A href='' onclick=""return updateTable('nombre_cuenta_auxiliar');"">Nombre</A></TD></TR>" & VbCrLf

	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
  sTemp = oGralLedgerUS.TblSubsidiaryAccounts(Session("sAppServer"), CLng(nSubsLedgerId), CLng(nSubsAccountId), CStr(sOrderBy))
  Set oGralLedgerUS = Nothing	
  If Len(sTemp) <> 0 Then
		sTemp = sHeader & sTemp & VbCrLf & "</TABLE>"
	Else
		sTemp = sHeader & "<TR>El mayor auxiliar no tiene cuentas auxiliares.</TR>" & VbCrLf & "</TABLE>"
	End If
	TblSubsidiaryAccounts = sTemp
End Function

Function TransactionStatus(nTransactionId)
	Dim oVoucher, sTemp
	'***************************************
	On Error Resume Next	
	Set oVoucher = Server.CreateObject("AOGLVoucherUS.CVoucher")
	
	sTemp = oVoucher.TransactionStatus(Session("sAppServer"), CLng(nTransactionId))
	Set oVoucher = Nothing	
	TransactionStatus = sTemp	
End Function 

Function TransactionType(nVoucherId)
  Dim oVoucher, sTemp	
	'*********************************
	On Error Resume Next
	Set oVoucher = Server.CreateObject("AOGLVoucher.CServer")
  TransactionType = oVoucher.GetTransactionType(Session("sAppServer"), CLng(nVoucherId))
  Set oVoucher = Nothing
End Function

Function ValidateAreas(sAreasList)
  Dim oStdAccount
  '*******************************
	On Error Resume Next	
	Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
	ValidateAreas = oStdAccount.ValidateAreas(CStr(Session("sAppServer")), CStr(sAreasList))
	Set oStdAccount = Nothing
End Function

Function ValidateTransaction(nTransactionId)
  Dim oVoucherBS
	'*****************************************
	On Error Resume Next		
	Set oVoucherBS = Server.CreateObject("AOGLVoucher.CServer")
	ValidateTransaction = oVoucherBS.ValidateTransaction(Session("sAppServer"), CLng(nTransactionId))
	Set oVoucherBS = Nothing
End Function

Function VoucherId(nGralLedgerId, sVoucherNumber)
  Dim oVoucher, sTemp	
	'**********************************************
	On Error Resume Next
	Set oVoucher = Server.CreateObject("AOGLVoucher.CServer")
  VoucherId = oVoucher.GetTransactionId(Session("sAppServer"), CLng(nGralLedgerId), CStr(sVoucherNumber))
  Set oVoucher = Nothing
End Function

</SCRIPT>