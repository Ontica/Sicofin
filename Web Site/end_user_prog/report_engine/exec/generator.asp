<%
  Option Explicit     
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1		
		
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	'On Error Resume Next
	
	Dim gnReportId, gsFileName, nScriptTimeout, gsWorkerTable
	
	Response.Buffer = False
	
	gnReportId = CLng(Request.Form("txtReportId"))	

	If (gnReportId <> 0) Then
		Call Main()
	End If
	
	Sub Main()
		Dim oReportsEngine, sAdditionalEntries
		'*************************************
		On Error Resume Next
		Set oReportsEngine = Server.CreateObject("EUPReportBuilder.CBuilder")
		sAdditionalEntries = GetAdditionalEntries()
		gsWorkerTable = oReportsEngine.WorkerTable(Session("sAppServer"), CLng(gnReportId), CStr(sAdditionalEntries))
		Set oReportsEngine = Nothing
		If (Err.number <> 0) Then					
			Session("errNumber") = Err.number
			Session("errDesc")   = Err.description
			Session("errSource") = Err.source
			Err.Clear
		End If		
	End Sub
	
	Sub GenerateReport()
		Dim oReportsEngine, oReport, oData, vGralLedgers
		Dim nDataSourceId, nDataFilterId, nRuleId, sFixedPars, bGenerateInCascade, sFileName 
		'***********************************************************************************
		On Error Resume Next
		Set oReportsEngine = Server.CreateObject("EUPReportBuilder.CBuilder")
		Set oReport = oReportsEngine.Report(Session("sAppServer"), CLng(gnReportId))
		nDataSourceId = CLng(oReport("reportDataClassId"))
		nDataFilterId = CLng(oReport("reportDataSubClassId"))
		oReport.Close
		Set oReport = Nothing
		Select Case nDataSourceId
			Case 167:
				Set oData = GetBalancesData()
			Case 186, 236, 237, 238, 239, 240:
				If (nDataFilterId = 205) Or (nDataFilterId = 256) Then
				  vGralLedgers = Array(BancariaA(), BancariaA(), FideosUDISA(), FideosUDISA())
				  Set oData = GetGLRulesEngineData(Array(6, 7, 8, 9), vGralLedgers)									
				Else
				  nRuleId  = oReportsEngine.GetDataSourceRuleId(Session("sAppServer"), CLng(nDataFilterId))
				  vGralLedgers = GeneralLedgersArray()
				  Set oData = GetGLRulesEngineData(nRuleId, vGralLedgers)					
				End If
			Case Else:				
				Set oData = GetDictionaryData(nDataSourceId, nDataFilterId)
		End Select
		sFixedPars = GetFixedPars()
		bGenerateInCascade = False
		If Not bGenerateInCascade Then
			sFileName = oReportsEngine.BuildReport(Session("sAppServer"), CLng(gnReportId), (oData), CStr(sFixedPars))			
		Else			
			sFileName = oReportsEngine.BuildReport(Session("sAppServer"), CLng(gnReportId), (oData), CStr(sFixedPars), sFileName, True)
		End If
		gsFileName = oReportsEngine.URLFilesPath & sFileName
	End Sub
		

  Function GetFixedPars()
		Dim oParameters
		'********************
		Set oParameters = Server.CreateObject("EFAParameters.CParameters")
		GetFixedPars = oParameters.GeneratedReportsPars(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), _
																										CLng(Request.Form("cboGralLedgers")), _ 
																										Request.Form("txtInitialDate1"), Request.Form("txtFinalDate1"))
		Set oParameters = Nothing		
  End Function
											
  Function GetBalancesData()
  	Dim oBalances, oDataRS
		'***********************
		On Error Resume Next		
		Set oBalances = Server.CreateObject("EOGLBalances.CBalances")
		Set oDataRS = Server.CreateObject("ADODB.Recordset")
		
		'Set oDataRS = oBalances.BalanceGeneral
  	Set GetBalancesData = oDataRS
  End Function 
   
  Function GetDictionaryData(nDataClassId, nDataSubClassId)
  	Dim oDictionary, oDataRS
		'*****************************************************
		On Error Resume Next
		Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")		
		Set oDataRS = Server.CreateObject("ADODB.Recordset")
		If nDataSubClassId <> 0 Then
			Set oDataRS = oDictionary.ItemData(Session("sAppServer"), CLng(nDataSubClassId))
		Else
			Set oDataRS = oDictionary.ItemData(Session("sAppServer"), CLng(nDataClassId))
		End If		
  	Set GetDictionaryData = oDataRS  
  End Function
  
  Function GetGLRulesEngineData(vaRules, vaGralLedgers)
		Dim oGLRulesEngine, oDataRS
		Dim vnBalancesTypes, vsBalancesPeriods, vsBalancesCurrencies, nTransactionType, nVoucherType
		Dim vnExchangeRateTypes, vnExchangeRateCurrencies, vdExchangeRateDates, vbRoundBalancesTo		
		Dim sInterpretationPars, dRuleDate
		'*******************************************************************************************
		On Error Resume Next
		Set oGLRulesEngine = Server.CreateObject("EFARulesEngine.CServer")
		Set oDataRS = Server.CreateObject("ADODB.Recordset")
	
		vnBalancesTypes          = Array("G")
		vsBalancesPeriods        = Array(CDate(Request.Form("txtInitialDate1")), CDate(Request.Form("txtFinalDate1")))
		vsBalancesCurrencies     = Array(0)
		vnExchangeRateTypes      = Array(CLng(Request.Form("cboExchangeRateTypes")))
		vnExchangeRateCurrencies = Array(CLng(Request.Form("cboExchangeRateCurrencies")))
		vdExchangeRateDates      = Array(CDate(Request.Form("txtFinalDate1")))
		vbRoundBalancesTo        = CLng(Request.Form("cboRoundBalancesTo"))
		
		If Len(Request.Form("chkTransactionTypes")) <> 0 Then
			nTransactionType = -1 * Request.Form("cboTransactionTypes")
		Else
			nTransactionType = Request.Form("cboTransactionTypes")
		End If
		
		If Len(Request.Form("chkVoucherTypes")) <> 0 Then
			nVoucherType = -1 * CLng(Request.Form("cboVoucherTypes"))
		Else
			nVoucherType = CLng(Request.Form("cboVoucherTypes"))
		End If
				
		sInterpretationPars = "groupingLevel=0;showZeros=false;collapseGralLedgers=false;"
		dRuleDate			     = Date()
		Set oDataRS = oGLRulesEngine.InterpretRules(Session("sAppServer"), vaRules, vaGralLedgers, _
													vnBalancesTypes, vsBalancesPeriods, CLng(nTransactionType), CLng(nVoucherType), _
													vsBalancesCurrencies, vnExchangeRateTypes, vnExchangeRateCurrencies, vdExchangeRateDates, _
													vbRoundBalancesTo, CStr(sInterpretationPars), CDate(dRuleDate))
		Set GetGLRulesEngineData = oDataRS
		If (Err.number <> 0) Then					
			Session("errNumber") = Err.number
			Session("errDesc")   = Err.description
			Session("errSource") = Err.source
			Err.Clear
		End If		
	End Function

	Function GetAdditionalEntries() 
		Dim sTemp
		'****************************
		If Len(Request.Form("txtInitialDate1")) <> 0 Then
			sTemp = "Período:|" & Request.Form("txtInitialDate1") & " al " & Request.Form("txtFinalDate1")
		End If
		GetAdditionalEntries = sTemp
	End Function
		
	Function GeneralLedgersArray()
		Dim oVoucherUS, vGralLedgers, sTemp
		'**********************************
		On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		If (Len(Request.Form("cboGralLedgers")) <> 0 ) Then
			If (Len(Request.Form("txtFromGL")) = 0) Then
				If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
					sTemp = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), ",")
					vGralLedgers = Split(sTemp, ",")
				Else
					vGralLedgers = Array(CLng(Request.Form("cboGralLedgers")))
				End If
			Else
				vGralLedgers = oVoucherUS.GetGLRangeArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), _
														  CLng(Request.Form("txtFromGL")), CLng(Request.Form("txtToGL")))
			End If
		End If
		GeneralLedgersArray = vGralLedgers
		Set oVoucherUS = Nothing		
	End Function	
	
	Function BancariaA()
		BancariaA = Array(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32)
	End Function 
	
	Function FideosUDISA()
		FideosUDISA = Array(264,265,266,267,268,269,270)
	End Function
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<TITLE>Trabajando...</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function gotoDownload(sPage) {
	window.location.href = '/empiria/general/worker_wizard/download.asp?type=1&file=' + sPage;
}

function gotoError() {
	window.location.href = '/empiria/central/exceptions/exception.asp';
}

function window_onload() {
	document.all.clicker.click();
}

//-->
</SCRIPT>
</head>
<body rightmargin=3 leftMargin=3 topmargin=3 bottommargin=3 onload="return window_onload()" style='cursor:wait;background-color:white;'>
<div id=divProcessing>
<table width=100% border=0>	
	<tr>
		<td valign=top>
			<img src="/empiria/images/central/working.gif" style="cursor:wait;">
			<table class=applicationTable>
				<tr>
					<td>
					<INPUT type="checkbox" name=chkSendTo style='cursor:hand;'>
					Al finalizar, cerrar esta ventana y enviar el reporte a mi bandeja de documentos.
					<br>
					<INPUT class=cmdSubmit type="button" value="Cancela el trabajo" name=cmdCancel onclick='window.close();'>
					</td>
				</tr>
			</table>
		</td>
		<td width=100% nowrap valign=top>
			<table class=applicationTable height=100%>
				<tr><td width=100% colspan=2><b>Generando el reporte ...</b></td></tr>
				<%=gsWorkerTable%>
			</table>
		</td>
	</tr>
</table>
</div>
<%
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	If CLng(gnReportId) <> 0 Then
		Call GenerateReport()
	End If
	Server.ScriptTimeout = nScriptTimeout
	If Err.number = 0 Then
		Response.Write ("<A id=clicker onclick='gotoDownload(""" & gsFileName & """)'></A>")
	Else				
		Set Session("oError") = Err
		Response.Write ("<A id=clicker onclick='gotoError()'></A>")
	End If	
%>
</body>
</html>