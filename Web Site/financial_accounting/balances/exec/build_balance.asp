<%
  Option Explicit     
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim oReports, sFileName, nScriptTimeout
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
					
	Sub Main()
		Dim oVoucherUS, vGralLedgers, sTemp, dExcRateDate
		Dim nTransactionTypePar, nVoucherTypes, nCurrencyPar, bPrintInCascade, bCascadeDates, bConsolidateExchangeRateCurrency
		
		'On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")		
		Set oReports = Server.CreateObject("EFABalanceReporter.CReporter")
				
		If (Len(Request.Form("cboGralLedgers")) <> 0 ) Then		
			If (Len(Request.Form("txtFromGL")) = 0) Then
				If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
					sTemp = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), ",")
					vGralLedgers = Split(sTemp, ",")
				Else
					vGralLedgers = CLng(Request.Form("cboGralLedgers"))
				End If
			Else
				vGralLedgers = oVoucherUS.GetGLRangeArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), CLng(Request.Form("txtFromGL")), CLng(Request.Form("txtToGL")))
			End If
		End If
			
		If Len(Request.Form("txtExchangeRateDate")) <> 0 Then
			dExcRateDate = Request.Form("txtExchangeRateDate")
		Else
			dExcRateDate = Date()
		End If
		
		If Len(Request.Form("chkTransactionTypes")) <> 0 Then
			nTransactionTypePar = -1 * Request.Form("cboTransactionTypes")
		Else
			nTransactionTypePar = Request.Form("cboTransactionTypes")
		End If
		
		If Len(Request.Form("chkVoucherTypes")) <> 0 Then
			nVoucherTypes = -1 * CLng(Request.Form("cboVoucherTypes"))
		Else
			nVoucherTypes = CLng(Request.Form("cboVoucherTypes"))
		End If
		
		If CLng(Request.Form("cboExchangeRateTypes")) <> 0 Then
			If Len(Request.Form("chkCurrencies")) <> 0 Then
				nCurrencyPar = -1 * CLng(Request.Form("cboCurrencies"))
			Else
				nCurrencyPar = CLng(Request.Form("cboCurrencies"))
			End If
		Else					
			nCurrencyPar = 0
		End If
		
		If Len(Request.Form("chkPrintInCascade")) <> 0 Then
			bPrintInCascade = True
		Else
			bPrintInCascade = False
		End If
										
		If Len(Request.Form("chkCascadeDates")) <> 0 Then
			bCascadeDates = True
		Else
			bCascadeDates = False
		End If
		If Len(Request.Form("chkConsolidateExchangeRateCurrency")) <> 0 Then
			bConsolidateExchangeRateCurrency = True
		Else
			bConsolidateExchangeRateCurrency = False
		End If
		
		Select Case CLng((Request.Form("cboBalanceFormat")))
			Case 1
				sFileName = Report_16(Request.Form("cboGLGroups"), vGralLedgers, Request.Form("cboPatterns"), _
														  Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														  Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
														  nTransactionTypePar, nVoucherTypes, _
														  nCurrencyPar, Request.Form("cboExchangeRateTypes"), _
														  Request.Form("cboExchangeRateCurrencies"), _
														  dExcRateDate, bPrintInCascade, bCascadeDates, Request.Form("cboBalanceType"), False, _
														  bConsolidateExchangeRateCurrency)
			Case 2
				sFileName = Report_16_Comp(Request.Form("cboGLGroups"), vGralLedgers, Request.Form("cboPatterns"), _
														       Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														       Request.Form("txtInitialDate2"), Request.Form("txtFinalDate2"), _
														       Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
														       nTransactionTypePar, nVoucherTypes, _
														       nCurrencyPar, Request.Form("cboExchangeRateTypes"), _
														       Request.Form("cboExchangeRateCurrencies"), _
														       dExcRateDate, bPrintInCascade, Request.Form("cboBalanceType"), False)
			Case 3
				sFileName = Report16_Subs(vGralLedgers, Request.Form("cboPatterns"), _
																	Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
																	Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
																	nTransactionTypePar, nVoucherTypes, _
																	nCurrencyPar, Request.Form("cboExchangeRateTypes"), _
																	Request.Form("cboExchangeRateCurrencies"), _
																	dExcRateDate, bPrintInCascade, Request.Form("cboBalanceType"), _
																	bConsolidateExchangeRateCurrency)
			Case 4
				sFileName = Report_16(Request.Form("cboGLGroups"), vGralLedgers, Request.Form("cboPatterns"), _
														  Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														  Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
														  nTransactionTypePar, nVoucherTypes, _
														  nCurrencyPar, Request.Form("cboExchangeRateTypes"), _
														  Request.Form("cboExchangeRateCurrencies"), _
														  dExcRateDate, bPrintInCascade, bCascadeDates, Request.Form("cboBalanceType"), True, _
														  bConsolidateExchangeRateCurrency)
			Case 5
				sFileName = Report_19(vGralLedgers, Request.Form("cboPatterns"), _
														  Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														  Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
														  nTransactionTypePar, nVoucherTypes, Request.Form("cboExchangeRateCurrencies"), _
														  Request.Form("cboBalanceType"))
		End Select
		sFileName = oReports.URLFilesPath & sFileName
		'If Err.number <> 0 Then
		'   Response.Write Err.Number & " " & Err.description & " " & Err.source
		'End If	
		Set oReports = Nothing
	End Sub  

	Function Report_16(nGralLedgerGroup, aGL, sPattern, dFromDate, dToDate, _
										 sFromAccount, sToAccount, nTransactionTypeId, nVoucherTypeId, _
										 nCurrencyType, nExchangeRateType, nExcRateCurrency, dExcRateDate, _
										 bPrintInCascade, bCascadeDates, nBalanceType, bShowAverageColumn, _
										 bConsolidateExchangeRateCurrency)

		Report_16 = oReports.Balances(Session("sAppServer"), CLng(nGralLedgerGroup), aGL, _
																	CStr(sPattern), CDate(dFromDate), CDate(dToDate), _
																	CStr(sFromAccount), CStr(sToAccount), _
																	CLng(nTransactionTypeId), CLng(nVoucherTypeId), _
																	CLng(nCurrencyType), CLng(nExchangeRateType), CLng(nExcRateCurrency), _
																	CDate(dExcRateDate), CBool(bPrintInCascade), CBool(bCascadeDates), CLng(nBalanceType), _
																	False, CBool(bShowAverageColumn), CBool(bConsolidateExchangeRateCurrency))
	End Function
	

	Function Report_16_Comp(nGralLedgerGroup, aGL, sPattern, dFromDate, dToDate, dFromDate2, dToDate2, _
													sFromAccount, sToAccount, nTransactionTypeId, nVoucherTypeId, _
													nCurrencyType, nExchangeRateType, nExcRateCurrency, dExcRateDate, bPrintInCascade, nBalanceType, bShowAverageColumn)

		Report_16_Comp = oReports.BalancesComparative(Session("sAppServer"), CLng(nGralLedgerGroup), aGL, _
																						      CStr(sPattern), CDate(dFromDate), CDate(dToDate), _
																									CDate(dFromDate2), CDate(dToDate2), _
																									CStr(sFromAccount), CStr(sToAccount), _
																									CLng(nTransactionTypeId), CLng(nVoucherTypeId), _
																									CLng(nCurrencyType), CLng(nExchangeRateType), CLng(nExcRateCurrency), _
																									CDate(dExcRateDate), CBool(bPrintInCascade), CLng(nBalanceType), False, True, True)
	End Function
                                                 
	Function Report16_Subs(aGL, sPattern, dFromDate, dToDate, sFromAccount, sToAccount, _
												 nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExchangeRateType, _
												 nExcRateCurrency, dExcRateDate, bPrintInCascade, nBalanceType, _
												 bConsolidateExchangeRateCurrency)
												 
		Report16_Subs = oReports.Report111_112(Session("sAppServer"), aGL, CStr(sPattern), _
																					 CDate(dFromDate), CDate(dToDate), _
																					 CBool(bPrintInCascade), CStr(sFromAccount), CStr(sToAccount), _
																					 CLng(nTransactionTypeId), CLng(nVoucherTypeId), _
																					 CLng(nCurrencyType), CLng(nExchangeRateType), CLng(nExcRateCurrency), _
																					 CDate(dExcRateDate),  CLng(nBalanceType), , , _
																					 CBool(bConsolidateExchangeRateCurrency))
	End Function
	
	Function Report_19(aGL, sPattern, dFromDate, dToDate, sFromAccount, sToAccount, _
										 nTransactionTypeId, nVoucherTypeId, nCurrencyType, nBalanceType)
		Report_19 = oReports.Report19(Session("sAppServer"), CStr(sPattern), aGL, _
															    CDate(dFromDate), CDate(dToDate), _
															    CStr(sFromAccount), CStr(sToAccount), _
															    CLng(nTransactionTypeId), CLng(nVoucherTypeId), _
															    CLng(nCurrencyType), CLng(nBalanceType), False)
	End Function
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function showRightButtonMsg() {
  var sMsg;
  
  sMsg = "Para obtener una copia de la balanza en su equipo, se requiere hacer\n" +
         "clic con el botón derecho del ratón y seleccionar la opción\n" + 
         "'Guardar destino como...'\n\n" + 
         "Gracias."
	alert(sMsg);	
}

function showReportInBrowser() {	
	window.open('<%=sFileName%>', 'dummy', "menubar=yes,toolbar=yes,scrollbars=yes,status=yes,location=no");
	return true;
}

//-->
</SCRIPT>
</head>
<body>
<table bgColor="khaki" width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td><font size=2><b>La balanza solicitada está lista.</b></font></td>	
</tr>
<tr>
	<td><font size=2><b>¿Qué desea hacer?</b></font></td>	
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td>
		<a href="<%=sFileName%>" onclick="showRightButtonMsg();return false;">
			<img src="/empiria/images/download.jpg" border=0>
		</a>
	</td>	
	<td valign=middle>
		<a href="<%=sFileName%>" onclick="showRightButtonMsg();return false;">
			Si se desea obtener una copia de la balanza en su equipo, haga clic sobre esta liga 
			con el botón derecho del ratón y seleccione la opción 'Guardar destino como...'
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="" onclick="showReportInBrowser();return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="" onclick="showReportInBrowser();return false;">	
			Ver la balanza generada en una página nueva del navegador.
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<a href="" onclick="window.history.back();">
			Regresar al constructor de balanzas.
		</a>
		<br>
	</td>	
</tr>
</table>
</body>
</html>