<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	  
	If CLng(Request.Form("txtItemId")) = 0 Then
		Call SaveItem(0)
	Else
		Call SaveItem(CLng(Request.Form("txtItemId")))
  End If
      
  Sub SaveItem(nItemId)
		Dim oGralLedger, oRecordset
		'**************************
		'On Error Resume Next				
		Set oGralLedger = Server.CreateObject("AOGralLedger.CGralLedger")		
		Set oRecordset = Server.CreateObject("ADODB.Recordset")
		Set oRecordset = oGralLedger.GetGeneralLedgerRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("numero_mayor")        = Request.Form("txtGralLedgerNumber")		
		oRecordset("nombre_mayor")        = Request.Form("txtGralLedgerName")
		oRecordset("prefijo_cuentas_auxiliares") = Request.Form("txtSubsAccountsPrefix")
		oRecordset("id_moneda_base") = Request.Form("cboCurrencies")					
		oRecordset("id_calendario")	 = Request.Form("cboCalendars")
		'oRecordset("fecha_apertura") = CDate("29/12/2000")
		oGralLedger.Save Session("sAppServer"), (oRecordset), _
										 CLng(Request.Form("cboVouchersGroups")), _
										 CLng(Request.Form("cboReportGroups")), _
										 CLng(nItemId)
		
		oRecordset.Close
		Set oRecordset = Nothing
		Set oGralLedger = Nothing
		If (Err.number = 0) Then
			Set Session("oError") = Nothing
		Else
			Set Session("oError") = Err
		End If
  End Sub
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
</head>
<body onload="window.opener.location.href='../general_ledgers.asp?id=<%=Request.Form("cboReportGroups")%>';window.opener.location.href;window.close();">
</body>
</html>