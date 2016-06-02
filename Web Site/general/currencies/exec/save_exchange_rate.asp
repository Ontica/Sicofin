<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage
 
	gsReturnPage = "../pages/exchange_rate_editor.asp"
	
	If CLng(Request.QueryString("id")) = 0 Then
		Call SaveItem(0)
	Else
		Call SaveItem(CLng(Request.QueryString("id")))
  End If
  
    
  Sub SaveItem(nItemId)
		Dim oCurrencies, oRecordset
		'**************************		
		'On Error Resume Next
		Set oCurrencies = Server.CreateObject("AOCurrencies.CManager")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oCurrencies.GetExchangeRateRS(Session("sAppServer"), CLng(nItemId))
		'oRecordset("from_currency_id")			= 1
		'oRecordset("exchange_rate_type_id") = Request.Form("cboExchangeRateTypes")
		'oRecordset("to_currency_id")				= Request.Form("cboCurrencies")		
		'oRecordset("from_date")							= Request.Form("txtDate")
		'oRecordset("to_date")								= Request.Form("txtDate")
		oRecordset("exchange_rate")					= Request.Form("txtExchangeRate")
		oCurrencies.SaveExchangeRate Session("sAppServer"), (oRecordset), CLng(nItemId)		
		oRecordset.Close
		Set oRecordset = Nothing
		Set oCurrencies = Nothing
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
<body onload='window.close();'>
</body>
</html>