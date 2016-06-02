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
				
	Call SaveItems()  
    
	'Response.Write "Cuenta: " & Request.Form("txtExchangeRate").Count & "<br><br><br>"
	'Response.Write "Llave: " & Request.Form("txtExchangeRatesArray")
	
  Sub SaveItems()
		Dim aCurrencies, oCurrencies
		'***************************
		'On Error Resume Next
		Set oCurrencies = Server.CreateObject("AOCurrencies.CManager")
		
		oCurrencies.SaveExchangeRates Session("sAppServer"), 1, CLng(Request.Form("cboExchangeRateTypes")), _
																  CDate(Request.Form("txtDate")), CStr(Request.Form("txtExchangeRatesArray"))
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