<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Call DeleteItem(CLng(Request.QueryString("id")))    
	
  Sub DeleteItem(nItemId)
		Dim oCurrencies
		'********************
		On Error Resume Next
		Set oCurrencies = Server.CreateObject("AOCurrencies.CManager")
		oCurrencies.DeleteExchangeRate Session("sAppServer"), CLng(nItemId)
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