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
		Dim oGralLedger
		'********************
		On Error Resume Next
		Set oGralLedger = Server.CreateObject("AOGralLedger.CGralLedger")
		oGralLedger.Delete Session("sAppServer"), CLng(nItemId)
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
<body onload='window.opener.location.href=window.opener.location.href;window.close();'>
</body>
</html>