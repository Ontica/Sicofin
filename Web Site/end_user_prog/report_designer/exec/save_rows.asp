<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	Dim gsNewItemPage, gsEditItemPage
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
 
  gsNewItemPage = "../column_editor.asp"
	If CLng(Request.Form("txtRowId")) = 0 Then
		Call SaveItem(0)
	Else
		gsEditItemPage = "../row_editor.asp?id=" & Request.Form("txtRowId")
		Call SaveItem(Request.Form("txtRowId"))
  End If
   
  Sub SaveItem(nItemId)
		Dim oReportDesigner
		'******************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		If (nItemId <> 0) Then
			oReportDesigner.SaveRow Session("sAppServer"), CLng(Request.Form("txtReportId")), CLng(nItemId)
		Else
			oReportDesigner.AddRows Session("sAppServer"), CLng(Request.Form("txtReportId")), _
															CLng(Request.Form("txtFromPosition")), CLng(Request.Form("txtToPosition"))
		End If
		Set oReportDesigner = Nothing
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
<body onload='window.opener.location.href=window.opener.location.href; window.close();'>
</body>
</html>