<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
 
	Call SaveItem()
   
  Sub SaveItem()
		Dim oReportDesigner, oRecordset, nItemId
		'***************************************
		'On Error Resume Next
		nItemId = Request.Form("txtItemId")
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oReportDesigner.GetItemRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("itemName")     = Request.Form("txtName")
		oRecordset("itemLabel")    = Request.Form("txtName")		
		oRecordset("itemFilterId") = CLng(Request.Form("cboFilters"))
		oRecordset("itemPrintConditionId") = CLng(Request.Form("cboPrintConditions"))
		oRecordset("itemPrintLayout") = Request.Form("txtPrintLayout")
		oReportDesigner.SaveItem Session("sAppServer"), (oRecordset), CLng(nItemId)
		oRecordset.Close
		Set oRecordset = Nothing
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
<body onload='window.close();'>
</body>
</html>