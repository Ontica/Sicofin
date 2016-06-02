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
	If CLng(Request.Form("txtColumnId")) = 0 Then
		Call SaveItem(0)
	Else
		gsEditItemPage = "../column_editor.asp?id=" & Request.Form("txtColumnId")
		Call SaveItem(Request.Form("txtColumnId"))
  End If
   
  Sub SaveItem(nItemId)
		Dim oReportDesigner, oRecordset
		'************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oReportDesigner.GetColumnRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("columnId")            = nItemId
		oRecordset("reportId")            = Request.Form("txtReportId")
		oRecordset("columnName")          = Request.Form("txtName")		
		If Len(Request.Form("txtPosition")) <> 0 Then
			oRecordset("columnWorksheet")     = ""
			oRecordset("columnPosition")      = Request.Form("txtPosition")
			oRecordset("columnlength")        = Request.Form("txtLength")
		Else
			oRecordset("columnWorksheet")     = Request.Form("cboWorkSheet")	
			oRecordset("columnPosition")      = Request.Form("cboExcelColumns")
			oRecordset("columnlength")        = 0
		End If		
		If Len(Request.Form("chkPivotColumn")) <> 0 Then
			oRecordset("isPivotColumn")       = 1
		Else
			oRecordset("isPivotColumn")       = 0
		End If
		oRecordset("columnLayout")      = ""
		oRecordset("columnDataId")      = Request.Form("cboDataItems")
		oRecordset("columnFilterId")    = Request.Form("cboFilters")
		oRecordset("columnOrderId")     = 0
		oRecordset("columnOperationId") = Request.Form("cboOperations")		
		oReportDesigner.SaveColumn Session("sAppServer"), (oRecordset), CLng(nItemId)
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
<body onload='window.opener.location.href=window.opener.location.href; window.close();'>
</body>
</html>