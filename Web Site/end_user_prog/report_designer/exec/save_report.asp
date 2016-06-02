<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	Dim gsNewItemPage, gsEditItemPage
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
 
  gsNewItemPage = "../report_designer.asp"
	If CLng(Request.Form("txtReportId")) = 0 Then
		Call SaveItem(0)
	Else
		gsEditItemPage = "../report_designer.asp?id=" & Request.Form("txtReportId")
		Call SaveItem(Request.Form("txtReportId"))
  End If
   
  Sub SaveItem(nItemId)
		Dim oReportDesigner, oRecordset
		'************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oReportDesigner.GetReportRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("reportName")            = Request.Form("txtName")
		oRecordset("reportDescription")     = Request.Form("txtDescription")
		oRecordset("reportKeywords")        = Request.Form("txtKeywords")
		oRecordset("reportCategoryId")      = Request.Form("cboReportCategories")
		oRecordset("reportDataClassId")     = Request.Form("cboClasses")		
		oRecordset("reportDataSubclassId")  = Request.Form("cboSubClasses")
		oRecordset("reportDataOrderId")     = Request.Form("cboDataOrders")		
		oRecordset("reportTechnology")		  = Request.Form("cboReportTechnologies")		
		oRecordset("reportTemplateFile")    = Request.Form("txtTemplateFile")
		oRecordset("reportHelpFile")        = Request.Form("txtHelpFile")
		oRecordset("reportIconFile")        = Request.Form("txtIconFile")
		oRecordset("reportPrintLayout")		  = ""		
		oRecordset("authorId")              = CLng(Session("uid"))
		oRecordset("lastUpdate")            = Date()
		If (nItemId = 0) Then
			oRecordset("reportStatus")          = "S"
		End If
		oRecordset("historicReportId")      = 0		
		oRecordset("FromDate")              = CDate("31/12/2000")
		oRecordset("ToDate")                = CDate("31/12/2049")
		oReportDesigner.SaveReport Session("sAppServer"), (oRecordset), CLng(nItemId)
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