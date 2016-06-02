<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
		
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If

	Call CreateSection()
	   
  Sub CreateSection()
		Dim oReportDesigner, oRecordset
		'******************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oReportDesigner.GetSectionRS(Session("sAppServer"), 0)
		oRecordset("reportId")                = Request.Form("txtReportId")		
		oRecordset("sectionWorkSheet")        = Request.Form("txtWorkSheet") 
		oRecordset("sectionName")             = Request.Form("txtName")
		oRecordset("sectionShape")            = "H"
		If Request.Form("cboParametrizationModes") = "C" Then
			oRecordset("sectionType")       = "DE"
			oRecordset("sectionName")       = "Detalle"
			oRecordset("sectionPosition")   = 1
			oRecordset("sectionInitialRow")	= 1
			oRecordset("sectionRows")				= Request.Form("txtSectionRows") 
		ElseIf Request.Form("cboParametrizationModes") = "R" Then
			oRecordset("sectionType")       = "FX"
			oRecordset("sectionPosition")   = 1
			oRecordset("sectionInitialRow")	= Request.Form("txtInitialRow") 
			oRecordset("sectionRows")				= CLng(CLng(Request.Form("txtFinalRow")) - CLng(Request.Form("txtInitialRow")) + 1)
		ElseIf Request.Form("cboSectionTypes") = "FX" Then
			oRecordset("sectionType")       = "FX"
			oRecordset("sectionInitialRow")	= Request.Form("txtInitialRow")
			oRecordset("sectionRows")				= CLng(CLng(Request.Form("txtFinalRow")) - CLng(Request.Form("txtInitialRow")) + 1)
		ElseIf Len(Request.Form("cboInsertPosition")) <> 0 Then
			oRecordset("sectionPosition")   = CLng(Request.Form("cboInsertPosition"))
			oRecordset("sectionType")       = Request.Form("cboSectionTypes")
			oRecordset("sectionInitialRow")	= 1
			oRecordset("sectionRows")				= Request.Form("txtSectionRows")
		Else			
			oRecordset("sectionType")       = Request.Form("cboSectionTypes")
			oRecordset("sectionInitialRow")	= 1
			oRecordset("sectionRows")				= Request.Form("txtSectionRows")
		End If
		oRecordset("sectionCols")						= 0		
		oRecordset("sectionInitialColumn")  = 0
		oRecordset("sectionFinalColumn")    = 0
		If Len(Request.Form("txtDataGrouping")) <> 0 Then
			oRecordset("sectionDataGrouping")   = Request.Form("txtDataGrouping") & "|" & Request.Form("txtDataGroupingExp")
		End If
		If Len(Request.Form("txtDataOrder")) <> 0 Then
			oRecordset("sectionDataOrder")      = Request.Form("txtDataOrder") & "|" & Request.Form("txtDataOrderExp") 
		End If
		If Len(Request.Form("txtDataFilter")) <> 0 Then
			oRecordset("sectionDataFilter")     = Request.Form("txtDataFilter") & "|" & Request.Form("txtDataFilterExp")
		End If
		oRecordset("sectionPrintCondition") = 0
		oRecordset("sectionPrintLayout")    = ""
		oReportDesigner.SaveSection Session("sAppServer"), (oRecordset), 0
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