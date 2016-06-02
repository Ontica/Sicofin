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
		Dim oReportDesigner, oParentRowRS, oRecordset, nItemId, nParentRowId
		'***************************************
		'On Error Resume Next
		nItemId      = Request.Form("txtItemId")
		nParentRowId = Request.Form("cboSectionRows")
		
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")		
						
		Set oRecordset = oReportDesigner.GetItemRS(Session("sAppServer"), CLng(nItemId))
		If (CLng(nItemId) = 0) Then
			Set oParentRowRS = oReportDesigner.GetItem(Session("sAppServer"), CLng(nParentRowId))
			oRecordset("reportId")      = oParentRowRS("reportId")
			oRecordset("itemSectionId") = oParentRowRS("itemSectionId")
			oRecordset("itemWorkSheet") = oParentRowRS("itemWorkSheet")
			oRecordset("itemRow")       = oParentRowRS("itemRow")
			oParentRowRS.Close
			Set oParentRowRS = Nothing
		Else
			Set oParentRowRS = oReportDesigner.GetItem(Session("sAppServer"), CLng(nParentRowId))
			oRecordset("itemRow") = oParentRowRS("itemRow")
			oParentRowRS.Close
			Set oParentRowRS = Nothing
		End If
		oRecordset("itemType") = Request.Form("cboItemTypes")
		oRecordset("itemName") = Request.Form("txtName")
    If (Len(Request.Form("txtPosition")) <> 0) Then
			oRecordset("itemColumn") = CLng(Request.Form("txtPosition"))
			oRecordset("itemLength") = CLng(Request.Form("txtLength"))
		Else
			oRecordset("itemColumn") = CLng(Request.Form("cboExcelColumns"))
			oRecordset("itemLength") = 0
		End If		
		If (oRecordset("itemType") = "L") Then
			oRecordset("itemValue")       = Request.Form("txtLabel")
			oRecordset("itemFilter")      = ""
			oRecordset("ItemFilterDesc")  = ""
			oRecordset("itemOperationId") = 0
		ElseIf (oRecordset("itemType")  = "P") Then			
			oRecordset("itemValue")       = Request.Form("cboParameters")
			oRecordset("itemFilter")      = ""
			oRecordset("ItemFilterDesc")  = ""
			oRecordset("itemOperationId") = 0
		ElseIf (oRecordset("itemType")  = "F") Then			
			oRecordset("itemValue")       = Request.Form("cboDataItems")			
			oRecordset("itemFilter")      = Request.Form("txtFilterExp")
			oRecordset("ItemFilterDesc")  = Request.Form("txtFilterViewer")
			oRecordset("itemOperationId") = CLng(Request.Form("cboOperations"))
		ElseIf (oRecordset("itemType")  = "E") Then			
			oRecordset("itemValue")       = 0			
			oRecordset("itemFilter")      = ""
			oRecordset("ItemFilterDesc")  = ""
			oRecordset("itemOperationId") = 0
		End If
		oRecordset("itemTag")            = Request.Form("txtItemTag")
		oRecordset("itemPrintCondition") = Request.Form("txtPrintCondition")
		oRecordset("itemPrintLayout")    = Request.Form("txtPrintLayout")
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
<SCRIPT LANGUAGE=javascript>
<!--

function window_onload() {	
  if (window.opener != null || !window.opener.closed) {
		window.opener.location.href = window.opener.location.href;
	}
	window.close();
}

//-->
</SCRIPT>
<meta http-equiv="Pragma" content="no-cache">
</head>
<body onload='window_onload();'>
</body>
</html>