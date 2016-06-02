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
	If CLng(Request.Form("txtFilterId")) = 0 Then
		Call SaveItem(0)
	Else
		gsEditItemPage = "../column_editor.asp?id=" & Request.Form("txtColumnId")
		Call SaveItem(Request.Form("txtFilterId"))
  End If
   
  Sub SaveItem(nItemId)
		Dim oDictionary, oRecordset
		'************************
		'On Error Resume Next
		Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oDictionary.GetItemRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("itemType") = "F"
		oRecordset("itemName") = Request.Form("txtName")
		oRecordset("itemAlias") = ""
		oRecordset("itemIsClassId") = 0
		oRecordset("itemLinkedClassId") = 0
		oRecordset("itemLinkedClassAttrId") = 0
		oRecordset("itemIsHidden") = 0						
		oRecordset("itemDataType") = "S"
		oRecordset("itemDataTypeLength") = 0
		oRecordset("itemDataTypePrecision") = 0
		oRecordset("itemDataTypeFormat") = ""
		oRecordset("itemStringValue") = Request.Form("txtFilter")		
		oRecordset("itemNumericValue") = 0
		oRecordset("itemDateValue") = null
		oRecordset("itemPosition") = 9999
		oRecordset("itemParentId") = Request.Form("txtClassId")
		oRecordset("itemAuthorId") = Session("uid")				
		oDictionary.SaveItem Session("sAppServer"), (oRecordset), CLng(nItemId)
		oRecordset.Close
		Set oRecordset = Nothing
		Set oDictionary = Nothing
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