<% 
	Option Explicit	
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
%>

<!--#INCLUDE virtual="/empiria/bin/ms_scripts/rs.asp"-->

<% RSDispatch %>

<SCRIPT RUNAT=SERVER Language=javascript>

function IServerScripts() 	{
	this.ActivateReport			= Function('nReportId', 'return ChangeReportStatus(nReportId, "A")');	  
	this.CboSubClasses			= Function('nClassId', 'nSelectedItemId', 'return CboSubClasses(nClassId, nSelectedItemId)');
	this.CboDataItemOperations = Function('nClassId', 'sAttribute', 'nSelectedItemId', 'return CboDataItemOperations(nClassId, sAttribute, nSelectedItemId)');
	this.CboDataItemOrders  = Function('nDataItemId', 'nSelectedItemId', 'return CboDataItemOrders(nDataItemId, nSelectedItemId)');	  
	this.CboItemValues      = Function('nDictionaryItemId', 'nSelectedItemId', 'return CboItemValues(nDictionaryItemId, nSelectedItemId)');
	this.CboSubClasses      = Function('nDataItemId', 'nSelectedItemId', 'return CboSubClasses(nDataItemId, nSelectedItemId)');	
	this.DeleteItem			    = Function('nItemId', 'return DeleteItem(nItemId)');
	this.DeleteReport			  = Function('nReportId', 'return DeleteReport(nReportId)');	  	  
	this.DeleteRow			    = Function('nItemId', 'return DeleteRow(nItemId)');
	this.DeleteSection      = Function('nSectionId', 'return DeleteSection(nSectionId)');
	this.ExistsItemInPosition = Function('nSectionId', 'nRow', 'nColumn', 'nItemId', 'return ExistsItemInPosition(nSectionId, nRow, nColumn, nItemId)'); 
	this.InsertRow			    = Function('nItemId', 'nDirection', 'return InsertRow(nItemId, nDirection)');
	this.IsNumeric			    = Function('nValue', 'return IsNumericOK(nValue)');
	this.ItemDataType       = Function('nDictionaryItemId', 'return ItemDataType(nDictionaryItemId)');
	this.ItemDataType       = Function('nDictionaryItemId', 'return ItemDataType(nDictionaryItemId)');	  
	this.ItemAlias          = Function('nDictionaryItemId', 'return ItemAlias(nDictionaryItemId)');
	this.SuspendReport			= Function('nReportId', 'return ChangeReportStatus(nReportId, "S")');
	this.TraslapingRows     = Function('nReportId', 'nFromRow', 'nToRow', 'sWorksheet', 'return TraslapingRows(nReportId, nFromRow, nToRow, sWorksheet)');
}

public_description = new IServerScripts();  

</SCRIPT>

<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">

Function CboSubClasses(nClassId, nSelectedItemId)
	Dim oDictionary, sTemp
	'*****************************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	sTemp = oDictionary.CboSubClasses(Session("sAppServer"), CLng(nClassId), CLng(nSelectedItemId))
	Set oDictionary = Nothing
	If Len(sTemp) <> 0 Then
		sTemp = "<SELECT name=cboSubClasses style='width:330'>" & sTemp & "</SELECT>"
	Else
		sTemp = "<SELECT name=cboSubClasses style='width:330'>" & _
						"<OPTION value=0>-- No existen subestructuras para el elemento seleccionado --</OPTION></SELECT>"
	End If
	CboSubClasses = sTemp
End Function

Function CboDataItemOperations(nClassId, sAttribute, nSelectedItemId)
	Dim oDictionary, sTemp
	'********************************************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	sTemp = oDictionary.CboDataItemOperations(Session("sAppServer"), CLng(nClassId), CStr(sAttribute), CLng(nSelectedItemId))
	Set oDictionary = Nothing
	If Len(sTemp) <> 0 Then
		sTemp = "<SELECT name=cboOperations style='width:320'>" & VbCrLf & _
					  "<OPTION value=0>-- No aplicar ninguna operación sobre el elemento --</OPTION>" & VbCrLf & _
					  sTemp & "</SELECT>" & VbCrLf
	Else
		sTemp = "<SELECT name=cboOperations style='width:320'>" & _
						"<OPTION value=0>-- No existen operaciones para el elemento seleccionado --</OPTION></SELECT>"
	End If
	CboDataItemOperations = sTemp
End Function

Function CboDataItemOrders(nDataItemId, nSelectedItemId)
	Dim oDictionary, sTemp
	'*****************************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	sTemp = oDictionary.CboDataItemOrders(Session("sAppServer"), CLng(nDataItemId), CLng(Session("uid")), CLng(nSelectedItemId))
	Set oDictionary = Nothing
	If Len(sTemp) <> 0 Then
		sTemp = "<SELECT name=cboDataOrders style='width:330'>" & VbCrLf & _
						"<OPTION value=0>-- No aplicar ningún ordenamiento sobre el elemento --</OPTION>" & VbCrLf & _
						sTemp & "</SELECT>" & VbCrLf
	Else
		sTemp = "<SELECT name=cboDataOrders style='width:330'>" & VbCrLf & _
						"<OPTION value=0>-- No existen ordenamientos para la estructura seleccionada --</OPTION></SELECT>" & VbCrLf
	End If
	CboDataItemOrders = sTemp
End Function

Function CboItemValues(nDictionaryItemId, nSelectedItemId) 
	Dim oDictionary
	'*******************************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	CboItemValues = oDictionary.CboItemValues(Session("sAppServer"), CLng(nDictionaryItemId), CLng(nSelectedItemId))
	Set oDictionary = Nothing
End Function

Function ChangeReportStatus(nReportId, sNewStatus)
	Dim oReportDesigner
	'***********************************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.ChangeStatus Session("sAppServer"), CLng(nReportId), CStr(sNewStatus)
	Set oReportDesigner = Nothing
End Function

Function DeleteItem(nItemId)
	Dim oReportDesigner
	'*****************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.DeleteItem Session("sAppServer"), CLng(nItemId)
	Set oReportDesigner = Nothing
End Function

Function DeleteReport(nReportId)
	Dim oReportDesigner
	'*****************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.DeleteReport Session("sAppServer"), CLng(nReportId)
	Set oReportDesigner = Nothing
End Function

Function DeleteRow(nItemId)
	Dim oReportDesigner
	'************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.DeleteRow Session("sAppServer"), CLng(nItemId)
	Set oReportDesigner = Nothing
End Function

Function DeleteSection(nSectionId)
	Dim oReportDesigner
	'************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.DeleteSection Session("sAppServer"), CLng(nSectionId)
	Set oReportDesigner = Nothing
End Function

Function ExistsItemInPosition(nSectionId, nRow, nColumn, nItemId)
	Dim oReportDesigner
	'**************************************************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  ExistsItemInPosition = oReportDesigner.ExistsItemInPosition(Session("sAppServer"), CLng(nSectionId), CLng(nRow), CLng(nColumn), CLng(nItemId))
	Set oReportDesigner = Nothing
End Function

Function InsertRow(nItemId, nDirection)
	Dim oReportDesigner
	'************************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  oReportDesigner.InsertRow Session("sAppServer"), CLng(nItemId), CLng(nDirection)
	Set oReportDesigner = Nothing
End Function
	  
Function IsNumericOK(nValue) 
	IsNumericOK = IsNumeric(nValue)
End Function

Function ItemAlias(nDictionaryItemId) 
	Dim oDictionary
	'************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	ItemAlias = oDictionary.ItemAlias(Session("sAppServer"), CLng(nDictionaryItemId))
	Set oDictionary = Nothing
End Function


Function ItemDataType(nDictionaryItemId) 
	Dim oDictionary
	'************************************
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")	
	ItemDataType = oDictionary.ItemDataType(Session("sAppServer"), CLng(nDictionaryItemId))
	Set oDictionary = Nothing
End Function

Function TraslapingRows(nReportId, nFromRow, nToRow, sWorksheet)
	Dim oReportDesigner
	'***********************************************
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
  TraslapingRows = oReportDesigner.RowsInSection(Session("sAppServer"), CLng(nReportId), _
																								 CLng(nFromRow), CLng(nToRow), CStr(sWorksheet))
	Set oReportDesigner = Nothing
End Function

</SCRIPT>