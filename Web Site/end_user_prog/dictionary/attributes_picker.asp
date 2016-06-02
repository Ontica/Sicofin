<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTemplateFile, gsTitle
	Dim gnClassId, gsClassName, gsPickerType, gsAttributesTable, gsReportDescription, gsTackedWindows	
	Dim gsColName, gnReportDataId, gsColDescription, gsColPosition, gsCboWorkSheets
	Dim gsFilterDataType, gsPosition, gsLength, gsPivotColumn, gnSubClassId
	Dim oDictionary, gsCboAttributes, gscboFilters, gsCboOperations, gsCboExcelColumns
	
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")
	
	
	gnClassId    = Request.QueryString("classId")
	gnSubClassId = Request.QueryString("subclassId")
	gsPickerType = Request.QueryString("pickerType")
	
	Call Main()
	
	Sub Main() 
		Dim oRecordset
		'***************
		Set oRecordset = oDictionary.GetItem(Session("sAppServer"), CLng(gnClassId))
		gsClassName   = oRecordset("itemName")
		If Len(oRecordset("itemName")) > 64 Then
			gsTitle	= gsClassName & "..."
		Else
			gsTitle	= gsClassName
		End If
		'gsAttributesTable = oDictionary.GetAttributesFiltersTable(Session("sAppServer"), CLng(gnClassId), CLng(gnSubClassId))
		Set oRecordset = Nothing
		'Set oRecordset  = oDictionary.GetItem(Session("sAppServer"), CLng(gnFilterId))		
		gsCboAttributes	= oDictionary.CboItemAttributes(Session("sAppServer"), CLng(gnClassId), CLng(gnSubClassId))
		'gsFilterDataType = oRecordset("itemDataType")
		'gnReportDataId = CLng(oRecordset("reportDataId"))
		Set oRecordset = Nothing		
	End Sub
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Diseñador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oBuildWindow = null;
var gbSended = false, nAttrVisible = null;

function itemAlias(nItemId) {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "ItemAlias", nItemId);
	return (obj.return_value);
}

function refreshValues(nType) {
	document.all.divNoValues.style.display = 'none';	
	document.all.divValue.style.display = 'none';	 
	document.all.divComboValues.style.display = 'none';
	document.all.divBetweenValues.style.display = 'none';
	nAttrVisible = nType;
	switch (nType) {
		case 0:
			document.all.divNoValues.style.display = 'inline';			
			return false;
		case 1:
			document.all.divValue.style.display = 'inline';	 
			return false;
		case 2:
			document.all.divComboValues.style.display = 'inline';
			return false;
		case 3:
			document.all.divBetweenValues.style.display = 'inline';	
			return false;
	}
}

function addFilter_(sOperator) {
	var oObj, sTemp;
	
	oObj = document.all.cboAttributes;
	if (oObj.value == 0) {
		alert("Requiero la selección de un atributo de la lista.");
		oObj.focus();
		return false;
	}
	
	sTemp = itemAlias(oObj.value);	
	if (document.all.txtExpression.value.indexOf(sTemp) != -1) {
		alert("El atributo seleccionado ya está incluido en la lista.");
		oObj.focus();
		return false;
	}	
	if (document.all.txtExpression.value != '') {
		document.all.txtExpression.value += ' ' + sOperator + ' ' + sTemp;
	} else {	
		document.all.txtExpression.value = sTemp;
	}
			
	sTemp = oObj.options[oObj.selectedIndex].text;
	if (document.all.txtViewer.value != '') {
		document.all.txtViewer.value += ' ' + sOperator + ' ' + sTemp;
	} else {	
		document.all.txtViewer.value = sTemp;
	}	
	return false;
}

function addOrder(sMode) {
	var oObj, sTemp;
	
	oObj = document.all.cboAttributes;
	if (oObj.value == 0) {
		alert("Requiero la selección de un atributo de la lista.");
		oObj.focus();
		return false;
	}
	
	sTemp = itemAlias(oObj.value);	
	if (document.all.txtExpression.value.indexOf(sTemp) != -1) {
		alert("El atributo seleccionado ya está incluido en la lista.");
		oObj.focus();
		return false;
	}
	sTemp += ' ' + sMode;
	if (document.all.txtExpression.value != '') {
		document.all.txtExpression.value += ', ' + sTemp;
	} else {	
		document.all.txtExpression.value = sTemp;
	}
			
	sTemp = oObj.options[oObj.selectedIndex].text + ' ' + sMode;	
	if (document.all.txtViewer.value != '') {
		document.all.txtViewer.value += ', ' + sTemp;
	} else {	
		document.all.txtViewer.value = sTemp;
	}	
	return false;
}

function addGrouping() {
	var oObj, sTemp;
	
	oObj = document.all.cboAttributes;
	if (oObj.value == 0) {
		alert("Requiero la selección de un atributo de la lista.");
		oObj.focus();
		return false;
	}
	
	sTemp = itemAlias(oObj.value);	
	if (document.all.txtExpression.value.indexOf(sTemp) != -1) {
		alert("El atributo seleccionado ya está incluido en la lista.");
		oObj.focus();
		return false;
	}
	if (document.all.txtExpression.value != '') {
		document.all.txtExpression.value += ', ' + sTemp;
	} else {	
		document.all.txtExpression.value = sTemp;
	}
	
	sTemp = oObj.options[oObj.selectedIndex].text;
	if (document.all.txtViewer.value != '') {
		document.all.txtViewer.value += ', ' + sTemp;
	} else {	
		document.all.txtViewer.value = sTemp;
	}	
	return false;
}

function refreshSelection() {
	document.all.txtViewer.value = '';
	document.all.txtExpression.value = '';
	return false;
}

function oData() {
	var dataExpression; 
	var dataViewer; 
}

function pickData() {
	if (gbSended) {
		return false;
	}
	gbSended = true;
	
	oData.dataExpression = document.all.txtExpression.value;
	oData.dataViewer     = document.all.txtViewer.value;
  if (window.dialogArguments != null) {
		window.dialogArguments.dataExpression = oData.dataExpression;
		window.dialogArguments.dataViewer     = oData.dataViewer;
  }
  window.returnValue = true;
  window.close();
}

function loadArguments() {
	if (window.dialogArguments != null) {
		oData.dataExpression = window.dialogArguments.dataExpression;
		oData.dataViewer     = window.dialogArguments.dataViewer;		
	}
  document.all.txtExpression.value = oData.dataExpression;
  document.all.txtViewer.value     = oData.dataViewer;  
  window.returnValue = false;
  return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload='loadArguments();'>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
		<% If gsPickerType = "dataGrouping" Then %>
			Definición del agrupamiento
		<% ElseIf gsPickerType = "dataOrdering" Then %>
			Definición del ordenamiento
		<% ElseIf gsPickerType = "dataFiltering" Then %>
			Definición del filtro
		<% End If %>
		</TD>
	  <TD align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=fullScrollMenu>
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle nowrap>
						<%=gsTitle%>
					</TD>										
					<TD align=right nowrap>
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='document.all.frmSend.reset();' alt="Refrescar">
					</TD>					
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
        <TR class=applicationTableRowDivisor>
					<% If gsPickerType = "dataGrouping" Then %>
						<TD colspan=2><b>Campos disponibles para el agrupamiento</b></TD>
					<% ElseIf gsPickerType = "dataOrdering" Then %>
						<TD colspan=2><b>Campos disponibles para el ordenamiento</b></TD>
					<% ElseIf gsPickerType = "dataFiltering" Then %>
						<TD colspan=2><b>Campos disponibles para el filtrado</b></TD>
					<% End If %>					
			  </TR>			  
			  <TR nowrap>
					<TD valign=top>Atributo:</TD>
					<TD>
						<SELECT name=cboAttributes style='width:320'>
						 <OPTION value=0>--Seleccionar un atributo de la lista--</OPTION>
						 <%=gsCboAttributes%>						 
						</SELECT>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<% If gsPickerType = "dataGrouping" Then %>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdAddGroup style='width:70' value="Agregar" onclick='addGrouping();'>
						<% ElseIf gsPickerType = "dataOrdering" Then %>
						<INPUT class=cmdSubmit type=button name=cmdAddOrderAsc style='width:70' value="Ascendente" onclick='addOrder("ASC");'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdAddOrderDesc style='width:70' value="Descendente" onclick='addOrder("DESC");'>
						<% ElseIf gsPickerType = "dataFiltering" Then %>
						<INPUT class=cmdSubmit type=button name=cmdAddFilterAnd style='width:70' value="Agregar 'Y'" onclick='addFilter_("AND");'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdAddFilterOr style='width:70' value="Agregar 'O'" onclick='addFilter_("OR");'>
						<% End If %>
					</TD>
			  </TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2>
						<b>Visor de la expresión construida</b>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<a href='' onclick='return(refreshSelection());'>Restablecer</a>
					</TD>
			  </TR>			  
			  <TR nowrap>
					<TD valign=top>&nbsp;</TD>
					<TD>
						<TEXTAREA rows=3 name=txtViewer style="width:320" readonly></TEXTAREA>						
					</TD>
			  </TR> 
			</TABLE>
		</TD>
	</TR>
	<TR>
	  <td colspan=4 nowrap align=right>
		 <INPUT type="hidden" name=txtExpression>
		 <INPUT type="hidden" name=txtClassId value="<%=gnClassId%>">
		 <INPUT TYPE=hidden name=txtTackedWindows>						
			<INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="pickData();">						
			&nbsp; &nbsp;
	   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
	   &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
	  </td>
	</TR>	
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>