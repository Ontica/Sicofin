<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTemplateFile, gsTitle
	Dim gnFilterId, gnClassId, gsClassName, gsAttributesTable, gsReportDescription, gsTackedWindows	
	Dim gsColName, gnReportDataId, gsColDescription, gsColPosition, gsCboWorkSheets
	Dim gsFilterDataType, gsPosition, gsLength, gsPivotColumn, gnSubClassId
	Dim oDictionary, gsCboAttributes, gscboFilters, gsCboOperations, gsCboExcelColumns
	
	Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")
	
	
	gnClassId    = Request.QueryString("classId")
	gnSubClassId = Request.QueryString("subClassId")

	Call FilterValues()
	
	Sub FilterValues() 
		Dim oRecordset
		'***************
		Set oRecordset = oDictionary.GetItem(Session("sAppServer"), CLng(gnClassId))
		gsClassName   = oRecordset("itemName")
		If Len(oRecordset("itemName")) > 46 Then
			gsTitle	= gsClassName & "..."
		Else
			gsTitle	= gsClassName
		End If
		'gsAttributesTable = oDictionary.GetAttributesFiltersTable(Session("sAppServer"), CLng(gnClassId), CLng(gnSubClassId))
		Set oRecordset = Nothing
		Set oRecordset  = oDictionary.GetItem(Session("sAppServer"), CLng(gnFilterId))		
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

function oData() {
	var dataExpression; 
	var dataViewer; 
}

function openWindow(sWindowName) {
	var sURL, sPars;
	
	switch (sWindowName) {
		case 'createFilter':
			if (document.all.cboDataItems.value == 0 || document.all.cboDataItems.value == '') {
				alert("Para crear el filtro, primero requiero la selección del elemento de información que presentará la columna.");
				document.all.cboDataItems.focus();
				return false;
			}
			sURL = 'build_filter.asp?reportId=<%=gnClassId%>&id=' + arguments[1];
			sPars = 'height=380px,width=480px,resizable=no,scrollbars=no,status=no,location=no';
			oBuildWindow = createWindow(oBuildWindow, sURL, sPars);
			return false;
	}	
	return false;	
}

function openFile(sFileName) {
	return false;
}

function buildOperation() {
	notAvailable();	
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
	var sTemp, sTemp2;
	
	switch (nAttrVisible) {
		case 1:	
			if (document.all.txtAttrValue.value == '') {	
				alert("Requiero el valor del filtro.");
				return false;
			}
			sTemp  = '(' + itemAlias(document.all.cboAttributes.value) + ' ';
			sTemp += document.all.cboOperators.value + " '" + document.all.txtAttrValue.value + "')";
			sTemp2 = sTemp;
			break;
		case 2:
			sTemp  = '(' + itemAlias(document.all.cboAttributes.value) + ' ';
			sTemp += document.all.cboOperators.value + " '";
			sTemp2 = sTemp;
			
			sTemp  += document.all.cboAttrValues.value + "')";
			sTemp2 += document.all.cboAttrValues.options[document.all.cboAttrValues.selectedIndex].text + "')";
			break;
		case 3:
			if (document.all.txtAttrValueA.value == '' || document.all.txtAttrValueB.value == '') {	
				alert("Necesito los valores entre los que estará el atributo.");
				return false;
			}		
			sTemp  = "('" + document.all.txtAttrValueA.value + "' <= " + itemAlias(document.all.cboAttributes.value); 
			sTemp += "<= '"	+ document.all.txtAttrValueB.value + "')";			
			sTemp2 = sTemp;
			break;
	}
	
	if (document.all.txtExpression.value != '') {
		document.all.txtExpression.value += ' ' + sOperator + ' ' + sTemp;
	} else {	
		document.all.txtExpression.value = sTemp;
	}
		
	if (document.all.txtViewer.value != '') {
		document.all.txtViewer.value += ' ' + sOperator + ' ' + sTemp2;
	} else {
		document.all.txtViewer.value = sTemp2;
	}
					
	return false;
}

function itemAlias(nItemId) {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "ItemAlias", nItemId);
	return (obj.return_value);
}

function cboAttributes_onchange() {
	var obj, nItemDataType;	
		
	if (document.all.cboAttributes.value == 0) {
		refreshValues(0);
		return false;
	}
	
	obj = RSExecute("../end_user_prog_scripts.asp", "ItemDataType", document.all.cboAttributes.value);
	nItemDataType = obj.return_value;
	if (nItemDataType == 'L' || nItemDataType == 'K') {
		obj = RSExecute("../end_user_prog_scripts.asp", "CboItemValues", document.all.cboAttributes.value);
		document.all.divCboAttrValues.innerHTML = "<SELECT name=cboAttrValues style='width:320'>" + obj.return_value + '</SELECT>';
		refreshValues(2);	
	} else {
		refreshValues(1);	
	}
}

function cboOperators_onchange() {
	cboAttributes_onchange();
	if (document.all.cboAttributes.value == 0) {
		return false;
	}
	if (document.all.cboOperators.value == "/") {
		refreshValues(3);
	}	
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

function cleanFilter() {
	document.all.txtViewer.value = '';
	document.all.txtExpression.value = '';
	return false;
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
		<% If gnFilterId = 0 Then %>
			Crear filtro
		<% Else %>
			Editar filtro
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
					<TD colspan=2><b>Construcción del filtro</b></TD>
			  </TR>			  
			  <TR nowrap>
					<TD valign=top>Atributo:</TD>
					<TD>
						<SELECT name=cboAttributes style='width:320' LANGUAGE=javascript onchange="return cboAttributes_onchange()">
						 <OPTION value=0>--Seleccionar un atributo de la lista--</OPTION>
						 <%=gsCboAttributes%>						 
						</SELECT>
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>Operador:</TD>
					<TD>
					 	<SELECT name=cboOperators style='width:160' LANGUAGE=javascript onchange="return cboOperators_onchange()">
						  <OPTION value="=">Igual a (=)</OPTION>
						  <OPTION value="<>">Distinto de (<>)</OPTION>
						  <OPTION value="Like">Parecido a ('Like')</OPTION>
						  <OPTION value=">">Mayor que (>)</OPTION>
						  <OPTION value=">=">Mayor o igual que (>=)</OPTION>
						  <OPTION value="<">Menor que (<)</OPTION>
						  <OPTION value="<=">Menor o igual que (<=)</OPTION>
						  <OPTION value="/">Entre los valores A y B </OPTION>						  
						</SELECT>
					</TD>
			  </TR>
				<TR id=divNoValues nowrap style='display:inline;'>
					<TD valign=top>Valor:</TD>
					<TD><b>Primero se debe seleccionar un atributo de la lista</b><br>&nbsp;<br>&nbsp;<br>&nbsp;<br>&nbsp;<br>&nbsp;</TD>
			  </TR>			  
				<TR id=divBetweenValues nowrap style='display:none;'>
					<TD valign=top>Entre:</TD>
					<TD>
						A: <INPUT name=txtAttrValueA style="width:145">
						B: <INPUT name=txtAttrValueB style="width:145">
						<br>&nbsp;<br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'Y'" onclick='addFilter_("AND");'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'O'" onclick='addFilter_("OR");'>		
					</TD>			
			  </TR>			  
			  <TR id=divComboValues nowrap style='display:none;'>
					<TD valign=top>Valor:</TD>
					<TD>
						<div id=divCboAttrValues>
						<SELECT name=cboAttrValues style="width:320">
						</SELECT>
						</div>
						<br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;						
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'Y'" onclick='addFilter_("AND");'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'O'" onclick='addFilter_("OR");'>
						<br>&nbsp;	
					</TD>
			  </TR>
			  <TR id=divValue nowrap style='display:none;'>
					<TD valign=top>Valor:</TD>
					<TD>
						<INPUT name=txtAttrValue style="width:160">						
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'Y'" onclick='addFilter_("AND");'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Agregar 'O'" onclick='addFilter_("OR");'>
						<br>&nbsp;<br>&nbsp;
					</TD>
			  </TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2>
						<b>Visor del filtro construido</b>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<a href='' onclick='return(cleanFilter());'>Limpiar</a>
					</TD>
			  </TR>			  
			  <TR nowrap>
					<TD valign=top>Filtro:</TD>
					<TD>
						<INPUT TYPE=hidden name=txtExpression>
						<TEXTAREA rows=3 name=txtViewer style="width:320"  readonly></TEXTAREA>
					</TD>
			  </TR> 
			</TABLE>
		</TD>
	</TR>
	<TR>
	  <td colspan=4 nowrap align=right>	   
		 <INPUT type="hidden" name=txtClassId value="<%=gnClassId%>">			
			<INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="pickData();">						
			&nbsp; &nbsp;
	   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
	   &nbsp; &nbsp; &nbsp;
	  </td>
	</TR>	
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>