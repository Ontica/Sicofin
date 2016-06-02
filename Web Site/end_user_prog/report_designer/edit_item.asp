<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnItemId, gnParentRowId, gnReportDataClassId, gnReportDataSubClassId, gsReportTechnology
	Dim gsCboSectionRows, gsCboExcelColumns, gnSectionId, gsFilterViewer, gsFilterExpression
	Dim gsCboParameters, gsCboDataItems, gsCboOperations, gsCboItemTypes, gsPrintLayout
	Dim gsName, gsColumn, gsLength, gsItemType, gsItemValue, gsItemTag, gsFiltersBox
	
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oDictionary, oRecordset
		Dim nReportId, nItemDataId, nItemOperationId
		'**********************************************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		If (Len(Request.QueryString("id")) <> 0) Then
			gnItemId = Request.QueryString("id")
			gnParentRowId = GetParentRowId(gnItemId)
			Set oRecordset = oReportDesigner.GetItem(Session("sAppServer"), CLng(gnItemId))
			nReportId				 = oRecordset("reportId")
			gnSectionId		   = oRecordset("itemSectionId")
			gsName					 = oRecordset("itemName")			
			gsColumn				 = oRecordset("itemColumn")
			gsLength				 = oRecordset("itemLength")
			gsItemValue			 = oRecordset("itemValue")
			gsItemTag        = oRecordset("itemTag")
			nItemOperationId = oRecordset("itemOperationId")			
			gsItemType       = oRecordset("itemType")
			gsCboItemTypes   = oReportDesigner.CboItemTypes(CStr(gsItemType))
			gsFilterExpression = oRecordset("ItemFilter")
			gsFilterViewer   = oRecordset("ItemFilterDesc")			
			'FilterBox Perdido
			'gsFiltersBox     = oReportDesigner.SectionFiltersBox(Session("sAppServer"), CLng(gnSectionId))
			oRecordset.Close
			Set oRecordset = Nothing
			Set oRecordset         = oReportDesigner.GetReport(Session("sAppServer"), CLng(nReportId))
			gnReportDataClassId    = oRecordset("reportDataClassId")
			gnReportDataSubClassId = oRecordset("reportDataSubClassId")
			gsReportTechnology     = oRecordset("reportTechnology")
			oRecordset.Close
			Set oRecordset = Nothing			
			Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")
			nItemDataId     = oDictionary.ItemId(Session("sAppServer"), CLng(gnReportDataClassId), CStr(gsItemValue))
			gsCboParameters = oDictionary.CboParameters(Session("sAppServer"), CLng(gnReportDataClassId), CStr(gsItemValue))
			gsCboDataItems	= oDictionary.CboDataItemAttributes(Session("sAppServer"), CLng(gnReportDataClassId), CStr(gsItemValue))
			'gsCboFilters    = oDictionary.CboDataItemFilters(Session("sAppServer"), CLng(gnReportDataClassId), CLng(Session("uid")), CLng(nItemFilterId))
			gsCboOperations = oDictionary.CboDataItemOperations(Session("sAppServer"), CLng(gnReportDataClassId), CStr(gsItemValue), CLng(nItemOperationId))
		Else
			gnItemId      = 0
			gnParentRowId = CLng(Request.QueryString("rowId"))
			gsColumn    = 0
			If (Len(Request.QueryString("col")) <> 0) Then
				gsColumn    = CLng(Request.QueryString("col"))
			End If
			Set oRecordset = oReportDesigner.GetItem(Session("sAppServer"), CLng(gnParentRowId))
			nReportId      = oRecordset("reportId")
			gnSectionId		 = oRecordset("itemSectionId")
			gsItemType     = ""
			gsCboItemTypes   = oReportDesigner.CboItemTypes()
			gsFilterExpression = ""
			gsFilterViewer     = ""
			gsFiltersBox       = oReportDesigner.SectionFiltersBox(Session("sAppServer"), CLng(gnSectionId))
			oRecordset.Close
			Set oRecordset = Nothing
			Set oDictionary = Server.CreateObject("AOReportsDesigner.CDictionary")
			Set oRecordset = oReportDesigner.GetReport(Session("sAppServer"), CLng(nReportId))
			gnReportDataClassId = oRecordset("reportDataClassId")
			gnReportDataSubClassId = oRecordset("reportDataSubClassId")
			gsReportTechnology = oRecordset("reportTechnology")
			oRecordset.Close
			Set oRecordset = Nothing
			gsCboParameters = oDictionary.CboParameters(Session("sAppServer"), CLng(gnReportDataClassId))
			gsCboDataItems	= oDictionary.CboDataItemAttributes(Session("sAppServer"), CLng(gnReportDataClassId))
			'gsCboFilters    = oDictionary.CboDataItemFilters(Session("sAppServer"), CLng(gnReportDataClassId), CLng(Session("uid")))
			'gsCboOperations = oDictionary.CboDataItemOperations(Session("sAppServer"), CLng(gnReportDataClassId), CLng(nItemDataId))
		End If
		gsCboSectionRows  = oReportDesigner.CboSectionRows(Session("sAppServer"), CLng(gnSectionId), CLng(gnParentRowId))
		If (gsReportTechnology = "E") Then			
			gsCboExcelColumns = oReportDesigner.CboExcelColumns(CLng(gsColumn))			
		End If
		Set oRecordset = Nothing
	
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("./exec/exception.asp")
		End If		  
	End Sub
	
	Function GetParentRowId(nChildId)
		Dim oReportDesigner
		'******************************
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		GetParentRowId = oReportDesigner.GetParentRowId(Session("sAppServer"), CLng(nChildId))
		Set oReportDesigner = Nothing
	End Function
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
var gbSended = false;

function deleteItem() {
	var obj;
	if (confirm('¿Elimino el elemento del reporte?')) {
		obj = RSExecute("../end_user_prog_scripts.asp", "DeleteItem", <%=gnItemId%>);	
		window.opener.location.href = window.opener.location.href;
		window.close();
	}
	return false;
}

function existsItemInPosition() {
	var obj;
	var nRow, nColumn;

	nRow = document.all.cboSectionRows.options[document.all.cboSectionRows.selectedIndex].text;
	<% If (gsReportTechnology = "E") Then %>
		nColumn = document.all.cboExcelColumns.value;
	<% Else %>
		nColumn = document.all.txtPosition.value;
	<% End If %>
	obj = RSExecute("../end_user_prog_scripts.asp", "ExistsItemInPosition", <%=gnSectionId%>, nRow, nColumn, <%=gnItemId%>);
	return (obj.return_value);
}


function updateCboFilters() {
	var obj;
	
	//obj = RSExecute("../end_user_prog_scripts.asp", "CboDataItemFilters", document.all.cboDataItems.value, 0);
	//document.all.divDataFilters.innerHTML = obj.return_value;
}

function updateCboOperations() {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "CboDataItemOperations", <%=gnReportDataClassId%>, document.all.cboDataItems.value, 0);
	document.all.divDataOperations.innerHTML = obj.return_value;
}

function setHiddenControls() {
	switch(document.all.cboItemTypes.value) {
		case 'L':
			document.all.txtItemTag.value = document.all.txtLabel.value;
			break;
		case 'P':
			document.all.txtItemTag.value = document.all.cboParameters.options[document.all.cboParameters.selectedIndex].text;
			break;
		case 'F':
			document.all.txtItemTag.value = document.all.cboDataItems.options[document.all.cboDataItems.selectedIndex].text;
			break;				
		case 'E':
			document.all.txtItemTag.value = 'Expresión ???';
			break;
	}
}

function checkFiltering() {	
	if (typeof(document.all.cboFilters) == 'undefined') {
		return true;
	}		
	//if (document.all.cboFilters.length != null) {
		//for (i = 0 ; i < document.all.cboFilters.length ; i++) {
		//	if (document.all.cboFilters[i].value == 0) {
		//		alert("Requiero el valor del filtro.");
		//		document.all.cboFilters[i].focus();		
		//		return false;
		// }
		//}
	//} else {
		if (document.all.cboFilters.value == 0) {			
			alert("Requiero el valor del filtro.");
			return false;
		}
	//}		
	return true;
}

function constructFilter() {	
	if (typeof(document.all.cboFilters) == 'undefined') {
		return true;
	}	
	//if (document.all.cboFilters.length != null) {
//		for (i = 0 ; i < document.all.cboFilters.length ; i++) {
//			if (document.all.txtFilterExp.value != '') {				
//				document.all.txtFilterExp.value += ' AND ';
//			}			
//			document.all.txtFilterExp.value += '(' + document.all.cboFilters[i].tag + ' = ' + document.all.cboFilters[i].value + ')';
//		}		
//	} else {		
		if (document.all.txtFilterExp.value != '') {
			document.all.txtFilterExp.value += ' AND ';
		}	
		document.all.txtFilterExp.value += '(' + document.all.cboFilters.tag + ' = ' + document.all.cboFilters.value + ')';
	//}
	return false;
}


function doSubmit() {
	var sMsg, nVoucherId;

  if (gbSended) {
		return false;
	}
	<% If (gsReportTechnology = "T") Then %>
	if (document.all.txtPosition.value == '') {
		alert('Requiero la posición de la columna.');
		document.all.txtPosition.focus();
		return false;
	}	
	if (document.all.txtLength.value == '') {
		alert('Requiero la longitud de la columna.');
		document.all.txtLength.focus();
		return false;
	}	
	<% End If %>
	if (existsItemInPosition()) {
		alert("Ya existe un elemento en la posición seleccionada.");
		return false;		
	}
	if (document.all.cboItemTypes.value == '') {
		alert('Requiero el tipo de datos que desplegará el elemento.');
		document.all.cboItemTypes.focus();
		return false;
	}
	if (document.all.cboItemTypes.value == 'L' && document.all.txtLabel.value == '') {
		alert('Requiero la etiqueta que desplegará el elemento.');
		document.all.txtLabel.focus();
		return false;
	}
	if (document.all.cboItemTypes.value == 'P' && document.all.cboParameters.value == '') {
		alert('Requiero el parámetro que se mostrará en el elemento.');
		document.all.cboParameters.focus();
		return false;
	}	
	if (document.all.cboItemTypes.value == 'F' && document.all.cboDataItems.value == '') {
		alert('Requiero la selección del elemento de información o campo que se presentará.');
		document.all.cboDataItems.focus();
		return false;
	}
	if (document.all.cboItemTypes.value == 'E') {
		alert('Por el momento no está disponible el manejo de expresiones.\n\nGracias ...');
		document.all.cboItemTypes.focus();
		return false;
	}
	if (document.all.cboItemTypes.value == 'F') {
		if (!checkFiltering()) {
			return false;
		}
	}
	constructFilter();
	setHiddenControls();
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function oData() {
	var dataExpression; 
	var dataViewer; 
}

function pickData(sPickerName) {
	var sURL, sPars;
	
	sURL  = '../dictionary/filter_picker.asp?';
	sURL += 'classId=<%=gnReportDataClassId%>&subclassId=<%=gnReportDataSubClassId%>';	
	sPars = "dialogHeight:410px;dialogWidth:400px;resizable:no;scroll:no;status:no;help:no;";	
	switch (sPickerName) {
		case 'dataFiltering':
			oData.dataExpression = document.all.txtFilterExp.value;
			oData.dataViewer     = document.all.txtFilterViewer.value;
			if (window.showModalDialog(sURL, oData, sPars)) {
				document.all.txtFilterExp.value    = oData.dataExpression;
				document.all.txtFilterViewer.value = oData.dataViewer;				
			}
			return false;
	}	
}

function buildOperation() {
	notAvailable();	
}

function cboItemTypes_onchange() {
	document.all.divNone.style.display    = 'none';
	document.all.divLabel.style.display   = 'none';
	document.all.divParameters.style.display = 'none';
	document.all.divFields1.style.display = 'none';
	document.all.divFields2.style.display = 'none';
	document.all.divFields3.style.display = 'none';	
	document.all.divFields4.style.display = 'none';	
	document.all.divFields5.style.display = 'none';	
	switch(document.all.cboItemTypes.value) {
		case '':
			document.all.divNone.style.display = 'inline';
			break;
		case 'L':
			document.all.divLabel.style.display = 'inline';	
			break;
		case 'P':
			document.all.divParameters.style.display = 'inline';
			break;			
		case 'F':
			document.all.divFields1.style.display = 'inline';
			document.all.divFields2.style.display = 'inline';
			document.all.divFields3.style.display = 'inline';
			document.all.divFields4.style.display = 'inline';
			document.all.divFields5.style.display = 'inline';
			break;
		case 'E':
			break;
	}
}

function window_onload() {
	document.all.cboItemTypes.value = '<%=gsItemType%>';
	cboItemTypes_onchange();
}

function cboDataItems_onchange() {
	updateCboOperations();
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmSend action='./exec/save_item.asp' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			<% If gnItemId = 0 Then %>
				Nuevo elemento o celda
			<% Else %>
				Edición del elemento o celda
			<% End If %>
		</TD>
	  <TD align=right nowrap>			<A href='' onclick="return(deleteItem());">Eliminar</A>			
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;
		</TD>
	</TR>
	<TR>
		<TD colspan=2 nowrap>
			<TABLE class=applicationTable>
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Identificación y posición del elemento</b></TD>
			  </TR>			
			  <TR>
					<TD valign=top nowrap>Nombre del elemento:</TD>
			    <TD colspan=3 width=430>
						<INPUT name=txtName style="height:20px;width:320px;" value='<%=gsName%>'>	
			    </TD>	    
			  </TR>
			  <% If (gsReportTechnology  = "E") Then %>
			  <TR>
					<TD valign=top nowrap>Posición en donde se colocará:</TD>					
			    <TD colspan=3 nowrap>
						Renglón: &nbsp;
						<SELECT name=cboSectionRows style="width:60">
							<%=gsCboSectionRows%>
						</SELECT>
						&nbsp; &nbsp;
						Columna: &nbsp;
						<SELECT name=cboExcelColumns style="width:60" >
							<%=gsCboExcelColumns%>
						</SELECT>
			    </TD>
			  </TR>
			  <% Else %>
			  <TR nowrap>
					<TD valign=top>Columna en donde se colocará:</TD>
					<TD colspan=3>
						Renglón: &nbsp;
						<SELECT name=cboSectionRows style="width:60">
							<%=gsCboSectionRows%>
						</SELECT>
						&nbsp; &nbsp;&nbsp;			
						Posición:&nbsp;
						<INPUT name=txtPosition style="height:20px;width:35px;" maxlength=4 value='<%=gsColumn%>'>
						&nbsp; &nbsp;&nbsp;
						Longitud:&nbsp;
						<INPUT name=txtLength style="height:20px;width:35px;" maxlength=4 value='<%=gsLength%>'>
					</TD>
			  </TR>	  
			  <% End If %>
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Tipo de datos del elemento</b></TD>
			  </TR>
			  <TR>
					<TD valign=top nowrap>Tipo de datos que desplegará el elemento:</TD>
			    <TD colspan=3 width=430>
						<SELECT name=cboItemTypes style="width:320" onchange="return cboItemTypes_onchange()">
							<OPTION value="">-- Seleccionar el tipo de datos del elemento--</OPTION>
							<%=gsCboItemTypes%>				
						</SELECT>
			    </TD>
			  </TR>
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Información que presentará el elemento:</b></TD>
			  </TR>
				<TR id=divNone style='display:inline;'>
					<TD valign=top></TD>
			    <TD colspan=3 width=430>
						Primero se debe seleccionar de la lista anterior el tipo de datos que presentará el elemento.
			    </TD>
			  </TR>
			  <TR id=divLabel nowrap style='display:none;'>
					<TD valign=top>Etiqueta (o texto) que mostrará el elemento:</TD>
					<TD colspan=3>
						<INPUT name=txtLabel style="height:20px;width:320px;" value='<%=gsItemValue%>'>
					</TD>
			  </TR>
			  <TR id=divParameters nowrap style='display:none;'>
					<TD valign=top>Colocar el siguiente parámetro en tiempo de ejecución:</TD>
					<TD colspan=3 width=430>
						<SELECT name=cboParameters style="width:320">
							<OPTION value=''>-- Seleccionar el parámetro--</OPTION>
							<%=gsCboParameters%>
						</SELECT>						
			    </TD>					
			  </TR>			      		  
			  <TR id=divFields1 style='display:none;'>
					<TD valign=top>Elemento de información:</TD>
			    <TD colspan=3 width=430>
						<SELECT name=cboDataItems style="width:320" onchange="return cboDataItems_onchange()">
							<OPTION value=''>-- Seleccionar el elemento de información--</OPTION>
							<%=gsCboDataItems%>
						</SELECT>						
			    </TD>
			  </TR>
			  <TR id=divFields2 nowrap style='display:none;'>
					<TD valign=top>Aplicar la siguiente operación al elemento seleccionado:</TD>
					<TD colspan=3>
					  <div id=divDataOperations>
						<SELECT name=cboOperations style="width:320">
							<OPTION value=0>-- No aplicar niguna operación --</OPTION>
							<%=gsCboOperations%>
						</SELECT>
						</div>
					</TD>
			  </TR>
			  <TR id=divFields3 class=applicationTableRowDivisor>
					<TD colspan=4><b>Filtrar la información del elemento por:</b></TD>
			  </TR>			 
			  <TR id=divFields4 nowrap style='display:none;'>
					<TD colspan=4>
						<DIV STYLE="overflow:auto; float:bottom; width=100%; height=72px">
						<TABLE width=100% border=0>
							<%=gsFiltersBox%>
						</TABLE>
						</DIV>
					</TD>
			   <TR>
			   <TR id=divFields5>
					<TD valign=top>Filtro adicional:</TD>
					<TD colspan=3>
						<INPUT name=txtFilterViewer style="width:320" readonly value="<%=gsFilterViewer%>">
						<br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='height:20;width:85;' value="Editar filtro ..." onclick="pickData('dataFiltering');">						
					</TD>
			  </TR>  
			</TABLE>
		</TD>
	</TR>
	<TR>
	  <td colspan=2 nowrap align=right>
		 <INPUT type="hidden" name=txtItemId value="<%=gnItemId%>">
	   <INPUT type="hidden" name=txtPrintLayout value="<%=gsPrintLayout%>">	   
	   <INPUT type="hidden" name=txtFilterExp value="<%=gsFilterExpression%>">
	   <INPUT type="hidden" name=txtItemTag value="<%=gsItemTag%>">
		 <INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="doSubmit();">
			&nbsp; &nbsp; &nbsp;
	   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>	
	   &nbsp; &nbsp; &nbsp; &nbsp;
	  </td>
	</TR>	
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>