<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTemplateFile, gsTitle
	Dim gnReportId, gsReportName, gsSectionName, gnSectionsCount, gsReportTechnology, gsWorksheet
	Dim gsReportDataClassId, gsReportDataSubClassId, gsCboSectionTypes, gbHasFixedSections
		
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oRecordset
		'******************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		gnReportId = Request.QueryString("reportId")
		Set oRecordset         = oReportDesigner.GetReport(Session("sAppServer"), CLng(gnReportId))
		gsReportName           = oRecordset("reportName")
		gsReportDataClassId    = oRecordset("reportDataClassId")
		gsReportDataSubClassId = oRecordset("reportDataSubClassId")
		If Len(oRecordset("reportName")) > 48 Then
			gsTitle	= gsReportName & "..."
		Else
			gsTitle	= gsReportName
		End If		
		gsReportTechnology = oRecordset("reportTechnology")		
		If Len(Request.QueryString("worksheet")) <> 0 Then
			gsWorksheet = Request.QueryString("worksheet")
		Else
			gsWorksheet = ""
		End If
		gnSectionsCount     = oReportDesigner.SectionsCount(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))
		gsCboSectionTypes   = oReportDesigner.CboSectionsTypes(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))
		gbHasFixedSections  = oReportDesigner.HasFixedSections(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))
		Set oReportDesigner = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("./exec/exception.asp")
		End If		  
	End Sub
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Dise�ador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var gbSended = false;

function oData() {
	var dataExpression; 
	var dataViewer; 
}

function pickData(sDataSource) {
	var sURL, sPars;
	
	sURL  = '../dictionary/attributes_picker.asp?';
	sURL += 'classId=<%=gsReportDataClassId%>&subclassId=<%=gsReportDataSubClassId%>';
	sURL += '&pickerType=' + sDataSource;
	sPars = "dialogHeight:285px;dialogWidth:420px;resizable:no;scroll:no;status:no;help:no;";	
	switch (sDataSource) {
		case 'dataGrouping':		
			oData.dataExpression = document.all.txtDataGroupingExp.value;
			oData.dataViewer     = document.all.txtDataGrouping.value;			
			if (window.showModalDialog(sURL, oData, sPars)) {				
				document.all.txtDataGroupingExp.value = oData.dataExpression;
				document.all.txtDataGrouping.value = oData.dataViewer;
			}
			return false;	
		case 'dataOrdering':
			oData.dataExpression = document.all.txtDataOrderExp.value;
			oData.dataViewer     = document.all.txtDataOrder.value;
			if (window.showModalDialog(sURL, oData, sPars)) {
				document.all.txtDataOrderExp.value = oData.dataExpression;
				document.all.txtDataOrder.value = oData.dataViewer;
			}			
			return false;
		case 'dataFiltering':
			oData.dataExpression = document.all.txtDataFilterExp.value;
			oData.dataViewer     = document.all.txtDataFilter.value;
			if (window.showModalDialog(sURL, oData, sPars)) {				
				document.all.txtDataFilterExp.value = oData.dataExpression;
				document.all.txtDataFilter.value = oData.dataViewer;				
			}
			return false;
	}	
}

function isNumeric(nValue) {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "IsNumeric", nValue, 0);
	return(obj.return_value);
}

function traslapingRows(nFromRow, nToRow) {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "TraslapingRows", <%=gnReportId%>, nFromRow, nToRow, "<%=gsWorksheet%>");
	return(obj.return_value);	
}

function doSubmit() {
	var sMsg, nVoucherId;

	if (gbSended) {
		return false;
	}
	<% If gnSectionsCount = 0 Then %>
	if(document.all.cboParametrizationModes.value == '') {
		alert('Requiero se elija la forma en que ser� parametrizada la hoja de trabajo o reporte.');
		document.all.cboParametrizationModes.focus();
		return false;
	}
	if(document.all.cboParametrizationModes.value == 'C') {
		if (!verifyRowsNumber()) {
			return false;
		}
	}
	if(document.all.cboParametrizationModes.value == 'R') {
		if (!verifyInitAndFinalRows()) {
			return false;
		}
	}
	<% ElseIf (gnSectionsCount <> 0) AND gbHasFixedSections Then %>
	if (!verifyInitAndFinalRows()) {
		return false;
	}
	<% ElseIf (gnSectionsCount <> 0) AND (NOT gbHasFixedSections) Then %>
	if (!verifyRowsNumber()) {
		return false;
	}
	<% End If %>	
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function verifyRowsNumber() {
	if (document.all.txtSectionRows.value == '') {
		alert('Requiero el n�mero de renglones que tendr� la secci�n.');
		document.all.txtSectionRows.focus();
		return false;
	}
	if (!isNumeric(document.all.txtSectionRows.value)) {
		alert('No reconozco el n�mero de renglones que tendr� la secci�n.');
		document.all.txtSectionRows.focus();
		return false;
	}	
	return true;
}	
		
function verifyInitAndFinalRows() {
	if (document.all.txtInitialRow.value == '') {
		alert('Requiero el n�mero de rengl�n en donde empezar� la secci�n.');
		document.all.txtInitialRow.focus();
		return false;
	}
	if (!isNumeric(document.all.txtInitialRow.value)) {
		alert('No reconozco el n�mero de rengl�n en donde empezar� la secci�n.');
		document.all.txtInitialRow.focus();
		return false;
	}			
	if (document.all.txtFinalRow.value == '') {
		alert('Requiero el n�mero de rengl�n en donde terminar� la secci�n.');
		document.all.txtFinalRow.focus();
		return false;			
	}
	if (!isNumeric(document.all.txtFinalRow.value)) {
		alert('No reconozco el n�mero de rengl�n en donde terminar� la secci�n.');
		document.all.txtFinalRow.focus();
		return false;
	}	
	if (Number(document.all.txtFinalRow.value) < Number(document.all.txtInitialRow.value)) {
		alert('El n�mero de rengl�n donde termina la secci�n debe ser mayor que el n�mero de rengl�n en donde empieza.');
		document.all.txtFinalRow.focus();
		return false;
	}
	if (traslapingRows(document.all.txtInitialRow.value, document.all.txtFinalRow.value)) {
		alert('Los renglones de inicio y t�rmino de secci�n se traslapan con los de otra secci�n de la hoja de trabajo o reporte.');
		document.all.txtFinalRow.focus();
		return false;
	}	
	return true;
}

function cboParametrizationModes_onchange() {
	document.all.divMsg.style.display = 'none';
	document.all.divFirstRow.style.display = 'none';
	document.all.divLastRow.style.display  = 'none';
	document.all.divNumberOfRows.style.display  = 'none';
	document.all.divDataGrouping.style.display = 'none';		
	document.all.divDataOrder.style.display  = 'none';	
	switch (document.all.cboParametrizationModes.value) {
		case '':
			document.all.divMsg.style.display = 'inline';
			break;
		case 'C':
			document.all.divNumberOfRows.style.display  = 'inline';
			document.all.divDataGrouping.style.display = 'inline';		
			document.all.divDataOrder.style.display  = 'inline';				
			break;
		case 'R':
			document.all.divFirstRow.style.display = 'inline';
			document.all.divLastRow.style.display  = 'inline';			
			break;
	}
}

function window_onload() {
<% If (gnSectionsCount <> 0) AND gbHasFixedSections Then %>	
	document.all.divFirstRow.style.display = 'inline';
	document.all.divLastRow.style.display  = 'inline';
	document.all.divNumberOfRows.style.display = 'none';
	document.all.divDataGrouping.style.display = 'none';		
	document.all.divDataOrder.style.display  = 'none';	
<% ElseIf (gnSectionsCount <> 0) AND (Not gbHasFixedSections) Then %>
	document.all.divFirstRow.style.display = 'none';
	document.all.divLastRow.style.display  = 'none';
	document.all.divNumberOfRows.style.display  = 'inline';	
	document.all.divDataGrouping.style.display = 'inline';		
	document.all.divDataOrder.style.display  = 'inline';	
<% End If %>
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox LANGUAGE=javascript onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>		
			Nueva secci�n
		</TD>
	  <TD align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;
		</TD>
	</TR>
	<FORM name=frmSend action='./exec/create_section.asp' method=post>
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
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Identificaci�n de la secci�n</b></TD>
				</TR>
				<TR>					
					<TD nowrap>Nombre de la secci�n:</TD>
			    <TD colspan=3 width=100%>
						<INPUT type="text" name=txtName value='<%=gsSectionName%>' style='width:280;'>
						<br>
						<% If (gsReportTechnology = "E") Then %>
							(En la hoja de trabajo <b><%=gsWorksheet%></b>)
						<% End If %>
			    </TD>
			  </TR>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Tipo de la secci�n</b></TD>
				</TR>
				<TR>
					<TD nowrap>Tipo de la secci�n:</TD>
			    <TD colspan=3 width=100%>						
						<SELECT name=cboSectionTypes style="width:280">
							<%=gsCboSectionTypes%>
						</SELECT>
							<% If gnSectionsCount = 0 Then %>
							<br><br>
							Debido a que todav�a no hay secciones definidas, el �nico tipo de secci�n que 
							se puede agregar es una secci�n de detalle y adem�s deber� contestarse por �nica vez
							a la siguiente pregunta:
							<% End If %>							
			    </TD>
			  </TR>	
			  <% If (gnSectionsCount = 0) Then %>
				<TR>
					<% If (gsReportTechnology = "E") Then %>
						<TD nowrap>�En qu� forma ser� parametrizada la hoja de trabajo?:</TD>
					<% Else %>
					<TD nowrap>�En qu� forma ser� parametrizado el reporte?:</TD>
					<% End If %>
			    <TD colspan=3 width=100%>						
						<SELECT name=cboParametrizationModes style="width:280" onchange="return cboParametrizationModes_onchange()">
							<OPTION value="">-- Forma de parametrizaci�n --</OPTION>
							<OPTION value="C">Por columnas</OPTION>
							<OPTION value="R">Por celdas o posiciones fijas</OPTION>							
						</SELECT>	
			    </TD>
			  </TR>			  
			  <% End If %>
				<TR class=applicationTableRowDivisor nowrap>
					<% If (gsReportTechnology = "E") Then %>
						<TD valign=top colspan=4><b>Tama�o y posici�n de la secci�n en la hoja de trabajo</b></TD>
					<% Else %>
						<TD valign=top colspan=4><b>Tama�o y posici�n de la secci�n en el reporte</b></TD>
					<% End If %>					
				</TR>
				<% If (gnSectionsCount = 0) Then %>
				<TR id=divMsg style='display:inline;height=59;'>
					<% If (gsReportTechnology = "E") Then %>
					<TD colspan=4>Primero debe elegirse la forma en que ser� parametrizada la hoja de trabajo.</TD>
					<% Else %>
					<TD colspan=4>Primero debe elegirse la forma en que ser� parametrizado el reporte.</TD>
			    <% End If %>
			    </TD>
			  </TR>
			  <% End If %>			  
				<TR id=divFirstRow style='display:none;'>
					<TD nowrap>La secci�n empieza en el rengl�n: </TD>
			    <TD colspan=3 width=100%>
						<INPUT name=txtInitialRow maxlength=4 style='height:20px;width:75px;'>
			    </TD>
			  </TR>
				<TR id=divLastRow style='display:none;'>
					<TD nowrap>La secci�n termina en el rengl�n: </TD>
			    <TD colspan=3 width=100%>
						<INPUT name=txtFinalRow maxlength=4 style='height:20px;width:75px;'>
			    </TD>
			  </TR>
			  <TR id=divNumberOfRows style='display:none;height=59;'>
					<TD nowrap>N�mero de renglones que tendr� la secci�n: &nbsp; &nbsp;</TD>
			    <TD colspan=3>						
						<INPUT name=txtSectionRows maxlength=4 style='height:20px;width:75px;'>						
			    </TD>
			  </TR>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Agrupaci�n, ordenamiento y filtrado de la informaci�n</b></TD>
				</TR>
				<TR id=divDataGrouping style='display:none;'>
					<TD nowrap>Agrupar los elementos de la secci�n por:</TD>
			    <TD colspan=3 nowrap>
						<INPUT type=hidden name=txtDataGroupingExp>
						<INPUT name=txtDataGrouping maxlength=4 style='width:210;' readonly>
						<INPUT class=cmdSubmit type=button name=cmdDataGrouping style='height:20;width:65;' value="Editar ..." onclick="pickData('dataGrouping');">
			    </TD>
			  </TR>
				<TR id=divDataOrder style='display:none;'>
					<TD nowrap>Ordenar los elementos de la secci�n por:</TD>
			    <TD colspan=3 nowrap>
						<INPUT type=hidden name=txtDataOrderExp>
						<INPUT name=txtDataOrder maxlength=4 style='width:210;' readonly>
						<INPUT class=cmdSubmit type=button name=cmdDataOrdering style='height:20;width:65;' value="Editar ..." onclick="pickData('dataOrdering');">
			    </TD>
			  </TR>			  
				<TR>
					<TD nowrap>Filtrar los elementos de la secci�n por:</TD>
			    <TD colspan=3 nowrap>
						<INPUT type=hidden name=txtDataFilterExp>
						<INPUT name=txtDataFilter maxlength=4 style='width:210;' readonly>
						<INPUT class=cmdSubmit type=button name=cmdDataFiltering style='height:20;width:65;' value="Editar ..." onclick="pickData('dataFiltering');">
			    </TD>
			  </TR>			  
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 align=right>			
			<INPUT type="hidden" name=txtReportId value="<%=gnReportId%>">
			<INPUT type="hidden" name=txtWorkSheet value="<%=gsWorksheet%>">
			 <INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="doSubmit();">
			 &nbsp; &nbsp;
			<INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
			&nbsp; &nbsp;	&nbsp; &nbsp;&nbsp; 	
		</TD>
	</TR>
	</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>