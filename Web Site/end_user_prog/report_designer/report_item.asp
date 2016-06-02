<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsCell, gnReportItemId, gsName, gsCboOperators, gsCboGroups, gsCboPrintCond, gsCboSheets, gsMark, gsGroupsTable
		
	Call Main()

	Sub Main()
		Dim oReportMgr, oRecordset, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'****************************************************************
		gnReportItemId   = CLng(Request.QueryString("id"))
		Set oReportMgr   = Server.CreateObject("AOReportsMgr.CManager")
		If CLng(gnReportItemId) <> 0 Then 		
			Set oRecordset  = oReportMgr.RuleRS(Session("sAppServer"), CLng(gnReportItemId))
			gsName					= oRecordset("schemaName")
			gsCell					= oRecordset("schemaCell")
			gsCboGroups	    = oReportMgr.CboGroups(Session("sAppServer"), 6)
			gsCboOperators	= oReportMgr.CboOperators(Session("sAppServer"), "+")
			gsCboPrintCond	= oReportMgr.CboPrintConditions(Session("sAppServer"), CLng(oRecordset("printCondition")))
			gsCboGroups			= oReportMgr.CboPrintConditions(Session("sAppServer"), CLng(oRecordset("printCondition")))
			gsCboSheets     = "<OPTION value='Hoja1'>Hoja1</OPTION>"
			gsGroupsTable		= oReportMgr.GroupsTable(Session("sAppServer"), CLng(gnReportItemId))
			gsMark					= oRecordset("schemaMark")
		Else 
			gsName					= ""
			gsCboGroups	    = oReportMgr.CboGroups(Session("sAppServer"), 6)
			gsCboOperators	= oReportMgr.CboOperators(Session("sAppServer"), "+")
			gsCboPrintCond	= oReportMgr.CboPrintConditions(Session("sAppServer"), 1)
			gsCboSheets     = "<OPTION value='Hoja1'>Hoja1</OPTION>"
			gsGroupsTable		= oReportMgr.GroupsTable(Session("sAppServer"), 0)
		End If
		Set oReportMgr   = Nothing
		Set oRecordset  = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Esquema del reporte</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("./exec/server_scripts.asp","IsNumeric", sNumber, nDecimals);
	return obj.return_value;
}

function addGroup() {
	var oTable = document.all.groupsTable;
	var oRow, oCell;
	
	if (document.all.txtFactor.value == '') {
		alert('Requiero el factor que se aplicará a la integración.');
		document.all.txtFactor.focus();
		return false;
	}	
	if (!isNumeric(document.all.txtFactor.value, 6)) {
		alert('No entiendo el factor de integración proporcionado.');
		document.all.txtFactor.focus();
		return false;
	}
	oRow = oTable.insertRow();
	oCell = oRow.insertCell();
	oCell.innerHTML = '<INPUT type="radio" name=optRuleId value=' + getRowData() + '>';
	oCell = oRow.insertCell();
	oCell.innerText = document.all.cboGroups.options(document.all.cboGroups.selectedIndex).text
	oCell = oRow.insertCell();
	oCell.innerText = document.all.txtFactor.value;
	oCell = oRow.insertCell();
	oCell.innerText = document.all.cboOperators.value;
	return false;
}

function delGroup() {
	var i;
	var oRows = document.all.groupsTable.tBodies[0].rows;
	for (i = 0; i < oRows.length; i++) {
		if (document.all['optRuleId'].length != null) {		
			if (document.all.optRuleId[i].checked) {
				alert(document.all.optRuleId[i].value);
				document.all.groupsTable.tBodies[0].deleteRow(i); 
				break;
			}
		} else {
			if (document.all.optRuleId.checked) {
				document.all.groupsTable.tBodies[0].deleteRow(i); 
				break;
			}
		}
	}	
	return false;
}

function getRowData() {
	return (document.all.cboGroups.value + '|' + document.all.txtFactor.value + '|' + document.all.cboOperators.value);
}

function validate() {
	if (document.all.txtDescription.value == '') {
		alert('Requiero el nombre del elemento.');
		document.all.txtDescription.focus();
		return false;		
	}
	if (document.all.txtCell.value == '') {
		alert('Requiero la celda en donde se ubicará el elemento.');
		document.all.txtCell.focus();
		return false;		
	}
	if (document.all.groupsTable.tBodies[0].rows == 0) {
		alert('El elemento debe tener al menos una agrupación que lo integre.');
		document.all.txtCell.focus();
		return false;		
	}
	return true;
}

function sendInfo() {
	if (validate()) {
		document.frmEditor.action = './exec/save_report_item.asp';
		document.frmEditor.submit();
	}	
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Agregar elemento
		</TD>
		<TD align=right nowrap>
			<img align=absMiddle src="/empiria/images/help_red.gif" onclick='notAvailable();' alt="Ayuda">	
			<img align=absMiddle src="/empiria/images/invisible.gif">
			<img align=absMiddle src="/empiria/images/close_red.gif" onclick='window.close();' alt="Cerrar">
		</TD>
	</TR>
  <TR>
		<TD colspan=2 nowrap>
			<FORM name=frmEditor method='post' target='_top'>
			<TABLE class=applicationTable cellpadding=1>
				<TR>
				  <TD valign=top>Nombre:</TD>
				  <TD valign=top>
						<TEXTAREA name=txtDescription ROWS=2 style="width:300px"><%=gsName%></TEXTAREA>
					</TD>
				</TR>
				<TR>
				  <TD valign=top>Posición:</TD>
				  <TD valign=top>
						Hoja: &nbsp;
						<SELECT name=cboSheets style="WIDTH: 100px">
						  <%=gsCboSheets%>
						</SELECT>
						Celda: &nbsp;
						<INPUT name=txtCell value="<%=gsCell%>" style="width:85px">&nbsp;(e.g., A4, E42, BC12)					
					</TD>
				</TR>
				<TR>
				  <TD>Imprimir:</TD>
				  <TD>
						<SELECT name=cboPrintCond style="WIDTH: 265px">
							<%=gsCboPrintCond%>
						</SELECT>
				  </TD>
				</TR>
				<TR>
					<TD><INPUT name=txtReportItemId type=hidden value=<%=gnReportItemId%>></TD>
					<TD>
						<TABLE class=fullScrollMenu>			
							<TR class=fullScrollMenuHeader>
								<TD class=fullScrollMenuTitle nowrap>
									Integración del elemento
								</TD>
							</TR>
						</TABLE>
						Agrupación:
						<SELECT name=cboGroups style="WIDTH: 265px">
						  <%=gsCboGroups%>
						</SELECT>
						&nbsp; &nbsp; <a href='' onclick='return(addGroup());'>Agregar</a>
						<br>
						Factor: &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT name=txtFactor value="1" style="width:75px">
						&nbsp; &nbsp;
						Operación:			
						<SELECT name=cboOperators style="WIDTH: 118px">
						  <%=gsCboOperators%>
						</SELECT>
						&nbsp; &nbsp; <a href='' onclick='return(delGroup());'>Eliminar</a>
						<SPAN id=pendingPostingsTable STYLE="overflow:auto; float:bottom; width=99%; height=140px">
						<TABLE id=groupsTable class=applicationTable>
							<THEAD>
								<TR class=applicationTableHeader>
									<TD>&nbsp;</TD>
									<TD nowrap>Agrupación</TD>
									<TD nowrap>Factor</TD>
									<TD nowrap>Op</TD>
								</TR>
							</THEAD>
							<TBODY>
							<%=gsGroupsTable%>
							</TBODY>
						</TABLE>
						</SPAN>						
					</TD>
				</TR>
			  <TR>
				  <TD colspan=2 align=right nowrap>
						<INPUT class=cmdSubmit name=cmdOk type=button value="Aceptar" style="WIDTH: 60px" onclick='sendInfo();'>&nbsp;&nbsp;&nbsp;
						<INPUT class=cmdSubmit name=cmdCancel type=button value="Cancelar" style="WIDTH: 60px" onclick='window.close();'>
						&nbsp; &nbsp; &nbsp;
					</TD>		
				</TR>
			</TABLE>
			</FORM>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>