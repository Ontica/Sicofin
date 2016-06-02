<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTemplateFile, gsTitle
	Dim gnColumnId, gnReportId, gsReportName, gsReportDescription, gsTackedWindows	
	Dim gsColName, gnReportDataId, gsColDescription, gsColPosition, gsCboWorkSheets
	Dim gsReportTechnology, gsPosition, gsLength, gsPivotColumn
	Dim oReportDesigner, gsCboDataAttributes, gscboFilters, gsCboOperations, gsCboExcelColumns
	
	Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
	
	
	If Len(Request.QueryString("id")) <> 0 Then
		gnColumnId = Request.QueryString("id")
	Else
		gnColumnId = 0
	End If
	
	Call Main()
	
	Sub Main()
		Dim oRecordset, nSelectedVoucherType
		'***********************************************
		'On Error Resume Next						
		
		Call ReportValues(Request.QueryString("reportId"))		
				
		If (gnColumnId <> 0) Then			
			Set oRecordset = oReportDesigner.GetColumn(Session("sAppServer"), CLng(gnColumnId))			
			gsColName			    = oRecordset("columnName")
			gsCboDataAttributes	 = oReportDesigner.CboDataItemAttributes(Session("sAppServer"), CLng(gnReportDataId), CLng(oRecordset("columnDataId")))
			gscboFilters = oReportDesigner.CboDataRestrictions(Session("sAppServer"), CLng(gnReportDataId), CLng(Session("uid")), CLng(oRecordset("columnFilterId")))
			gsCboOperations = oReportDesigner.CboDataOperations(Session("sAppServer"), CLng(gnReportDataId), CLng(oRecordset("columnOperationId")))
			gsCboExcelColumns   = oReportDesigner.CboExcelColumns(CLng(oRecordset("columnPosition")))
			If CLng(oRecordset("isPivotColumn")) = 1 Then
				gsPivotColumn = "checked"
			End If
			gsPosition  = oRecordset("columnPosition")
			gsLength    = oRecordset("columnLength")
			If Len(gsTemplateFile) <> 0 Then
				If Not IsNull(oRecordset("columnWorkSheet")) Then
					gsCboWorkSheets = oReportDesigner.CboTemplateWorksheets(Session("sAppServer"), CLng(gnReportId), CStr(oRecordset("columnWorkSheet")))
				Else
					gsCboWorkSheets = oReportDesigner.CboTemplateWorksheets(Session("sAppServer"), CLng(gnReportId))
				End If
			Else
				gsCboWorkSheets = "<OPTION value='Hoja1'>Hoja predeterminada</OPTION>"
			End If
		Else
			gsCboDataAttributes	= oReportDesigner.CboDataItemAttributes(Session("sAppServer"), CLng(gnReportDataId))
			gsCboFilters        = oReportDesigner.CboDataRestrictions(Session("sAppServer"), CLng(gnReportDataId), CLng(Session("uid")))
			gsCboOperations     = oReportDesigner.CboDataOperations(Session("sAppServer"), CLng(gnReportDataId))
			gsCboExcelColumns   = oReportDesigner.CboExcelColumns()			
			If Len(gsTemplateFile) <> 0 Then				
				gsCboWorkSheets = oReportDesigner.CboTemplateWorksheets(Session("sAppServer"), CLng(gnReportId))				
			Else
				gsCboWorkSheets = "<OPTION value='Hoja1'>Hoja predeterminada</OPTION>"
			End If			
		End If
		gsTackedWindows = Request.Form("txtTackedWindows")		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("./exec/exception.asp")
		End If		  
	End Sub		

	Sub ReportValues(nReportId) 
		Dim oRecordset
		'************************
		gnReportId		 = nReportId
		Set oRecordset = oReportDesigner.GetReport(Session("sAppServer"), CLng(gnReportId))
		gsReportName   = oRecordset("reportName")
		If Len(oRecordset("reportName")) > 36 Then
			gsTitle	= gsReportName & "..."
		Else
			gsTitle	= gsReportName
		End If
		If oRecordset("reportStatus") = "S" Then
			gsTitle = gsTitle & " (suspendido)"
		End If
		If IsNull(oRecordset("reportTemplateFile")) Then
			gsTemplateFile = "" 
		Else
			gsTemplateFile = oRecordset("reportTemplateFile")
		End If
		gsReportTechnology       = oRecordset("reportTechnology")
		gnReportDataId = CLng(oRecordset("reportDataId"))
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
var gbSended = false;

function doSubmit() {
	var sMsg, nVoucherId;

	if (document.all.txtName.value == '') {
		alert('Requiero el título de la columna.');
		document.all.txtName.focus();
		return false;
	}
	if (document.all.cboDataItems.value == '') {
		alert('Requiero la selección de el elemento de información que se desplegará en la columna.');
		document.all.cboDataItems.focus();
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
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function openWindow(sWindowName) {
	var sURL, sPars;
	
	switch (sWindowName) {
		case 'createFilter':
			sURL = '../dictionary/edit_filter.asp?id=0&classId=<%=gnReportDataId%>';
			sPars = 'height=380px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			oBuildWindow = createWindow(oBuildWindow, sURL, sPars);
			return false;
		case 'createOperation':
			notAvailable();
/*			
			if (oReportItemWindow == null || oReportItemWindow.closed) {
				oReportItemWindow = window.open(sURL, '_blank', sPars);
			} else {
				oReportItemWindow.focus();
				oReportItemWindow.navigate(sURL);
			}
*/
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

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmSend action='./exec/save_column.asp' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
		<% If gnColumnId = 0 Then %>
			Nueva columna
		<% Else %>
			Edición de la columna
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
					<TD colspan=4><b>Posición de la columna</b></TD>
			  </TR>  
			  <% If (gsReportTechnology  = "E") Then %>
			  <TR nowrap>
					<TD valign=top>Hoja de trabajo:</TD>
					<TD colspan=3>
						&nbsp;
						<SELECT name=cboWorkSheet style="width:200">
						  <%=gsCboWorkSheets%>
						</SELECT>
						&nbsp; &nbsp;	&nbsp; &nbsp;	
						Columna:
						<SELECT name=cboExcelColumns style="width:60">
							<%=gsCboExcelColumns%>
						</SELECT>									
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>¿Es columna pivote?:</TD>
					<TD colspan=3>
						<INPUT type="checkbox" name=chkPivotColumn <%=gsPivotColumn%> value=true>
					</TD>
			  </TR>
			  <% ElseIf (gsReportTechnology  = "T") Then %>
			  <TR nowrap>
					<TD valign=top>Posición horizontal:</TD>
					<TD colspan=3>
						&nbsp;<INPUT name=txtPosition style="height:20px;width:50px;" value='<%=gsPosition%>'>
						&nbsp; &nbsp; &nbsp; &nbsp;
						Longitud:
						<INPUT name=txtLength style="height:20px;width:50px;" value='<%=gsLength%>'>
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>¿Es columna pivote?:</TD>
					<TD colspan=3>
						<INPUT type="checkbox" name=chkPivotColumn <%=gsPivotColumn%> value=true>
					</TD>
			  </TR>
			  <% End If %>			
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Información que se presentará en la columna</b></TD>
			  </TR>			
			  <TR>
					<TD valign=top nowrap>Título de la columna:</TD>
			    <TD colspan=3 width=430>
						<INPUT name=txtName style="height:20px;width:400px;" value='<%=gsColName%>'>						
						<br>&nbsp;	
			    </TD>	    
			  </TR>			
			  <TR>
					<TD valign=top>Información a presentar en la columna:</TD>
			    <TD colspan=3 width=430>
						<SELECT name=cboDataItems style="width:400">
							<OPTION value=''>-- Seleccionar el elemento de información que se presentará en la columna--</OPTION>
							<OPTION value=0>(Dejar vacía)</OPTION>
							<%=gsCboDataAttributes%>
						</SELECT>
						<br>&nbsp;
			    </TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>Filtrar la información de la columna mediante:</TD>
					<TD colspan=3>
						<SELECT name=cboFilters style="width:400">
							<OPTION value=0>-- No aplicar ningún filtro --</OPTION>
							<%=gscboFilters%>
						</SELECT>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='width:100' value="Crear filtro ..." onclick="openWindow('createFilter');">
						<br>&nbsp;
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>Aplicar la siguiente operación sobre los elementos de la columna:</TD>
					<TD colspan=3>
						<SELECT name=cboOperations style="width:400">
							<OPTION value=0>-- No aplicar niguna operación --</OPTION>
							<%=gsCboOperations%>
						</SELECT>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdRestriction style='width:100' value="Crear operación ..." onclick="openWindow('createOperation');">
						<br>&nbsp;
					</TD>
			  </TR>			  
			</TABLE>
		</TD>
	</TR>
	<TR>
	  <td colspan=4 nowrap align=right>
	   <INPUT type="hidden" name=txtColumnId value="<%=gnColumnId%>">
		 <INPUT type="hidden" name=txtReportId value="<%=gnReportId%>">
		 <INPUT TYPE=hidden name=txtTackedWindows>						
		 <INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="doSubmit();">						
			&nbsp; &nbsp;
	   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
	   &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
	  </td>
	</TR>	
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>