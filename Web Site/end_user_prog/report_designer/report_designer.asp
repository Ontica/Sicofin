<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim oReportRows, gsReportName, gsReportTemplateFile, gsCboWorkSheets, gsWorksheet
	Dim gsReportSections, gsReportColumns 'gsReportRows
	Dim gnReportId, gsTackedWindows, gnSelectedColumn, gsReportTechnology
		
	Call Main()

	Sub Main()
		Dim oReportDesigner, oRecordset, nSelectedVoucherType
		'***********************************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")			
		If (Len(Request.QueryString("id")) <> 0) Then
			gnReportId = CLng(Request.QueryString("id"))
			Set oRecordset     = oReportDesigner.GetReport(Session("sAppServer"), CLng(gnReportId))
			gsReportName	     = oRecordset("reportName")
			gsReportTechnology = oRecordset("reportTechnology")
			If Not IsNull(oRecordset("reportTemplateFile")) Then
				gsReportTemplateFile = oReportDesigner.FilesPath & "templates/" & oRecordset("reportTemplateFile")
			End If			
			If (gsReportTechnology = "E") Then
				If Len(Request.QueryString("worksheet")) <> 0 Then
					gsWorksheet = Request.QueryString("worksheet")
				Else
					gsWorksheet = oReportDesigner.DefaultWorksheet(Session("sAppServer"), CLng(gnReportId))
				End If				
				gsCboWorkSheets  = oReportDesigner.CboWorksheets(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))
				gsReportColumns  = oReportDesigner.ColumnsHeader(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))
				gsReportSections = oReportDesigner.SectionsTable(Session("sAppServer"), CLng(gnReportId), CStr(gsWorksheet))				
			Else
				gsReportColumns  = oReportDesigner.ColumnsHeader(Session("sAppServer"), CLng(gnReportId))
				gsReportSections = oReportDesigner.SectionsTable(Session("sAppServer"), CLng(gnReportId))
			End If
			
			'gsReportColumns = oReportDesigner.ColumnsHeader(Session("sAppServer"), CLng(gnReportId))
			'gsReportRows   = oReportDesigner.RowsTable(Session("sAppServer"), CLng(gnReportId)						
			oRecordset.Close
		Else
			gnReportId = 0
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

%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>La Aldea Ontica® / Diseñador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var gbSended = false;

var oReportItemWindow = null, oColumnEditorWindow = null, oSectionEditorWindow = null;

function createWindow(oWindow, sURL, sPars) {
	var oTempWindow;

	if (oWindow == null || oWindow.closed) {
		oTempWindow = window.open(sURL, '_blank', sPars);
		return oTempWindow;
	} else {
		oWindow.focus();
		oWindow.navigate(sURL);
		return oWindow;
	}
}

function openWindow(sWindowName) {
	var sURL, sPars;
	
	switch (sWindowName) {
		case 'createSection':
			sURL  = 'create_section.asp?id=' + arguments[1] + '&reportId=<%=gnReportId%>';
			<% If (gsReportTechnology = "E") Then %>				
				sURL += '&worksheet=' + document.all.cboWorkSheets.value;
			<% End If %>
			sPars = 'height=480px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oSectionEditorWindow = createWindow(oSectionEditorWindow, sURL, sPars);
			return false;
		case 'addItem':		
			sURL  = 'edit_item.asp?rowId=' + arguments[1] + '&col=' + arguments[2];
			sPars = 'height=420px,width=560px,resizable=no,scrollbars=no,status=no,location=no';
			oReportItemWindow = createWindow(oReportItemWindow, sURL, sPars);
			window.close();
			return false;		
		case 'editItem':
			sURL  = 'edit_item.asp?id=' + arguments[1];
			sPars = 'height=420px,width=560px,resizable=no,scrollbars=no,status=no,location=no';
			oReportItemWindow = createWindow(oReportItemWindow, sURL, sPars);
			window.close();
			return false;
		case 'rowOptions':
			sURL = 'row_options.asp?id=' + arguments[1];
			sPars = 'height=280px,width=380px,resizable=no,scrollbars=no,status=no,location=no';
			oReportItemWindow = createWindow(oReportItemWindow, sURL, sPars);
			return false;
		case 'sectionOptions':
			sURL  = 'section_options.asp?id=' + arguments[1];
			sPars = 'height=200px,width=300px,resizable=no,scrollbars=no,status=no,location=no';
			oSectionEditorWindow = createWindow(oSectionEditorWindow, sURL, sPars);
			return false;
		case 'editColumn':
			sURL = 'column_editor.asp?id=' + arguments[1] + '&reportId=<%=gnReportId%>';
			sPars = 'height=460px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oColumnEditorWindow = createWindow(oColumnEditorWindow, sURL, sPars);
			return false;
/*			
			if (oReportItemWindow == null || oReportItemWindow.closed) {
				oReportItemWindow = window.open(sURL, '_blank', sPars);
			} else {
				oReportItemWindow.focus();
				oReportItemWindow.navigate(sURL);
			}
*/			
	}	
	return false;	
}

function cboWorkSheets_onchange() {
	var sURL;
	
	sURL  = 'report_designer.asp?id=<%=gnReportId%>&worksheet=';
	sURL += document.all.cboWorkSheets.value;
	window.location.href = sURL;
}

function unloadWindows(oWindow) {
	if (oReportItemWindow != null && !oReportItemWindow.closed) {
		oReportItemWindow.close();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY onunload="unloadWindows(oReportItemWindow)">
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			<%=gsReportName%>			
		</TD>
	  <TD align=right nowrap>
			<A align=absmiddle href='designed_reports.asp'>Otros reportes diseñados</A>
			<img align=absmiddle src='/empiria/images/invisible.gif'>			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = 'designed_reports.asp';" alt="Cerrar y regresar a la página principal">
		</TD>
	</TR>
	<TR id=divTasksOptions style='display:none'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Tareas
					</TD>
					<TD nowrap align=left>
						<A href="" onclick="return(notAvailable());">Lista de tareas</A>
						&nbsp; | &nbsp
						<A href="" onclick="return(notAvailable());">Mi lista de tareas pendientes</A>
					</TD>
					<TD nowrap align=right>
					  <img id=cmdTasksOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divTasksOptions, this)' alt='Fijar la ventana'>					
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='/empiria/images/invisible.gif'>					  
						<img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divTasksOptions, document.all.cmdTasksOptionsTack)" alt='Cerrar'>
					</TD>
				</TR>
				<TR>
					<TD nowrap colspan=3>
						<A href="voucher_explorer.asp">Explorador de pólizas</A>
						&nbsp;&nbsp;&nbsp;&nbsp;						
						<A href="transaction_selector.asp">Consulta de saldos</A>			
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="transaction_selector.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="transaction_selector.asp">Reportes</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=fullScrollMenu>			
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle nowrap>
						Vista en modo de diseño
					</TD>
					<TD align=right nowrap>
					<% If (gsReportTechnology = "E") Then %>
						Hoja de trabajo:&nbsp;
						<SELECT name=cboWorkSheets style="width:180" onchange="return cboWorkSheets_onchange()">
						  <%=gsCboWorkSheets%>
						</SELECT>
						&nbsp; &nbsp; 				
					<% End If %>
						<A align=absmiddle href='' onclick="return(openWindow('createSection', 0));">Agregar sección</A>
	  			<% If Len(gsReportTemplateFile) <> 0 Then %>	  				&nbsp; | &nbsp;
						<A align=absmiddle href='<%=gsReportTemplateFile%>' target='_blank'>Archivo prediseñado</A>											<% End If %>
						&nbsp;
						<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick="window.location.href=window.location.href;" alt="Refrescar">
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
				<TR class=applicationTableHeader>
					<%=gsReportColumns%>
				</TR>				
				<%=gsReportSections%>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>				
</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>