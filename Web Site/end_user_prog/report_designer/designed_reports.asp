<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReportsList, gsTackedWindows
	
	Call Main()
	
	Sub Main()
		Dim oReportsDesigner
		'********************************
		'On Error Resume Next
			
		Set oReportsDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		gsReportsList = oReportsDesigner.ReportsTable(Session("sAppServer"), CLng(Session("uid")))
		Set oReportsDesigner = Nothing
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
<TITLE>La Aldea Ontica® / Diseñador de reportes</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oEditWindow = null;

function openWindow(sWindowName) {
	var sURL, sPars;
	
	switch (sWindowName) {
		case 'createReport':
			sURL  = "report_editor.asp";
			sPars = 'height=460px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oEditWindow = createWindow(oEditWindow, sURL, sPars);
			return false;
		case 'editReport':
			sURL = 'report_editor.asp?id=' + arguments[1];
			sPars = 'height=460px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oEditWindow = createWindow(oEditWindow, sURL, sPars);
			return false;
		case 'paramReport':			
			sURL = 'report_designer.asp?id=' + arguments[1];
			window.location.href = sURL;			
			//oEditWindow = createWindow(oEditWindow, sURL, sPars);
			return false;	
		case 'shareReport':
		  notAvailable();
		  return false;
			sURL = 'report_permissions.asp?id=' + arguments[1];			
			sPars = 'height=460px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oEditWindow = createWindow(oEditWindow, sURL, sPars);
			return false;
	}	
	return false;	
}

function deleteReport(nReportId, sReportName) {
	if (confirm('¿Elimino el reporte "' + sReportName + '" del sistema?')) {
	
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload="unloadWindows(oEditWindow)">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Diseñador de reportes
		</TD>
		<TD colspan=3 align=right nowrap>
			<A href='../engine/selector.asp'>Ir al generador de reportes</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>
			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
	</TR>
	<TR id=divTasksOptions style='display:none'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class='fullScrollMenuHeader'>
					<TD class='fullScrollMenuTitle' nowrap>
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
						<A href="../../contabilidad/reports/balances.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../contabilidad/reports/financial_statements.asp">Estados financieros</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../contabilidad/reports/other_reports.asp">Otros reportes</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../contabilidad/balances/balance_explorer.asp">Explorador de saldos</A>
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
					<TD class=fullScrollMenuTitle nowrap width=315>
						Lista de reportes diseñados
					</TD>
					<TD colspan=1 align=right nowrap>	
						<A href='' onclick="return(openWindow('createReport'))">Diseñar un nuevo reporte</A>
						&nbsp; | &nbsp;
						<A href='' onclick="return(notAvailable())">Buscar reporte</A>
						<img align=absmiddle src='/empiria/images/invisible4.gif'>
						<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick="window.location.href='designed_reports.asp'" alt='Actualizar ventana'>
						<img align=absmiddle src='/empiria/images/invisible.gif'>							
						<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
				<TR class=applicationTableHeader>
					<TD nowrap>Nombre</TD>
					<TD nowrap>Origen de la información</TD>
					<TD nowrap>Filtrado por</TD>
          <TD nowrap>Tecnología</TD>
          <TD nowrap>Actualizado el</TD>
          <TD nowrap>Estado</TD>
          <TD nowrap>&nbsp;</TD>
         </TR>
				<%=gsReportsList%>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>