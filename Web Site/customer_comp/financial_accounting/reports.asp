<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
	  Response.Redirect Application("exit_page")
	End If
	
	Dim gsUserReportsTable
	
	Call Main()
		
	Sub Main()
		Dim oReports
		'*************
		On Error Resume Next
		Set oReports = Server.CreateObject("SCFFixedReports.CReports")
		gsUserReportsTable = oReports.ReportsTable(Session("sAppServer"), Session("uid"))
		Set oReports = Nothing				
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("../programs/exception.asp")
		End If
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function callEditor(nItemId) {
	if (nItemId != 2) {
		window.location.href = "set_parameters.asp?id=" + nItemId;
		return false;
	} else {
		window.location.href = "set_parameters.asp?id=" + nItemId;
		return false;
	}	
}

function refreshPage(nOrderId) {
  if (nOrderId == 0) {
		window.location.href = "reports.asp";
	} else {	
		window.location.href = "reports.asp" + '?order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Reportes contables fijos
		</TD>
	  <TD align=right nowrap>
			<img align=middle src='/empiria/images/invisible8.gif'>
			<img align=middle src='/empiria/images/invisible8.gif'>			<img align=middle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=middle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=middle src='/empiria/images/invisible.gif'>
			<img align=middle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">
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
						<A href="balances.asp">Balanzas de comprobación</A>
						&nbsp; &nbsp; &nbsp;
						<A href="other_reports.asp">Estados financieros</A>
						&nbsp; &nbsp; &nbsp;
						<A href="../balances/balance_explorer.asp">Explorador de saldos</A>
						&nbsp; &nbsp; &nbsp;
						<A href="../balances/balance_explorer.asp">Explorador de pólizas</A>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD colspan=4>			
			<TABLE class=applicationTable>
				<THEAD>
					<TR class=fullScrollMenuHeader>
						<TD class=fullScrollMenuTitle colspan=2>Lista de reportes fijos del sistema de contabilidad financiera</TD>
					</TR>						
					<TR class=applicationTableHeader>
					  <TD nowrap><A href='' onclick='return refreshPage(1);'>Nombre</A></TD>
					  <TD nowrap align=center><A href='' onclick='return refreshPage(2);'>Descripción</A></TD>
					</TR>	
				</THEAD>
				<% If Len(gsUserReportsTable) <> 0 Then %>
					<%=gsUserReportsTable%>
				<% Else %>
					<TR><TD colspan=2 align=center>El catálogo de reportes está vacío.</TD></TR>
				<% End If %>
			</TABLE>			
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>