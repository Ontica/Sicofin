<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsCboCategories, gsGralLedgersTable, gnSelectedCategory
	
	Dim gsTackedWindows
	
	Call Main()
	
	Sub Main()
		Dim oObject
		'*****************
		'On Error Resume Next
		Set oObject = Server.CreateObject("AOGLVoucherUS.CVoucher")
		If (Len(Request.QueryString("id")) <> 0) Then
			gnSelectedCategory = CLng(Request.QueryString("id"))			
			gsCboCategories = oObject.CboGeneralLedgerCategories(Session("sAppServer"), 0, 2, CLng(gnSelectedCategory))
		Else	
			gnSelectedCategory = 0
			gsCboCategories = oObject.CboGeneralLedgerCategories(Session("sAppServer"), 0, 2)
		End If
		Set oObject = Nothing
		
		Set oObject = Server.CreateObject("AOGralLedgerUS.CServer")
		If (gnSelectedCategory <> 0) Then						
			gsGralLedgersTable = oObject.GetGralLedgersHTMLTable(Session("sAppServer"), CLng(gnSelectedCategory))
		Else			
			gsGralLedgersTable = ""
		End If
		Set oObject = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
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

var oEditWindow = null, oAccountsWindow = null;

function cboCategories_onchange() {	
	if (document.all.cboCategories.value != 0) {
		window.location.href = "general_ledgers.asp?id=" + document.all.cboCategories.value;
	}
	return false;
}

function callEditor(sOperation, nItemId) {
	var sURL, sOpt;
					
  switch (sOperation) {
    case 'editItem':
			sURL  = 'general_ledger_editor.asp?id=' + nItemId;
			sOpt = 'height=410px,width=520px,resizable=no,scrollbars=no,status=no,location=no';
			if (oEditWindow == null || oEditWindow.closed) {
				oEditWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oEditWindow.focus();
				oEditWindow.navigate(sURL);
			}	  	
			return false;
		case 'accounts':
			sURL  = 'general_ledger_accounts.asp?id=' + nItemId;
			sOpt = 'height=360px,width=500px,resizable=no,scrollbars=no,status=no,location=no';
			if (oAccountsWindow == null || oAccountsWindow.closed) {
				oAccountsWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oAccountsWindow.focus();
				oAccountsWindow.navigate(sURL);
			}	  	
			return false;
		case 'users':
			alert("Por el momento esta opción no está disponible.\n\nGracias.");
			return false;
	}
}		

//-->
</SCRIPT>
</HEAD>
<BODY onload='showTackedWindows(Array(<%=gsTackedWindows%>));' onunload="unloadWindows(oEditWindow, oAccountsWindow)">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Contabilidades
		</TD>
		<TD nowrap>
			Grupo de contabilidades: &nbsp
			<SELECT name=cboCategories onchange="return(cboCategories_onchange());">
				<%=gsCboCategories%>
			</SELECT>			
		</TD>
		<TD colspan=2 align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>
			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
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
						Contabilidades en el grupo seleccionado
					</TD>
					<TD colspan=1 align=right nowrap>
						<% If Len(gsCboCategories) <> 0 Then %>
						<A href='' onclick="return(callEditor('editItem', 0));">Crear contabilidad</A>
						<% End If %>
						<img align=absmiddle src='/empiria/images/invisible4.gif'>
						<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick='window.location.href=window.location.href;' alt='Actualizar ventana'>
						<img align=absmiddle src='/empiria/images/invisible.gif'>							
						<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
				<TR class=applicationTableHeader>					
					<TD nowrap>Contabilidad</TD>
					<TD nowrap>Prefijo auxiliares</TD>
					<TD nowrap>Moneda base</TD>					
          <TD nowrap>Calendario</TD>          
          <TD nowrap>&nbsp;</TD>
         </TR>
         <% If (gnSelectedCategory) > 0 Then %>
					<% If Len(gsGralLedgersTable) <> 0 Then %>
						<%=gsGralLedgersTable%>
					<% Else %>
						<TR><TD colspan=5><b>No hay contabilidades definidas en el grupo de contabilidades seleccionado.</b></TD></TR>
					<% End If %>
				<% Else %>
					<TR><TD colspan=5><b>Necesito la selección de un grupo de contabilidades.</b></TD></TR>
				<% End If %>
			</TABLE>			
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>