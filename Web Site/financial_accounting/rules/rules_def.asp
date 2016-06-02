<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsRulesDefinitionHeader, gsRulesDefinitionBody, gsTackedwindows, gnSelectedColumn
	
	Call Main()
	
	Sub Main()
		Dim oRulesMgr
		'*************
		On Error Resume Next
		
		Call SetGlobalValues()
		
		Set oRulesMgr = Server.CreateObject("EFARulesMgrBS.CManager")
		gsRulesDefinitionHeader = oRulesMgr.RuleDefinitionsHeader(Session("sAppServer"), CLng(gnSelectedColumn))
		gsRulesDefinitionBody = oRulesMgr.RuleDefinitionsBody(Session("sAppServer"), Session("uid"), CLng(gnSelectedColumn))				
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If
	End Sub
	
  Sub SetGlobalValues()
		If Request.Form.Count Then
		 gnSelectedColumn = Request.Form("txtSelectedColumn")
		 gsTackedWindows  = Request.Form("txtTackedWindows")
		Else
		 gnSelectedColumn = 1
		 gsTackedWindows  = Request.Form("txtTackedWindows")
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

function orderBy(nSelectedColumn) {
  var nTemp = document.all.txtSelectedColumn.value;

  if (nTemp == '') {
		document.all.txtSelectedColumn.value = nSelectedColumn;
	} else {
		if ((nTemp == nSelectedColumn) || (Math.abs(nSelectedColumn) == nSelectedColumn)) {
			if (nTemp == nSelectedColumn) {
				document.all.txtSelectedColumn.value = (-1 * nSelectedColumn);
			} else {
				document.all.txtSelectedColumn.value = nSelectedColumn;
			}
		} else {
			document.all.txtSelectedColumn.value = nSelectedColumn;
		}
  }
  document.all.frmSend.action = '';
  document.all.frmSend.submit();
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Experto contable
		</TD>
	  <TD colspan=3 align=right nowrap>
			<A href="add_rule_def.asp">Crear nueva regla</A>&nbsp; &nbsp; &nbsp;
			<img align=absbottom src='/empiria/images/refresh_red.gif' onclick='document.all.frmSend.submit();' alt="Refrescar">
			<img align=absbottom src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absbottom src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absbottom src='/empiria/images/invisible.gif'>
			<img align=absbottom src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">
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
						<A href="../reports/designed_reports.asp">Diseñador de reportes</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../balances/balance_explorer.asp">Explorador de saldos</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../reports/balances.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;	
						<A href="../reports/other_reports.asp">Reportes contables</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable> 
				<THEAD>
					<TR class=fullScrollMenuHeader valign=center>
						<TD class=fullScrollMenuTitle colspan=5>Reglas definidas en el experto contable</TD>
					</TR>									
					<TR class=applicationTableHeader valign=center>
						<%=gsRulesDefinitionHeader%>
					</TR>				
				</THEAD>
				<% If (Len(gsRulesDefinitionBody) <> 0) Then %>
					<%=gsRulesDefinitionBody%>
				<% Else %>
				<TBODY>
					<TR>
						<TD colspan=5 align=middle><b>La base de conocimiento contable está vacía.</b></TD>
					</TR>
				</TBODY>
			<% End If %>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<FORM name=frmSend method=post>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows value='<%=gsTackedWindows%>'>
</FORM>
</BODY>
</HTML>