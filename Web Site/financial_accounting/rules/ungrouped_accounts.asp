<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnRuleDefId, gsRuleDefName, gsUngroupedAccountsTable, gnRuleDefType, nScriptTimeout, gsUngroupedAccountsHeader
	Dim gnSelectedColumn

	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
	
	Sub Main()
		Dim oUngroupedItems, oRecordset
		'******************************
		'On Error Resume Next
		
		Call SetGlobalValues()
		
		Set oUngroupedItems = Server.CreateObject("EFARulesMgrBS.CUngroupedItems")						
		gsUngroupedAccountsHeader = oUngroupedItems.UngroupedAccountsHeader(CLng(gnSelectedColumn))
		gsUngroupedAccountsTable  = oUngroupedItems.UngroupedAccountsTbl(Session("sAppServer"), CLng(gnRuleDefId), Date)
		Set oUngroupedItems = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If
	End Sub
	
  Sub SetGlobalValues()
		Dim oRuleDef
		'******************
		gnRuleDefId	 = Request.QueryString("id")
		If Request.Form.Count Then
		 gnSelectedColumn = Request.Form("txtSelectedColumn")		 
		Else
		 gnSelectedColumn = 1		 
		End If		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gsRuleDefName = oRuleDef.RuleDefName(Session("sAppServer"), CLng(gnRuleDefId))
		gnRuleDefType = oRuleDef.RuleDefType(Session("sAppServer"), CLng(gnRuleDefId))
		Set oRuleDef = Nothing
	End Sub		
%>
<HTML>
<HEAD>
<TITLE>Base de conocimiento contable</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oGroupUngroupedRuleWindow = null;

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

function callEditor(sAccount, nSector, nCurrency) {
	var sURL, sOpt;    
		
	sURL  = 'group_ungrouped_rule.asp?ruleDefId=<%=gnRuleDefId%>&account=' + sAccount;
	sURL += '&sectorId=' + nSector + '&currencyId=' + nCurrency;
	sOpt  = 'height=510,width=450,location=0,resizable=0';
	if (oGroupUngroupedRuleWindow == null || oGroupUngroupedRuleWindow.closed) {
		oGroupUngroupedRuleWindow = window.open(sURL, '_blank', sOpt);
	} else {
		oGroupUngroupedRuleWindow.focus();
		oGroupUngroupedRuleWindow.navigate(sURL);
	}
	return false;
}

function window_onunload() {
	if (oGroupUngroupedRuleWindow != null && !oGroupUngroupedRuleWindow.closed) {
		oGroupUngroupedRuleWindow.close();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onunload='window_onunload();'>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Cuentas pendientes de incorporar
		</TD>
	  <TD colspan=3 align=right nowrap>
			<img align=absbottom src='/empiria/images/refresh_red.gif' onclick='window.location.href="ungrouped_accounts.asp?id=<%=gnRuleDefId%>"' alt="Refrescar">						<img align=absbottom src='/empiria/images/help_red.gif' onclick='notAvailable();' alt='Ayuda'>			<img align=absbottom src='/empiria/images/invisible.gif'>
			<img align=absbottom src='/empiria/images/close_red.gif' onclick='window.close();' alt='Cerrar y regresar a la página principal'>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						<%=gsRuleDefName%>
					</TD>		
					<TD nowrap align=right>						
						<A href='' onclick='return(notAvailable());'>Cuentas duplicadas</A>&nbsp; | &nbsp;
						<A href="" onclick='return notAvailable();'> Buscar</A>&nbsp; | &nbsp;
						<A href='' onclick='return(notAvailable());'>Imprimir</A>&nbsp; &nbsp;
					</TD>
				</TR>
			</TABLE>
			<DIV STYLE="overflow:auto;float:bottom;width=100%; height=335px">
			<TABLE class=applicationTable>
				<% If Len(gsUngroupedAccountsTable) <> 0 Then %>					
					<THEAD>
						<TR class=applicationTableHeader valign=center>
							<%=gsUngroupedAccountsHeader%>
						</TR>
					</THEAD>
					<%=gsUngroupedAccountsTable%>
				<% Else %>
					<THEAD>
						<TR>
							<TD colspan=6 align=center>
								<b>No hay cuentas pendientes de incorporar para esta regla.</b>
							</TD>
						</TR>
					</THEAD>	
				<% End If %>
			</TABLE>
			</DIV>
		</TD>
	</TR>
</TABLE>
<FORM name=frmSend method=post>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
</FORM>
</BODY>
</HTML>
