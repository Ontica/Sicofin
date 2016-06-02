<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnRuleDefId, gsRuleDefName, gsRulesTable, gnRuleDefType, nScriptTimeout, gsRuleDefHeader
	Dim gsTackedWindows, gnSelectedColumn

	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
	
	Sub Main()
		Dim oRuleDef, oRecordset
		'************************
		'On Error Resume Next
		
		Call SetGlobalValues()
		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gnRuleDefId	 = Request.QueryString("id")
		gsRuleDefName = oRuleDef.RuleDefName(Session("sAppServer"), CLng(gnRuleDefId))
		gnRuleDefType = oRuleDef.RuleDefType(Session("sAppServer"), CLng(gnRuleDefId))
		gsRuleDefHeader = oRuleDef.Header(Session("sAppServer"), CLng(gnRuleDefType), CLng(gnSelectedColumn))
		Select Case Request.QueryString("order")
			Case ""
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "")
			Case "1"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "fecha_afectacion")
			Case "2"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "fecha_registro")
			Case "3"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "tipo_transaccion, fecha_afectacion")
			Case "4"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "tipo_poliza, fecha_afectacion")
			Case "5"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "concepto_transaccion, fecha_afectacion")
			Case "6"
				gsRulesTable = oRuleDef.RulesTable(Session("sAppServer"), CLng(gnRuleDefId), "concepto_transaccion, fecha_afectacion, nombre_autorizada_por")
		End Select		
		Set oRuleDef = Nothing
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
<TITLE>Base de conocimiento contable</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/rules.js"></script>
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oUngroupedRulesWindow = null;

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

function insertSectionInfo(oSection, nObjectId) {
  var oRow, i, j;
	var obj, aRows, aCols, sTemp;

  obj = RSExecute("../financial_accounting_scripts.asp", "RuleItemsSection", <%=gnRuleDefId%>, nObjectId);  
  sTemp = obj.return_value;
  if (sTemp == '') {
    return false;
   }
	aRows = sTemp.split("¦");
  for (i = 0; i < aRows.length; i++) {
		aCols = aRows[i].split("|");
		oRow = oSection.insertRow();
		for (j = 0; j < aCols.length; j++) {
			oRow.insertCell().innerHTML = aCols[j];
		}
  }	
}

function callEditor(nOperation, nItemId) {
	var sURL, sOpt;
	
  switch (nOperation) {
    case 1:		//Edit rule
      <% If (gnRuleDefType = 1) Then %> 
			sURL = "edit_voucher_rule.asp?id=" + nItemId;
	  	window.open(sURL, null, "height=440,width=450,location=0,resizable=0");
	  	<% ElseIf (gnRuleDefType = 2) Then %>
			sURL = "edit_rule.asp?id=" + nItemId;
	  	window.open(sURL, null, "height=440,width=450,location=0,resizable=0");	  	
	  	<% ElseIf (gnRuleDefType = 3) Then %>
			sURL = "edit_posting_rule.asp?id=" + nItemId;
	  	window.open(sURL, null, "height=440,width=450,location=0,resizable=0");	  	
	  	<% End If %>
			return false;
		case 2:	  //Edit rule group
			<% If (gnRuleDefType = 1) Then %> 
			sURL = 'rule_voucher_group_options.asp?ruleDefId=<%=gnRuleDefId%>&id=' + nItemId;
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");
			<% ElseIf (gnRuleDefType = 2) Then %>
			sURL = 'rule_group_options.asp?ruleDefId=<%=gnRuleDefId%>&id=' + nItemId;
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			<% ElseIf (gnRuleDefType = 3) Then %> 
			sURL = 'rule_posting_group_options.asp?ruleDefId=<%=gnRuleDefId%>&id=' + nItemId;
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			<% End If %>
			return false;
		case 3:		// Pendiente
			<% If (gnRuleDefType = 1) Then %> 
				sURL = 'add_voucher_rule.asp?ruleDefId=<%=gnRuleDefId%>&id=' + nItemId;
				window.open(sURL, null, "height=440,width=450,location=0,resizable=0");
			<% ElseIf (gnRuleDefType = 2) Then %>
			
			<% ElseIf (gnRuleDefType = 3) Then %>
				sURL = 'add_posting_rule.asp?ruleDefId=<%=gnRuleDefId%>';
				window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			<% End If %>
			return false;					
		case 4:   //Add group
			<% If (gnRuleDefType = 1) Then %> 
			sURL = 'add_voucher_rule_group.asp?id=0&ruleDefId=<%=gnRuleDefId%>&derivated=false';
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");
			<% ElseIf (gnRuleDefType = 2) Then %>
			sURL = 'add_rule_group.asp?id=0&ruleDefId=<%=gnRuleDefId%>&derivated=false';
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			<% ElseIf (gnRuleDefType = 3) Then %>
			sURL = 'add_posting_rule_group.asp?id=0&ruleDefId=<%=gnRuleDefId%>&derivated=false';
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			<% End If %>
			return false;
		case 5:   //Edit group reference			
			sURL = 'edit_group_reference.asp?id=' + nItemId + '&ruleDefId=<%=gnRuleDefId%>';
			window.open(sURL, null, "height=440,width=450,location=0,resizable=0");			
			return false;
		case 6:	  //Ungrouped accounts
			sURL = 'ungrouped_accounts.asp?id=<%=gnRuleDefId%>';
			sOpt = 'height=400,width=550,location=0,resizable=0,scrollbars=0';
			if (oUngroupedRulesWindow == null || oUngroupedRulesWindow.closed) {
				oUngroupedRulesWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oUngroupedRulesWindow.focus();
			}
			return false;
	}
	return false;
}

function window_onunload() {
	if (oUngroupedRulesWindow != null && !oUngroupedRulesWindow.closed) {
		oUngroupedRulesWindow.close();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload='window_onunload();'>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Definición de reglas contabilizadoras
		</TD>
	  <TD colspan=3 align=right nowrap>	  	  
			<% If (Len(gsRulesTable) = 0) AND (gnRuleDefType <> 0) Then %>
				<A href="" onclick='callEditor(4,0);return false'>Insertar agrupación</A>&nbsp; | &nbsp;
			<% End If %>
			<A href="rules_def.asp">Base de conocimiento contable</A>&nbsp; &nbsp; 
			<img align=absbottom src='/empiria/images/refresh_red.gif' onclick='window.location.href="rules.asp?id=<%=gnRuleDefId%>"' alt="Refrescar">
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
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						<%=gsRuleDefName%>
					</TD>		
					<TD nowrap align=right>
						<% If (gnRuleDefType = 2) Then %>
						<A href='' onclick='return(callEditor(6,0));'>Cuentas pendientes de incorporar</A>&nbsp; | &nbsp;
						<A href='' onclick='return(notAvailable());'>Cuentas duplicadas</A>&nbsp; | &nbsp;
						<% End If %>
						<A href="" onclick='return notAvailable();'> Buscar</A>&nbsp; | &nbsp;
						<A href='' onclick='return(notAvailable());'>Imprimir</A>&nbsp; &nbsp;
						<img align=absbottom src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>				
				<% If (gnRuleDefType = 0) Then %>				  
					<THEAD>
						<TR>
							<TD colspan=6 align=center>
								<b>Esta regla no acepta agrupaciones de cuentas.</b>
							</TD>
						</TR>
					</THEAD>
				<% End If %>
				<% If Len(gsRulesTable) <> 0 Then %>					
					<THEAD>
						<TR class=applicationTableHeader valign=center>
							<%=gsRuleDefHeader%>
						</TR>
					</THEAD>
					<%=gsRulesTable%>
				<% Else %>
					<THEAD>
						<TR>
							<TD colspan=6 align=center>
								<b>No hay cuentas definidas para esta regla.</b>
							</TD>
						</TR>
					</THEAD>	
				<% End If %>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<FORM name=frmSend method=post>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
