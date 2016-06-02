<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If

	Dim gsCboSubsidiaryLedgers, gnSubsidiaryAccountId
	Dim gnGralLedgerId, gnSubsidiaryLedgerId, gsSubsidiaryAccount, gsGralLedgerName, gsSubsAccountsTable

	gnGralLedgerId			 = Request.QueryString("gralLedgerId")
	gsSubsidiaryAccount  = Request.QueryString("subsAccount")
	If Len(Request.QueryString("id")) <> 0 Then
		gnSubsidiaryLedgerId = CLng(Request.QueryString("id"))
	Else
		gnSubsidiaryLedgerId = 0
	End If
		
	Call Main(gnGralLedgerId, gsSubsidiaryAccount)

	Sub Main(nGralLedgerId, sSubsidiaryAccount)
		Dim oGralLedgerUS, nSubsidiaryLedgerId 
		Dim nGLAccountId, nSectorId
		'***********************************************************
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
				
		'gsGralLedgerName = oGralLedgerUS.GeneralLedgerName(Session("sAppServer"), CLng(nGralLedgerId))
		gnSubsidiaryAccountId = 0
		If (Len(sSubsidiaryAccount) <> 0) Then
			nSubsidiaryLedgerId		  = oGralLedgerUS.SubsidiaryAccountLedgerId(Session("sAppServer"), CLng(nGralLedgerId), CStr(sSubsidiaryAccount))
			gnSubsidiaryAccountId   = oGralLedgerUS.SubsidiaryAccountId(Session("sAppServer"), CLng(nGralLedgerId), CStr(sSubsidiaryAccount))
			gsSubsAccountsTable     = oGralLedgerUS.TblSubsidiaryAccounts(Session("sAppServer"), CLng(nSubsidiaryLedgerId), CLng(gnSubsidiaryAccountId))
			gsCboSubsidiaryLedgers  = oGralLedgerUS.CboSubsidiaryLedgers(Session("sAppServer"), CLng(nGralLedgerId), CLng(nSubsidiaryLedgerId))
		Else
			nSubsidiaryLedgerId = gnSubsidiaryLedgerId
			gsCboSubsidiaryLedgers  = oGralLedgerUS.CboSubsidiaryLedgers(Session("sAppServer"), CLng(nGralLedgerId), CLng(nSubsidiaryLedgerId))
			gsSubsAccountsTable     = oGralLedgerUS.TblSubsidiaryAccounts(Session("sAppServer"), CLng(nSubsidiaryLedgerId))
		End If		
		
		Set oGralLedgerUS = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Selector de cuentas auxiliares</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oSubsAccountWindow = null;

window.returnValue = '';	

function openWindow(sWindowName) {
	var sURL, sPars;
	
	if (document.all.cboSubsidiaryLedgers.value == '0') {
		alert("Requiero se seleccione el mayor auxiliar");
		return false;
	}
	
	sURL = 'subsidiary_account_editor.asp?gralLedger=<%=gnGralLedgerId%>&subsLedgerId=' + document.all.cboSubsidiaryLedgers.value;
	sPars = 'height=410px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
	switch (sWindowName) {
		case 'add':
			oSubsAccountWindow = createWindow(oSubsAccountWindow, sURL, sPars);
			return false;
		case 'edit':
			sURL += '&id=' + arguments[1];
			oSubsAccountWindow = createWindow(oSubsAccountWindow, sURL, sPars);
			return false;
	}	
	return false;	
}

function updateTable(sOrderBy) {
	var obj;
	
	obj = RSExecute("../financial_accounting_scripts.asp", "TblSubsidiaryAccounts", document.all.cboSubsidiaryLedgers.value, <%=gnSubsidiaryAccountId%>, sOrderBy);
	document.all.divSubsidiaryAccountsTable.innerHTML = obj.return_value;		
	window.event.returnValue = false;	
}

function window_onload() {	
	document.all.cboSubsidiaryLedgers.focus();
}

function pickData(nItemId) {	
  var sTemp;	
	//sTemp = new String(sValue);
  //sTemp = sTemp.substring(sTemp.length - 16, sTemp.length);
	window.returnValue = nItemId;
	window.close();
	window.event.returnValue = false;
	return false;
}

function searchItem() {
	alert('Por el momento esta opción no está disponible.');
	window.event.returnValue = false;	
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()" onunload="unloadWindows(oSubsAccountWindow)">
<FORM name=frmEditor action="exec/assign_subsidiary_ledgers.asp" method="post">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Cuentas auxiliares
		</TD>		<TD colspan=3 align=right nowrap>			<a href='' onclick="return(openWindow('add'));">Agregar</a>&nbsp; | &nbsp;
			<a href='' onclick="searchItem();">Buscar</a>&nbsp;&nbsp;						<img align=absmiddle src='/empiria/images/refresh_red.gif' onclick="updateTable('');" alt="Actualizar">
			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">								</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=fullScrollMenu>			
				<TR class=fullScrollMenuHeader>
					<TD nowrap>
						Mayor auxiliar:
					</TD>
					<TD align=right>
  					<SELECT name=cboSubsidiaryLedgers style="width:200;height:18;" onchange="updateTable('');">
  						<%=gsCboSubsidiaryLedgers%>
						</SELECT>&nbsp;
					</TD>
				</TR>
			</TABLE>
			<SPAN id=divSubsidiaryAccountsTable>
			<TABLE class=applicationTable>
				<TR class=applicationTableHeader>
					<TD nowrap><A href='' onclick="return updateTable('numero_cuenta_auxiliar');">Auxiliar</A></TD>
					<TD><A href='' onclick="return updateTable('nombre_cuenta_auxiliar');">Nombre</A></TD>
				</TR>
				<% If Len(gsSubsAccountsTable) <> 0 Then %>
					<%=gsSubsAccountsTable%>
				<% Else %>
				<TR><TD>Primero se debe seleccionar un mayor auxiliar</TD></TR>
				<% End If %>
			</TABLE>
			</SPAN>		
			</TABLE>
		<TD>
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>