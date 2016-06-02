<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnGralLedgerId, gnCategoryId, gsGralLedgerAccountsTable, gsGralLedgerName
	
	Call Main()
	
	Sub Main()
		Dim oGralLedgerUS
		'****************
		On Error Resume Next
		gnGralLedgerId = CLng(Request.QueryString("id"))
		gnCategoryId = CLng(Request.QueryString("categoryId"))
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")		
		
		gsGralLedgerName = oGralLedgerUS.GeneralLedgerName(Session("sAppServer"), CLng(gnGralLedgerId))
		If Len(gsGralLedgerName) > 90 Then
			gsGralLedgerName = Left(oGralLedgerUS.GeneralLedgerName(Session("sAppServer"), CLng(gnGralLedgerId)), 84)
		End If
		gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "")
		Set oGralLedgerUS = Nothing
		'If (Err.number <> 0) Then
		'	Session("nErrNumber") = "&H" & Hex(Err.number)
		'	Session("sErrSource") = Err.source
		'	Session("sErrDescription") = Err.description			
		'	Session("sErrPage") = Request.ServerVariables("URL")		  
		 ' Response.Redirect("/empiria/central/exceptions/exception.asp")
		'End If
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function callEditor(nOperation, nItemId) {
	var sURL = 'general_ledger_account_editor.asp?generalLedgerId=<%=gnGralLedgerId%>';
  switch (nOperation) {
    case 1:		//Add
			window.open(sURL, null, "height=280,width=400,location=0,resizable=0");
			return false;
    case 2:		//Edit
			window.open(sURL + '&id=' + nItemId, null, "height=280,width=400,location=0,resizable=0");
			return false;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmEditor action="" method="post">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Cuentas en la contabilidad
		</TD>
		<TD colspan=2 align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>						<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.close();" alt="Cerrar">		</TD>
	</TR>
	<TR>
		<TD colspan=3>	    
			<TABLE class=fullScrollMenu>
				<TR class=fullScrollMenuHeader>
					<TD colspan=4><%=gsGralLedgerName%></TD>
					<TD align=right>
						<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick='window.location.href=window.location.href;' alt='Actualizar ventana'>
					</TD>
				</TR>
			</TABLE>
			<% If Len(gsGralLedgerAccountsTable) <> 0 Then %>
			<DIV STYLE='overflow:auto; float:bottom; width=100%; height=300'>
			<% End If %>
			<TABLE class=applicationTable>			
				<TR class=applicationTableHeader>
				  <TD nowrap>Cuenta</TD>
				  <TD>Nombre</TD>
				  <TD>Rol</TD>
				  <TD nowrap>Tipo de cuenta</TD>
				  <TD>Naturaleza</TD>
				</TR>
				<% If Len(gsGralLedgerAccountsTable) <> 0 Then %>
				<%=gsGralLedgerAccountsTable%>
				<% Else %>
				<TR><TD colspan=5><b>Esta contabilidad no tiene cuentas asignadas.</b></TD></TR>
				<% End If %>
			</TABLE>
			<% If Len(gsGralLedgerAccountsTable) <> 0 Then %>
			</DIV>
			<% End If %>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>