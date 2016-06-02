<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnSubsidiaryLedgerId, gnEntityId, gsSubsidiaryAccountsTable, gsSubsidiaryLedgerName
	
	Call Main()
	
	Sub Main()
		Dim oGralLedgerUS
		'****************
		On Error Resume Next
		gnSubsidiaryLedgerId = CLng(Request.QueryString("id"))
		gnEntityId = CLng(Request.QueryString("entityId"))
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		gsSubsidiaryLedgerName = oGralLedgerUS.SubsidiaryLedgerName(Session("sAppServer"), CLng(gnSubsidiaryLedgerId))
		Select Case Request.QueryString("order")
			Case ""
				gsSubsidiaryAccountsTable = oGralLedgerUS.GetSubsidiaryAccountsHTMLTable(Session("sAppServer"), CLng(gnSubsidiaryLedgerId), "")
			Case "1"
				gsSubsidiaryAccountsTable = oGralLedgerUS.GetSubsidiaryAccountsHTMLTable(Session("sAppServer"), CLng(gnSubsidiaryLedgerId), "Numero_Cuenta_Auxiliar")
			Case "2"
				gsSubsidiaryAccountsTable = oGralLedgerUS.GetSubsidiaryAccountsHTMLTable(Session("sAppServer"), CLng(gnSubsidiaryLedgerId), "Nombre_Cuenta_Auxiliar, Numero_Cuenta_Auxiliar")
			Case "3"
				gsSubsidiaryAccountsTable = oGralLedgerUS.GetSubsidiaryAccountsHTMLTable(Session("sAppServer"), CLng(gnSubsidiaryLedgerId), "Descripcion, Numero_Cuenta_Auxiliar")
		End Select		
		Set oGralLedgerUS = Nothing
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
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function callEditor(nOperation, nItemId) {
	var sURL = 'subsidiary_account_editor.asp?subsidiaryLedgerId=<%=gnSubsidiaryLedgerId%>';
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

function refreshPage(nOrderId) {
	var sURL = 'subsidiary_accounts.asp?entityId=<%=gnEntityId%>&id=<%=gnSubsidiaryLedgerId%>';
  if (nOrderId == 0) {
		window.location.href = sURL;
	} else {	
		window.location.href = sURL + '&order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=82px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="90%">
	<TR>
		<TD colspan= 3 rowspan=2><FONT face=Arial size=3 color=maroon><STRONG>Cuentas auxiliares en:&nbsp;</STRONG><FONT size=1><%=gsSubsidiaryLedgerName%></FONT></FONT></TD>
	</TR>
	<TR>
	  <TD colspan=3 valign=top align=right nowrap>
			<INPUT type="button" name=cmdAddItem value="Agregar cuenta auxiliar" style="WIDTH:160px" LANGUAGE=javascript onclick="return callEditor(1,0);">&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdRefresh value="Actualizar vista" style="WIDTH:120px" LANGUAGE=javascript onclick="window.location.href=window.location.href;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdReturn value="Cerrar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href = 'subsidiary_ledgers.asp?id=<%=gnEntityId%>';">
		</TD>
	</TR>
	<TR><TD>&nbsp;</TD></TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=85%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="90%">
<% If Len(gsSubsidiaryAccountsTable) <> 0 Then %>	
	<TR>
	  <TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Número de cuenta</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Nombre</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Descripción</b></FONT></A></TD>	  
	</TR>
	<%=gsSubsidiaryAccountsTable%>
<% Else %>
	<TR><TD colspan=4 align=center>El catálogo de cuentas auxiliares está vacío.</TD></TR>
<% End If %>
</TABLE>
</DIV>
</BODY>
</HTML>