<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnGralLedgerId, gnCategoryId, gsGralLedgerAccountsTable, gsCboGLCategories, gsCboGralLedgers
	
	
	Call Main()
	
	Sub Main()
		Dim oGralLedgerUS
		'****************************
		On Error Resume Next
		gnCategoryId = CLng(Request.QueryString("categoryId"))
		gnGralLedgerId = CLng(Request.QueryString("id"))		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")		
		gsCboGLCategories = oGralLedgerUS.CboGralLedgerCategories(Session("sAppServer"), 1)
		gsCboGralLedgers  = oGralLedgerUS.CboGralLedgers(Session("sAppServer"), Session("uid"), 1)
		If (gnGralLedgerId <> 0) Then
			Select Case Request.QueryString("order")
				Case ""
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "")
				Case "1"
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "Numero_Cuenta_Estandar")
				Case "2"
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "Nombre_Cuenta_Estandar, Numero_Cuenta_Estandar")
				Case "3"
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "Rol_Cuenta, Numero_Cuenta_Estandar")
				Case "4"
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "Object_Name, Numero_Cuenta_Estandar")
				Case "5"
					gsGralLedgerAccountsTable = oGralLedgerUS.GetGralLedgerAccountsHTMLTable(Session("sAppServer"), CLng(gnGralLedgerId), "Naturaleza DESC, Numero_Cuenta_Estandar")
			End Select
		End If
		Set oGralLedgerUS = Nothing
		'If (Err.number <> 0) Then
		'	Session("nErrNumber") = "&H" & Hex(Err.number)
		'	Session("sErrSource") = Err.source
		'	Session("sErrDescription") = Err.description			
		'	Session("sErrPage") = Request.ServerVariables("URL")		  
		 ' Response.Redirect("/empiria/central/e|xceptions/exception.asp")
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

function updateCboGralLedgers() {

}

function refreshPage(nOrderId) {
	var sURL = 'general_ledger_accounts.asp?categoryId=<%=gnCategoryId%>&id=<%=gnGralLedgerId%>';
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
<BODY SCROLL=NO>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=82px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="98%">
	<TR>
		<TD colspan= 3 rowspan=3 valign=top><FONT face=Arial size=3 color=maroon><STRONG>Visor de contabilidades</STRONG></TD>
	</TR>
	<TR>
	  <TD colspan=3 valign=top align=right nowrap>
			Tipo de contabilidad:
			<SELECT name=cboGLCategories style="WIDTH: 520px" onchange='return updateCboGralLedgers()'> 
				<%=gsCboGLCategories%>
      </SELECT>	  				
		</TD>
	</TR>
	<TR>
	  <TD colspan=3 valign=top align=right nowrap>
			Contabilidad
			<SELECT name=cboGralLedgers style="WIDTH: 520px" onchange='return updateCboGralLedgers()'> 
				<%=gsCboGralLedgers%>
      </SELECT>	  				
		</TD>
	</TR>	
	<TR><TD colspan=6 align=right><A href="" onclick="window.history.back();">Cerrar</A></TD></TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=85%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="98%">
<% If Len(gsGralLedgerAccountsTable) <> 0 Then %>
	<A href="#SCROLLABLE_DIV_TOP"></A>
	<TR>
	  <TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Número de cuenta</b></FONT></A></TD>
	  <TD align=center><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Nombre</b></FONT></A></TD>
	  <TD nowrap><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Rol</b></FONT></A></TD>
	  <TD><A href="" onclick="return refreshPage(4);"><FONT color=maroon><b>Tipo de cuenta</b></FONT></A></TD>
	  <TD nowrap><A href="" onclick="return refreshPage(5);"><FONT color=maroon><b>Naturaleza</b></FONT></A></TD>
	</TR>
	<%=gsGralLedgerAccountsTable%>
	<TR>
	  <TD nowrap colspan=5 align=right><A href="#SCROLLABLE_DIV_TOP">Subir</A></TD>
	</TR>	
<% Else %>
	<TR><TD colspan=5 align=center>Este mayor no tiene cuentas registradas.</TD></TR>
<% End If %>
</TABLE>
<BR>&nbsp;
</DIV>
</BODY>
</HTML>