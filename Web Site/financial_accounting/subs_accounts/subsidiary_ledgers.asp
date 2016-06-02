<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsCboEntities, gsSubsidiaryLedgersTable, gnSelectedEntity
	
	Call Main()
	
	Sub Main()
		Dim oGralLedgersUS
		'*****************
		On Error Resume Next
		Set oGralLedgersUS = Server.CreateObject("AOGralLedgerUS.CServer")
		If (Len(Request.QueryString("id")) <> 0) Then			
			gnSelectedEntity = CLng(Request.QueryString("id"))
			gsCboEntities = oGralLedgersUS.CboSubsidiaryLedgerEntities(Session("sAppServer"), CLng(gnSelectedEntity))
			Select Case Request.QueryString("order")
				Case ""
					gsSubsidiaryLedgersTable = oGralLedgersUS.GetSubsidiaryLedgersHTMLTable(Session("sAppServer"), CLng(gnSelectedEntity), "")
				Case "1"
					gsSubsidiaryLedgersTable = oGralLedgersUS.GetSubsidiaryLedgersHTMLTable(Session("sAppServer"), CLng(gnSelectedEntity), "Nombre_Mayor, Nombre_Mayor_Auxiliar")
				Case "2"
					gsSubsidiaryLedgersTable = oGralLedgersUS.GetSubsidiaryLedgersHTMLTable(Session("sAppServer"), CLng(gnSelectedEntity), "Descripcion, Nombre_Mayor, Nombre_Mayor_Auxiliar")
				Case "3"
					gsSubsidiaryLedgersTable = oGralLedgersUS.GetSubsidiaryLedgersHTMLTable(Session("sAppServer"), CLng(gnSelectedEntity), "Descripcion, Nombre_Mayor, Nombre_Mayor_Auxiliar")
			End Select			
		Else
			gnSelectedEntity = 0
			gsCboEntities = oGralLedgersUS.CboSubsidiaryLedgerEntities(Session("sAppServer"))			
			gsSubsidiaryLedgersTable = ""
		End If
		Set oGralLedgersUS = Nothing
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

function callEditor(nOperation, nItemId) {
	var sPars, sURL, retValue;	
	
	sURL  = "subsidiary_account_picker.asp?gralLedgerId=" + nItemId;
	sURL += '&id=' + arguments[2];
	sPars = "resizable:0;status:0;dialogHeight:420px;dialogWidth:400px;";
	
	switch (nOperation) {  
		case 1:		//Add			
			//window.open(sURL, null, "height=280,width=460,location=0,resizable=0");
			//return false;
		case 2:		//Edit
			//window.open(sURL +  '&id=' + nItemId, null, "height=280,width=460,location=0,resizable=0");
			//return false;
	}			
	retValue = window.showModalDialog(sURL, "", sPars); 
	return false;
}

function refreshPage(nOrderId) {
	<% If (gnSelectedEntity	> 0) Then %>
    if (nOrderId == 0) {
		  window.location.href = "subsidiary_ledgers.asp?id=<%=gnSelectedEntity%>";
	  } else {	
		  window.location.href = "subsidiary_ledgers.asp?id=<%=gnSelectedEntity%>" + '&order=' + nOrderId;
	  }
	<% Else %>
    if (nOrderId == 0) {
		  window.location.href = "subsidiary_ledgers.asp";
	  } else {
		  window.location.href = "subsidiary_ledgers.asp" + '?order=' + nOrderId;
	  }
	<% End If %>
	return false;
}

function cboEntities_onchange() {
	if (window.document.all.cboEntities.value != 0) {
		window.location.href = "subsidiary_ledgers.asp?id=" + window.document.all.cboEntities.value;
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Mayores auxiliares
		</TD>		<TD colspan=3 align=right nowrap>
			Ver los mayores auxiliares de: &nbsp;
			<SELECT name=cboEntities style='width:280;' onchange="return cboEntities_onchange()">
				<%=gsCboEntities%>
			</SELECT>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href='<%=Application("main_page")%>';" alt="Cerrar">								</TD>
	</TR>
	<TR>
	  <TD colspan=4>			
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle colspan=4>Mayores auxiliares del tipo seleccionado</TD>					
				</TR>				
				<TR class=applicationTableHeader>
					<TD><A href="" onclick="return refreshPage(1);">Contabilidad a la que pertence</A></TD>
					<TD><A href="" onclick="return refreshPage(2);">Nombre del mayor auxiliar</A></TD>
					<TD><A href="" onclick="return refreshPage(3);">Descripción del mayor auxiliar</A></TD>	  
					<TD>&nbsp;</TD>
				</TR>
				<% If (gnSelectedEntity > 0) Then %>				
					<% If Len(gsSubsidiaryLedgersTable) <> 0 Then %>
						<%=gsSubsidiaryLedgersTable%>
					<% Else %>
						<TR><TD colspan=5 align=center>No hay mayores auxiliares definidos en esta categoría.</TD></TR>
					<% End If %>
				<% End If %>
			</TABLE>			
		</TD>
	</TR>	
</TABLE>
</BODY>
</HTML>