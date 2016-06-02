<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If	
	
	Dim gsReturnPage, gsCancelPage
	Dim gsErrNumber, gsErrSource, gsErrDescription

	Select Case CLng(Request.Form("cboTransactionTypes"))
		Case 38
			'Compraventa de moneda extranjera
			gsReturnPage = "../transacciones/compraventa_divisas.asp"
		Case 39
			'Compraventa de dólares
			gsReturnPage = "../transacciones/compraventa_dolares.asp"
		Case 40
			'Concentración fiduciaria
			gsReturnPage = "../transacciones/concentracion_fiduciaria.asp"
		Case 41
			'Traspaso de deficientes o remanentes por fideicomiso
			gsReturnPage = "../traspaso_remanentes.asp"
		Case 42
			'Cancelación de cuentas de resultados
			gsReturnPage = "../cancelacion_resultados.asp"
		Case Else
			gsReturnPage = "../../main.asp"
		End Select
		gsCancelPage = Application("main_page")  
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
  <% If Len(gsErrSource) = 0 Then %>
	window.document.all.frmSend.submit();
	<% End If %>
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<% If Len(gsErrSource) = 0 Then %>
	<FORM name=frmSend action="<%=gsReturnPage%>" method="post">		
		<INPUT type="hidden" name=txtDescription value="<%=Request.Form("txtDescription")%>">		
	</FORM>
<% Else %>
<TABLE align=center border=1 cellPadding=1 cellSpacing=1 width="60%">  
  <TR>
    <TD colspan=2 align=center>Tengo un problema</TD>
  </TR>
  <TR>
    <TD colspan=2 align=center>
    <%=gsErrSource%>
    </TD>
	</TR>
  <TR>
    <TD colspan=2 align=center>
    Fuente: <%=gsErrSource%> &nbsp; Número: <%=gsErrNumber%>
    </TD>
  </TR>
  <TR>
    <TD align=right><INPUT type="button" value="Reintentar" name=cmdReturn LANGUAGE=javascript onclick="window.location.href ="<%=gsReturnPage%>""></TD>
    <TD><INPUT type="button" value="Cancelar" name=cmdCancel LANGUAGE=javascript onclick="window.location.href ="<%=gsCancelPage%>""></TD>
  </TR>
</TABLE>
<% End If %>
</BODY>
</HTML>