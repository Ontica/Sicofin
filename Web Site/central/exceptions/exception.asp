<%
  Option Explicit
		
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsErrNumber, gsErrSource, gsErrDescription, gsErrPage

	gsErrNumber = "&H" & Hex(Session("errNumber"))
	gsErrSource = Session("errSource")
	gsErrDescription = Session("errDesc")
	gsErrPage		= Request.ServerVariables("HTTP_REFERER")
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<TITLE>Manejador de excepciones</TITLE>
</head>
<body class=bdyDialogBox>
<table class=fullScrollMenu>
	<TR class=fullScrollMenuHeader>
		<TD class=fullScrollMenuTitle nowrap colspan=3>
			Ocurrió un problema técnico
		</TD>
		<TD align=right nowrap> 
			<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>			
			<img align=absmiddle src='/empiria/images/close_white.gif' onclick="window.close();">
		</TD>
	</tr>
</table>
<table class=applicationTable>	
<FORM name=frmSend action="./exec/send_to_support.asp" method=post>
	<tr>
		<td nowrap><b>Descripción: &nbsp;</b></td>
		<td><%=gsErrDescription%>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Origen: &nbsp;</b></td>
		<td><%=gsErrPage%></td>
	</tr>		
	<tr>
		<td nowrap><b>Componente: &nbsp;</b></td>
		<td><%=gsErrSource%></td>
	</tr>
	<tr>
		<td nowrap><b>Identificador: &nbsp;</b></td>
		<td><%=gsErrNumber%><br></td>
	</tr>
	<tr>
		<td nowrap colspan=2>
			<b>¿Cómo ocurrió el problema?:&nbsp;</b><br><br>
			<TEXTAREA rows=5 cols=73 name=txtDescription></TEXTAREA>
		</td>		
	</tr>
	<tr>
		<td nowrap colspan=2 align=right>			
			<INPUT class=cmdSubmit  name=cmdSend type="submit" value="Enviar a soporte técnico">&nbsp;&nbsp;
		</td>		
	</tr>	
	<tr>
		<td colspan=2>
			Para soporte telefónico marcar la extensión 4367 (Jaime Méndez), o bien, a la 6151 (Apolinar Riojas).<br><br>
		</td>
	</tr>
</FORM>
</table>
</body>
</html>
