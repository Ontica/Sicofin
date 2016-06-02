<%
  Option Explicit
  Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1		
		
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	Response.Buffer = False
	
	On Error Resume Next
	
	Dim gsImportedFilesTable	
		
	Call Main()
		
	Sub Main()
		Dim iGemPyC
		'**********
		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")				
		gsImportedFilesTable = iGemPyC.ImportedFilesTable()
		Set iGemPyC = Nothing
	End Sub
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<BASE target="_blank">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<TITLE>Bandeja de archivos</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function gotoDownload(sPage) {
	window.location.href = '../../../pages/download.asp?type=2&file=' + sPage;
}

function gotoError() {
	window.location.href = '../../../pages/exception.asp';
}

function window_onload() {
	document.all.clicker.click();
}

function window_onbeforeunload() {
	var sMsg;
	
	sMsg = "El proceso de importación continuará aunque se cierre esta ventana.\n\n¿Cierro esta ventana?";	
	if (!confirm(sMsg)) {
		event.cancelBubble = true;
		return false;
	}
	return true;
}

//-->
</SCRIPT>
</head>
<body class=bdyDialogBox>
<table class=fullScrollMenu>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap>
			<font size=2><b>Archivos con los resultados de las importaciones GEM</b></font>
		</TD>
		<TD colspan=3 align=right nowrap>			
			<img align=absmiddle src='/empiria/images/invisible.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='return(notAvailable());' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">								</TD>
	</TR>
</table>
<table class=fullScrollMenu>
	<TR class=fullScrollMenuHeader>
		<TD nowrap>
			<b>Selección:</b> &nbsp; &nbsp;
			<a href='' onclick='return(notAvailable());'>Enviar a mi bandeja</a>
			&nbsp; | &nbsp; 
			<a href='' onclick='return(notAvailable());'>Enviar a otro participante</a>
			&nbsp; | &nbsp; 
			<a href='' onclick='return(notAvailable());'>Eliminar</a>			
		<TD align=right nowrap>
			<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick='window.location.href=window.location.href;' alt='Actualizar ventana'>		
		</TD>
	</TR>
</table>
<div style="overflow:auto;float:bottom;width=100%; height=247px">
<table class=applicationTable border=0>
	<tr class=applicationTableHeader>
		<td><INPUT type="checkbox" name=chkItem></td>
		<td width=100% nowrap>Fecha en que se realizó la importación</td>
	</tr>
	<%=gsImportedFilesTable%>
</table>
</div>
</body>
</html>