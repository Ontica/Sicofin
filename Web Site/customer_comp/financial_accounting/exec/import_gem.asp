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
	
	Dim gsLogFile, gnImportVouchers, gnEstimatedMinutes, gnInitialTime
	Dim sSourceAppServer, nScriptTimeout
	
	sSourceAppServer = "GemPyC"
		
	Call Main()
		
	Sub Main()
		Dim iGemPyC
		'**********
		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")		
		gnImportVouchers    = iGemPyC.GEMPendingTransactionsCount(CStr(sSourceAppServer))
		gnEstimatedMinutes  = iGemPyC.EstimatedImportTime(CStr(sSourceAppServer), CLng(gnImportVouchers))
		gnInitialTime       = Time()
		Set iGemPyC = Nothing
	End Sub	
	
	Sub ImportGEM()
		Dim iGemPyC
		'**********		
		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")
		gsLogFile = iGemPyC.ImportTransactions(Session("sAppServer"), CStr(sSourceAppServer), CLng(Session("uid")))
		gsLogFile = iGemPyC.URLFilesPath & gsLogFile
		Set iGemPyC = Nothing
  End Sub
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<TITLE>Trabajando...</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function gotoDownload(sPage) {
	window.location.href = '/empiria/general/worker_wizard/download.asp?type=2&file=' + sPage;
}

function gotoError() {
	window.location.href = '/empiria/central/exceptions/exception.asp';
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
<body rightmargin=3 leftMargin=3 topmargin=3 bottommargin=3 style='cursor:wait;background-color:white;' onload="return window_onload()">
<div id=divProcessing>
<table width=100% border=0>	
	<tr>
		<td valign=top>
			<img src="/empiria/images/central/working.gif" style="cursor:wait;">
			<table class=applicationTable>
				<tr>
					<td>
					<INPUT type="checkbox" name=chkSendTo style='cursor:hand;'>
					Al finalizar, cerrar esta ventana y enviar el reporte a mi bandeja de documentos
					<br>
					</td>
				</tr>
			</table>
		</td>
		<td width=100% nowrap valign=top>
			<table class=applicationTable height=100%>
				<tr><td width=100% colspan=2><b>Importando pólizas...<b></td></tr>
				<tr><td width=100% colspan=2>Importando las pólizas del sistema GEM hacia el sistema de contabilidad financiera.</td></tr>
				<tr><td><b>El proceso intentará importar:</b></td><td width=40%><%=gnImportVouchers%> pólizas</td></tr>				
				<% If gnEstimatedMinutes = 1 Then %>
					<tr><td><b>Tiempo estimado:</b></td><td>Un minuto.</td></tr>
				<% Else %>
					<tr><td><b>Tiempo estimado:</b></td><td><%=gnEstimatedMinutes%> minutos.</td></tr>
				<% End If %>
				<tr><td><b>La importación inició a las:</b></td><td><%=gnInitialTime%></td></tr>
				<tr><td><b>Se estima que el proceso termine a las:</b></td><td><%=DateAdd("n", gnEstimatedMinutes,gnInitialTime)%></td></tr>
			</table>
		</td>
	</tr>
</table>
</div>
<%
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600			
	Call ImportGEM()
	Server.ScriptTimeout = nScriptTimeout
	If Err.number = 0 Then
		Response.Write ("<A id=clicker onclick='gotoDownload(""" & gsLogFile & """)'></A>")
	Else			
		Session("errNumber") = Err.number
		Session("errDesc")   = Err.description
		Session("errSource") = Err.source
		Err.Clear
		Response.Write ("<A id=clicker onclick='gotoError()'></A>")
	End If	
%>
</body>
</html>