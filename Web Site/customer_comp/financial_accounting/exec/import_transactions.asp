<%
  Option Explicit
	Response.Buffer = False	
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage, gsFilePath, gsFileName, gnImportedVouchers, gsLogFile
	Set Session("oError") = Nothing
	
	gsFilePath = Request.Form("txtPath")
	gsFileName = Request.Form("txtFileName")
	gnImportedVouchers = ImportOrabanksFile()
	
	Function ImportOrabanksFile()
		Dim iOrabanks, nScriptTimeout
		On Error Resume Next
		Set iOrabanks = Server.CreateObject("SCFIVouchersTextFile.CServer")
				
		nScriptTimeout  = Server.ScriptTimeout
		Server.ScriptTimeout = 3600
		gsLogFile = iOrabanks.Import(Session("sAppServer"), _
																 CStr(gsFilePath & gsFileName), _
																 CLng(Request.Form("cboStdAccountTypes")), _
																 CDate(Request.Form("txtElaborationDate")), _
																 CLng(Request.Form("cboVoucherTypes")), _
																 CLng(Session("uid")), gnImportedVouchers, _
																 CBool(Request.Form("chkForwardToUsers")), _
																 Not CBool(Request.Form("chkProtectPostings")), _
																 CBool(Request.Form("chkAutoGenerateSubsidiaryAccounts")))
		ImportOrabanksFile = gnImportedVouchers
		gsLogFile = iOrabanks.URLFilesPath & gsLogFile
		Server.ScriptTimeout = nScriptTimeout
  End Function
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function showRightButtonMsg() {
  var sMsg;
  
  sMsg = "Para obtener una copia del reporte en su equipo, se requiere hacer\n" +
         "clic con el botón derecho del ratón y seleccionar la opción\n" + 
         "'Guardar destino como...'\n\n" + 
         "Gracias."
	alert(sMsg);	
}

function showReportInBrowser(sFileName) {	
	window.open(sFileName, 'dummy', "menubar=yes,toolbar=yes,scrollbars=yes,status=yes,location=no");
	return true;
}

//-->
</SCRIPT>
</head>
<body bgcolor="#E7EFE7">
&nbsp;<br>
<table width="65%" border="1" align=center cellspacing="0" cellpadding="4">
<tr>
	<td bgColor="khaki"><FONT face=Arial size=3 color=maroon><b>Resultado de la importación de pólizas</b></FONT></td>
</tr>
<% If (gnImportedVouchers > 1) Then %>
<tr>
	<td bgColor="khaki"><FONT face=Arial size=2>La importación concluyó con éxito. Fueron incorporadas <b><%=gnImportedVouchers%></b> pólizas al sistema.</FONT></td>
</tr>
<% ElseIf (gnImportedVouchers = 1) Then %>
<tr>
	<td bgColor="khaki"><FONT face=Arial size=2>La importación concluyó con éxito. Se incorporó una póliza al sistema.</FONT></td>
</tr>
<% ElseIf (gnImportedVouchers < 1) Then %>
<tr>
	<td bgcolor=LightCoral><FONT face=Arial size=2>Tuve problemas al importar las pólizas del archivo <b><%=gsFileName%></b>. No se incorporó ninguna póliza al sistema.</FONT></td>
</tr>
<% End If %>
<tr><td>
	<table width="100%" border="0" cellspacing="0" cellpadding="4">
		<tr>
			<td>
				<a href="<%=gsLogFile%>" onclick="showRightButtonMsg();return false;">
					<img src="/empiria/images/download.jpg" border=0>
				</a>
			</td>	
			<td valign=middle>
				<a href="<%=gsLogFile%>" onclick="showRightButtonMsg();return false;">
					Para obtener una copia del <b>reporte de importación</b>, basta con hacer clic sobre esta liga 
					con el botón derecho del ratón y seleccionar la opción <b>'Guardar destino como...'</b>
				</a>
				<br><br>
			</td>	
		</tr>
		<tr>
			<td>
				<a href="" onclick="showReportInBrowser('<%=gsLogFile%>');return false;">
					<img src="/empiria/images/view.jpg" border=0>
				</a>
			</td>
			<td valign=middle>
				<a href="" onclick="showReportInBrowser('<%=gsLogFile%>');return false;">	
					Ver el <b>reporte de importación</b> en una página del navegador.
				</a>
				<br><br>
			</td>	
		</tr>
	</table>
</td></tr>
<% If (gnImportedVouchers <> 0) Then %>
<tr>
	<td><FONT face=Arial size=2><a href="../pages/pending_vouchers.asp">Ir a mis pólizas pendientes</a></FONT></td>
</tr>
<% End If %>
<tr>
	<td><FONT face=Arial size=2><a href="../pages/import.asp">Regresar al importador de pólizas</a></FONT></td>
</tr>
<tr>
	<td><FONT face=Arial size=2><a href="<%=Session("main_page")%>">Ir a la página principal</a></FONT></td>
</tr>
</table>
</body>
</html>