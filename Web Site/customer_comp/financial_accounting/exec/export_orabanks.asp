<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

    If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnTimeout, gsFileName1, gsFileName2, gsLogFile1, gsLogFile2
	Dim gnVouchersCount1, gnVouchersCount2
	
	gnTimeout = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
	Call Main()
	Server.ScriptTimeout = gnTimeout
	
	Sub Main()
		Dim iVouchersTextFile
		'*******************
		On Error Resume Next		
		Set iVouchersTextFile = Server.CreateObject("SCFIVouchersTextFile.CServer")
		'Archivo contabilidad Bancaria
		gsFileName1 = iVouchersTextFile.Export(Session("sAppServer"), 1, CDate(Request.Form("txtOrabanksDate")), _
																	 CStr(Request.Form("txtFromHour")), CDate(Request.Form("txtToHour")), _
																	 gnVouchersCount1, gsLogFile1)
		If Len(gsFileName1) <> 0 Then
			gsFileName1 = iVouchersTextFile.URLFilesPath & gsFileName1
		End If		
		gnVouchersCount1 = CLng(gnVouchersCount1)
		gsLogFile1 = iVouchersTextFile.URLFilesPath & gsLogFile1

		gsFileName2 = iVouchersTextFile.Export(Session("sAppServer"), 2, CDate(Request.Form("txtOrabanksDate")), _
																   CStr(Request.Form("txtFromHour")), CDate(Request.Form("txtToHour")), _
																	 gnVouchersCount2, gsLogFile2)
		If Len(gsFileName2) <> 0 Then
			gsFileName2 = iVouchersTextFile.URLFilesPath & gsFileName2
		End If
		gnVouchersCount2 = CLng(gnVouchersCount2)
		gsLogFile2 = iVouchersTextFile.URLFilesPath & gsLogFile2
		
		Set iVouchersTextFile = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("../programs/exception.asp")
		End If
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
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
<body>
&nbsp;<br>
<table width="65%" border="0" align=center cellspacing="0" cellpadding="4">
<tr>
	<td bgColor="khaki" colspan=2><FONT face=Arial size=3 color=maroon><b>Resultado de la exportación de pólizas</b></FONT></td>
</tr>
<tr>
	<td colspan=2 bgcolor=LightCoral><font size=2><b>Contabilidad bancaria</b></font></td>
</tr>
<tr>
	<td colspan=2 bgcolor=LightCoral align=right nowrap><font size=2>Pólizas exportadas: <b><%=gnVouchersCount1%></b></font></td>
</tr>
<% If (gnVouchersCount1 <> 0) Then %>
<tr>
	<td>
		<a href="<%=gsFileName1%>" onclick="showRightButtonMsg();return false;">
			<img src="/empiria/images/download.jpg" border=0>
		</a>
	</td>	
	<td valign=middle>
		<a href="<%=gsFileName1%>" onclick="showRightButtonMsg();return false;">
			Para obtener el archivo con las pólizas exportadas de la <b>contabilidad bancaria</b>, basta con hacer 
			clic sobre esta liga con el botón derecho del ratón y seleccionar la opción <b>'Guardar destino como...'</b>
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="" onclick="showReportInBrowser('<%=gsLogFile1%>');return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="" onclick="showReportInBrowser('<%=gsLogFile1%>');return false;">	
			Archivo con los detalles del proceso de exportación de las pólizas de la <b>contabilidad bancaria</b>.
		</a>
		<br><br>
	</td>
</tr>
<% Else %>
<tr>
	<td colspan=2><font size=2>No encontré pólizas en la <b>contabilidad bancaria</b> elaboradas el día <b><%=Request.Form("txtDate")%></b>.</font></td>	
</tr>
<% End If %>
<tr>
	<td colspan=2 bgcolor=LightCoral><font size=2><b>Contabilidad fiduciaria</b></font></td>
</tr>
<tr>
	<td colspan=2 bgcolor=LightCoral align=right nowrap><font size=2>Pólizas exportadas: <b><%=gnVouchersCount2%></b></font></td>
</tr>
<% If (gnVouchersCount2 <> 0) Then %>
<tr>
	<td>
		<a href="<%=gsFileName2%>" onclick="showRightButtonMsg();return false;">
			<img src="/empiria/images/download.jpg" border=0>
		</a>
	</td>	
	<td valign=middle>
		<a href="<%=gsFileName2%>" onclick="showRightButtonMsg();return false;">
			Para obtener el archivo con las pólizas exportadas de la <b>contabilidad fiduciaria</b>, basta con hacer 
			clic sobre esta liga con el botón derecho del ratón y seleccionar la opción <b>'Guardar destino como...'</b>
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="" onclick="showReportInBrowser('<%=gsLogFile2%>');return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="" onclick="showReportInBrowser('<%=gsLogFile2%>');return false;">	
			Archivo con los detalles del proceso de exportación de las pólizas de la <b>contabilidad fiduciaria</b>.
		</a>
		<br><br>
	</td>
</tr>
<% Else %>
<tr>
	<td colspan=2><font size=2>No encontré pólizas en la <b>contabilidad fiduciaria</b> elaboradas el día <b><%=Request.Form("txtDate")%></b>.</font></td>	
</tr>
<% End If %>
<tr>
	<td colspan=2 bgColor="khaki">&nbsp;</td>
</tr>
<tr>
	<td colspan=2><FONT face=Arial size=2><a href="../export.asp">Regresar al exportador de pólizas</a></FONT></td>
</tr>
<tr>
	<td colspan=2><FONT face=Arial size=2><a href="<%=Session("main_page")%>">Ir a la página principal</a></FONT></td>
</tr>
</table>
</body>
</HTML>