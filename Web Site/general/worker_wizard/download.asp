<%
  Option Explicit     
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1		
		
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsFileName, gsTitle, gsMsgOpen, gsMsgSave, gsMsgSend
	
	Call Main()
						
	Sub Main()
		Dim gnFileType, oReportsEngine
		'*****************
		'On Error Resume Next
		'Set oReportsEngine = Server.CreateObject("AOReportsEngine.CEngine")
		gsFileName = Request.QueryString("file")
		Select Case CLng(Request.QueryString("type"))
			Case 1			'Reporte del generador de reportes
				gsTitle   = "El reporte está listo"
				gsMsgOpen = "Abrir el reporte."
				gsMsgSave = "Guardar el reporte en mi bandeja de archivos."
				gsMsgSend = "Enviar el reporte a otro participante."
			Case 2			'Importación GEM
				gsTitle   = "Terminé con la importación de las pólizas del sistema GEM."
				gsMsgOpen = "Abrir el archivo con los resultados de la importación."
				gsMsgSave = "Guardar el archivo con los resultados de la importación en mi bandeja de archivos."
				gsMsgSend = "Enviar el archivo con los resultados de la importación a otro participante."
			Case 3			'Importación de pólizas
			Case 4			'Exportación de pólizas			
			Case 5			'Balanzas de comprobación
			Case 6			'Otros reportes contables						
		End Select		
		
		'Set oReportsEngine = Nothing		
	End Sub
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<TITLE>Resultado del proceso</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function showRightButtonMsg() {
  var sMsg;
  
  sMsg = "Para obtener una copia del archivo en su equipo, se debe hacer\n" +
         "clic con el botón derecho del ratón y seleccionar la opción\n" + 
         "'Guardar destino como...'\n\n" + 
         "Gracias."
	alert(sMsg);	
}

function showReportInBrowser() {	
	window.open('<%=gsFileName%>', 'dummy', "menubar=yes,toolbar=yes,scrollbars=yes,status=yes,location=no");
	return true;
}


//-->
</SCRIPT>
</head>
<body class=bdyDialogBox>
<table class=fullScrollMenu>
	<TR class=fullScrollMenuHeader>
		<TD class=fullScrollMenuTitle nowrap colspan=3>
			<%=gsTitle%>
		</TD>
		<TD align=right nowrap>		  
			<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>			
			<img align=absmiddle src='/empiria/images/close_white.gif' onclick="window.close();">
		</TD>
	</tr>
</table>
<table class=applicationTable border=0>
	<tr>
		<td valign=top nowrap>
		  <img src='/empiria/images/invisible4.gif'>
			<a href="<%=gsFileName%>" onclick="showReportInBrowser();return false;">
				<img src="/empiria/images/download.gif" border=0>
			</a>
		</td>
		<td valign=valign=top>
			<a href="<%=gsFileName%>" onclick="showReportInBrowser();return false;">	
				<%=gsMsgOpen%>				
			</a>
			<br><br>
			(Para obtener una copia del archivo, haga clic sobre la liga con el botón derecho
			del ratón y seleccione la opción 'Guardar destino como...')
			<br>&nbsp;
		</td>
	</tr>
	<tr>
		<td valign=top nowrap>
			<img src='/empiria/images/invisible.gif'>
			<a href="<%=gsFileName%>" onclick="return(notAvailable());">
				<img src="/empiria/images/save.gif" border=0>
			</a>
		</td>
		<td valign=valign=top>			
			<a href="<%=gsFileName%>" onclick="return(notAvailable());">
				<%=gsMsgSave%>
			</a>
			<br><br>			
		</td>	
	</tr>	
	<tr>
		<td valign=top nowrap>
			<img src='/empiria/images/invisible4.gif'>
			<a href="<%=gsFileName%>" onclick="return(notAvailable());">
				<img src="/empiria/images/workflow/users.gif" border=0 width=40>
			</a>
		</td>
		<td valign=valign=top>
			<a href="<%=gsFileName%>" onclick="return(notAvailable());">
				<%=gsMsgSend%>
			</a>
			<br><br>&nbsp;
		</td>	
	</tr>		
</table>
</body>
</html>
