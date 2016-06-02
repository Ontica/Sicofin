<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnTimeout, sFileName
	
	gnTimeout = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
	Call Main()
	Server.ScriptTimeout = gnTimeout
	
	Sub Main()
		'On Error Resume Next
		
		Select Case CLng(Request.QueryString("id"))
			Case 1
				sFileName = ExportPyCTransactions(Request.Form("txtTargetAppServer"))
			Case 2
				sFileName = ExportPyCBalances(Request.Form("txtTargetAppServer"), Date())
			Case 3
				Call ExportSigro(Request.Form("txtSigroDate"))
			Case 4
				Call ExportBalances(Request.Form("txtSigroDate"), Request.Form("txtFromDate"), Request.Form("txtToDate"))
		End Select
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("../programs/exception.asp")
		End If		
	End Sub
	
	Function ExportPyCTransactions(sTargetAppServer)
		Dim iGemPyC
		'*********************************************
		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")
		ExportPyCTransactions = iGemPyC.ExportTransactions(Session("sAppServer"), CStr(sTargetAppServer), 1)
		Set iGemPyC = Nothing
	End Function
	
	Function ExportPyCBalances(sTargetAppServer, dToDate)
		Dim iGemPyC
		'**************************************************
		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")
		ExportPyCBalances = iGemPyC.ExportBalances(Session("sAppServer"), CStr(sTargetAppServer), CDate(dToDate))
		Set iGemPyC = Nothing
	End Function
			
	Sub ExportSigro(dToDate)
		Dim iSigro
		'*********************
		Set iSigro = Server.CreateObject("SCFISigro.CServer")
		iSigro.CreateGralBalance Session("sAppServer"), CLng(nStdAccountTypeId), CDate(dFromDate), CDate(dToDate)
		Set iSigro = Nothing
	End Sub

	Sub ExportBalances(nStdAccountTypeId, dFromDate, dToDate)
		Dim iSigro
		'*****************************************************
		Set iSigro = Server.CreateObject("SCFISigro.CServer")
		iSigro.CreateGralBalance Session("sAppServer"), CLng(nStdAccountTypeId), CDate(dFromDate), CDate(dToDate)
		Set iSigro = Nothing
	End Sub
		
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
</head>
<body>
<br>
<table width="65%" border="0" align=center cellspacing="0" cellpadding="4">
<tr>
	<% If CLng(Request.QueryString("id")) = 1 Then %>
		<td bgColor="khaki" colspan=2><FONT face=Arial size=3 color=maroon><b>Exportación de pólizas hacia el sistema PyC</b></FONT></td>
	<% ElseIf CLng(Request.QueryString("id")) = 2 Then %>
		<td bgColor="khaki" colspan=2><FONT face=Arial size=3 color=maroon><b>Exportación de saldos hacia el sistema PyC</b></FONT></td>	
	<% ElseIf CLng(Request.QueryString("id")) = 3 Then %>
		<td bgColor="khaki" colspan=2><FONT face=Arial size=3 color=maroon><b>Exportación de saldos para el Sigro (reportes regulatorios)</b></FONT></td>		
	<% ElseIf CLng(Request.QueryString("id")) = 4 Then %>
		<td bgColor="khaki" colspan=2><FONT face=Arial size=3 color=maroon><b>Exportación de saldos (Subdirección de Informática)</b></FONT></td>	
	<% End If %>
</tr>
<tr>
	<td colspan=2><FONT face=Arial size=2><b>La operación concluyó satisfactoriamente.</b></FONT><br><br></td>
</tr>
<tr>
	<td colspan=2><FONT face=Arial size=2><a href="../pages/export.asp">Regresar al exportador</a></FONT></td>
</tr>
<tr>
	<td colspan=2><FONT face=Arial size=2><a href="<%=Session("main_page")%>">Ir a la página principal</a></FONT></td>
</tr>
</table>
</body>
</HTML>