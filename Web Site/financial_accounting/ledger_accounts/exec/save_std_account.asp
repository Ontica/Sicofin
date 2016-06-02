<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim nScriptTimeout, gsTryAgainPage
 
	gsTryAgainPage = "../standard_account_editor.asp?" & Request.QueryString
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call SaveItem(CLng(Request.QueryString("id")))
	Server.ScriptTimeout = nScriptTimeout	
 
  Sub SaveItem(nItemId)
		Dim oStdAccount, sLogFile
		'************************
		On Error Resume Next
		Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
		sLogFile = oStdAccount.ChangeRole(Session("sAppServer"), _
																			CLng(nItemId), _															
																			CStr(Request.Form("txtStdAccountRole")), _
																			CStr(Request.Form("txtSectors")), _
																			CStr(Request.Form("txtSectorRoles")), _
																			CLng(Session("uid")))

		Set oStdAccount = Nothing
		If (Err.number = 0) Then
			Set Session("oError") = Nothing
		Else
			Set Session("oError") = Err
		End If
  End Sub
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
</head>
<body>
<% If Session("oError") Is Nothing Then %>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td bgColor="khaki"><b>La cuenta <b><%=Request.Form("txtStdAccountNumber")%></b> fue modificada satisfactoriamente.</b></td>
</tr>
<tr><td><br><b>¿Qué desea hacer?</b></td></tr>
<tr>
	<td>
		<a href='<%=gsTryAgainPage%>'>Editar la cuenta</a>
	</td>	
</tr>
<tr>
	<td>
		<a href="" onclick='window.close();'>Cerrar esta ventana</a>
	</td>
</tr>
</table>
<% Else %>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td bgColor="khaki"><b>Ocurrió el siguiente problema:</b></td>
</tr>
<tr>
	<td bgColor="khaki"><b><%=Session("oError").Description%></b></td>
</tr>
<tr>
	<td bgColor="khaki"><b><%=Session("oError").Source%>&nbsp;(<%="H" & Hex(Session("oError").Number)%>)</b></td>
</tr>
<tr><td><a href="" onclick='window.close();'>Cerrar esta ventana</a></td></tr>
</table>
<%	
	Set Session("oError") = Nothing
  End If 
%>
</body>
</html>