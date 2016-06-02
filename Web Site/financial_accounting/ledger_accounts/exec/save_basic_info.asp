<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Call SaveBasicInfo()
	Response.Redirect "../standard_account_editor.asp?id=" & Request.QueryString("id")
  
  Sub SaveBasicInfo()
		Dim oStdAccountMgr, sLogFile
		'***************************
		'On Error Resume Next
		Set oStdAccountMgr = Server.CreateObject("EFAStdActBS.CStdAccount")
		oStdAccountMgr.SaveBasicInfo Session("sAppServer"), _
																 CLng(Request.Form("txtItemId")), _
																 CStr(Request.Form("txtStdAccountName")), _
																 CStr(Request.Form("txtStdAccountDescription")), _
																 CLng(Request.Form("cboStdAccountTypes")), _
																 CStr(Request.Form("cboStdAccountNature"))		
		Set oStdAccountMgr = Nothing

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
	<td bgColor="khaki"><b>La cuenta estándar fue eliminada satisfactoriamente.</b></td>
</tr>
<tr>
	<td>
		<br>
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
  End If %>
</body>
</html>