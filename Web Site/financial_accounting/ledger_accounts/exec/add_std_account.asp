<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim nScriptTimeout, gsTryAgainPage
 
	gsTryAgainPage = "../standard_account_editor.asp?type_id=" & Request.Form("txtStdAccountTypeId")
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call AddItem()
	Server.ScriptTimeout = nScriptTimeout	
 
  Sub AddItem()
		Dim oStdAccount, sLogFile
		'************************
		On Error Resume Next
		Set oStdAccount = Server.CreateObject("EFAStdActBS.CStdAccount")
		sLogFile = oStdAccount.Insert(Session("sAppServer"), _
																	CLng(Request.Form("txtStdAccountTypeId")), _
																	CStr(Request.Form("txtStdAccountNumber")), _
																	CStr(Request.Form("txtStdAccountName")), _
																	CStr(Request.Form("txtStdAccountDescription")), _
																	CStr(Request.Form("txtStdAccountRole")), _
																	CLng(Request.Form("txtStdAccountType")), _
																	CStr(Request.Form("txtStdAccountNature")), _
																	CStr(Request.Form("txtCurrencies")), _
																	CStr(Request.Form("txtAreas")), _
																	CStr(Request.Form("txtSectors")), _
																	CStr(Request.Form("txtSectorRoles")), _
																	CLng(Session("uid")), CDate(Request.Form("txtFromDate")))

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
	<td bgColor="khaki"><b>La cuenta <b><%=Request.Form("txtStdAccountNumber")%></b> fue agregada satisfactoriamente.</b></td>
</tr>
<tr><td><br><b>¿Qué desea hacer?</b></td></tr>
<tr>
	<td>
		<a href='<%=gsTryAgainPage%>'>Agregar otra cuenta en la <%=LCase(Request.Form("txtStdAccountTypeName"))%></a>
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