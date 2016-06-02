<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Call DeleteItem(CLng(Request.QueryString("id")))    
	
  Sub DeleteItem(nItemId)
		Dim oCalendar
		'********************
		On Error Resume Next
		Set oCalendar = Server.CreateObject("AOCalendar.CManager")
		oCalendar.DeleteHoliday Session("sAppServer"), CLng(nItemId)
		Set oCalendar = Nothing
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
	<td bgColor="khaki"><b>El día festivo fue eliminado satisfactoriamente.</b></td>
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