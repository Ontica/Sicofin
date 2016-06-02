<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage
 
	gsReturnPage = "../pages/holiday_editor.asp?calendar_id=" & Request.Form("txtCalendarId")
	
	If CLng(Request.Form("txtItemId")) = 0 Then
		Call SaveItem(0)
	Else
		gsReturnPage = gsReturnPage & "&id=" & Request.Form("txtItemId")
		Call SaveItem(CLng(Request.Form("txtItemId")))
  End If
  
    
  Sub SaveItem(nItemId)
		Dim oCalendar, oRecordset
		'************************
		On Error Resume Next
		Set oCalendar  = Server.CreateObject("AOCalendar.CManager")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")
		Set oRecordset = oCalendar.GetHolidayRS(Session("sAppServer"), CLng(nItemId))
		oRecordset("calendar_id") = Request.Form("txtCalendarId")
		oRecordset("holiday")	= Request.Form("txtHoliday")
		oRecordset("holiday_name")		= Request.Form("txtHolidayName")
		oCalendar.SaveHoliday Session("sAppServer"), (oRecordset), CLng(nItemId)		
		oRecordset.Close
		Set oRecordset = Nothing
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
	<td bgColor="khaki"><b>La información del día festivo fue enviada satisfactoriamente.</b></td>
</tr>
<tr><td><br><b>¿Qué desea hacer?</b></td></tr>
<tr>
	<% If CLng(Request.Form("txtItemId")) = 0 Then %>
	<td>
		<a href='<%=gsReturnPage%>'>Agregar otro día festivo</a>
	</td>	
	<% Else %>
	<td>
		<a href='<%=gsReturnPage%>'>Regresar al editor</a>
	</td>
	<% End If %>
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