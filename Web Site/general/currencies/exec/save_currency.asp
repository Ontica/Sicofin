<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		  
	Dim gsReturnPage
 
	gsReturnPage = "../pages/currency_editor.asp"
	
	If CLng(Request.Form("txtItemId")) = 0 Then
		Call SaveItem(0)
	Else
		gsReturnPage = gsReturnPage & "?id=" & Request.Form("txtItemId")
		Call SaveItem(CLng(Request.Form("txtItemId")))
  End If
  
    
  Sub SaveItem(nItemId)
		Dim oCurrencies, oRecordset
		'**************************
		On Error Resume Next
		Set oCurrencies = Server.CreateObject("AOCurrencies.CManager")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oCurrencies.GetCurrencyRS(Session("sAppServer"), CLng(nItemId))
			oRecordset("currency_name") = Request.Form("txtName")
			oRecordset("abbrev")				= Request.Form("txtAbbrev")
			oRecordset("symbol")				= Request.Form("txtSymbol")
			If Request.Form("chkReportOnly") = "on" Then
				oRecordset("Report_Only") = 1
			Else
				oRecordset("Report_Only") = 0
			End If
			If Request.Form("chkEditExchangeRate") = "on" Then
				oRecordset("Edit_Exchange_Rate") = 1
			Else
				oRecordset("Edit_Exchange_Rate") = 0
			End If			
			oRecordset("O_ID_MONEDA")		= Request.Form("txtClave")
		oCurrencies.SaveCurrency Session("sAppServer"), (oRecordset), CLng(nItemId)
		oRecordset.Close
		Set oRecordset = Nothing
		Set oCurrencies = Nothing	
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
	<td bgColor="khaki"><b>La información de la moneda fue enviada satisfactoriamente.</b></td>
</tr>
<tr><td><br><b>¿Qué desea hacer?</b></td></tr>
<tr>
	<% If CLng(Request.Form("txtItemId")) = 0 Then %>
	<td>
		<a href='<%=gsReturnPage%>'>Agregar otra moneda</a>
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