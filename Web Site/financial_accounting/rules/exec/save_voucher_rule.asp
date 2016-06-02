<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	  
	Dim gsReturnPage
			
	If CLng(Request.Form("txtRuleId")) = 0 Then
		gsReturnPage = ""
		Call SaveItem(0)
	Else
		gsReturnPage = "../edit_voucher_rule.asp?id=" & Request.Form("txtRuleId")
		Call SaveItem(CLng(Request.Form("txtRuleId")))
  End If
      
  Sub SaveItem(nRuleId)
		Dim oRule, oRecordset, nInsertAfterItemId
		'**********************************************************
		'On Error Resume Next
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")
		Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(nRuleId), False)
		If CLng(nRuleId) = 0 Then
			nInsertAfterItemId = Request.Form("cboInsertAfterItems")			
			oRecordset("id_regla_contable")	  = Request.Form("txtRuleDefId")
			oRecordset("tipo_grupo_cuenta")	  = Request.Form("txtRuleTypeId")
			oRecordset("id_entidad_agrupador")  = CLng(Request.Form("txtGroupEntityId"))
			oRecordset("id_agrupador_origen")   = CLng(Request.Form("txtRuleGroupId"))
		End If
		oRecordset("numero_grupo_cuenta")	  = ""
		oRecordset("nombre_grupo_cuenta")   = ""
		oRecordset("cuenta_origen_inicial") = Request.Form("txtFromAccount")
		oRecordset("cuenta_origen_final")   = Request.Form("txtToAccount")
		oRecordset("filtro_cuentas_origen") = Request.Form("txtRestriction")
		oRecordset("auxiliar_origen")				= Request.Form("txtSubsidiaryAccount")		
		If Len(Request.Form("chkSectors")) <> 0 Then
			oRecordset("id_sector_origen") = -1 * Request.Form("cboSectors")
		Else
			oRecordset("id_sector_origen") = Request.Form("cboSectors")
		End If
		If Len(Request.Form("chkCurrencies")) <> 0 Then
			oRecordset("id_moneda_origen") = -1 * Request.Form("cboCurrencies")
		Else
			oRecordset("id_moneda_origen") = Request.Form("cboCurrencies")
		End If			
		oRecordset("operador") = "+"
		oRecordset("factor")	 = 1
		If (Len(Request.Form("txtTargetAccount")) <> 0) Then
			oRecordset("ID_CUENTA_DESTINO")   = CLng(oRule.GetAccountId(Session("sAppServer"), Request.Form("txtTargetAccount")))			
			If (Len(Request.Form("txtTargetSubsidiaryAccount")) <> 0) Then
				oRecordset("AUXILIAR_DESTINO")    = CStr(Request.Form("txtTargetSubsidiaryAccount"))
				oRecordset("ID_AUXILIAR_DESTINO") = oRule.GetSubsidiaryAccountId(Session("sAppServer"), 9, CStr(Request.Form("txtTargetSubsidiaryAccount")))
			Else
				oRecordset("AUXILIAR_DESTINO")    = ""
				oRecordset("ID_AUXILIAR_DESTINO") = 0
			End If
			oRecordset("ID_SECTOR_DESTINO")   = CStr(Request.Form("cboTargetSectors"))
			oRecordset("ID_CUENTA_DESTINO_SOBREGIRO")   = CLng(oRule.GetAccountId(Session("sAppServer"), Request.Form("txtTargetOBAccount")))
			oRecordset("ID_AUXILIAR_DESTINO_SOBREGIRO") = 0
			oRecordset("ID_SECTOR_DESTINO_SOBREGIRO")   = 0 
		Else
			oRecordset("ID_CUENTA_DESTINO")   = 0
			oRecordset("AUXILIAR_DESTINO")    = ""
			oRecordset("ID_AUXILIAR_DESTINO") = 0
			oRecordset("ID_SECTOR_DESTINO")   = 0
			oRecordset("ID_CUENTA_DESTINO_SOBREGIRO")   = 0
			oRecordset("ID_AUXILIAR_DESTINO_SOBREGIRO") = 0
			oRecordset("ID_SECTOR_DESTINO_SOBREGIRO")   = 0 			
		End If
    If CLng(nRuleId) = 0 Then			
			oRule.AddVoucherRule Session("sAppServer"), (oRecordset), _
															 CStr(Request.Form("txtTargetAccount")), _
															 CStr(Request.Form("txtTargetSubsidiaryAccount")), _
															 CStr(Request.Form("txtTargetOBAccount")), _
															 CStr(Request.Form("txtTargetOBSubsidiaryAccount"))
															 
		Else
			oRule.SaveRule Session("sAppServer"), (oRecordset), CLng(nRuleId)		
		End If
		oRecordset.Close
		Set oRecordset = Nothing
		Set oRule = Nothing
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
	<td bgColor="khaki"><b>La agrupación de cuentas fue guardada satisfactoriamente.</b></td>
</tr>
<tr><td><br><b>¿Qué desea hacer?</b></td></tr>
<% If Len(gsReturnPage) <> 0 Then %>
<tr>
	<td>
		<a href='<%=gsReturnPage%>'>Regresar al editor</a>
	</td>
</tr>
<% End If %>
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

