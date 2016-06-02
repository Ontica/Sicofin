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
		gsReturnPage = "../edit_rule_group.asp?id=" & Request.Form("txtRuleId")
		Call SaveItem(CLng(Request.Form("txtRuleId")))
  End If
        
  Sub SaveItem(nRuleId)
		Dim oRule, oRecordset, nRuleGroupId, nInsertAfterItemId, bSaveAsChild
		'*************************************************************************
		'On Error Resume Next
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")
		Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(nRuleId), False)
		If CLng(nRuleId) = 0 Then
			nInsertAfterItemId = Request.Form("cboInsertAfterItems")
			nRuleGroupId       = Request.Form("txtRuleGroupId")
			bSaveAsChild		   = Request.Form("txtSaveAsChild")
			oRecordset("id_regla_contable")	  = Request.Form("txtRuleDefId")
			oRecordset("tipo_grupo_cuenta")	  = 0
		End If
		oRecordset("numero_grupo_cuenta")	  = Request.Form("txtNumber")
		oRecordset("nombre_grupo_cuenta")   = Request.Form("txtDescription")
		oRecordset("id_entidad_agrupador")  = 1
		oRecordset("id_agrupador_origen")   = 1
		oRecordset("cuenta_origen_inicial") = ""
		oRecordset("cuenta_origen_final")   = ""
		If CLng(Request.Form("cboFilters")) = 1 Then
			oRecordset("filtro_cuentas_origen") = "(saldo_actual > 0)"
		ElseIf CLng(Request.Form("cboFilters")) = -1 Then
			oRecordset("filtro_cuentas_origen") = "(saldo_actual < 0)"
		Else
			oRecordset("filtro_cuentas_origen") = ""
		End If
		oRecordset("auxiliar_origen")     = ""
		oRecordset("id_sector_origen")    = 0
		oRecordset("id_moneda_origen")    = 0
		oRecordset("operador")					  = Request.Form("cboOperators")
		oRecordset("factor")						  = Request.Form("txtFactor")
		oRecordset("ID_CUENTA_DESTINO")   = 0
    oRecordset("ID_AUXILIAR_DESTINO") = 0
    oRecordset("ID_SECTOR_DESTINO")   = 0
    oRecordset("ID_CUENTA_DESTINO_SOBREGIRO")   = 0
    oRecordset("ID_AUXILIAR_DESTINO_SOBREGIRO") = 0
    oRecordset("ID_SECTOR_DESTINO_SOBREGIRO")   = 0    
    If CLng(nRuleId) = 0 Then    
			oRule.AddRuleGroup Session("sAppServer"), (oRecordset), CLng(nRuleGroupId), CLng(nInsertAfterItemId), CBool(bSaveAsChild)
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

