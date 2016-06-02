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
		gsReturnPage = "../edit_rule.asp?id=" & Request.Form("txtRuleId")
		Call SaveItem(CLng(Request.Form("txtRuleId")))
  End If
      
  Sub SaveItem(nRuleId)
		Dim oRule, oRecordset, nRuleGroupId, nInsertAfterItemId
		'**********************************************************
		'On Error Resume Next
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")
		Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(nRuleId), False)
		If CLng(nRuleId) = 0 Then
			nInsertAfterItemId = Request.Form("cboInsertAfterItems")
			nRuleGroupId       = Request.Form("txtRuleGroupId")		
			oRecordset("id_regla_contable")	  = Request.Form("txtRuleDefId")
			oRecordset("tipo_grupo_cuenta")	  = Request.Form("txtRuleTypeId")
		End If
		oRecordset("numero_grupo_cuenta")	  = Request.Form("txtNumber")
		oRecordset("nombre_grupo_cuenta")   = Request.Form("txtDescription")
		oRecordset("id_entidad_agrupador")  = 1
		oRecordset("id_agrupador_origen")   = 1
		oRecordset("cuenta_origen_inicial") = Request.Form("txtFromAccount")
		oRecordset("cuenta_origen_final")   = Request.Form("txtToAccount")
		oRecordset("filtro_cuentas_origen") = Request.Form("txtRestriction")
		oRecordset("auxiliar_origen")				= Request.Form("txtSubsidiaryAccount")
		oRecordset("clip_cuenta_origen")		= CLng(Request.Form("cboClips"))
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
		oRecordset("operador")					  = Request.Form("cboOperators")
		oRecordset("factor")						  = Request.Form("txtFactor")
		oRecordset("ID_CUENTA_DESTINO")   = 0
    oRecordset("ID_AUXILIAR_DESTINO") = 0
    oRecordset("ID_SECTOR_DESTINO")   = 0
    oRecordset("ID_CUENTA_DESTINO_SOBREGIRO")   = 0
    oRecordset("ID_AUXILIAR_DESTINO_SOBREGIRO") = 0
    oRecordset("ID_SECTOR_DESTINO_SOBREGIRO")   = 0    
    If CLng(nRuleId) = 0 Then			
			oRule.AddRule Session("sAppServer"), (oRecordset), CLng(nRuleGroupId), CLng(nInsertAfterItemId)
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
<body onload='window.close();'>
</body>
</html>

