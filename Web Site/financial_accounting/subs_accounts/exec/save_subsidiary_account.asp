<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		 		
	If CLng(Request.Form("txtItemId")) = 0 Then
		Call SaveItem(0)
	Else	
		Call SaveItem(CLng(Request.Form("txtItemId")))
  End If
    
  Sub SaveItem(nItemId)
		Dim oGralLedger, oRecordset
		'**************************
		'On Error Resume Next
		Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
		Set oRecordset = Server.CreateObject("ADODB.Recordset")	
		Set oRecordset = oGralLedger.GetSubsidiaryAccountRS(Session("sAppServer"), CLng(nItemId))					
		If (CLng(nItemId) = 0) Then			
			oRecordset("numero_cuenta_auxiliar") = Request.Form("txtSubsidiaryAccountPrefix") & Request.Form("txtSubsidiaryAccountNumber")
		Else
			If CLng(oRecordset("id_mayor_auxiliar")) <> CLng(Request.Form("cboSubsidiaryLedgers")) Then
				Call DeleteExtendedAttributes(oRecordset("id_mayor_auxiliar"), CLng(nItemId))
			End If
		End If
		oRecordset("id_mayor_auxiliar")			 = Request.Form("cboSubsidiaryLedgers")
		oRecordset("nombre_cuenta_auxiliar") = Request.Form("txtSubsidiaryAccountName")
		oRecordset("descripcion")						 = Request.Form("txtSubsidiaryAccountDescription")		
		nItemId = oGralLedger.SaveSubsidiaryAccount(Session("sAppServer"), (oRecordset), CLng(nItemId))
		oRecordset.Close
		Set oRecordset = Nothing
		Set oGralLedger = Nothing
		Call SaveExtendedAttributes(Request.Form("cboSubsidiaryLedgers"), CLng(nItemId))
		If (Err.number = 0) Then
			Set Session("oError") = Nothing
		Else
			Set Session("oError") = Err
			Response.Redirect Application("error_page")
		End If
  End Sub
  
  Sub DeleteExtendedAttributes(nSubsLedgerId, nItemId)
		Dim oGralLedger, oExtendedAttrs, oRecordset, nSubsLedgerType, sItem, i
		'*********************************************************************
		Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
		Set oRecordset = oGralLedger.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(nSubsLedgerId))
		nSubsLedgerType = oRecordset("id_tipo_mayor_auxiliar")
		oRecordset.Close
		Set oRecordset = Nothing
		Set oExtendedAttrs = oGralLedger.SubsidiaryAccountExtendedAttrs(Session("sAppServer"), CLng(nSubsLedgerType), CLng(nItemId))
		For i = 0 To oExtendedAttrs.Count - 1
			oExtendedAttrs(oExtendedAttrs.Key(i)) = ""
		Next
		oExtendedAttrs.Save CLng(nItemId)
  End Sub
  
  
  Sub SaveExtendedAttributes(nSubsLedgerId, nItemId)
		Dim oGralLedger, oExtendedAttrs, oRecordset, nSubsLedgerType, sItem, i
		'*********************************************************************
		Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
		Set oRecordset = oGralLedger.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(nSubsLedgerId))
		nSubsLedgerType = oRecordset("id_tipo_mayor_auxiliar")
		oRecordset.Close
		Set oRecordset = Nothing
		Set oExtendedAttrs = oGralLedger.SubsidiaryAccountExtendedAttrs(Session("sAppServer"), CLng(nSubsLedgerType), CLng(nItemId))
		For i = 0 To oExtendedAttrs.Count - 1
			sItem = "txtExtAttr" & oExtendedAttrs.Key(i)
			If Len(Request.Form(sItem)) <> 0 Then
				oExtendedAttrs(oExtendedAttrs.Key(i)) = Request.Form(sItem)				
			End If
		Next
		oExtendedAttrs.Save CLng(nItemId)
  End Sub
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
</head>
<body onload="window.close();">
</body>
</html>