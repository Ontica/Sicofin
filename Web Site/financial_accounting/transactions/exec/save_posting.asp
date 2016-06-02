<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage
	
	gsReturnPage = "../posting_editor.asp?transactionId=" & Request.QueryString("transactionId")
	If CLng(Request.QueryString("postingId")) <> 0 Then
		gsReturnPage = gsReturnPage & "&id=" & Request.QueryString("postingId")
	End If
	
	Call SaveItem(CLng(Request.QueryString("transactionId")), CLng(Request.QueryString("postingId")))
 
	Sub SaveItem(nTransactionId, nPostingId)
    Dim oVoucherBS
  	Dim sStdAccountNumber, sSubsidiaryAccount, sSectorNumber, sResponsibilityArea
  	Dim nReferencePostingId, sBudgetKey, sDisponibilityKey, sVerificationNumber
		Dim sPostingType, dPostingDate, sDescription, nCurrencyId, nAmount, nBaseAmount
		
		On Error Resume Next		
		sStdAccountNumber		= Request.Form("txtAccount")
		If Len(Request.Form("txtSubsidiaryAccount")) <> 0 Then
			sSubsidiaryAccount	= Request.Form("txtSubsidiaryAccountP") & Request.Form("txtSubsidiaryAccount")
		Else
			sSubsidiaryAccount = ""
		End If
		sSectorNumber				= Request.Form("txtSector")
		nReferencePostingId = Request.Form("txtPostingReferenceId")
		sResponsibilityArea = Request.Form("txtResponsibilityArea")		
		sBudgetKey					= Request.Form("txtBudgetKey")
		sDisponibilityKey		= Request.Form("txtDisponibilityKey")
		sVerificationNumber = Request.Form("txtVerificationNumber")		
		sPostingType				= Request.Form("cboPostingType")
		dPostingDate = Request.Form("txtPostingDate")
		sDescription = Request.Form("txtDescription")		
		nCurrencyId = Request.Form("cboCurrencies")
		nAmount = Request.Form("txtAmount")
		If Len(Request.Form("txtBaseAmount")) <> 0 Then
			nBaseAmount = Request.Form("txtBaseAmount")
		Else
			nBaseAmount = nAmount
		End If
		Set oVoucherBS = Server.CreateObject("AOGLVoucher.CServer")
		
		If CLng(nPostingId) = 0 Then
			oVoucherBS.AddPostingValues Session("sAppServer"), CLng(nTransactionId), _
																	CStr(sStdAccountNumber), CStr(sSubsidiaryAccount), _
																	CStr(sSectorNumber), CLng(nReferencePostingId), _
																	CStr(sResponsibilityArea), CStr(sBudgetKey), CStr(sDisponibilityKey), _
																	CStr(sVerificationNumber), CStr(sPostingType), dPostingDate, _
																	CStr(sDescription), CLng(nCurrencyId), _
																	nAmount, nBaseAmount, False
		Else			
			oVoucherBS.EditPostingValues Session("sAppServer"), CLng(nTransactionId), CLng(nPostingId), _
																	 CStr(sStdAccountNumber), CStr(sSubsidiaryAccount), _
																	 CStr(sSectorNumber), CLng(nReferencePostingId), _
																	 CStr(sResponsibilityArea), CStr(sBudgetKey), CStr(sDisponibilityKey), _
																	 CStr(sVerificationNumber), CStr(sPostingType), dPostingDate, _
																	 CStr(sDescription), CLng(nCurrencyId), _
																	 nAmount, nBaseAmount, False																	 
		End If
								
		If (Err.number = 0) Then
			Set Session("oError") = Nothing
			Response.Redirect gsReturnPage
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
%>
</body>
</html>