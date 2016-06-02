<% 
	Option Explicit	
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If	
%>
<% RSDispatch %>

<!--#INCLUDE FILE="rs.asp"-->

<SCRIPT RUNAT=SERVER Language=javascript>

	function IServerScripts()
	{	  
	  //this.CboAccountCurrencies = Function('nAccountId', 'nSelCurrencyId', 'return CboAccountCurrencies(nAccountId, nSelCurrencyId)');
		//this.CboAccountSubsidiaryLedgers = Function('nAccountId', 'nSectorId', 'nSubsidiaryLedgerId', 'return CboAccountSubsidiaryLedgers(nAccountId, nSectorId, nSubsidiaryLedgerId)');
	  //this.CboGLInCategory = Function('nCategoryId', 'return CboGLInCategory(nCategoryId)');	  	  	  
	  //this.CboGralLedgersInGroup = Function('nGroupId', 'nSelectedItem', 'return CboGralLedgersInGroup(nGroupId, nSelectedItem)');	  
	  //this.CboSectorsInAccount = Function('nAccountId', 'nSectorId', 'return CboSectorsInAccount(nAccountId, nSectorId)');	  
		//this.CboSubsidiaryAccounts = Function('nSubsidiaryLedgerId', 'nSubsidiaryAccountId', 'return CboSubsidiaryAccounts(nSubsidiaryLedgerId, nSubsidiaryAccountId)');
	 // this.GetAccountProps = Function('nAccountId', 'return GetAccountProps(nAccountId)');	  	  	  
	}
	public_description = new IServerScripts();  

</SCRIPT>

<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">

Function CboAccountCurrencies(nAccountId, nSelCurrencyId)
  Dim oGLVoucherUS, sTemp	
	'******************************************************
	On Error Resume Next
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboAccountCurrencies(Session("sAppServer"), CLng(nAccountId), CLng(nSelCurrencyId))
  Set oGLVoucherUS = Nothing
    
	sTemp = "<SELECT name=cboCurrencies style=""WIDTH: 100%"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboAccountCurrencies = sTemp
End Function

Function CboAccountSubsidiaryLedgers(nAccountId, nSectorId, nSubsidiaryLedgerId)
  Dim oGLVoucherUS, sTemp	
	'*****************************************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
  sTemp = oGLVoucherUS.CboSubsidaryLedgers(Session("sAppServer"), CLng(nAccountId), CLng(nSectorId), CLng(nSubsidiaryLedgerId))
  Set oGLVoucherUS = Nothing  
	sTemp = "<SELECT name=cboSubsidiaryLedgers style=""WIDTH: 100%"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf	
	CboAccountSubsidiaryLedgers = sTemp
End Function

Function CboGLInCategory(nCategoryId)
  Dim oGLVoucherUS, sTemp	
	'**********************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGralLedgers(Session("sAppServer"), CLng(Session("uid")), CLng(nCategoryId))  
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGralLedgers style=""WIDTH: 520px"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboGLInCategory = sTemp
End Function

Function CboGralLedgersInGroup(nGroupId, nSelectedItem)
  Dim oGLVoucherUS, sTemp	
	'****************************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboGralLedgers(Session("sAppServer"), CLng(Session("uid")), CLng(nGroupId), CLng(nSelectedItem))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboGralLedgers style='WIDTH: 100%'>" & VbCrLf & _
					"<OPTION value=0>-- Todas las contabilidades en el grupo seleccionado--</OPTION>" & VbCrLf & _
					sTemp & _
			    "</SELECT>" & VbCrLf
	CboGralLedgersInGroup = sTemp
End Function

Function CboSectorsInAccount(nAccountId, nSectorId)
  Dim oGLVoucherUS, sTemp	
	'**********************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     
  sTemp = oGLVoucherUS.CboAccountSectors(Session("sAppServer"), CLng(nAccountId),CLng(nSectorId))
  Set oGLVoucherUS = Nothing
	sTemp = "<SELECT name=cboSectors style=""WIDTH: 100%"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboSectorsInAccount = sTemp
End Function

Function CboSubsidiaryAccounts(nSubsidiaryLedgerId, nSubsidiaryAccountId)
  Dim oGralLedgerUS, sTemp	
	'**********************************************************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
  sTemp = oGralLedgerUS.CboSubsidiaryAccounts(Session("sAppServer"), CLng(nSubsidiaryLedgerId), CLng(nSubsidiaryAccountId))
  Set oGralLedgerUS = Nothing	
	sTemp = "<SELECT name=cboSubsidiaryAccounts style=""WIDTH: 100%"" onchange=""return setSubsidiaryAccountName();"">" & VbCrLf & _
			     sTemp & _
			    "</SELECT>" & VbCrLf
	CboSubsidiaryAccounts = sTemp
End Function

Function GetAccountProps(nAccountId)
  Dim oGralLedgerUS
  '********************************
	On Error Resume Next
	Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
	GetAccountProps = oGralLedgerUS.GetAccountProps(Session("sAppServer"), CLng(nAccountId))
	Set oGralLedgerUS = Nothing
End Function

</SCRIPT>