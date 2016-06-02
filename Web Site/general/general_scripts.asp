<% 
	Option Explicit	
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
%>

<!--#INCLUDE virtual="/empiria/bin/ms_scripts/rs.asp"-->

<% RSDispatch %>

<SCRIPT RUNAT=SERVER Language=javascript>

	function IServerScripts()
	{
		this.CurrencyName = Function('nCurrencyId', 'return CurrencyName(nCurrencyId)');
		this.IsDate = Function('sDate', 'return IsDate_(sDate)');
		this.IsNumeric = Function('sNumber', 'nDecimals', 'return IsNumeric_(sNumber, nDecimals)');
		this.MoveFileFromFTPDir = Function('sTargetDir', 'sFileName', 'return MoveFileFromFTPDir(sTargetDir, sFileName)');
	}
	public_description = new IServerScripts();  

</SCRIPT>

<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">

Function CurrencyName(nCurrencyId)
  Dim oCurrenciesUS
	'*******************************
	On Error Resume Next	
	Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")		
  CurrencyName = oCurrenciesUS.CurrencyName(Session("sAppServer"), CLng(nCurrencyId), True)
  Set oCurrenciesUS = Nothing
End Function

Function IsDate_(sDate)
	IsDate_ = IsDate(sDate)
End Function

Function IsNumeric_(sNumber, nDecimals)
  Dim oGLVoucherUS
  '************************************
	On Error Resume Next	
	Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	IsNumeric_ = oGLVoucherUS.IsNumericOK(CStr(sNumber), CLng(nDecimals))
	Set oGLVoucherUS = Nothing
End Function

Function MoveFileFromFTPDir(sTargetDir, sFileName) 
	Dim oFileMgr
	'**********************************************
	Set oFileMgr = Server.CreateObject("EGEFileManager.CFile")
	MoveFileFromFTPDir = oFileMgr.MoveFromFTPDir(CStr(sTargetDir), CStr(sFileName), True)
	Set oFileMgr = Nothing
End Function

</SCRIPT>