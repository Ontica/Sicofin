<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Call Main()
	
	Sub Main()	
		Dim oVoucherBS, oRecordset, dApplicationDate
		'*************************************************
		On Error Resume Next
		If Len(Request.Form("txtAlternativeDate")) <> 0 Then
			dApplicationDate = Request.Form("txtAlternativeDate")
		Else
			dApplicationDate = Request.Form("cboApplicationDates")			
		End If
		
		Set oVoucherBS = Server.CreateObject("AOGLVoucher.CServer")
		oVoucherBS.SaveTransaction Session("sAppServer"), CLng(Request.QueryString("id")), _
															 Request.Form("txtDescription"), CLng(Request.Form("cboVoucherTypes")), _
															 CDate(dApplicationDate), CLng(Request.Form("cboSources"))
		
		
		If (Err.number = 0) Then			
		'	Response.Redirect("../voucher_editor.asp?id=" & Request.QueryString("id"))
		Else		
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If
	End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
function refreshVoucher() {
	window.opener.document.all.ancRefreshAll.click();
	return false;
}
//-->
</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="refreshVoucher();window.close();">
</BODY>
</HTML>