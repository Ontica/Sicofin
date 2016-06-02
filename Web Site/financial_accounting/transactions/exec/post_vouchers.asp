<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage, gsCancelPage, nScriptTimeout
	Dim gsErrNumber, gsErrSource, gsErrDescription

	gsReturnPage = "../add_voucher.asp"
	gsCancelPage = Application("main_page")
	 
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
  Call PostTransactions()
  Server.ScriptTimeout = nScriptTimeout
   
  Sub PostTransactions()
		Dim aVouchersIds, oTransaction, oVoucherUS, nTransactionId, gnDateStatus, i
		'************************************************************
		Set oTransaction = Server.CreateObject("AOGLVoucher.CServer")
		Set oVoucherUS   = Server.CreateObject("AOGLVoucherUS.CVoucher")
		
		aVouchersIds = Split(Request.Form("txtPendingVouchers"), ",")
    For i = LBound(aVouchersIds) To UBound(aVouchersIds)
			nTransactionId = aVouchersIds(i)			
			gnDateStatus	= oVoucherUS.DateStatus(Session("sAppServer"), CLng(nTransactionId))
			If gnDateStatus = 0 Then			
				If (oTransaction.ValidateTransaction(Session("sAppServer"), CLng(nTransactionId)) = 0) Then
					oTransaction.PostTransaction Session("sAppServer"), CLng(nTransactionId), CLng(Session("uid"))
				End If
			ElseIf gnDateStatus < 0 Then
				If oVoucherUS.UserIsSupervisor(Session("sAppServer"), CLng(nTransactionId), CLng(Session("uid"))) Then
					If (oTransaction.ValidateTransaction(Session("sAppServer"), CLng(nTransactionId)) = 0) Then
						oTransaction.PostTransaction Session("sAppServer"), CLng(nTransactionId), CLng(Session("uid"))
					End If
				Else
					If (oTransaction.ValidateTransaction(Session("sAppServer"), CLng(nTransactionId)) = 0) Then
						oTransaction.SendTransactionToCheck Session("sAppServer"), CLng(nTransactionId)
					End If					
				End If
			End If			
		Next
		Set oVoucherUS = Nothing
		Set oTransaction = Nothing		
		If (Err.number = 0) Then
			Response.Redirect("../voucher_explorer.asp")
		Else
			gsErrNumber = Err.number
			gsErrSource = Err.source
			gsErrDescription = Err.description
		End If
  End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
</HEAD>
<BODY>
<TABLE align=center border=1 cellPadding=1 cellSpacing=1 width="60%">  
  <TR>
    <TD colspan=2 align=center>Tengo un problema</TD>
  </TR>
  <TR>
    <TD colspan=2 align=center>
    <%=gsErrSource%>
    </TD>
	</TR>
  <TR>
    <TD colspan=2 align=center>
    Fuente: <%=gsErrSource%> &nbsp; Número: <%=gsErrNumber%>
    </TD>
  </TR>
  <TR>
    <TD align=right><INPUT type="button" value="Reintentar" name=cmdReturn LANGUAGE=javascript onclick="window.location.href ="<%=gsReturnPage%>""></TD>
    <TD><INPUT type="button" value="Cancelar" name=cmdCancel LANGUAGE=javascript onclick="window.location.href ="<%=gsCancelPage%>""></TD>
  </TR>    
</TABLE>
</BODY>
</HTML>