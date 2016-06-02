<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage, gsCancelPage
	Dim gsErrNumber, gsErrSource, gsErrDescription

	gsReturnPage = "../add_voucher.asp"
	gsCancelPage = Application("main_page")
	 	 
  Call PostTransaction()
   
  Sub PostTransaction()
		Dim oTransaction, nTransactionId
		'*******************************
		nTransactionId = Request.QueryString("id")
		If Len(nTransactionId) <> 0 Then
			Set oTransaction = Server.CreateObject("AOGLVoucher.CServer")
			oTransaction.PostTransaction Session("sAppServer"), CLng(nTransactionId), CLng(Session("uid"))
			Set oTransaction = Nothing
		End If
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