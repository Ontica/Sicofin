<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim oVoucherUS
	Dim aVouchersIds
	Dim gsVoucherHeader, gsVoucherPostings, gsVoucherStatus, nScriptTimeout
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
  Call Main()
  Server.ScriptTimeout = nScriptTimeout
	
	Sub Main()
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")		
		If Len(Request.QueryString("id")) <> 0 Then
			Redim aVouchersIds(0)
			aVouchersIds(0) = Request.QueryString("id")
		Else
			aVouchersIds = Split(Request.Form("txtPendingVouchers"), ",")
		End If
	End Sub	
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/voucher.css">
</HEAD>
<BODY bgcolor=white>
<% Dim i %>
<% For i = LBound(aVouchersIds) To UBound(aVouchersIds)
     gsVoucherHeader = oVoucherUS.Header(Session("sAppServer"), CLng(aVouchersIds(i)))
     gsVoucherPostings = oVoucherUS.GetPostings(Session("sAppServer"), CLng(aVouchersIds(i)), False)
     gsVoucherStatus = oVoucherUS.TransactionStatus(Session("sAppServer"), CLng(aVouchersIds(i)))
%>
<% If i > 0 Then %>
<P STYLE="page-break-before:always"></P>
<% End If %>
<TABLE border=0 cellPadding=1 cellSpacing=1 height=40 width="100%">
  <TR>
		<TD width="80%" align=center><FONT size=4><STRONG>Póliza no actualizada</STRONG></FONT></TD>
	</TR>
  <TR>
		<TD width="80%" align=center><FONT size=2><STRONG>Banco Nacional de Obras y Servicios Públicos, S.N.C.</STRONG></FONT></TD>		
	</TR>
</TABLE>
<br>
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<%=gsVoucherHeader%>
</TABLE>
<br>
<br>
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<TR>
	  <TD nowrap width=120><b>Núm. de cuenta</b></TD>
	  <TD><b>Sec</b></TD>
	  <TD width=40%><b>Descripción</b></TD>
	  <TD><b>Verif</b></TD>
	  <TD><b>Area</b></TD>
		<TD align=center><b>Moneda</b></TD>
	  <TD align=center nowrap><b>T. de cambio</b></TD>
	  <TD colspan=3 align=center width=30%><b>Importes</b></TD>
	</TR>
	<TR>
	  <TD><b><i>Auxiliar</i></b></TD>
	  <TD>&nbsp;</TD>
	  <TD><b><i>Concepto</i></b></TD>
	  <TD colspan=3>&nbsp;</TD>
	  <TD align=center>&nbsp;</TD>
	  <TD align=center><b>Parcial</b></TD>
	  <TD align=center><b>Debe</b></TD>
	  <TD align=center><b>Haber</b></TD>
	</TR>
	<%=gsVoucherPostings%>
</TABLE>
<% If Len(gsVoucherStatus) <> 0 Then %>
	<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
		<TR>
			<TD width=100% align=right><font color="red" size=2><%=gsVoucherStatus%></font></TD>
		</TR>
	</TABLE>
<% End If %>
<br>
<% Next %>
<% Set oVoucherUS = Nothing %>
<% Server.ScriptTimeout = nScriptTimeout %>
</BODY>
</HTML>