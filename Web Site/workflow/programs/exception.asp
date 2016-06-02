<%
  Option Explicit     
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsErrPage, gsCancelPage, gsSupportPage
	Dim gsErrNumber, gsErrSource, gsErrDescription
	  	 
  Call Main()

	Sub Main()		
		gsErrNumber = Session("nErrNumber")
		gsErrSource = Session("sErrSource")
		gsErrDescription = Session("sErrDescription")
		gsErrPage = Session("sErrPage")		
		gsCancelPage = "../../main.asp"
		gsSupportPage = "support.asp"
		Session("nErrNumber") = ""
		Session("sErrSource") = ""
		Session("sErrDescription") = ""
		Session("sErrPage") = ""		
	End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="../../resources/pages_style.css">
<TITLE>Banobras - Intranet corporativa</TITLE>
</HEAD>
<BODY>
<BR>
<TABLE align=CENTER border=0 cellPadding=1 cellSpacing=1 width="50%">
	<TR>
	<TD colspan=2 align=center>
		<H2>Ocurrió el siguiente problema</H2>
	</TD>
	</TR>
  <TR>
		<TD valign=top>Descripción: &nbsp;&nbsp;</TD>
		<TD valign=top><B><%=gsErrDescription%></B><BR>&nbsp;</TD>
	</TR>
  <TR>
    <TD valign=top>Fuente: &nbsp;&nbsp;</TD>
    <TD valign=top><B><%=gsErrSource%></B><BR>&nbsp;</TD>    
  </TR>
  <TR>
    <TD valign=top>Número: &nbsp;&nbsp;</TD>
    <TD valign=top><B><%=gsErrNumber%></B><BR>&nbsp;</TD>
  </TR>    
  <TR>
    <TD>&nbsp;</TD>    
		<TD nowrap>
			<form name=frmSend action="../../../mahler/pages/support.asp" method=post>
			<INPUT type="hidden" name="txtErrNumber" value="<%=gsErrNumber%>">
			<INPUT type="hidden" name="txtErrDescription" value="<%=gsErrDescription%>">
			<INPUT type="hidden" name="txtErrSource" value="<%=gsErrSource%>">
			<INPUT type="hidden" name="txtErrPage" value="<%=gsErrPage%>">
			<INPUT type="button" value="Reintentar" name=cmdReturn LANGUAGE=javascript onclick="window.location.href ='<%=gsErrPage%>'"> &nbsp;&nbsp;&nbsp;
      <INPUT type="button" value="Cancelar" name=cmdCancel LANGUAGE=javascript onclick="window.location.href ='<%=gsCancelPage%>'"> &nbsp;&nbsp;&nbsp;
      <INPUT type="submit" value="Enviar a soporte" name=cmdSend">
      </form>
		</TD>
  </TR>    
</TABLE>
</BODY>
</HTML>