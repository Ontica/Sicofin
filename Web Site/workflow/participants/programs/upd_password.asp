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
	 
	gsReturnPage = Application("main_page")
	gsCancelPage = Application("main_page")
	 	 
  Call Main(Request.Form("txtPassword"), Request.Form("txtNewPassword"))
   
  Sub Main(sPassword, sNewPassword)
		Dim oIdentity
		'************
		On Error Resume Next	
		Set oIdentity = Server.CreateObject("AOIdentity.CServices")
		oIdentity.ChangePassword Session("sAppServer"), Session("uid"), CStr(sPassword), CStr(sNewPassword), Session.SessionID
		Set oIdentity = Nothing
		If (Err.number = 0) Then
			Response.Redirect(gsReturnPage)
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