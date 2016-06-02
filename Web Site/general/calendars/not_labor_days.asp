<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsNotLaborDaysTable, gnCalendarId, gsCalendarName
	
	Call Main()
	
	Sub Main()
		Dim oCalendarUS
		'************
		On Error Resume Next
		gnCalendarId = CLng(Request.QueryString("Id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsCalendarName = oCalendarUS.CalendarName(Session("sAppServer"), CLng(gnCalendarId))
		gsNotLaborDaysTable = oCalendarUS.GetNotLaborWeekDaysHTMLTable(Session("sAppServer"), CLng(gnCalendarId))
		Set oCalendarUS = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("./exec/exception.asp")
		End If
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
</HEAD>
<BODY SCROLL=NO>
<FORM name=frmSend action="./exec/save_not_labor_days.asp?id=<%=gnCalendarId%>" method=post>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=78px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="70%">
	<TR>
		<TD nowrap colspan=3><FONT face=Arial size=3 color=maroon><STRONG>Días no laborables en <%=gsCalendarName%></STRONG></FONT></TD>
	</TR>
	<TR>	
	  <TD colspan=3 align=right nowrap>
			<INPUT type="submit" name=cmdAddItem value="Guardar" style="WIDTH:80px">&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdRefresh value="Actualizar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href='not_labor_days.asp?id=<%=gnCalendarId%>'">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdReturn value="Cerrar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href = 'calendars.asp';">
		</TD>
	</TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=90%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="70%">
<% If Len(gsNotLaborDaysTable) <> 0 Then %>
	<%=gsNotLaborDaysTable%>
<% Else %>
	<TR><TD colspan=4 align=center>ERROR GRAVE DEL SISTEMA. ¡Este calendario no tiene días de la semana!.</TD></TR>
<% End If %>
</TABLE>
<BR>
</DIV>
</FORM>
</BODY>
</HTML>