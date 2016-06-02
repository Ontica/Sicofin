<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsCalendarsTable
	
	Call Main()
	
	Sub Main()
		Dim oCalendarUS
		'************
		On Error Resume Next
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		Select Case Request.QueryString("order")
			Case ""
				gsCalendarsTable = oCalendarUS.GetCalendarHTMLTable(Session("sAppServer"), "")
			Case "1"
				gsCalendarsTable = oCalendarUS.GetCalendarHTMLTable(Session("sAppServer"), "calendar_name")
			Case "2"
				gsCalendarsTable = oCalendarUS.GetCalendarHTMLTable(Session("sAppServer"), "calendar_year, calendar_name")
		End Select		
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
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function callEditor(nOperation, nItemId) {
  switch (nOperation) {
    case 1:		//Add
			window.open("calendar_editor.asp", null, "height=180,width=340,location=0,resizable=0");
			return false;
    case 2:		//Edit
			window.open("calendar_editor.asp?id=" + nItemId, null, "height=180,width=340,location=0,resizable=0");
			return false;
	}
	return false;
}

function refreshPage(nOrderId) {
  if (nOrderId == 0) {
		window.location.href = "calendars.asp";
	} else {	
		window.location.href = "calendars.asp" + '?order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=52px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="70%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Calendarios</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>
			<INPUT type="button" name=cmdAddItem value="Agregar" style="WIDTH:80px" LANGUAGE=javascript onclick="return callEditor(1,0);">&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdRefresh value="Actualizar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href=window.location.href;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdReturn value="Cerrar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href = '<%=Session("main_page")%>';">
		</TD>
	</TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=90%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="70%">
<A href="#SCROLLABLE_DIV_TOP"></A>
<% If Len(gsCalendarsTable) <> 0 Then %>
	<TR>
	  <TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Calendario</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Año</b></FONT></A></TD>
	  <TD colspan=3 align=center><FONT color=maroon><b>Asignación de</b></FONT></TD>
	</TR>
	<%=gsCalendarsTable%>
	<TR>
	  <TD nowrap colspan=5 align=right><A href="#SCROLLABLE_DIV_TOP">Subir</A></TD>
	</TR>	
<% Else %>
	<TR><TD colspan=5 align=center>No hay calendarios definidos.</TD></TR>
<% End If %>
</TABLE>
<BR>&nbsp;
</DIV>
</BODY>
</HTML>