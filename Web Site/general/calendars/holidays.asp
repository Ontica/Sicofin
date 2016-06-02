<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsHolidaysTable, gnCalendarId, gsCalendarName
	
	Call Main()
	
	Sub Main()
		Dim oCalendarUS
		'************
		On Error Resume Next
		gnCalendarId = CLng(Request.QueryString("Id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsCalendarName = oCalendarUS.CalendarName(Session("sAppServer"), CLng(gnCalendarId))
		Select Case Request.QueryString("order")
			Case ""
				gsHolidaysTable = oCalendarUS.GetHolidaysHTMLTable(Session("sAppServer"), CLng(gnCalendarId))
			Case "1"
				gsHolidaysTable = oCalendarUS.GetHolidaysHTMLTable(Session("sAppServer"), CLng(gnCalendarId), "holiday")				
			Case "2"
				gsHolidaysTable = oCalendarUS.GetHolidaysHTMLTable(Session("sAppServer"), CLng(gnCalendarId), "holiday_name")
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
	var sURL = "holiday_editor.asp?calendar_id=<%=gnCalendarId%>";
  switch (nOperation) {
    case 1:		//Add
			window.open(sURL, null, "height=200,width=360,location=0,resizable=0");
			return false;
    case 2:		//Edit
			window.open(sURL + "&id=" + nItemId, null, "height=200,width=360,location=0,resizable=0");
			return false;
	}
	return false;
}

function refreshPage(nOrderId) {	
  if (nOrderId == 0) {
		window.location.href = "holidays.asp?id=<%=gnCalendarId%>";
	} else {	
		window.location.href = "holidays.asp?id=<%=gnCalendarId%>" + '&order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=78px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="70%">
	<TR>
		<TD nowrap colspan=3><FONT face=Arial size=3 color=maroon><STRONG>Días festivos en <%=gsCalendarName%></STRONG></FONT></TD>
	</TR>
	<TR>
	  <TD colspan=3 align=right nowrap>
			<INPUT type="button" name=cmdAddItem value="Agregar" style="WIDTH:80px" LANGUAGE=javascript onclick="return callEditor(1,0);">&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdRefresh value="Actualizar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href=window.location.href;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdReturn value="Cerrar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href = 'calendars.asp';">
		</TD>
	</TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=90%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="70%">
<% If Len(gsHolidaysTable) <> 0 Then %>
	<A href="#SCROLLABLE_DIV_TOP"></A>
	<TR>
		<TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Día festivo</b></FONT></A></TD>
	  <TD nowrap><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Se conmemora</b></FONT></A></TD>	  
	</TR>
	<%=gsHolidaysTable%>
	<TR>
	  <TD nowrap colspan=2 align=right><A href="#SCROLLABLE_DIV_TOP">Subir</A></TD>
	</TR>	
<% Else %>
	<TR><TD colspan=4 align=center>Este calendario no tiene días festivos.</TD></TR>
<% End If %>
</TABLE>
<BR>&nbsp;
</DIV>
</BODY>
</HTML>