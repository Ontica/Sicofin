<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gnCalendarId, gsHolidayName, gsHoliday, gsYear
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oCalendarUS
		'**************
		gbEdit = False
		gsTitle = "Agregar día festivo"
		gnItemId = 0
		gnCalendarId = CLng(Request.QueryString("calendar_id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsHolidayName = ""
		gsYear = oCalendarUS.CalendarYear(Session("sAppServer"), CLng(gnCalendarId))
		gsHoliday   = "/" & gsYear
		Set oCalendarUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oCalendarUS, oRecordset
		'**************************
		gbEdit = True
		gsTitle = "Editar día festivo"
		gnItemId = CLng(nItemId)		
		gnCalendarId = CLng(Request.QueryString("calendar_id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")		
		Set oRecordset = oCalendarUS.GetHoliday(Session("sAppServer"), CLng(gnItemId))		
		gsHolidayName = oRecordset("holiday_name")
		gsHoliday = oRecordset("holiday")
		gsYear = oCalendarUS.CalendarYear(Session("sAppServer"), CLng(gnCalendarId))
		oRecordset.Close
		Set oRecordset = Nothing
		Set oCalendarUS = Nothing	
	End Sub	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function isDate(sDate) {
	var obj;
	obj = RSExecute("../general_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function validate() {
	var dDocument = window.document.all;
	
	if (dDocument.txtHoliday.value == '') {
		alert("Requiero el día festivo.");
		dDocument.txtHoliday.focus();
		return false;
	}	
	if (!isDate(dDocument.txtHoliday.value)) {
		alert("No reconozco el día festivo proporcionado.");
		dDocument.txtHoliday.focus();
		return false;
	}		
	if (dDocument.txtHolidayName.value == '') {
		alert("Requiero el nombre del día festivo.");
		dDocument.txtHolidayName.focus();
		return false;
	}	
	return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el día festivo "<%=gsHolidayName%>"?')) {
		window.document.frmEditor.action = "./exec/delete_holiday.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();		
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "./exec/save_holiday.asp";
		window.document.frmEditor.submit();
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO>
<FORM name=frmEditor action="" method="post">
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%" height="100%">
<TR>  
  <TD colspan=2 bgcolor=khaki><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>
<% If gbEdit Then %>
	<TD colspan=2 bgcolor=khaki align=right><INPUT name=cmdDelete type=button value="Eliminar" LANGUAGE=javascript onclick="return cmdDelete_onclick()"></TD>
<% Else %>
	<TD colspan=2 bgcolor=khaki>&nbsp;</TD>
<% End If %>
</TR>
<TR>
  <TD valign=middle>Día festivo:</TD>
  <TD valign=middle colspan=2><INPUT name=txtHoliday value=<%=gsHoliday%> maxlength=11 style="HEIGHT: 22px; WIDTH: 100%"></TD>
  <TD valign=middle>&nbsp;&nbsp;( día / mes / <%=gsYear%>)</TD>
</TR>
<TR>
  <TD valign=top>Se conmemora:</TD>
  <TD valign=top colspan=3><TEXTAREA name=txtHolidayName rows=3 style="WIDTH: 100%"><%=gsHolidayName%></TEXTAREA></TD>
</TR>
<TR>
  <TD><INPUT name=txtItemId type="hidden" value="<%=gnItemId%>"><INPUT name=txtCalendarId type="hidden" value="<%=gnCalendarId%>"></TD>
<% If gbEdit Then %>
	<TD><INPUT name=cmdEditItem type=button value="Aceptar" LANGUAGE=javascript onclick="return saveItem()"></TD>
<% Else %>
  <TD><INPUT name=cmdAddItem type=button value="Agregar" LANGUAGE=javascript onclick="return saveItem()"></TD>
<% End If %>
  <TD colspan=2 align=right><INPUT name=cmdCancel type=button value="Cancelar" onclick="window.close();"></TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
