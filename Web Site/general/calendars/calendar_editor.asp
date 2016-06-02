<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsCalendarName, gsCboYears	
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oCalendarUS
		'**************
		gbEdit = False
		gsTitle = "Agregar calendario"
		gnItemId = 0
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsCalendarName = ""
		gsCboYears = oCalendarUS.CboYears()
		Set oCalendarUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oCalendarUS
		'******************
		gbEdit = True
		gsTitle = "Editar calendario"
		gnItemId = CLng(nItemId)		
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsCalendarName = oCalendarUS.CalendarName(Session("sAppServer"), CLng(gnItemId))
		gsCboYears = oCalendarUS.CboYears(,  , oCalendarUS.CalendarYear(Session("sAppServer"), CLng(gnItemId)))
		Set oCalendarUS = Nothing	
	End Sub	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function validate() {
	var dDocument = window.document.all;
	if (dDocument.txtCalendarName.value == '') {
		alert("Requiero el nombre del calendario.");
		dDocument.txtCalendarName.focus();
		return false;
	}
	return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el calendario "<%=gsCalendarName%>"?')) {
		window.document.frmEditor.action = "./exec/delete_calendar.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "./exec/save_calendar.asp";
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
  <TD valign=top>Nombre:</TD>
  <TD valign=top colspan=3><TEXTAREA name=txtCalendarName rows=3 style="WIDTH: 100%"><%=gsCalendarName%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=center>Año:</TD>
	<TD colSpan=3>
		<SELECT name=cboYears>
			<%=gsCboYears%>
		</SELECT>
	</TD>
</TR>
<TR>
  <TD><INPUT name=txtItemId type="hidden" value="<%=gnItemId%>"></TD>
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
</HTML>