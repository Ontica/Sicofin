<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gnCalendarId, gsPeriodName, gsFromDate, gsToDate, gsCboPeriodStatus
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oCalendarUS, sYear
		'*********************
		gbEdit = False
		gsTitle = "Agregar período"
		gnItemId = 0
		gnCalendarId = CLng(Request.QueryString("calendar_id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		gsPeriodName = ""
		sYear   = oCalendarUS.CalendarYear(Session("sAppServer"), CLng(gnCalendarId))
		gsFromDate   = "/" & sYear
		gsToDate     = "/" & sYear
		gsCboPeriodStatus = oCalendarUS.CboPeriodStatus()
		Set oCalendarUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oCalendarUS, oRecordset
		'**************************
		gbEdit = True
		gsTitle = "Editar período"
		gnItemId = CLng(nItemId)		
		gnCalendarId = CLng(Request.QueryString("calendar_id"))
		Set oCalendarUS = Server.CreateObject("AOCalendarUS.CServer")
		Set oRecordset = oCalendarUS.GetPeriod(Session("sAppServer"), CLng(gnItemId))
		gsCboPeriodStatus = oCalendarUS.CboPeriodStatus()
		gsPeriodName = oRecordset("period_name")
		gsFromDate = oRecordset("from_date")
		gsToDate = oRecordset("to_date")
		gsCboPeriodStatus = oCalendarUS.CboPeriodStatus(oRecordset("is_open"))
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
	if (dDocument.txtPeriodName.value == '') {
		alert("Requiero el nombre del período.");
		dDocument.txtPeriodName.focus();
		return false;
	}
	if (dDocument.txtFromDate.value == '') {
		alert("Requiero la fecha de inicio del período.");
		dDocument.txtFromDate.focus();
		return false;
	}	
	if (!isDate(dDocument.txtFromDate.value)) {
		alert("No reconozco la fecha de inicio proporcionada.");
		dDocument.txtFromDate.focus();
		return false;
	}
	if (dDocument.txtToDate.value == '') {
		alert("Requiero la fecha de finalización del período.");
		dDocument.txtToDate.focus();
		return false;
	}			
	if (!isDate(dDocument.txtToDate.value)) {
		alert("No reconozco la fecha de finalización proporcionada.");
		dDocument.txtToDate.focus();
		return false;
	}
	return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el período "<%=gsPeriodName%>"?')) {
		window.document.frmEditor.action = "./exec/delete_period.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();		
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "./exec/save_period.asp";
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
  <TD valign=top colspan=3><TEXTAREA name=txtPeriodName rows=3 style="WIDTH: 100%"><%=gsPeriodName%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=top>Del día:</TD>
  <TD valign=top colspan=2><INPUT name=txtFromDate value=<%=gsFromDate%> maxlength=11 style="HEIGHT: 22px; WIDTH: 100%"></TD>
  <TD>&nbsp;&nbsp;( día / mes /año )</TD>
</TR>
<TR>
  <TD valign=top>Al día:</TD>
  <TD valign=top colspan=2><INPUT name=txtToDate value=<%=gsToDate%> maxlength=11 style="HEIGHT: 22px; WIDTH: 100%"></TD>
  <TD>&nbsp;&nbsp;( día / mes /año )</TD>
</TR>
<TR>
  <TD valign=center>Estado:</TD>
	<TD colSpan=2>
		<SELECT name=cboIsOpen style="HEIGHT: 22px; WIDTH: 100%">			
			<%=gsCboPeriodStatus%>
		</SELECT>
	</TD>
	<TD>&nbsp;</TD>
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
