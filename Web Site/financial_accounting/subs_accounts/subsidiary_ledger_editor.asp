<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gnEntityId, gsName, gsDescription, gsCboEntities
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oGralLedgerUS
		gbEdit = False
		gsTitle = "Agregar mayor auxiliar"	
		gnItemId = 0
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		gsCboEntities = oGralLedgerUS.CboSubsidiaryLedgerEntities(Session("sAppServer"), CLng(Request.QueryString("entityId")))
		Set oGralLedgerUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oGralLedgerUS, oRecordset
		'****************************
		gbEdit = True
		gsTitle = "Editar mayor auxiliar"
		gnItemId = CLng(nItemId)		
		Set oGralLedgerUS = Server.CreateObject("AoGralLedgerUS.CServer")	
		Set oRecordset  = oGralLedgerUS.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(nItemId))
		gsCboEntities = oGralLedgerUS.CboSubsidiaryLedgerEntities(Session("sAppServer"), CLng(Request.QueryString("entityId")))
		Set oGralLedgerUS = Nothing
		gsName				= oRecordset("nombre_mayor_auxiliar")
		gsDescription	= oRecordset("descripcion")
		oRecordset.Close
		Set oRecordset = Nothing		
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
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "IsNumeric" , sNumber, nDecimals);
	return obj.return_value;
}

function compareDates(sDate1, sDate2) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CompareDates" , sDate1, sDate2);
	return obj.return_value;
}

function validate() {
	var dDocument = window.document.all;
	
	if (dDocument.txtName.value == '') {
		alert("Requiero el nombre del mayor auxiliar.");
		dDocument.txtFromDate.focus();
		return false;		
	}
	return true;
}

function txtFromDate_onblur() {
	var dDocument = window.document.all;
	
	if ((dDocument.txtFromDate.value != '') && (dDocument.txtToDate == '')) {
		if (isDate(dDocument.txtFromDate.value)) {
			dDocument.txtToDate.value = dDocument.txtFromDate.value;
		}
	}
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el mayor auxiliar "<%=gsName%>"?')) {		
		window.document.frmEditor.action = "exec/delete_subsidiary_ledger.asp?id=<%=gnItemId%>";		
		window.document.frmEditor.submit();		
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "exec/save_subsidiary_ledger.asp";
		window.document.frmEditor.submit();
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO LANGUAGE=javascript onload="window.document.all.txtName.focus();">
<FORM name=frmEditor method="post">
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%" height="100%">
<TR>  
  <TD colSpan=3 bgcolor=khaki><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>
<% If gbEdit Then %>
	<TD bgcolor=khaki align=right><INPUT name=cmdDelete type=button value="Eliminar" LANGUAGE=javascript onclick="return cmdDelete_onclick()"></TD>
<% Else %>
	<TD bgcolor=khaki>&nbsp;</TD>
<% End If %>
</TR>
<TR>
  <TD valign=middle nowrap>Tipo de mayor auxiliar:</TD>
  <TD colspan=3>
		<SELECT name=cboEntities style="HEIGHT: 22px; WIDTH: 100%">
			<%=gsCboEntities%>
		</SELECT>	  
	</TD>  
</TR>
<TR>
  <TD valign=top nowrap>Nombre:</TD>
  <TD valign=top colspan=3 width=100%><TEXTAREA name=txtName rows=3 style="WIDTH: 100%"><%=gsName%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=top nowrap>Descripción:</TD>
  <TD valign=top colspan=3><TEXTAREA name=txtDescription rows=5 style="WIDTH: 100%"><%=gsDescription%></TEXTAREA></TD>
</TR>
<TR>
<TD colspan=2><INPUT name=txtItemId type="hidden" value="<%=gnItemId%>"></TD>
<% If gbEdit Then %>
	<TD><INPUT name=cmdEditItem type=button value="Aceptar" LANGUAGE=javascript onclick="return saveItem();"></TD>
<% Else %>
  <TD><INPUT name=cmdAddItem type=button value="Agregar" LANGUAGE=javascript onclick="return saveItem();"></TD>
<% End If %>
  <TD align=right><INPUT name=cmdCancel type=button value="Cancelar" onclick="window.close();"></TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
