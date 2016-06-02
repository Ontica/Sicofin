<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsName, gsDescription
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		gbEdit = False
		gsTitle = "Agregar tipo mayor auxiliar"		
		gnItemId = 0
	End Sub
	
	Sub EditItem(nItemId)
		Dim oGralLedgerUS, oRecordset
		'****************************
		gbEdit = True
		gsTitle = "Editar tipo mayor auxiliar"		
		gnItemId = CLng(nItemId)		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		Set oRecordset = oGralLedgerUS.GetEntityRS(Session("sAppServer"), CLng(nItemId))
		Set oGralLedgerUS = Nothing
		gsName   = oRecordset("entity_name")
		gsDescription = oRecordset("description")
		oRecordset.Close
		Set oRecordset = Nothing		
	End Sub	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="../resources/pages_style.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function validate() {
	var dDocument = window.document.all;
	if (dDocument.txtName.value == '') {
		alert("Requiero el nombre del tipo de mayor auxiliar.");
		dDocument.txtName.focus();
		return false;
	}
  return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el tipo de mayor auxiliar "<%=gsName%>"?')) {
		window.document.frmEditor.action = "../programs/delete_entity.asp?id=<%=gnItemId%>";		
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "../programs/save_entity.asp";
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
  <TD colSpan=2 bgcolor=khaki><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>
<% If gbEdit Then %>
	<TD colSpan=2 bgcolor=khaki align=right><INPUT name=cmdDelete type=button value="Eliminar" LANGUAGE=javascript onclick="return cmdDelete_onclick()"></TD>
<% Else %>
	<TD colSpan=2 bgcolor=khaki>&nbsp;</TD>
<% End If %>
</TR>
<TR>
  <TD valign=top>Nombre:</TD>
  <TD colSpan=3><TEXTAREA name=txtName rows=2 style="WIDTH: 100%"><%=gsName%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=top>Descripción:</TD>
  <TD colSpan=3><TEXTAREA name=txtDescription rows=4 style="WIDTH: 100%"><%=gsDescription%></TEXTAREA></TD>
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
