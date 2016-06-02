<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsTitle, gbEdit
	Dim gnItemId, gsName, gsClave, gsDescription
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		gbEdit = False
		gsTitle = "Agregar sector"		
		gnItemId = 0
	End Sub
	
	Sub EditItem(nItemId)
		Dim oGralLedgerUS, oRecordset
		'******************
		gbEdit = True
		gsTitle = "Editar sector"		
		gnItemId = CLng(nItemId)		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		Set oRecordset = oGralLedgerUS.GetSectorRS(Session("sAppServer"), CLng(nItemId))
		Set oGralLedgerUS = Nothing
		gsName        = oRecordset("nombre_sector")
		gsClave       = oRecordset("clave_sector")
		gsDescription = oRecordset("descripcion")
		oRecordset.Close
		Set oRecordset = Nothing
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
	if (dDocument.txtName.value == '') {
		alert("Requiero el nombre del sector.");
		dDocument.txtName.focus();
		return false;
	}
	if (dDocument.txtClave.value == '') {
		alert("Requiero la clave o número correspondiente al sector.");
		dDocument.txtClave.focus();
		return false;
	}		
  return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el sector "<%=gsName%>"?')) {
		window.document.frmEditor.action = "exec/delete_sector.asp?id=<%=gnItemId%>";		
		window.document.frmEditor.submit();		
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "exec/save_sector.asp";
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
  <TD valign=center>Nombre:</TD>
  <TD colSpan=3><INPUT name=txtName value="<%=gsName%>" maxlength=64 style="HEIGHT: 22px; WIDTH: 100%"></TD>
</TR>
<TR>
  <TD valign=center>Clave:</TD>
  <TD colSpan=3><INPUT name=txtClave value="<%=gsClave%>" maxlength=3 style="HEIGHT: 22px; WIDTH: 50px"></TD>
</TR>
<TR>
  <TD valign=top>Descripción:</TD>
  <TD colSpan=3><TEXTAREA name=txtDescription rows=4 style="width:100%"><%=gsDescription%></TEXTAREA></TD>
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
