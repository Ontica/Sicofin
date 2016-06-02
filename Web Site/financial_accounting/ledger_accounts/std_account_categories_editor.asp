<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsCategoryName, gsCategoryDescription, gsAccountPattern
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oGralLedgerUS
		'**************************
		gbEdit = False
		gsTitle = "Agregar catálogo estándar"
		gnItemId = 0		
	End Sub
	
	Sub EditItem(nItemId)
		Dim oGralLedgerUS, oRecordset
		'************************************
		gbEdit = True
		gsTitle = "Editar catálogo estándar"
		gnItemId = CLng(nItemId)		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		Set oRecordset = oGralLedgerUS.GetCategoryRS(Session("sAppServer"), CLng(nItemId))
		gsAccountPattern = oRecordset("object_key")
		Set oGralLedgerUS = Nothing
		gsCategoryName				= oRecordset("object_name")
		gsCategoryDescription = oRecordset("object_description")
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
	if (dDocument.txtCategoryName.value == '') {
		alert("Requiero el nombre del catálogo estándar.");
		dDocument.txtAccountNumber.focus();
		return false;
	}
	if (dDocument.txtAccountPattern.value == '') {
		alert("Requiero la estructura que tendrán las cuentas de este catálogo estándar.");
		dDocument.txtName.focus();
		return false;
	}
  return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino el catálogo estándar "<%=gsCategoryName%>"?')) {
		window.document.frmEditor.action = "exec/delete_std_account_categ.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "exec/save_std_account_categ.asp";
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
  <TD valign=top>Nombre del catálogo:</TD>
	<TD colSpan=3><TEXTAREA name=txtCategoryName rows=2 style="width:100%"><%=gsCategoryName%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=top>Descripción:</TD>
  <TD colSpan=3><TEXTAREA name=txtCategoryDescription rows=4 style="width:100%"><%=gsCategoryDescription%></TEXTAREA></TD>
</TR>
<TR>
  <TD valign=center>Estructura de las cuentas:</TD>
  <TD colspan=3><INPUT name=txtAccountPattern value="<%=gsAccountPattern%>" maxlength=256 style="HEIGHT: 22px; WIDTH: 100%"></TD>
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
