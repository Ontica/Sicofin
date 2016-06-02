<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsName, gsAbbrev, gsSymbol, gsReportOnly, gsClave, gsEditExchangeRate
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		gbEdit = False
		gsTitle = "Agregar moneda"		
		gnItemId = 0
	End Sub
	
	Sub EditItem(nItemId)
		Dim oCurrenciesUS, oRecordset
		'******************
		gbEdit = True
		gsTitle = "Editar moneda"		
		gnItemId = CLng(nItemId)		
		Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")
		Set oRecordset = oCurrenciesUS.GetCurrencyRS(Session("sAppServer"), CLng(nItemId))
		Set oCurrenciesUS = Nothing
		gsName   = oRecordset("currency_name")
		gsAbbrev = oRecordset("abbrev")
		gsSymbol = oRecordset("symbol")
		If CLng(oRecordset("Report_Only")) <> 0 Then
			gsReportOnly  = "checked"
		End If
		If CLng(oRecordset("Edit_Exchange_Rate")) <> 0 Then
			gsEditExchangeRate  = "checked"
		End If				
		gsClave  = oRecordset("O_ID_MONEDA")
		oRecordset.Close
		Set oRecordset = Nothing		
	End Sub	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function existsCurrencyKey(sKey) {
	var obj;
	obj = RSExecute("../general_scripts.asp", "ExistsCurrencyKey" , sKey);
	return obj.return_value;
}

function validate() {
	var dDocument = window.document.all;
	if (dDocument.txtName.value == '') {
		alert("Requiero el nombre de la moneda.");
		dDocument.txtName.focus();
		return false;
	}
	if (dDocument.txtAbbrev.value == '') {
		alert("Requiero la abreviatura que corresponde a la moneda.");
		dDocument.txtAbbrev.focus();
		return false;
	}	
	if (dDocument.txtClave.value == '') {
		alert("Requiero la clave que corresponderá a la moneda.");
		dDocument.txtClave.focus();
		return false;
	}
	<% If Not gbEdit Then %>
	if (existsCurrencyKey(dDocument.txtClave.value)) {
		alert("La clave proporcionada corresponde a otra moneda.");
		dDocument.txtClave.focus();
		return false;
	}
	<% End If %>
  return true;
}

function cmdDelete_onclick() {
	if (confirm('¿Elimino la moneda "<%=gsName%>"?')) {		
		window.document.frmEditor.action = "./exec/delete_currency.asp?id=<%=gnItemId%>";		
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "./exec/save_currency.asp";
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
  <TD colSpan=3><INPUT name=txtName value="<%=gsName%>" maxlength=48 style="HEIGHT: 22px; WIDTH: 100%"></TD>
</TR>
<TR>
  <TD valign=center>Abreviatura:</TD>
  <TD colSpan=3><INPUT name=txtAbbrev value="<%=gsAbbrev%>" maxlength=4 style="HEIGHT: 22px; WIDTH: 50px"></TD>
</TR>
<TR>
  <TD valign=center>Símbolo:</TD>
  <TD colSpan=3><INPUT name=txtSymbol value="<%=gsSymbol%>" maxlength=2 style="HEIGHT: 22px; WIDTH: 50px"></TD>
</TR>
<TR>
  <TD valign=center>Clave:</TD>
  <TD colSpan=3><INPUT name=txtClave value="<%=gsClave%>" maxlength=2 style="HEIGHT: 22px; WIDTH: 50px"></TD>
</TR>
<TR>
  <TD valign=center>Deshabilitada:</TD>
  <TD><INPUT type="checkbox" name=chkReportOnly <%=gsReportOnly%>></TD>
  <TD valign=center>Editar tipo de cambio:</TD>
  <TD><INPUT type="checkbox" name=chkEditExchangeRate <%=gsEditExchangeRate%>></TD>  
</TR>
<TR>

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
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
