<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsStdAccountNumber, gsStdAccountName, gsStdAccountDescription
	Dim gsStdAccountTypeName, gsAreasList, gsRole
	
  If (Len(Request.QueryString("id")) = 0) Then 
    Call AddItems()
  Else
		Call EditItems(Clng(Request.QueryString("id")))
	End If
	
	Sub AddItems()
		Dim oStdAccount, nParentId
		'*************************
		gsTitle  = "Asistente para agregar cuentas: áreas"
		gbEdit   = False
		gnItemId = 0

	  If Len(Request.Form("txtAreas")) <> 0 Then
			gsAreasList = Request.Form("txtAreas")
		ElseIf Len(Request.Form("txtAreas")) = 0 Then 
			If CBool(Request.Form("txtUseConvertion")) Then
				Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
				nParentId = oStdAccount.GetParentIdWithNumber(Session("sAppServer"), _
																										  CLng(Request.Form("txtStdAccountTypeId")), _
																										  CStr(Request.Form("txtStdAccountNumber")))
				gsAreasList = oStdAccount.ResponsibilityAreasList(Session("sAppServer"), CLng(nParentId))
				Set oStdAccount = Nothing
			Else
				gsAreasList = ""
			End If
		End If
		
		gsRole							 = Request.Form("txtStdAccountRole")
		gsStdAccountTypeName = Request.Form("txtStdAccountTypeName")
		gsStdAccountNumber   = Request.Form("txtStdAccountNumber")
		gsStdAccountName     = Request.Form("txtStdAccountName")
	End Sub
	
	Sub EditItems(nItemId)
		Dim oStdAccount
		'*******************
		gsTitle  = "Edición de áreas de responsabilidad"
		gbEdit = True
		gnItemId = CLng(nItemId)
		Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		gsAreasList = oStdAccount.ResponsibilityAreasList(Session("sAppServer"), CLng(nItemId))
		Set oStdAccount = Nothing
		gsRole							 = Request.Form("txtStdAccountRole")
		gsStdAccountTypeName = Request.Form("txtStdAccountTypeName")
		gsStdAccountNumber   = Request.Form("txtStdAccountNumber")
		gsStdAccountName     = Request.Form("txtStdAccountName")
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

function countCheckBoxes(sCheckBoxName) {
	var i= 0, counter = 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				counter++;
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			counter++;
		}
	}
	return counter;
}

function selectCheckBoxes(sCheckBoxName, bCheck) {
	var i= 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
		  if (!document.all[sCheckBoxName](i).disabled) {
				document.all[sCheckBoxName](i).checked = bCheck;
			}
		}		
	} else {
		if (!document.all[sCheckBoxName](i).disabled) {
			document.all[sCheckBoxName].checked = bCheck;
		}
	}
	return true;	
}

function checkedBoxes(sCheckBoxName) {
	var sTemp = '', i = 0;
	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				if (document.all[sCheckBoxName](i).disabled) {
				sTemp += '  ' + document.all[sCheckBoxName](i).tagName + '\n';
				} else {
				sTemp += '* ' + document.all[sCheckBoxName](i).tagName + '\n';
				}
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			if (document.all[sCheckBoxName](i).disabled) {
				sTemp += '  ' + document.all[sCheckBoxName](i).tagName + '\n';
			} else {
				sTemp += '* ' + document.all[sCheckBoxName](i).tagName + '\n';
			}
		}
	}
	return sTemp;
}

function setFormArray(sCheckBoxName) {
	var sTemp = '', i = 0;
	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
			   if (sTemp == '') {
						sTemp = document.all[sCheckBoxName](i).value;
				 } else {
						sTemp += ',' + document.all[sCheckBoxName](i).value;				 
				 }
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			if (sTemp == '') {
			 	sTemp = document.all[sCheckBoxName](i).value;
			} else {
			 	sTemp += ',' + document.all[sCheckBoxName](i).value;				 
			}		
		}
	}
	return sTemp;
}

function validateAreas() {
	var obj, sAreas = '';
	
	refreshAreasList(document.frmEditor.txtAreas.value, false);
	obj = RSExecute("../financial_accounting_scripts.asp", "ValidateAreas", document.frmEditor.txtAreas.value);
	return (obj.return_value);
}

function refreshAreasList(sList, bIncludeCheckBoxes) {
	var obj;
	if (sList != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatWildCharsList", sList);
		sList = obj.return_value;
	}
	document.frmEditor.txtAreas.value = sList;
	obj = RSExecute("../financial_accounting_scripts.asp", "GetAreasList", sList, bIncludeCheckBoxes);
	if (bIncludeCheckBoxes) {
		document.frmEditor.txtAreas.value = '';
	}
	window.divAreasList.innerHTML = obj.return_value;
	return true;
}

function cmdPrevious_onclick() {
  <% If gsRole <> "X" Then %>
		document.frmEditor.action = 'add_currencies.asp';
	<% Else %>
		document.frmEditor.action = 'add_sectors_roles.asp';
	<% End If %>
	document.frmEditor.submit();
}

function cmdNext_onclick() {
  if (!validateAreas()) {
    alert("problems.");
		return false;
  }
<% If (gnItemId = 0) Then %> 
	 document.frmEditor.action = 'last_step.asp';
<% Else %>
  document.frmEditor.action = 'exec/save_areas.asp?id=<%=gnItemId%>';
<% End If %>
	document.frmEditor.submit();
	return true;
}

function checkList() {
  var sMsg = '';
  
  refreshAreasList(document.frmEditor.txtAreas.value, false);
  if (document.frmEditor.txtAreas.value == '') {
		sMsg = 'La cuenta no manejará áreas de responsabilidad.';
		alert(sMsg);
		return false;
	}
  if (document.frmEditor.txtAreas.value == '*') {
		sMsg = 'La cuenta manejará todas las áreas de responsabilidad.';
		alert(sMsg);
		return false;
  }
  sMsg = 'La cuenta manejará sólo las áreas de responsabilidad que aparecen en la lista de la derecha.';
  alert(sMsg);
	return false;
}

function chkAreas_onclick() {
  var oSource = window.event.srcElement;
  
  if (oSource.checked) {
    if (document.frmEditor.txtAreas.value == '') {
			document.frmEditor.txtAreas.value = oSource.value;
		} else {
			document.frmEditor.txtAreas.value += "," + oSource.value;
		}
	} else {
		document.frmEditor.txtAreas.value += oSource.value;
	}
}

function window_onload() {
	refreshAreasList('<%=gsAreasList%>', false);
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FORM name=frmEditor action="" method="post">
<TABLE align=center border=0 cellPadding=3 cellSpacing=0 width="550px">
<TR bgcolor=khaki>  
  <TD	nowrap><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>
	<TD colspan=3 align=right nowrap>		
		<A href='' onclick='window.close();return false;'>Cerrar</A>
	</TD>	
</TR>
</TABLE>
<TABLE align=center border=1 cellPadding=3 cellSpacing=0 width="550px">
<TR>
  <TD nowrap>Catálogo de cuentas:</TD>  
  <TD colspan=3><b><%=gsStdAccountTypeName%></b></TD>
</TR>
<TR>
  <TD nowrap>Número de cuenta:</TD>
  <TD colspan=3><b><%=gsStdAccountNumber%></b></TD>
</TR>
<TR>
  <TD valign=top>Nombre:</TD>
	<TD colSpan=3 valign=top><b><%=gsStdAccountName%></b></TD>  
</TR>
<TR>
  <TD valign=top>
		Areas de responsabilidad:<br>
	</TD>
	<TD colSpan=3 valign=top>
		<INPUT name=txtAreas style="WIDTH: 250px" value="<%=gsAreasList%>">
		<A href='' onclick='refreshAreasList(document.frmEditor.txtAreas.value, false);return false;'>Actualizar la lista de áreas</A><br>
		<b>Tips:</b> Si fuera el caso, emplear comas para separar las áreas (sin dejar espacios).<br> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		 Es posible emplear los <A href ='hola.html' target='_blank'>comodines</A> ( * y ? ) para describir múltiples áreas.
	</TD>  
</TR>
<TR>
  <TD valign=top rowspan=2 nowrap>
    <br><br>
		<A href='' onclick='refreshAreasList("*", false); return false;'>Todas las áreas</A><br><br>
		<A href='' onclick='refreshAreasList("*", true); return false'>Selecccionar de una lista</A><br><br>
		<A href='' onclick='refreshAreasList("", false); return false;'>Ningún área</A><br><br><br>
		<A href='' onclick='checkList();return false;'>Comprobar selección</A><br><br>
	</TD>
	<TD nowrap width=100%>
		 <img src='/empiria/images/help.gif' align=top> &nbsp;&nbsp;&nbsp;&nbsp;Las áreas que aparecen en la siguiente lista son las que manejará la cuenta.
	</TD>
</TR>
<TR>
	<TD colSpan=3 valign=top>
	   <DIV ID=divAreasList STYLE="overflow:auto; float:bottom; width=100%; height=150px">
	   <TABLE border=0 cellPadding=1 cellSpacing=0 width=100%>
	      
	   </TABLE>
	   </DIV>
	</TD>
</TR>
<TR>
	<TD>
		<INPUT name=txtItemId type="hidden" value="<%=gnItemId%>">
		<INPUT name=txtStdAccountTypeId type="hidden" value="<%=Request.Form("txtStdAccountTypeId")%>">
		<INPUT name=txtStdAccountTypeName type="hidden" value="<%=gsStdAccountTypeName%>">
		<INPUT name=txtStdAccountNumber type="hidden" value="<%=gsStdAccountNumber%>">
		<INPUT name=txtStdAccountName type="hidden" value="<%=gsStdAccountName%>">
		<INPUT name=txtStdAccountDescription type="hidden" value="<%=Request.Form("txtStdAccountDescription")%>">
		<INPUT name=txtStdAccountRole type="hidden" value="<%=Request.Form("txtStdAccountRole")%>">
		<INPUT name=txtStdAccountType type="hidden" value="<%=Request.Form("txtStdAccountType")%>">
		<INPUT name=txtStdAccountNature type="hidden" value="<%=Request.Form("txtStdAccountNature")%>">
		<INPUT name=txtUseConvertion type="hidden" value="<%=Request.Form("txtUseConvertion")%>">
		<INPUT name=txtCurrencies type="hidden" value="<%=Request.Form("txtCurrencies")%>">
		<INPUT name=txtSectors type="hidden" value="<%=Request.Form("txtSectors")%>">
		<INPUT name=txtSectorRoles type="hidden" value="<%=Request.Form("txtSectorRoles")%>">
		<INPUT name=txtFromDate type="hidden" value="<%=Request.Form("txtFromDate")%>">
	</TD>
	<TD colspan=3 nowrap align=right>
	  <% If gbEdit Then %>
	  <INPUT name=cmdCancel type=button style="WIDTH:100px" value="Cancelar" LANGUAGE=javascript onclick="window.location.href = 'standard_account_editor.asp?<%=Request.QueryString%>'">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdSave type=button style="WIDTH:100px" value="Aceptar" LANGUAGE=javascript onclick="return cmdNext_onclick()">
		<% Else %>
		<INPUT name=cmdPrevious style="WIDTH:100px" type=button value="<< Anterior" LANGUAGE=javascript onclick="return cmdPrevious_onclick()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdNext type=button style="WIDTH:100px" value="Siguiente >>" LANGUAGE=javascript onclick="return cmdNext_onclick()">
    <% End If %>		
	</TD>  
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>