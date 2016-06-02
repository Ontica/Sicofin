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
	Dim gsStdAccountTypeName, gsSectorsList
	
  If (Len(Request.QueryString("id")) = 0) Then

  Else
		Call EditItems(CLng(Request.QueryString("id")))
	End If
	
	Sub EditItems(nItemId)
		Dim oStdAccount
		'*******************
		gsTitle  = "Conversión del rol de la cuenta: sectores"
		gbEdit = True
		gnItemId = CLng(nItemId)
		Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		gsSectorsList = oStdAccount.ChildsSectorsList(Session("sAppServer"), CLng(nItemId), True)
		Set oStdAccount = Nothing
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

function countCheckBoxes(sCheckBoxName, bCountDisabled) {
	var i= 0, counter = 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				counter++;
				if (!bCountDisabled && document.all[sCheckBoxName](i).disabled) {
					counter--;	
				}
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			counter++;
			if (!bCountDisabled && document.all[sCheckBoxName].disabled) {
				counter--;	
			}			
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
		if (!document.all[sCheckBoxName].disabled) {
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
				sTemp += '   ' + document.all[sCheckBoxName](i).tagName + '\n';
				} else {
				sTemp += '*  ' + document.all[sCheckBoxName](i).tagName + '\n';
				}
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			if (document.all[sCheckBoxName].disabled) {
				sTemp += '   ' + document.all[sCheckBoxName].tagName + '\n';
			} else {
				sTemp += '* ' + document.all[sCheckBoxName].tagName + '\n';
			}
		}
	}
	return sTemp;
}

function  addFormArrayItem(sInput, sValue) {
	if (sInput == '') {
		return sValue;
	} else {
		return (sInput + ',' + sValue);
	}
}

function setFormArray(sCheckBoxName, bCountDisabled) {
	var sTemp = '', i = 0;
	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				if (bCountDisabled) {
				  sTemp = addFormArrayItem(sTemp, document.all[sCheckBoxName](i).value);
				} else {
					if (!bCountDisabled && !document.all[sCheckBoxName](i).disabled) {
						sTemp = addFormArrayItem(sTemp, document.all[sCheckBoxName](i).value);
				  }
				}
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			if (bCountDisabled) {
			  sTemp = addFormArrayItem(sTemp, document.all[sCheckBoxName].value);
			} else {
				if (!bCountDisabled && !document.all[sCheckBoxName].disabled) {
					sTemp = addFormArrayItem(sTemp, document.all[sCheckBoxName].value);
			  }
			}
		}
	}
	return sTemp;
}

function cmdPrevious_onclick() {
	document.frmEditor.action = 'change_role.asp?id=<%=gnItemId%>';
	document.frmEditor.submit();
}

function cmdNext_onclick() {
  var nCount = 0;
  
  nCount = countCheckBoxes('chkSectors', true);
	if (nCount == 0) {
	  alert("Requiero que la cuenta maneje al menos un sector.");
	  return false;
  }
  document.frmEditor.txtSectors.value = setFormArray('chkSectors', true);
  
	document.frmEditor.action = 'change_sectors_roles.asp?id=<%=gnItemId%>';
	document.frmEditor.submit();
return true;
}

function checkList() {
  var sMsg = '', nCount = 0;
  
<% If (gnItemId = 0) Then %>  
  nCount = countCheckBoxes('chkSectors', true);
<% Else %>
  nCount = countCheckBoxes('chkSectors', false);
<% End If %>
	if (nCount == 0) {
	<% If (gnItemId = 0) Then %>
		sMsg  = 'No se ha seleccionado ningun sector. \n\n';
		sMsg += 'Nota: Requiero que la cuenta maneje al menos un sector.';
  <% Else %>
		sMsg  = 'No se ha agregado ningun sector. \n\n';  
  <% End If %>
	  alert(sMsg);
	  return false;
  }
  nCount = countCheckBoxes('chkSectors', true);
  if (nCount > 1) {
		sMsg = 'La cuenta <%=gsStdAccountNumber%> manejará los siguientes sectores:\n\n';
	} else {
		sMsg = 'La cuenta <%=gsStdAccountNumber%> manejará el sector:\n\n';
	}
	
  sMsg += checkedBoxes('chkSectors') + '\n';
  sMsg += '(*) Sectores nuevos.\n\n';
  alert(sMsg);
}
//-->
</SCRIPT>
</HEAD>
<BODY>
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
  <TD valign=top rowspan=2>
    <br><br><br><br>
		<A href='' onclick='selectCheckBoxes("chkSectors", true); return false'>Seleccionar todos</A><br><br>
		<A href='' onclick='selectCheckBoxes("chkSectors", false); return false'>Deseleccionar todos</A><br><br>
		<A href='' onclick='checkList();return false;'>Comprobar selección</A><br><br>
	</TD>
	<TD>
		<% If CBool(Request.Form("txtUseConvertion")) Then %>
		 <img src='/empiria/images/help.gif' align=top> Los sectores preseleccionados serán empleados en el proceso de conversión.
		<% Else %>
		 <img src='/empiria/images/help.gif' align=top> Para eliminar los sectores de una cuenta se debe seleccionar la 
		 opción correspondiente en el menú tareas de la página anterior.
		<% End If %>
	</TD>
</TR>
<TR>
	<TD colSpan=3 valign=top>
	   <DIV STYLE="overflow:auto; float:bottom; width=100%; height=200px">
	   <TABLE border=0 cellPadding=1 cellSpacing=0 width=100%>
			 <TR bgcolor=khaki>
				<TD><font color=maroon><b> &nbsp;Sector</b></font></TD>
				<TD><font color=maroon nowrap><b>Rol &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></TD>
			 </TR>	
	      <%=gsSectorsList%>	
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
	</TD>
	<TD colspan=3 nowrap align=right>
	  <% If gbEdit Then %>
	  <INPUT name=cmdCancel type=button style="WIDTH:100px" value="Cancelar" LANGUAGE=javascript onclick="window.location.href = 'standard_account_editor.asp?<%=Request.QueryString%>'">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <INPUT name=cmdPrevious style="WIDTH:100px" type=button value="<< Anterior" LANGUAGE=javascript onclick="return cmdPrevious_onclick()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdSave type=button style="WIDTH:100px" value="Siguiente >>" LANGUAGE=javascript onclick="return cmdNext_onclick()">
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