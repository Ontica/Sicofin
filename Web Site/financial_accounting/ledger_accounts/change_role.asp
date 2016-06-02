<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gsStdAccountNumber, gsStdAccountName, gsStdAccountDescription, gsAditionalAccounts
	Dim gsStdAccountTypeName, gsRole, gsRoleDescription, gsStdAccountTypeDesc, gsStdAccountNatureName
	Dim gsSectors, gsCurrencies, gsAreas, gnConvertionType, gsParentAccount, gsParentRole, gsCboStdAccountRoles
	Dim gnStdAccountTypeId
	        
  If (Len(Request.QueryString("id")) = 0) Then

  Else
		Call EditItems(CLng(Request.QueryString("id")))
	End If
	
	Sub EditItems(nItemId)
		Dim oStdAccount, oRecordset
		'******************
		gsTitle  = "Conversión del rol de la cuenta"
		gbEdit = True
		gnItemId = CLng(nItemId)
		Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		Set oRecordset = oStdAccount.GetStdAccount(Session("sAppServer"), CLng(nItemId))
		gnStdAccountTypeId = oRecordset("id_tipo_cuentas_std")
		gsStdAccountTypeName = oRecordset("nombre_tipo_cuentas_std")
		gsRole							 = oRecordset("rol_cuenta")
		gsRoleDescription    = oStdAccount.RoleDescription(CStr(gsRole))
		gsCboStdAccountRoles = oStdAccount.CboStdAccountRoles()
		'Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		'gsSectorsList = oStdAccount.SectorsList(Session("sAppServer"), CLng(nItemId), True)
		'Set oStdAccount = Nothing
		gsStdAccountNumber   = Request.Form("txtStdAccountNumber")
		gsStdAccountName     = Request.Form("txtStdAccountName")
	End Sub	

	Function GetAreasListDescription()
		Select Case Request.Form("txtAreas")
			Case ""
				GetAreasListDescription = "La cuenta <font color=maroon><b>no</b></font> manejará áreas de responsabilidad."
			Case "*"
				GetAreasListDescription = "La cuenta manejará <font color=maroon><b>todas</b></font> las áreas de responsabilidad."
			Case Else
				GetAreasListDescription = Replace(Request.Form("txtAreas"), ",", ", ")
		End Select
	End Function
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
					sTemp += document.all[sCheckBoxName](i).tagName + '\n';
				} else {
					sTemp += document.all[sCheckBoxName](i).tagName + '\t(nueva)\n';
				}
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			if (document.all[sCheckBoxName](i).disabled) {
				sTemp += document.all[sCheckBoxName](i).tagName + '\n';
			} else {
				sTemp += document.all[sCheckBoxName](i).tagName + '\t(nueva)\n';
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

function cmdPrevious_onclick() {
	document.frmEditor.action = 'standard_account_editor.asp?id=<%=gnItemId%>';
	document.frmEditor.submit();
}

function cmdNext_onclick() {
  var sMsg = '', nCount = 0;
  
  document.frmEditor.txtStdAccountRole.value  = document.frmEditor.cboStdAccountRole.value;
	if (document.frmEditor.cboStdAccountRole.value == 'X') {
		document.frmEditor.action = 'change_sectors.asp?id=<%=gnItemId%>';
		document.frmEditor.submit();   
		return true;
	}
  if (document.frmEditor.cboStdAccountRole.value == '<%=gsRole%>') {
		alert('El rol actual y el nuevo no deben coincidir.');
		return false;
	}
  sMsg  = '¿Toda la información proporcionada es correcta?\n\n';
  if (!confirm(sMsg)) {
    return false;
  }
  sMsg  = 'IMPORTANTE: A partir de este punto, el proceso será irreversible.\n\n'
  sMsg += '¿Continúo con la inserción de la cuenta?\n\n';
  if (!confirm(sMsg)) {
    return false;
  }
	document.frmEditor.action = 'exec/save_std_account.asp?id=<%=gnItemId%>';
	document.frmEditor.submit();
	return true;
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
  <TD nowrap><b>Catálogo de cuentas:</b></TD>  
  <TD colspan=3><%=gsStdAccountTypeName%></TD>
</TR>
<TR>
  <TD nowrap valign=top><b>Número de cuenta:</b></TD>
  <TD colspan=3 valign=top><%=gsStdAccountNumber%></TD>
</TR>
<TR>
  <TD valign=top><b>Nombre:</b></TD>
	<TD colSpan=3 valign=top><%=gsStdAccountName%></TD>  
</TR>
<TR>
  <TD valign=top><b>Rol actual:</b></TD>
	<TD colSpan=3 valign=top><%=gsRoleDescription%></TD>  
</TR>
<TR>
  <TD valign=top><b>Nuevo rol:</b></TD>
	<TD colSpan=3 valign=top>
		<SELECT name=cboStdAccountRole style="WIDTH:250px">
			<%=gsCboStdAccountRoles%>
		</SELECT>
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
		<INPUT name=txtAreas type="hidden" value="<%=Request.Form("txtAreas")%>">
		<INPUT name=txtSectors type="hidden" value="<%=Request.Form("txtSectors")%>">
		<INPUT name=txtSectorRoles type="hidden" value="<%=Request.Form("txtSectorRoles")%>">
	</TD>
	<TD colspan=3 align=right>
	  <% If gbEdit Then %>
	  <INPUT name=cmdCancel type=button style="WIDTH:100px" value="Cancelar" LANGUAGE=javascript onclick="return cmdPrevious_onclick()">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdSave type=button style="WIDTH:100px" value="Aceptar" LANGUAGE=javascript onclick="return cmdNext_onclick()">
		<% Else %>
		<INPUT name=cmdPrevious style="WIDTH:100px" type=button value="<< Anterior" LANGUAGE=javascript onclick="return cmdPrevious_onclick()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdNext type=button style="WIDTH:100px" value="Aceptar" LANGUAGE=javascript onclick="return cmdNext_onclick()">
    <% End If %>		
	</TD>  
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>