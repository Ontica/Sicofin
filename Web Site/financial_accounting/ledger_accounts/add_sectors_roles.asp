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
    Call AddItems()
  Else
		Call EditItems(CLng(Request.QueryString("id")))
	End If
	
	Sub AddItems()
		Dim oStdAccount, nParentId
		'*************************
		gsTitle  = "Asistente para agregar cuentas: rol por sector"
		gbEdit   = False
		gnItemId = 0
	
	  Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		gsSectorsList = oStdAccount.SectorsListWithArray(Session("sAppServer"), CStr(Request.Form("txtSectors")))
		Set oStdAccount  = Nothing
		gsStdAccountTypeName = Request.Form("txtStdAccountTypeName")
		gsStdAccountNumber   = Request.Form("txtStdAccountNumber")
		gsStdAccountName     = Request.Form("txtStdAccountName")
	End Sub
	
	Sub EditItems(nItemId)
		Dim oStdAccount
		'*******************
		gsTitle  = "Agregar sectores a la cuenta: rol por sector"
		gbEdit = True
		gnItemId = CLng(nItemId)
		Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
    gsSectorsList = oStdAccount.SectorsListWithArray(Session("sAppServer"), CStr(Request.Form("txtSectors")))
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
				sTemp += document.all[sCheckBoxName](i).tagName + '\n';
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			sTemp = document.all[sCheckBoxName].tagName + '\n';
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
			 	sTemp = document.all[sCheckBoxName].value;
			} else {
			 	sTemp += ',' + document.all[sCheckBoxName].value;				 
			}		
		}
	}
	return sTemp;
}

function cmdPrevious_onclick() {
	document.frmEditor.action = 'add_sectors.asp';
	<% If (gnItemId <> 0) Then %>
		document.frmEditor.action += '?<%=Request.QueryString%>';
	<% End If %>
	document.frmEditor.submit();
}

function cmdNext_onclick() {
 <% If (gnItemId = 0) Then %>
 	 document.frmEditor.action = 'edit_areas.asp';
 <% Else %>
   document.frmEditor.action = 'exec/save_sectors.asp?id=<%=gnItemId%>';
 <% End If %>
 document.frmEditor.txtSectorRoles.value = setFormArray('chkSectors');
 document.frmEditor.submit();
 return true;
}

function checkList() {
  var sMsg = '', nCount = 0;
  
  nCount = countCheckBoxes('chkSectors');
  switch (nCount) {
		case 0: 
			sMsg = 'Ninguno de los sectores será manejado a nivel de auxiliar.\n\n';
			break;
		case 1:
			sMsg  = 'La cuenta <%=gsStdAccountNumber%> manejará a nivel de auxiliar el siguiente sector:\n\n';
			sMsg += checkedBoxes('chkSectors') + '\n';
			break;
		default:
			sMsg  = 'La cuenta <%=gsStdAccountNumber%> manejará los siguientes sectores a nivel de auxiliar:\n\n';
			sMsg += checkedBoxes('chkSectors') + '\n';
			break;
  }
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
  <TD nowrap>Catálogo de cuentas:&nbsp;&nbsp;</TD>  
  <TD colspan=3><b><%=gsStdAccountTypeName%></b></TD>
</TR>
<TR>
  <TD nowrap>Número de cuenta:</TD>
  <TD colspan=3><b><%=gsStdAccountNumber%></b></TD>
</TR>
<TR>
  <TD valign=top nowrap>Nombre:</TD>
	<TD colSpan=3 valign=top><b><%=gsStdAccountName%></b></TD>  
</TR>
<TR>
  <TD valign=top rowspan=2>
		<br><br><br><br>
		<A href='' onclick='selectCheckBoxes("chkSectors", true); return false'>Seleccionar todos</A><br><br>
		<A href='' onclick='selectCheckBoxes("chkSectors", false); return false'>Deseleccionar todos</A><br><br>
	  <A href='' onclick='checkList();return false;'>Comprobar selección</A><br><br>
	</TD>
	<TD width=100%>
		<img src='/empiria/images/exclamation.gif' align=top>
			<font color=maroon><b>¿Qué sectores se manejarán a nivel de auxiliar?</b></font>
	</TD>
</TR>
<TR>
	<TD colSpan=3 valign=top>
	   <DIV STYLE="overflow:auto; float:bottom; width=100%; height=200px">
	   <TABLE border=0 cellPadding=1 cellSpacing=0 width=100%>
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
		<INPUT name=txtSectorRoles type="hidden" value="<%=Request.Form("txtSectorRoles")%>">
		<INPUT name=txtFromDate type="hidden" value="<%=Request.Form("txtFromDate")%>">
	</TD>
	<TD colspan=3 nowrap align=right>
	  <% If gbEdit Then %>
	  <INPUT name=cmdCancel type=button style="WIDTH:100px" value="Cancelar" LANGUAGE=javascript onclick="window.location.href = 'standard_account_editor.asp?<%=Request.QueryString%>'">
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  	  <INPUT name=cmdPrevious style="WIDTH:100px" type=button value="<< Anterior" LANGUAGE=javascript onclick="return cmdPrevious_onclick()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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