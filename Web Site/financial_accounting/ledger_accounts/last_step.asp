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
	Dim gsSectors, gsCurrencies, gsAreas, gnConvertionType, gsParentAccount, gsParentRole
	        
  If (Len(Request.QueryString("id")) = 0) Then
    Call AddItems()
  Else
		Call EditItems(CLng(Request.QueryString("id")))
	End If
	
	Sub AddItems()
		Dim oStdAccount, oRecordset, nParentId
		'*************************************
		gsTitle  = "Asistente para agregar cuentas: lista de verificación"
		gbEdit   = False
		gnItemId = 0
	  Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		'If CBool(Request.Form("txtUseConvertion")) Then
		'	nParentId = oStdAccount.GetParentIdWithNumber(Session("sAppServer"), _
		'																								CLng(Request.Form("txtStdAccountTypeId")), _
		'																								CStr(Request.Form("txtStdAccountNumber")))
		'	gsSectorsList = oStdAccount.SectorsList(Session("sAppServer"), CLng(nParentId), True)
		'Else
		'	gsSectorsList = oStdAccount.SectorsList(Session("sAppServer"))			
		'End If
		
		gsStdAccountTypeName   = Request.Form("txtStdAccountTypeName")
		gsStdAccountNumber     = Request.Form("txtStdAccountNumber")
		gsStdAccountName       = Request.Form("txtStdAccountName")
		gsRole							   = Request.Form("txtStdAccountRole")
		gsRoleDescription      = oStdAccount.RoleDescription(CStr(gsRole))
		gsStdAccountTypeDesc   = oStdAccount.TypeDescription(CStr(Session("sAppServer")), CStr(Request.Form("txtStdAccountType")))
		gsStdAccountNatureName = oStdAccount.NatureDescription(CStr(Request.Form("txtStdAccountNature")))
		gsAditionalAccounts    = oStdAccount.NotAddedAccountsList(CStr(Session("sAppServer")), _
																															CLng(Request.Form("txtStdAccountTypeId")), _
																															CStr(gsStdAccountNumber))
		If Len(gsAditionalAccounts) <> 0 Then
			gsAditionalAccounts  = CStr(gsStdAccountNumber) & ", "
		Else
			gsAditionalAccounts  = CStr(gsStdAccountNumber)
		End If
		If gsRole <> "S" Then
			gsCurrencies = oStdAccount.FormatCurrenciesList(CStr(Session("sAppServer")), CStr(Request.Form("txtCurrencies")))
		  gsAreas      = GetAreasListDescription()
		  If gsRole <> "X" Then
				gsSectors = "La cuenta no manejará sectores debido a que no es sectorizada."
			Else
			  gsSectors = oStdAccount.FormatSectorsList(CStr(Session("sAppServer")), CStr(Request.Form("txtSectors")), CStr(Request.Form("txtSectorRoles")))
			End If
		Else
			gsCurrencies = "Por ser sumaria, la cuenta no manejará monedas."
			gsAreas      = "Por ser sumaria, la cuenta no manejará áreas de responsabilidad."
			gsSectors    = "Por ser sumaria, la cuenta no manejará sectores."
		End If
		
		Select Case gsRole
			Case "S"
				gnConvertionType = 0
			Case Else
				nParentId	= oStdAccount.GetParentIdWithNumber(CStr(Session("sAppServer")), _
																											CLng(Request.Form("txtStdAccountTypeId")), _
																											CStr(gsStdAccountNumber))
				Set oRecordset = oStdAccount.GetStdAccount(CStr(Session("sAppServer")), CLng(nParentId))
				gsParentAccount = oRecordset("numero_cuenta_estandar")
				gsParentRole    =	oRecordset("rol_cuenta")
				gnConvertionType = 1
		End Select
		Set oStdAccount  = Nothing
	End Sub
	
	Sub EditItems(nItemId)
		Dim oStdAccount
		'******************
		gsTitle  = "Edición de cuentas: lista de verificación"
		gbEdit = True
		gnItemId = CLng(nItemId)
		gsCurrencies = "Lista de monedas"
		gsSectors    = "Lista de sectores"
		'Set oStdAccount = Server.CreateObject("EFAStdActUS.CServer")
		'gsSectorsList = oStdAccount.SectorsList(Session("sAppServer"), CLng(nItemId), True)
		'Set oStdAccount = Nothing
		gsStdAccountTypeName = Request.Form("txtStdAccountTypeName")
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
<% If gsRole <> "S" Then %>
	document.frmEditor.action = 'edit_areas.asp';
<% Else %>
	document.frmEditor.action = 'standard_account_editor.asp';
<% End If %>	
	document.frmEditor.submit();
}

function cmdNext_onclick() {
  var sMsg = '', nCount = 0;
  
  sMsg  = '¿Toda la información proporcionada es correcta?\n\n';
  if (!confirm(sMsg)) {
    return false;
  }
  sMsg  = 'IMPORTANTE: A partir de este punto, el proceso será irreversible.\n\n'
  sMsg += '¿Continúo con la inserción de la cuenta?\n\n';
  if (!confirm(sMsg)) {
    return false;
  }
  <% If gbEdit Then %>
		document.frmEditor.action = 'exec/save_std_account.asp?id=<%=gnItemId%>';
	<% Else %>
		document.frmEditor.action = 'exec/add_std_account.asp';
	<% End If %>
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
  <TD nowrap valign=top><%=gsStdAccountNumber%></TD>
  <TD nowrap valign=top><b>Rol de la cuenta:</b></TD>
	<TD nowrap valign=top><%=gsRoleDescription%></TD>  
</TR>
<TR>
  <TD valign=top><b>Nombre:</b></TD>
	<TD colSpan=3 valign=top><%=gsStdAccountName%></TD>  
</TR>
<TR>
  <TD nowrap valign=top><b>Tipo de cuenta:</b></TD>
	<TD nowrap valign=top><%=gsStdAccountTypeDesc%></TD>
  <TD valign=top><b>Naturaleza:</b></TD>
	<TD valign=top><%=gsStdAccountNatureName%></TD>  
</TR>
<TR>
	<TD valign=top><b>Monedas:</b></TD>
	<TD colSpan=3 valign=top><%=gsCurrencies%></TD>
</TR>
<TR>
  <TD nowrap valign=top><b>Sectores:</b></TD>
	<TD colSpan=3 valign=top><%=gsSectors%></TD>  
</TR>
<TR>
  <TD nowrap valign=top><b>Areas de responsabilidad:</b></TD>
	<TD colSpan=3 valign=top><%=gsAreas%></TD>  
</TR>
<TR>
  <TD nowrap valign=top><b>Proceso de conversión:<b></TD>
	<TD colSpan=3 valign=top>
		<% If (gnConvertionType = 0) Then %>
			No será necesaria la conversión de cuentas.
		<% Else %>
			<TABLE align=center border=0 cellPadding=2 cellSpacing=0 width="100%">
			<TR>
				<TD><img src="/empiria/images/exclamation.gif" align=top></TD>
				<TD colspan=3>Se trasladarán los saldos en los mayores de la <font color=maroon><b><%=Lcase(gsStdAccountTypeName)%></b></font> de la siguiente manera:</font></TD>
			</TR>
			<TR>
			  <TD>&nbsp;</TD> 
				<TD valign=top align=right>
					<font color=maroon><b><%=gsParentAccount%></b></font>
					<% If gsParentRole = "X" Then %>
						(a nivel sector)
					<% ElseIf gsParentRole = "C" Then %>
						(a nivel auxiliar)
					<% End If %>
				 </TD>
				<TD valign=top align=center width=1%><img src="/empiria/images/right_arrow.gif" align=top></TD>
				<TD valign=top align=left><font color=maroon><b><%=gsStdAccountNumber%></b></font></TD>
			</TR>
			<TR>
				<TD>&nbsp;</TD>
				<TD colspan=3>Además, las cuentas <font color=maroon>X, T, U y V</font> serán convertidas a sumarias.</TD>
			</TR>
			</TABLE>
		<% End If %>
	</TD>  
</TR>

<TR>
	<TD nowrap valign=top><b>Tareas adicionales:</b></TD>
	<TD colspan=3 valign=top>
		<% If Len(gsAditionalAccounts) <> 0 Then %>
			Se agregarán las siguientes cuentas sumarias en forma automática:<br>
			<b><%=gsAditionalAccounts%></b>.<br><br>
			<img src="/empiria/images/help.gif" align=top> &nbsp;&nbsp;Después de ejecutar el proceso, será necesario cambiar los nombres de dichas cuentas en forma manual.
	  <% Else %>
			No hay tareas adicionales
	  <% End If %>
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
		<INPUT name=txtFromDate type="hidden" value="<%=Request.Form("txtFromDate")%>">
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