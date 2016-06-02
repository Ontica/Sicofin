<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsTitle, gbEdit
	Dim gnItemId, gnStdAccountTypeId, gsStdAccountNumber, gsStdAccountName, gsStdAccountDescription
	Dim gbCanDelete, gsRole, gsRoleDescription, gsStdAccountTypeName, gsCboStdAccountTypes, gsCboStdAccountNature, gsCboStdAccountRoles
	
  If (Len(Request.QueryString("id")) = 0) Then
		If (Request.ServerVariables("REQUEST_METHOD") <> "POST") Then
			Call AddStdAccount(CLng(Request.QueryString("type_id")), False)
		Else
			Call AddStdAccount(CLng(Request.Form("txtStdAccountTypeId")), True)
		End If
	Else
		Call EditStdAccount(CLng(Request.QueryString("id")))
	End If

	Sub AddStdAccount(nStdAccountCatId, bFromWizard)
		Dim oStdAccountMgr
		'**************************
		gbEdit = False
		gsTitle = "Asistente para agregar cuentas"
		gnItemId = 0
		gnStdAccountTypeId = nStdAccountCatId
		Set oStdAccountMgr    = Server.CreateObject("EFAStdActUS.CServer")
		If (Not bFromWizard) Then
			gsStdAccountTypeName  = oStdAccountMgr.GetStdAccountCategory(Session("sAppServer"), CLng(nStdAccountCatId)).Fields("object_name")
			gsCboStdAccountTypes  = oStdAccountMgr.CboStdAccountTypes(Session("sAppServer"))
			gsCboStdAccountNature = oStdAccountMgr.CboStdAccountNature()
			gsCboStdAccountRoles  = oStdAccountMgr.CboStdAccountRoles()
		ElseIf bFromWizard Then
			gsStdAccountNumber      = Request.Form("txtStdAccountNumber")
			gsStdAccountName        = Request.Form("txtStdAccountName")
			gsStdAccountDescription = Request.Form("txtStdAccountDescription")
			gsStdAccountTypeName    = oStdAccountMgr.GetStdAccountCategory(Session("sAppServer"), CLng(nStdAccountCatId)).Fields("object_name")
			gsCboStdAccountTypes    = oStdAccountMgr.CboStdAccountTypes(Session("sAppServer"), CLng(Request.Form("txtStdAccountType")))
			gsCboStdAccountNature   = oStdAccountMgr.CboStdAccountNature(CStr(Request.Form("txtStdAccountNature")))
			gsCboStdAccountRoles    = oStdAccountMgr.CboStdAccountRoles(CStr(Request.Form("txtStdAccountRole")))
		End If
		Set oStdAccountMgr = Nothing		
	End Sub
	
	Sub EditStdAccount(nItemId)
		Dim oStdAccountMgr, oRecordset
		'****************************
		gbEdit = True
		gsTitle = "Asistente para la edición de cuentas"
		gnItemId = CLng(nItemId)		
		Set oStdAccountMgr = Server.CreateObject("EFAStdActUS.CServer")
		Set oRecordset = oStdAccountMgr.GetStdAccount(Session("sAppServer"), CLng(nItemId))
		gnStdAccountTypeId = oRecordset("id_tipo_cuentas_std")
		gsStdAccountTypeName = oRecordset("nombre_tipo_cuentas_std")
		gsRole = oRecordset("rol_cuenta")
		gsRoleDescription = oStdAccountMgr.RoleDescription(CStr(gsRole))
		gsCboStdAccountTypes  = oStdAccountMgr.CboStdAccountTypes(Session("sAppServer"), oRecordset("id_tipo_cuenta"))
		gsCboStdAccountNature = oStdAccountMgr.CboStdAccountNature(oRecordset("naturaleza"))
		Set oStdAccountMgr = Nothing
		gsStdAccountNumber = oRecordset("numero_cuenta_estandar")
		gsStdAccountName   = oRecordset("nombre_cuenta_estandar")
		gsStdAccountDescription = oRecordset("descripcion")
		oRecordset.Close
		Set oRecordset = Nothing
		
		Set oStdAccountMgr = Server.CreateObject("AOGralLedger.CStandardAccount")
		gbCanDelete = oStdAccountMgr.ReadyForDelete(Session("sAppServer"), CLng(nItemId))
		Set oStdAccountMgr = Nothing
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
	obj = RSExecute("../financial_accounting_scripts.asp", "IsDate", sDate);
	return obj.return_value;
}

function isStdAccountNumberValid(sStdAccountNumber) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "IsStdAccountNumberValid", <%=gnStdAccountTypeId%>, sStdAccountNumber);
	return obj.return_value;
}

function validateAppend(sStdAccountNumber) {
	var obj, sMsg, sParentRole, sParentNumber;
	
	if (!isDate(document.all.txtFromDate.value)) {
		sMsg = "No reconozco la fecha de alta de la cuenta.";
		alert(sMsg);
		document.all.txtFromDate.focus();
		return false;
	}
		
	obj = RSExecute("../financial_accounting_scripts.asp", "GetStdAccountId", <%=gnStdAccountTypeId%>, sStdAccountNumber);	
	if (obj.return_value > 0) {
		sMsg = "El número de cuenta proporcionado ya existe en el catálogo\n'<%=gsStdAccountTypeName%>'." 
		alert(sMsg);
		return false;			
	}
	if (obj.return_value < 0) {
		sMsg  = 'El número de cuenta que se desea agregar alguna vez existió en éste catálogo estándar.\n\n';
		sMsg += '¿Recupero dicha cuenta de la historia?'
		if (confirm(sMsg)) {
			alert("Esta operación está en construcción. Gracias.")
			return false;   // regresar true al concluir.
			return true;
		} else {
			return false;
		}
	}
	obj = RSExecute("../financial_accounting_scripts.asp", "GetStdAccountParentNumber", <%=gnStdAccountTypeId%>, sStdAccountNumber);
	sParentNumber = obj.return_value;
	obj = RSExecute("../financial_accounting_scripts.asp", "GetStdAccountParentRole", <%=gnStdAccountTypeId%>, sStdAccountNumber);
	sParentRole = obj.return_value;
	
	if ((document.frmEditor.cboStdAccountRole.value == 'S') && (sParentRole != 'S') && (sParentRole != '')) {
		sMsg  = "La cuenta que se está insertando no puede ser sumaria, ya que la\n";
    sMsg += "cuenta de la que se derivará (la " + sParentNumber + ") no es sumaria.\n\n";
		sMsg += "Sugerencias:\n\n";
		sMsg += "1) Que la cuenta que se está agregando no sea sumaria.\n\n";
		sMsg += "2) Primero intentar el cambio de rol de la cuenta " + sParentNumber + " a sumaria\n"
		sMsg += "    y después agregar la cuenta " + sStdAccountNumber + ".\n\n";
		sMsg += "Gracias.";
		alert(sMsg);
		return false;
	} 
	
	switch (sParentRole) {
		case '':
	 		return true;
    case 'S':
			return true;
    default:
			document.frmEditor.txtUseConvertion.value = true;
			sMsg  = "Proceso de conversión de saldos:\n\n";
      sMsg += "Al final de esta operación, trasladaré los saldos de la cuenta " + sParentNumber + " a la cuenta\n";
      sMsg += sStdAccountNumber + ", en cada una de las contabilidades que empleen dicha cuenta.\n\n";
      sMsg += "¿Continúo con la inserción de la cuenta " +  sStdAccountNumber + "?";
      return (confirm(sMsg));
  }
  document.frmEditor.txtUseConvertion.value = false;
  return false;
}

function validate() {
	var dDocument = window.document.all;
	<% If (Not gbEdit) Then %>
	if (dDocument.txtStdAccountNumber.value == '') {
		alert("Requiero el número de cuenta.");
		dDocument.txtStdAccountNumber.focus();
		return false;
	}
	if (!isStdAccountNumberValid(dDocument.txtStdAccountNumber.value)) {
		alert("No reconozco la estructura del número de cuenta proporcionado.");
		dDocument.txtStdAccountNumber.focus();
		return false;
	}
	<% End If %>
	if (dDocument.txtStdAccountName.value == '') {
		alert("Requiero el nombre de la cuenta.");
		dDocument.txtStdAccountName.focus();
		return false;
	}
	<% If (Not gbEdit) Then %>
		if (!validateAppend(dDocument.txtStdAccountNumber.value)) {
			return false;
		}
	<% End If %>
	
  return true;
}
function txtStdAccountNumber_onblur() {
	var obj;
	gnAccountId = 0;
	if (document.all.txtStdAccountNumber.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountNumber", <%=gnStdAccountTypeId%>, document.all.txtStdAccountNumber.value);
		if (obj.return_value != '') {
			document.all.txtStdAccountNumber.value = obj.return_value;			
		} else {
			alert("No entiendo el formato de la cuenta proporcionada.");
		}
	}
	return true;
}

function cmdNext_onclick() {
	if (validate()) {
		document.frmEditor.txtStdAccountRole.value   = document.frmEditor.cboStdAccountRole.value;
		document.frmEditor.txtStdAccountType.value   = document.frmEditor.cboStdAccountTypes.value;
		document.frmEditor.txtStdAccountNature.value = document.frmEditor.cboStdAccountNature.value;
	  if (document.frmEditor.cboStdAccountRole.value != 'S') {
			document.frmEditor.action = "add_currencies.asp";
		} else {
			document.frmEditor.action = "last_step.asp";
		}
		document.frmEditor.submit();
	}
	return false;
}

function saveBasicInfo() {
	if (validate()) {
		document.frmEditor.action = 'exec/save_basic_info.asp?id=<%=gnItemId%>';
		document.frmEditor.submit();
		return true;
	}
	return false;
}

function viewSectors() {
   alert('Por el momento esta opción no está disponible. Gracias.');
   return false;
}

function cmdExecuteEditTask_onclick() {
   switch (document.frmEditor.cboOperation.value) {
      case '0':  
				alert('Requiero la selección de una tarea de la lista de la izquierda.');
				document.frmEditor.cboOperation.focus();
				return false;
      case '1':
				saveBasicInfo();
				return false;
      case '2':
				document.frmEditor.action = 'add_currencies.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
      case '3':
        alert("Por el momento esta opción no esta disponible. Gracias.");
        return false;
      	document.frmEditor.action = 'delete_currencies.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;  
      case '4':
				document.frmEditor.action = 'add_sectors.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
      case '5':
        alert("Por el momento esta opción no esta disponible. Gracias.");
        return false;
				document.frmEditor.action = 'delete_sectors.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
			case '6':
				alert("Por el momento esta opción no esta disponible. Gracias.");
        return false;
				document.frmEditor.action = 'change_sector_role.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
      case '7':
				document.frmEditor.action = 'change_role.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
      case '8':
				document.frmEditor.action = 'edit_areas.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;
      case '9':
        alert("Por el momento esta opción no esta disponible. Gracias.");
        return false;
 				document.frmEditor.action = 'delete.asp?id=<%=gnItemId%>';
				document.frmEditor.submit();
				return false;				
   }
   return false;
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
  <% If gbEdit Then %>
  <TD colspan=3><b><%=gsStdAccountNumber%></b></TD>
  <% Else %>
  <TD colspan=3><INPUT name=txtStdAccountNumber value="<%=gsStdAccountNumber%>" maxlength=256 style="WIDTH: 400px" LANGUAGE=javascript onblur="return txtStdAccountNumber_onblur()"></TD>
  <% End If %>
</TR>
<TR>
  <TD valign=top>Nombre:</TD>
	<TD colSpan=3 valign=top><TEXTAREA name=txtStdAccountName rows=2 style="WIDTH: 400px"><%=gsStdAccountName%></TEXTAREA></TD>  
</TR>
<TR>
  <TD valign=top>Descripción:</TD>
  <TD colSpan=3 valign=top><TEXTAREA name=txtStdAccountDescription rows=5 style="WIDTH: 400px"><%=gsStdAccountDescription%></TEXTAREA></TD>
</T>
<TR>
  <TD nowrap>Rol de la cuenta:</TD>
  <TD colspan=3>
		<% If gbEdit Then %>
			<b><%=gsRoleDescription%></b>
			<% If gsRole = "X" Then %>
		&nbsp;&nbsp;&nbsp;&nbsp;<A href='' onclick='return viewSectors(); return false'>Ver sectores</A>
			<% End If %>
		<% Else %>
		<SELECT name=cboStdAccountRole style="WIDTH:250px">
			<%=gsCboStdAccountRoles%>
		</SELECT>
		<% End If %>
	</TD>  
</TR>
<TR>
  <TD nowrap>Tipo de cuenta:</TD>
  <TD colspan=3>
		<SELECT name=cboStdAccountTypes style="WIDTH:250px">
			<%=gsCboStdAccountTypes%>
		</SELECT>
	</TD>  
</TR>
<TR>
  <TD>Naturaleza:</TD>
  <TD colspan=3>
		<SELECT name=cboStdAccountNature style="WIDTH:250px">
			<%=gsCboStdAccountNature%>
		</SELECT>
		<INPUT name=txtItemId type="hidden" value="<%=gnItemId%>">
		<INPUT name=txtStdAccountTypeId type="hidden" value="<%=gnStdAccountTypeId%>">
		<INPUT name=txtStdAccountTypeName type="hidden" value="<%=gsStdAccountTypeName%>">
		<% If gbEdit Then %>
		<INPUT name=txtStdAccountNumber type="hidden" value="<%=gsStdAccountNumber%>">
		<INPUT name=txtFromDate type="hidden" value="29/12/2000">
		<% End If %>
		<INPUT name=txtStdAccountRole type="hidden" value="<%=Request.Form("txtStdAccountRole")%>">
		<INPUT name=txtStdAccountType type="hidden" value="<%=Request.Form("txtStdAccountType")%>">
		<INPUT name=txtStdAccountNature type="hidden" value="<%=Request.Form("txtStdAccountNature")%>">
		<INPUT name=txtUseConvertion type="hidden" value="false">
	</TD>
</TR>
<% If (Not gbEdit) Then %>
<TR>
  <TD nowrap>Fecha de alta:</TD>
  <TD colspan=3>
		<INPUT name=txtFromDate value="<%=Date()%>" style="width:100px">
		&nbsp;(día / mes / año)
	</TD> 
</TR>
<% End If %>
<% If gbEdit Then %>
<TR>
	<TD>¿Qué se desea hacer?</TD>
	<TD colspan=3 nowrap>
	    <SELECT name=cboOperation style="WIDTH:300px">
	       <OPTION value=0>-- Lista de tareas --</OPTION>
	       <OPTION value=1>Guardar los cambios efectuados</OPTION>
	       <% If gsRole <> "S" Then %>
	       <OPTION value=2>Agregar monedas a la cuenta</OPTION>
	       <OPTION value=3>Eliminar monedas de la cuenta</OPTION>
	       <% End If %>
	       <% If gsRole = "X" Then %>
	       <OPTION value=4>Agregar sectores a la cuenta</OPTION>
	       <OPTION value=5>Eliminar sectores de la cuenta</OPTION>
	       <OPTION value=6>Modificar el rol de los sectores de la cuenta</OPTION>
	       <% End If %>
	       <OPTION value=7>Modificar el rol de la cuenta</OPTION>
	       <OPTION value=8>Editar las áreas de responsabilidad</OPTION>
	        <% If gbCanDelete Then %>
	       <OPTION value=9>Eliminar la cuenta del catálogo</OPTION>
	        <% End If %>
	    </SELECT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdExecuteEditTask type=button value="Ejecutar"  LANGUAGE=javascript onclick="return cmdExecuteEditTask_onclick()">
	</TD>  
</TR>
<% Else %>
<TR>
	<TD>&nbsp;</TD>
	<TD colspan=3 nowrap align=right>	
		<INPUT name=cmdNext type=button style="WIDTH:100px" value="Siguiente >>" LANGUAGE=javascript onclick="return cmdNext_onclick()">
	</TD>  
</TR>
<% End If %>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>