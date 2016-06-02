<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleId, gsNumber, gsName, gsFromAccount, gsToAccount, gsFromSubsidiaryAccount
	Dim gsRestriction, gsCboCurrencies, gsCboSectors, gsCboClips, gsCboOperators, chkCurrencies, chkSectors
	Dim gsFactor, gnStdAccountTypeId, gsTitle
		
	Call Main()

	Sub Main()
		Dim oRule, oRuleDef, oRecordset, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'*************************************************************************
		On Error Resume Next
		gnRuleId  = Request.QueryString("id")
		gsTitle	  = "Agrupación de cuentas"		
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(gnRuleId))
		gsNumber								= oRecordset("numero_grupo_cuenta")
		gsName									= oRecordset("nombre_grupo_cuenta")
		gsFromAccount						= oRecordset("cuenta_origen_inicial")
		gsToAccount							= oRecordset("cuenta_origen_final")
		gsRestriction					  = oRecordset("filtro_cuentas_origen")
		gsFromSubsidiaryAccount = oRecordset("auxiliar_origen")
		If CLng(oRecordset("id_moneda_origen")) < 0 Then
			chkCurrencies = "checked"
		End If
		If CLng(oRecordset("id_sector_origen")) < 0 Then
			chkSectors = "checked"
		End If
		gsCboCurrencies			= oRule.CboCurrencies(Session("sAppServer"), CLng(oRecordset("id_moneda_origen")))
		gsCboSectors				= oRule.CboSectors(Session("sAppServer"), CLng(oRecordset("id_sector_origen")))
		gsCboOperators			= oRule.CboOperators(Session("sAppServer"), CStr(oRecordset("operador")))
		If IsNull(oRecordset("clip_cuenta_origen")) Then
			gsCboClips					= oRule.CboClips(Session("sAppServer"))
		Else
			gsCboClips					= oRule.CboClips(Session("sAppServer"), CStr(oRecordset("clip_cuenta_origen")))

		End If
		gsFactor						= oRecordset("factor")
		Set oRule  = Nothing
		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gnStdAccountTypeId  = oRuleDef.RuleDefStdAccountTypeId(Session("sAppServer"), CLng(oRecordset("id_regla_contable")))
		Set oRecordset = Nothing
		Set oRuleDef = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Rango de cuentas</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function deleteRule() {
	if (confirm('¿Elimino el rango de cuentas?')) {
		obj = RSExecute("../financial_accounting_scripts.asp","DeleteRule", <%=gnRuleId%>);
		window.close();
	}
	return false
}

function displayPicker(sPickerName, oTarget) {
	var sURL = "../../pickers/", sPars = "resizable:0;status:0;";
	var retValue;
	
	switch (sPickerName) {
		case 'restriction':
			sURL  = 'restriction_picker.asp';
			sPars = 'dialogHeight:310px;dialogWidth:400px;';
			retValue = window.showModalDialog(sURL, '' , sPars);
			if (retValue != 'undefined') {				
				document.all.txtRestriction.value = retValue;
			}
			return false;
		default:
			alert('No tengo definda la ventana solicitada.'); 
			return false;
	}	
	//oTarget.value = window.showModalDialog(sURL, "" , sPars);
	return true;	
}

function setAccountNumber(oControl) {
	var obj;	
	if (oControl.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountNumber", <%=gnStdAccountTypeId%> , oControl.value);
		if (obj.return_value != '') {
			oControl.value = obj.return_value;
		} else {
			alert("No entiendo el formato de la cuenta proporcionada.");
		}
	}
	return true;
}

function txtFromAccount_onblur() {
	setAccountNumber(document.all.txtFromAccount);
}

function txtToAccount_onblur() {
	setAccountNumber(document.all.txtToAccount);
}

function validate() {
	if (document.all.txtFromAccount.value == '') {
		alert("Requiero al menos el número de cuenta inicial del rango.");
		document.all.txtFromAccount.focus;
		return false;
	}
	if (document.all.txtFactor.value == '') {
		alert("Requiero el factor que se le aplicará al rango");
		document.all.txtFactor.focus;
		return false;
	}
	return true;
}

function sendInfo() {
	if (validate()) {
		document.frmEditor.submit();
	}
}
//-->
</SCRIPT>
</HEAD>
<BODY scroll=no>
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
<TR bgcolor=khaki height=25>
  <TD nowrap>
		<FONT face=Arial color=maroon><STRONG><%=gsTitle%></STRONG></FONT>		
  </TD>
  <TD nowrap align=right>
		<A href='' onclick='sendInfo();return false;'>Aceptar</A>&nbsp;&nbsp;&nbsp;&nbsp;
 		<A href='' onclick='window.close();return false;'>Cancelar</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 		<A href='' onclick='deleteRule();return false;'>Eliminar</A>
  </TD>
</TR>
</TABLE>
<FORM name=frmEditor action="./exec/save_rule.asp" method="post">
<TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
<TR>
  <TD>Número:</TD>
  <TD><INPUT name=txtNumber maxlength=16 value="<%=gsNumber%>" style="width:100px"></TD>
</TR>
<TR>
  <TD valign=top>Nombre:</TD>
  <TD>
		<TEXTAREA name=txtDescription ROWS=2 style="width:300px"><%=gsName%></TEXTAREA>
	</TD>
</TR>
<TR>
  <TD>Desde la cuenta:</TD>
  <TD><INPUT name=txtFromAccount maxlength=255 value="<%=gsFromAccount%>" style="width:100%" LANGUAGE=javascript onblur="return txtFromAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Hasta la cuenta:</TD>
  <TD><INPUT name=txtToAccount maxlength=255 value="<%=gsToAccount%>" style="width:100%" LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Auxiliar:</TD>
  <TD><INPUT name=txtSubsidiaryAccount maxlength=255 value="<%=gsFromSubsidiaryAccount%>" style="width:100%"></TD>
</TR>
<TR>
  <TD valign=top>Moneda:</TD>
	<TD>
		<SELECT name=cboCurrencies style="WIDTH: 100%">
			<%=gsCboCurrencies%>
		</SELECT>
		<INPUT type="checkbox" name=chkCurrencies value="true" <%=chkCurrencies%>>Saldos de todas las monedas excepto la seleccionada
	</TD>
</TR>
<TR>
  <TD valign=top>Calificación de la moneda:</TD>
	<TD>
		<SELECT name=cboClips style="WIDTH: 100%">
			<%=gsCboClips%>
		</SELECT>
	</TD>
</TR>
<TR>
  <TD valign=top>Sectores:</TD>
	<TD>
		<SELECT name=cboSectors style="WIDTH: 100%"> 
			<%=gsCboSectors%>
		</SELECT>		
		<INPUT type="checkbox" name=chkSectors value="true" <%=chkSectors%>>Saldos de todos los sectores excepto el seleccionado
	</TD>
</TR>
<TR>
  <TD valign=top><A href='' onclick="displayPicker('restriction', this);return false;">Restricción:</A></TD>
  <TD>
		<TEXTAREA name=txtRestriction ROWS=2 style="width:100%" readonly><%=gsRestriction%></TEXTAREA>
	</TD>
</TR>
<TR>
  <TD>Operación dentro del grupo:</TD>
  <TD>
		<SELECT name=cboOperators style="WIDTH: 120px">
		  <%=gsCboOperators%>
		</SELECT>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;		
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		Factor:
		<INPUT name=txtFactor value="<%=gsFactor%>" style="width:60px">
		<INPUT name=txtRuleId type=hidden value=<%=gnRuleId%>>
  </TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>