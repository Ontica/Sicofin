<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleDefId, gnRuleId
	Dim gsCboCurrencies, gsCboSectors, gsCboOperators
	Dim gnStdAccountTypeId, gsRuleGroupName, gsGroupEntityId, gnRuleGroupId
	
	Dim gsFromAccount, gsToAccount, gsRestriction, gsFromSubsidiaryAccount
	Dim gsTargetAccount, gsTargetSubsidiaryAccount, gsCboTargetSectors
		
	Call Main()

	Sub Main()
		Dim oRule, oRuleDef, oRecordset, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'*************************************************************************
		'On Error Resume Next
		gnRuleId = CLng(Request.QueryString("Id"))
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")
		
		gsCboCurrencies			= oRule.CboCurrencies(Session("sAppServer"))
		gsCboSectors				= oRule.CboSectors(Session("sAppServer"))		
		gsCboOperators			= oRule.CboOperators(Session("sAppServer"))		
		If gnRuleId <> 0 Then
			Set oRecordset	= oRule.RuleRS(Session("sAppServer"), CLng(gnRuleId))
			gnRuleDefId			= oRecordset("id_regla_contable")
			gsRuleGroupName = oRecordset("nombre_grupo_cuenta")
			gsGroupEntityId = oRecordset("id_entidad_agrupador")
			gnRuleGroupId		= oRecordset("id_agrupador_origen")
			gsFromAccount		= oRecordset("cuenta_origen_inicial")
			gsToAccount			= oRecordset("cuenta_origen_final")
			gsRestriction		= oRecordset("filtro_cuentas_origen")
			gsCboTargetSectors  = oRule.CboTargetSectors(Session("sAppServer"), CLng(oRecordset("id_sector_destino")))
			gsFromSubsidiaryAccount = oRecordset("auxiliar_origen")
			gsTargetAccount = oRule.StdAccountNumber(Session("sAppServer"), CLng(oRecordset("id_cuenta_destino")))
			gsTargetSubsidiaryAccount = oRecordset("auxiliar_destino")
			Set oRecordset = Nothing
		Else
			gsCboTargetSectors  = oRule.CboTargetSectors(Session("sAppServer"))
		End If				
		Set oRule = Nothing
		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gnStdAccountTypeId  = oRuleDef.RuleDefStdAccountTypeId(Session("sAppServer"), CLng(gnRuleDefId))		
		Set oRuleDef  = Nothing		
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Rango de cuentas</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function deleteRule() {
	if (confirm('¿Elimino el rango de cuentas?')) {
		obj = RSExecute("../financial_accounting_scripts.asp","DeleteRule", <%=gnRuleId%>);
		window.close();
	}
	return false
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

function existsSubsidiaryAccount(sSubsAccount) {
	var obj;
	
	obj = RSExecute("../financial_accounting_scripts.asp", "SubsidiaryAccountName", 9, sSubsAccount);
	if (obj.return_value != '') {
		return true;
	}
	return false
}

function txtFromAccount_onblur() {
	setAccountNumber(document.all.txtFromAccount);
}

function txtToAccount_onblur() {
	setAccountNumber(document.all.txtToAccount);
}

function txtTargetAccount_onblur() {
	setAccountNumber(document.all.txtTargetAccount);
}

function txtTargetOBAccount_onblur() {
	setAccountNumber(document.all.txtTargetOBAccount);
}

function validate() {
	if (document.all.txtFromAccount.value == '') {
		alert("Requiero al menos el número de cuenta inicial del rango.");
		document.all.txtFromAccount.focus();
		return false;
	}
	if (document.all.txtFactor.value == '') {
		alert("Requiero el factor que se le aplicará al rango");
		document.all.txtFactor.focus;
		return false;
	}
	if (document.all.txtTargetSubsidiaryAccount.value != '') {
		if (!existsSubsidiaryAccount(document.all.txtTargetSubsidiaryAccount.value)) {
			alert("El auxiliar destino proporcionado no ha sido dado de alta en la contabilidad respectiva.");
			document.all.txtTargetSubsidiaryAccount.focus();
			return false;
		}
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
  <TD>
		<FONT face=Arial color=maroon><STRONG>Insertar rango de cuentas en <%=gsRuleGroupName%></STRONG></FONT>		
  </TD>
  <TD nowrap align=right>
		<A href='' onclick='sendInfo();return false;'>Aceptar</A>&nbsp;&nbsp;
 		<A href='' onclick='window.close();return false;'>Cancelar</A>&nbsp;&nbsp;
 		<A href='' onclick='deleteRule();return false;'>Eliminar</A>
  </TD>
</TR>
</TABLE>
<FORM name=frmEditor action="./exec/save_voucher_rule.asp" method="post">
<TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
<TR>
  <TD>Desde la cuenta (origen):</TD>
  <TD><INPUT name=txtFromAccount maxlength=255 style="width:100%" value='<%=gsFromAccount%>' LANGUAGE=javascript onblur="return txtFromAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Hasta la cuenta (origen):</TD>
  <TD><INPUT name=txtToAccount maxlength=255 style="width:100%" value='<%=gsToAccount%>' LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Auxiliar (origen):</TD>
  <TD><INPUT name=txtSubsidiaryAccount value='<%=gsFromSubsidiaryAccount%>'maxlength=255 style="width:100%"></TD>
</TR>
<TR>
  <TD valign=top>Moneda (origen):</TD>
	<TD>
		<SELECT name=cboCurrencies style="WIDTH: 100%">
			<%=gsCboCurrencies%>
		</SELECT>
		<INPUT type="checkbox" name=chkCurrencies value="true">Saldos de todas las monedas excepto la seleccionada
	</TD>
</TR>
<TR>
  <TD valign=top>Sector (origen):</TD>
	<TD>
		<SELECT name=cboSectors style="WIDTH: 100%"> 
			<%=gsCboSectors%>
		</SELECT>		
		<INPUT type="checkbox" name=chkSectors value="true">Saldos de todos los sectores excepto el seleccionado
	</TD>
</TR>
<TR>
  <TD valign=top>Restricción:</TD>
  <TD>
		<TEXTAREA name=txtRestriction ROWS=2 style="width:100%"><%=gsRestriction%></TEXTAREA>
	</TD>  
</TR>
<TR>
  <TD>Cuenta destino:</TD>
  <TD><INPUT name=txtTargetAccount value='<%=gsTargetAccount%>'maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtTargetAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Auxiliar destino:</TD>
  <TD><INPUT name=txtTargetSubsidiaryAccount value='<%=gsTargetSubsidiaryAccount%>' maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Sector destino:</TD>
  <TD>
		<SELECT name=cboTargetSectors style="WIDTH: 100%"> 
			<%=gsCboTargetSectors%>
		</SELECT>
	</TD>  
</TR>
<TR>
  <TD>Cuenta destino sobregiro:</TD>
  <TD><INPUT name=txtTargetOBAccount maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtTargetOBAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Auxiliar destino sobregiro:</TD>
  <TD><INPUT name=txtTargetOBSubsidiaryAccount maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
</TR>
<INPUT name=txtFactor type=hidden value=1>
<INPUT name=txtRuleId type=hidden value=<%=gnRuleId%>>
<INPUT name=txtRuleGroupId type=hidden value=<%=gnRuleGroupId%>>
<INPUT name=txtRuleDefId type=hidden value=<%=gnRuleDefId%>>
<INPUT name=txtRuleTypeId type=hidden value=3>
<INPUT name=txtGroupEntityId type=hidden value=<%=gsGroupEntityId%>>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>