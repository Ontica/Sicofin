<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleGroupId, gnRuleDefId
	Dim gsCboCurrencies, gsCboSectors, gsCboOperators
	Dim gnStdAccountTypeId, gsRuleGroupName, gsCboRuleChilds, gsCboClips
		
	Call Main()

	Sub Main()
		Dim oRule, oRuleDef, oRecordset, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'*************************************************************************
		gnRuleGroupId = CLng(Request.QueryString("Id"))
		gnRuleDefId   = CLng(Request.QueryString("ruleDefId"))
		Set oRule  = Server.CreateObject("EFARulesMgrBS.CRule")
		gsCboRuleChilds = oRule.CboRuleChilds(Session("sAppServer"), CLng(gnRuleGroupId))
		gsCboCurrencies			= oRule.CboCurrencies(Session("sAppServer"))
		gsCboSectors				= oRule.CboSectors(Session("sAppServer"))
		gsCboOperators			= oRule.CboOperators(Session("sAppServer"))
		gsCboClips					= oRule.CboClips(Session("sAppServer"))
		If gnRuleGroupId <> 0 Then
			Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(gnRuleGroupId))
			gsRuleGroupName = oRecordset("nombre_grupo_cuenta")
			Set oRecordset = Nothing
		End If
		Set oRule  = Nothing
		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gnStdAccountTypeId  = oRuleDef.RuleDefStdAccountTypeId(Session("sAppServer"), CLng(gnRuleDefId))
		Set oRuleDef = Nothing
				
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
  <TD>
		<FONT face=Arial color=maroon><STRONG>Insertar rango de cuentas en <%=gsRuleGroupName%></STRONG></FONT>		
  </TD>
  <TD nowrap align=right>
		<A href='' onclick='sendInfo();return false;'>Aceptar</A>&nbsp;&nbsp;
 		<A href='' onclick='window.close();return false;'>Cancelar</A>
  </TD>
</TR>
</TABLE>
<FORM name=frmEditor action="./exec/save_rule.asp" method="post">
<TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
<TR>
  <TD nowrap>Insertar después de:</TD>
  <TD nowrap>
		<SELECT name=cboInsertAfterItems style="WIDTH: 260px">
		  <OPTION value=<%=gnRuleGroupId%>>Insertar al principio del grupo</OPTION>
		  <%=gsCboRuleChilds%>
		</SELECT>
	</TD>
</TR>
<TR>
  <TD>Número:</TD>
  <TD><INPUT name=txtNumber maxlength=16 style="width:100px"></TD>
</TR>
<TR>
  <TD valign=top>Nombre:</TD>
  <TD>
		<TEXTAREA name=txtDescription ROWS=2 style="width:300px"></TEXTAREA>
	</TD>
</TR>
<TR>
  <TD>Desde la cuenta:</TD>
  <TD><INPUT name=txtFromAccount maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtFromAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Hasta la cuenta:</TD>
  <TD><INPUT name=txtToAccount maxlength=255 style="width:100%" LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
</TR>
<TR>
  <TD>Auxiliar:</TD>
  <TD><INPUT name=txtSubsidiaryAccount maxlength=255 style="width:100%"></TD>
</TR>
<TR>
  <TD valign=top>Moneda:</TD>
	<TD>
		<SELECT name=cboCurrencies style="WIDTH: 100%">
			<%=gsCboCurrencies%>
		</SELECT>
		<INPUT type="checkbox" name=chkCurrencies value="true">Saldos de todas las monedas excepto la seleccionada
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
		<INPUT type="checkbox" name=chkSectors value="true">Saldos de todos los sectores excepto el seleccionado
	</TD>
</TR>
<TR>
  <TD valign=top>Restricción:</TD>
  <TD>
		<TEXTAREA name=txtRestriction ROWS=2 style="width:100%"></TEXTAREA>
	</TD>  
</TR>
<TR>
  <TD>Operación:</TD>
  <TD>
		<SELECT name=cboOperators style="WIDTH: 120px">
		  <%=gsCboOperators%>
		</SELECT>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		Factor:
		<INPUT name=txtFactor style="width:60px" value=1>
		<INPUT name=txtRuleId type=hidden value=0>
		<INPUT name=txtRuleGroupId type=hidden value=<%=gnRuleGroupId%>>
		<INPUT name=txtRuleDefId type=hidden value=<%=gnRuleDefId%>>
    <INPUT name=txtRuleTypeId type=hidden value=3>    
  </TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>