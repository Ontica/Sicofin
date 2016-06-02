<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsRuleDefName, gnRuleDefId, gsFromAccount, gsToAccount, gsFromSubsidiaryAccount	
	Dim gsCboCurrencies, gsCboSectors, gsCboClips, gsCboOperators, gsCboRuleGroups, gsCboInsertAfterItems
	Dim gsFactor, gnStdAccountTypeId
	
	Call Main()

	Sub Main()
		Dim oRule, oRuleDef, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'*************************************************************************
		'On Error Resume Next
		gnRuleDefId   = Request.QueryString("ruleDefId")
		gsFromAccount	= Request.QueryString("account")
		If Right(gsFromAccount, 1) <> "*" Then
			gsToAccount = Request.QueryString("account")
		Else
			gsToAccount = ""
		End If
		gsFromSubsidiaryAccount = ""
		gsFactor								= 1
		Set oRule = Server.CreateObject("EFARulesMgrBS.CRule")		
		gsCboCurrencies			= oRule.CboCurrencies(Session("sAppServer"), CLng(Request.QueryString("currencyId")))
		gsCboSectors				= oRule.CboSectors(Session("sAppServer"), CLng(Request.QueryString("sectorId")))
		gsCboOperators			= oRule.CboOperators(Session("sAppServer"))
		gsCboClips					= oRule.CboClips(Session("sAppServer"))
		gsCboInsertAfterItems = "<OPTION value=0>Seleccionar un agrupador de último nivel</OPTION>"
		Set oRule  = Nothing
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gnStdAccountTypeId  = oRuleDef.RuleDefStdAccountTypeId(Session("sAppServer"), CLng(gnRuleDefId))		
		gsRuleDefName =  oRuleDef.RuleDefName(Session("sAppServer"), CLng(gnRuleDefId))
		gsCboRuleGroups		= oRuleDef.CboRules(Session("sAppServer"), CLng(gnRuleDefId))
		Set oRuleDef = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Base de conocimiento contable</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

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

function ruleChildsType() {
	obj = RSExecute("../financial_accounting_scripts.asp","RuleChildsType", document.all.cboRuleGroups.value);
	return(obj.return_value);	
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
	if (ruleChildsType() == 0) {
		alert("El grupo padre seleccionado no permite la agragación de grupos de cuentas.");
		document.all.cboRuleGroups.focus;
		return false;
	}
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
	document.all.txtRuleGroupId.value = document.all.cboRuleGroups.value;
	return true;
}

function sendInfo() {
	if (validate()) {
		document.frmEditor.submit();
	}
}

function cboRuleGroups_onchange() {
	var nRuleChildsType, sTemp, nSelectedGroup;
	
	nSelectedGroup = document.all.cboRuleGroups.value;
	nRuleChildsType = ruleChildsType()
	
	sTemp = "<SELECT name=cboInsertAfterItems style='width:96%'>";	
	if (nRuleChildsType != 0) {	
		obj = RSExecute("../financial_accounting_scripts.asp", "CboRuleChilds", nSelectedGroup);
		sTemp += '<OPTION value=' + nSelectedGroup + '>Insertar al principio del grupo</OPTION>';
		sTemp += obj.return_value;
	} else {
		sTemp += "<OPTION value=0>Seleccionar un agrupador de último nivel</OPTION>";
	}								
	sTemp += '</SELECT>';
	document.all.divCboInsertAfterItems.innerHTML = sTemp;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Incorporación de cuentas pendientes
		</TD>
	  <TD colspan=3 align=right nowrap>
			<img align=absbottom src='/empiria/images/refresh_red.gif' onclick='window.location.href=window.location.href;' alt="Refrescar">			<img align=absbottom src='/empiria/images/help_red.gif' onclick='notAvailable();' alt='Ayuda'>			<img align=absbottom src='/empiria/images/invisible.gif'>
			<img align=absbottom src='/empiria/images/close_red.gif' onclick='window.close();' alt='Cerrar y regresar a la página principal'>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						<%=gsRuleDefName%>
					</TD>
				</TR>
			</TABLE>			
			<TABLE class=applicationTable>
				<FORM name=frmEditor action="./exec/save_rule.asp" method="post">				
					<TR>
					  <TD>Grupo padre:</TD>
					  <TD>
							<SELECT name=cboRuleGroups style="WIDTH: 96%" LANGUAGE=javascript onchange="return cboRuleGroups_onchange()">
								<%=gsCboRuleGroups%>
							</SELECT>					  
					  </TD>
					</TR>
					<TR>
					  <TD>Insertar después de:</TD>
					  <TD>
							<div id=divCboInsertAfterItems>
								<SELECT name=cboInsertAfterItems style="width:96%">
								<%=gsCboInsertAfterItems%>
								</SELECT>
							</div>					  
					  </TD>
					</TR>									
					<TR>
					  <TD>Número:</TD>
					  <TD><INPUT name=txtNumber maxlength=16 style="width:100px"></TD>
					</TR>
					<TR>
					  <TD valign=top>Nombre:</TD>
					  <TD>
							<TEXTAREA name=txtDescription ROWS=2 style="width:96%"></TEXTAREA>
						</TD>
					</TR>
					<TR>
					  <TD>Desde la cuenta:</TD>
					  <TD><INPUT name=txtFromAccount maxlength=255 value="<%=gsFromAccount%>" style="width:96%" LANGUAGE=javascript onblur="return txtFromAccount_onblur()"></TD>
					</TR>
					<TR>
					  <TD>Hasta la cuenta:</TD>
					  <TD><INPUT name=txtToAccount maxlength=255 value="<%=gsToAccount%>" style="width:96%" LANGUAGE=javascript onblur="return txtToAccount_onblur()"></TD>
					</TR>
					<TR>
					  <TD>Auxiliar:</TD>
					  <TD><INPUT name=txtSubsidiaryAccount maxlength=255 value="<%=gsFromSubsidiaryAccount%>" style="width:96%"></TD>
					</TR>
					<TR>
					  <TD valign=top>Moneda:</TD>
						<TD>
							<SELECT name=cboCurrencies style="WIDTH: 96%">
								<%=gsCboCurrencies%>
							</SELECT>
							<INPUT type="checkbox" name=chkCurrencies value="true">Saldos de todas las monedas excepto la seleccionada
						</TD>
					</TR>
					<TR>
					  <TD valign=top>Calificación de la moneda:</TD>
						<TD>
							<SELECT name=cboClips style="WIDTH: 96%">
								<%=gsCboClips%>
							</SELECT>
						</TD>
					</TR>
					<TR>
					  <TD valign=top>Sectores:</TD>
						<TD>
							<SELECT name=cboSectors style="WIDTH: 96%"> 
								<%=gsCboSectors%>
							</SELECT>		
							<INPUT type="checkbox" name=chkSectors value="true">Saldos de todos los sectores excepto el seleccionado
						</TD>
					</TR>
					<TR>
					  <TD valign=top><A href='' onclick="displayPicker('restriction', this);return false;">Restricción:</A></TD>
					  <TD>
							<TEXTAREA name=txtRestriction ROWS=2 style="width:96%" readonly></TEXTAREA>
						</TD>
					</TR>
					<TR>
					  <TD>Operación dentro del grupo:</TD>
					  <TD>
							<SELECT name=cboOperators style="WIDTH: 120px">
							  <%=gsCboOperators%>
							</SELECT>
							&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
							Factor:
							<INPUT name=txtFactor value="<%=gsFactor%>" style="width:60px">
							<INPUT name=txtRuleId type=hidden value=0>
							<INPUT name=txtRuleTypeId type=hidden value=3>
							<INPUT name=txtRuleDefId type=hidden value=<%=gnRuleDefId%>>
							<INPUT name=txtRuleGroupId type=hidden>
					  </TD>
					</TR>					
				</FORM>
			</TABLE>			
		</TD>
	</TR>
	<TR>
		<TD colspan=4 align=right>
			<INPUT class=cmdSubmit name=cmdSave type=button value="Aceptar" style="width:75" onclick='sendInfo();'> &nbsp; &nbsp;
			<INPUT class=cmdSubmit name=cmdCancel  type=button value="Cancelar" style="width:75" onclick='window.close();'>&nbsp;&nbsp; &nbsp;
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/financial_accounting/")</script>
</HTML>