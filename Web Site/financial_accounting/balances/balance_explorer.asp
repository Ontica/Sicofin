<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsExplorerResultsHeader, gsExplorerResultsBody, gsTackedWindows, gnSelectedColumn, gnDisplayAdvancedSearchOptions
	Dim gsCboGralLedgerGroups, gsCboTransactionTypes, gnTransactionTypeId, gsCboVoucherTypes, gnVoucherTypeId
	Dim gsFromDate, gsToDate, gsToday, gsFromAccount, gsToAccount
	Dim gnGralLedgersCategory, gnGralLedger, gsShowInCascade, gbShowInCascade
	Dim gsExceptSelectedTransactionType, gsExceptSelectedVoucherType
	Dim gsCboAccountTypes, gnStdAccountTypeId, gsCboAccountNatures, gsAccountNature
	Dim gsCboAccountPatterns, gsAccountPattern, gsCboAccountBalanceType, gnAccountBalanceType
	Dim gsFromSubLedgerAccount, gsToSubLedgerAccount
	Dim	gbShowSubsAccounts, gsShowSubsAccounts, gbGroupBySubLedgerAccount, gsGroupBySubLedgerAccount
	Dim gbShowAverageColumn, gsShowAverageColumn
	Dim gbShowOnlyLastLevels, gsShowOnlyLastLevels
	Dim gsCboExchangeRateTypes, gnExchangeRateTypeId, gsCboCurrencies, gnExchangeRateCurrencyId, gsExchangeRateDate
	
  Dim oVoucherUS, nScriptTimeout
  Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")

	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call Main()
	Server.ScriptTimeout = nScriptTimeout  	

	Sub Main()
		Call SetGlobalValues()
		
		gsCboGralLedgerGroups	= oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), _
																															    CLng(Session("uid")), 2, _
																															    CLng(gnGralLedgersCategory))
		gsCboAccountPatterns    = oVoucherUS.CboStdAccountPatterns(Session("sAppServer"), 0, CStr(gsAccountPattern))
		gsCboAccountBalanceType = oVoucherUS.CboBalanceTypes(CLng(gnAccountBalanceType))
		gsCboAccountTypes			 = oVoucherUS.CboStdAccountTypes(Session("sAppServer"), CLng(gnStdAccountTypeId))
		gsCboAccountNatures		 = oVoucherUS.CboStdAccountNature(CStr(gsAccountNature))
		gsCboTransactionTypes  = oVoucherUS.CboTransactionTypes(Session("sAppServer"), _
																												   Abs(CLng(gnTransactionTypeId)))
		gsCboVoucherTypes      = oVoucherUS.CboVouchersTypes(Session("sAppServer"), _
																										     0, Abs(CLng(gnVoucherTypeId)))
		gsCboExchangeRateTypes = oVoucherUS.CboExchangeRateTypes(Session("sAppServer"), CLng(gnExchangeRateTypeId))
		gsCboCurrencies				 = oVoucherUS.CboCurrencies(Session("sAppServer"), CLng(gnExchangeRateCurrencyId))

		
		Call GetExplorerInformation()
		Set oVoucherUS = Nothing		
	End Sub
	
	Sub GetExplorerInformation()
		Dim oExplorer, vGralLedgers, bShowSubsAccounts, sWhere, sOrderBy
		'***************************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("EFABalanceExplorer.CExplorer")												
		gsExplorerResultsHeader = oExplorer.Header(CBool(gbShowSubsAccounts), _
																							 CBool(gbGroupBySubLedgerAccount), _
																							 CBool(gbShowAverageColumn), CLng(gnSelectedColumn))	  
		'sWhere = oExplorer.BuildSearchParametersString(CLng(gnGralLedgersCategory), _
		'																							 CStr(gsFromApplicationDate), _
		'																							 CStr(gsToApplicationDate), _
		'																							 CStr(gsFromElaborationDate), _
		'																							 CStr(gsToElaborationDate), _
		'																							 CStr(gsVoucherNumber), _
		'																							 CStr(gsVoucherConcept), _
		'																							 CStr(gsAccounts), CLng(gnTransactionTypeId), _
		'																							 CLng(gnVoucherTypeId), CLng(gnBalancingType))
		'sOrderBy = "numero_transaccion"

		If (CLng(Request.Form.Count) <> 0) Then
			vGralLedgers = GetGeneralLedgers()			
			gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), vGralLedgers, CBool(Not gbShowInCascade), _
																						 CStr(gsAccountPattern), CDate(gsFromDate), CDate(gsToDate), _
																						 CStr(gsFromAccount), CStr(gsToAccount), _
																						 CStr(gsFromSubLedgerAccount), CStr(gsToSubLedgerAccount), _
																						 CLng(gnTransactionTypeId), CLng(gnVoucherTypeId), _
																						 0, CLng(gnStdAccountTypeId), CStr(gsAccountNature), _
																						 CLng(gnExchangeRateTypeId), CLng(gnExchangeRateCurrencyId), _
																						 CDate(gsExchangeRateDate), CLng(gnAccountBalanceType), _
																						 CBool(gbShowSubsAccounts), CBool(gbGroupBySubLedgerAccount), _
																						 CBool(gbShowAverageColumn), CBool(gbShowOnlyLastLevels), _
																						 "", CLng(gnSelectedColumn), "")
		End If
		Set oExplorer = Nothing
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		
	End Sub

  Sub SetGlobalValues() 
		If (CLng(Request.Form.Count) <> 0) Then
			gsToday								 = Date()
			gnGralLedgersCategory  = Request.Form("cboGralLedgerGroups")
			gnGralLedger					 = Request.Form("cboGralLedgers")
			If Len(Request.Form("chkShowInCascade")) <> 0 Then
				gbShowInCascade = True
				gsShowInCascade = "checked"
			Else
				gbShowInCascade = False
				gsShowInCascade = ""
			End If			
			If Len(Request.Form("txtFromDate")) <> 0 Then
				gsFromDate = oVoucherUS.FormatDate(CDate(Request.Form("txtFromDate")))
			Else
				gsFromDate = oVoucherUS.FormatDate(Date())
			End If
			If Len(Request.Form("txtToDate")) <> 0 Then
				gsToDate = oVoucherUS.FormatDate(CDate(Request.Form("txtToDate")))
			Else
				gsToDate = oVoucherUS.FormatDate(Date())
			End If
			gsFromAccount				   = Request.Form("txtFromAccount")
			gsToAccount				     = Request.Form("txtToAccount")
			gsFromSubLedgerAccount = Request.Form("txtFromSubLedgerAccount")
			gsToSubLedgerAccount   = Request.Form("txtToSubLedgerAccount")
			gsAccountPattern		   = Request.Form("cboAccountPatterns")
			gnAccountBalanceType   = Request.Form("cboAccountBalanceType")
			gnStdAccountTypeId     = Request.Form("cboAccountTypes")
			gsAccountNature			   = Request.Form("cboAccountNatures")
			If Len(Request.Form("chkShowOnlyLastLevels")) <> 0 Then
				gbShowOnlyLastLevels = True
				gsShowOnlyLastLevels = "checked"
			Else
				gbShowOnlyLastLevels = False
				gsShowOnlyLastLevels = ""
			End If			
			If Len(Request.Form("chkShowSubsAccounts")) <> 0 Then
				gbShowSubsAccounts = True
				gsShowSubsAccounts = "checked"
			Else
				gbShowSubsAccounts = False
				gsShowSubsAccounts = ""
			End If
			If Len(Request.Form("chkGroupBySubLedgerAccounts")) <> 0 Then
				gbGroupBySubLedgerAccount = True
				gsGroupBySubLedgerAccount = "checked"
			Else
				gbGroupBySubLedgerAccount = False
				gsGroupBySubLedgerAccount = ""
			End If			
			If Len(Request.Form("chkShowAverageColumn")) <> 0 Then
				gbShowAverageColumn = True
				gsShowAverageColumn = "checked"
			Else
				gbShowAverageColumn = False
				gsShowAverageColumn = ""
			End If
			gnTransactionTypeId = Request.Form("cboTransactionTypes")
			If Len(Request.Form("chkTransactionTypes")) <> 0 Then
				gnTransactionTypeId  = -1 * CLng(gnTransactionTypeId)
				gsExceptSelectedTransactionType = "checked"
			End If			
			gnVoucherTypeId     = Request.Form("cboVoucherTypes")
			If Len(Request.Form("chkVoucherTypes")) <> 0 Then
				gnVoucherTypeId = -1 * CLng(gnVoucherTypeId)
				gsExceptSelectedVoucherType = "checked"
			End If
			gnExchangeRateTypeId     = Request.Form("cboExchangeRateTypes")
			gnExchangeRateCurrencyId = Request.Form("cboExchangeRateCurrencies")
			If Len(Request.Form("txtExchangeRateDate")) <> 0 Then
				gsExchangeRateDate     = oVoucherUS.FormatDate(CDate(Request.Form("txtExchangeRateDate")))
			End If
			gnSelectedColumn		= Request.Form("txtSelectedColumn")
			gsTackedWindows     = Request.Form("txtTackedWindows")
			gnDisplayAdvancedSearchOptions = Request.Form("txtDisplayAdvancedSearchOptions")
		Else
			gsToday								 = oVoucherUS.FormatDate(Date())
			gnGralLedgersCategory  = 0
			gnGralLedger					 = 0
			gbShowInCascade				 = False
			gsShowInCascade				 = ""
			gsFromDate						 = oVoucherUS.FormatDate(Date())
			gsToDate						   = oVoucherUS.FormatDate(Date())
			gsFromAccount					 = ""
			gsToAccount					   = ""
			gsAccountPattern			 = "&&&&-&&-&&-&&-&&-&&-&&"
			gnAccountBalanceType	 = 3
			gnStdAccountTypeId     = 0
			gsAccountNature			   = ""
			gbShowOnlyLastLevels = False
			gsShowOnlyLastLevels = ""			
			gbShowSubsAccounts     = False
			gsShowSubsAccounts     = ""
			gbShowAverageColumn    = False
			gsShowAverageColumn    = ""				
			gnTransactionTypeId    = 0			
			gnVoucherTypeId        = 0
			gnExchangeRateTypeId   = 0
			gnExchangeRateCurrencyId = 0
			gsExchangeRateDate     = ""			
			gnSelectedColumn			 = 1
			gsTackedWindows				 = ""
			gnDisplayAdvancedSearchOptions = 0
		End If
	End Sub	
	
	Function GetGeneralLedgers()
		Dim sTemp
		'*************************
		If (Len(Request.Form("cboGralLedgers")) <> 0 ) Then		
			If (Len(Request.Form("txtFromGL")) = 0) Then
				If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
					sTemp = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGralLedgerGroups")), ",")
					GetGeneralLedgers = Split(sTemp, ",")
				Else
					GetGeneralLedgers = CLng(Request.Form("cboGralLedgers"))
				End If
			Else
				GetGeneralLedgers = oVoucherUS.GetGLRangeArray(Session("sAppServer"), CLng(Request.Form("cboGralLedgerGroups")), _
																								  CLng(Request.Form("txtFromGL")), CLng(Request.Form("txtToGL")))
			End If
		End If
	End Function
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function formatAccount(sAccount) {
	var obj;
	if (sAccount != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountNumber", 1, sAccount);
		if (obj.return_value != '') {
			return (obj.return_value);
		} else {
			alert("No entiendo el formato de la cuenta por la que se desea hacer el filtrado.");
			return '';
		}
	} else {
		return '';
	}
}

function formatSubsAccount(oControl) {
	var obj, sPrefix;
  
	if (oControl.value == '') {
		return false;
	}
	if (document.all.cboGralLedgers.value != 0) {
		obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","GetGLSubsidiaryLedgerPrefix", document.all.cboGralLedgers.value);	
		sPrefix = obj.return_value;
	} else {
		sPrefix = '*';
	}		
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","FormatSubsAccount", oControl.value);	
	oControl.value = sPrefix + obj.return_value;
}

function isDate(sDate) {
	var obj;
	if (sDate != '') {
		obj = RSExecute("../financial_accounting_scripts.asp","IsDate", sDate);
		return obj.return_value;
	} else {
		return true;
	}
}

function setAccountNumber(oControl) {
	var obj;	
	if (oControl.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountNumber", 1 , oControl.value);
		if (obj.return_value != '') {
			oControl.value = obj.return_value;
		} else {
			alert("No entiendo el formato de la cuenta proporcionada.");
		}
	}
	return true;
}

function updateGralLedgers(nSelectedItem) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGralLedgerGroups.value, nSelectedItem);	
	document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
}

function orderBy(nSelectedColumn) {
  var nTemp = document.all.txtSelectedColumn.value;
  
  if (nTemp == '') {
		document.all.txtSelectedColumn.value = nSelectedColumn;
	} else {
		if ((nTemp == nSelectedColumn) || (Math.abs(nSelectedColumn) == nSelectedColumn)) {
			if (nTemp == nSelectedColumn) {
				document.all.txtSelectedColumn.value = (-1 * nSelectedColumn);
			} else {
				document.all.txtSelectedColumn.value = nSelectedColumn;
			}
		} else {
			document.all.txtSelectedColumn.value = nSelectedColumn;
		}
  }
  document.all.frmSend.action = '';
  document.all.frmSend.submit();
	return false;
}

function formatAccountRanges() {
	var sTemp;
	
	sTemp = document.frmSend.txtFromAccount.value;
	if (sTemp != '') {
		sTemp = formatAccount(sTemp);
		if (sTemp != '') {
			document.frmSend.txtFromAccount.value = sTemp;		
		} else { 
			document.frmSend.txtFromAccount.focus();
			return false
		}
	}
	sTemp = document.frmSend.txtToAccount.value;
	if (sTemp != '') {
		sTemp = formatAccount(sTemp);
		if (sTemp != '') {
			document.frmSend.txtToAccount.value = sTemp;		
		} else { 
			document.frmSend.txtToAccount.focus();
			return false
		}
	}		
	return true;
}

function validateForm() {
	if (document.all.cboGralLedgerGroups.value == '') {
		alert("Necesito la selección del grupo de contabilidades");
		document.all.cboGralLedgerGroups.focus();
		return false;
	}
	if (document.all.txtFromDate.value == '') {
		if (confirm('¿Obtengo los saldos al día de hoy?')) {
			document.all.txtFromDate.value = '<%=gsToday%>';
		}	else {
			document.all.txtFromDate.focus();
			return false;
		}
	}
	if (!isDate(document.all.txtFromDate.value)) {
		alert("No reconozco la fecha inicial para la consulta de los saldos.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (!isDate(document.all.txtToDate.value)) {
		alert("No reconozco la fecha final para la consulta de los saldos.");
		document.all.txtToDate.focus();
		return false;
	}
	if (document.all.txtExchangeRateDate.value != '') {
		if (!isDate(document.all.txtExchangeRateDate.value)) {
			alert("No reconozco la fecha del tipo de cambio.");
			document.all.txtExchangeRateDate.focus();
			return false;
		}
	}	
	if (!formatAccountRanges()) {
		return false;
	}	
	
	if ((document.all.txtFromAccount == '') && (document.all.txtToAccount == '') && (document.all.txtFromAccount == '') && (document.all.txtFromSubLedgerAccount == '') && (document.all.txtToSubLedgerAccount == '')) {
		alert("Requiero al menos un filtro por cuenta o por auxiliar.");
		document.all.txtFromAccount.focus();
		return false;
	}
	
	return true;	
}

function resetSearchOptions() {	
	return false;
}

function resetAdvancedSearchOptions() {
	return false;
}

function showAdvancedSearch() {
	showOptionsWindow(document.all.divAdvancedSearchOptions);
	if (document.all.divAdvancedSearchOptions.style.display == 'inline') {
		document.all.divAdvSearchLabel.innerText = 'Ocultar consulta avanzada';
		document.all.txtDisplayAdvancedSearchOptions.value = 1;
	} else {
		document.all.divAdvSearchLabel.innerText = 'Consulta avanzada';
		document.all.txtDisplayAdvancedSearchOptions.value = 0;
	}
	return false;
}

function getBalances() {	
	if (validateForm()) {
		document.all.frmSend.submit();
	}
	return false;
}

function frmSend_onsubmit() {
	var sOpt;
	alert("NOBBBBBEE");	
	sOpt = 'height=210,width=420,status=no,toolbar=no,menubar=no,location=no';		
	alert(document.all.frmSend.target);	
	if (document.all.frmSend.target == null || document.all.frmSend.target == '') {
		return (validateForm());
	}
	if (document.all.frmSend.target != '') {
		window.open(document.all.frmSend.action, 'oViewerWindow', sOpt);		
		return true;
	}
	return false;
}

function callViewer(nOperation) {
	var sURL, sOpt;
  switch (nOperation) {  
    case 1:		// View vouchers
			sURL = 'vouchers_detail.asp';
			document.all.frmSend.target = 'oViewerWindow';
			document.all.frmSend.action = sURL;
			document.all.frmSend.submit();
			document.all.frmSend.target = null;
			document.all.frmSend.action = '';
			return false;
    case 2:		// View averages detail
			sURL = 'average_datail.asp';
			document.all.frmSend.target = 'oViewerWindow';
			document.all.frmSend.action = sURL;
			document.all.frmSend.submit();
			document.all.frmSend.target = null;
			document.all.frmSend.action = '';
			return false;
			/*
			sOpt = 'height=465px,width=370px,resizable=no,scrollbars=no,status=no,location=no';
			if (oViewerWindow == null || oViewerWindow.closed) {
				oViewerWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oViewerWindow.focus();
				oViewerWindow.navigate(sURL);
			}
			*/
	}
	return false;
}

function showAverage(oSource) {
	alert("Por el momento la página con la integración de saldos promedio no está disponible.");
	return false;
}

function showIntegration(nGL, nCurcyId, nStdActId, nSubledgerActId, nSectorId) {	
	if (nGL == 0) {
		nGL = document.all.cboGralLedgers.value;
	}
	document.all.txtDetailGralLedgerId.value   = nGL;
	document.all.txtDetailCurrencyId.value     = nCurcyId;
	document.all.txtDetailStdActId.value       = nStdActId;
	document.all.txtDetailSubledgerActId.value = nSubledgerActId;
	document.all.txtDetailSectorId.value			 = nSectorId;
	alert(nGL + ', ' + nCurcyId + ', ' + nStdActId + ', ' + nSubledgerActId + ', ' + nSectorId);
	return(callViewer(1));
}

function window_onload() {
	updateGralLedgers(<%=gnGralLedger%>);
	showTackedWindows(Array(<%=gsTackedWindows%>));
	<% If gnDisplayAdvancedSearchOptions = 1 Then %>
		showAdvancedSearch();
	<% End If %>
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="window_onload();">
<FORM name=frmSend action='' method=post onsubmit="return frmSend_onsubmit()">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Explorador de saldos
		</TD>
		<TD colspan=3 align=right nowrap>						<img align=absbottom src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absbottom src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absbottom src='/empiria/images/invisible.gif'>
			<img align=absbottom src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">		</TD>
	</TR>
	<TR id=divTasksOptions style='display:none'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Tareas
					</TD>
					<TD nowrap align=left>
						<A href='' onclick="return(notAvailable());">Lista de tareas</A>
						&nbsp; | &nbsp
						<A href='' onclick="return(notAvailable());">Mi lista de tareas pendientes</A>
					</TD>
					<TD nowrap align=right>
					  <img id=cmdTasksOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divTasksOptions, this)' alt='Fijar la ventana'>					
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='/empiria/images/invisible.gif'>					  
						<img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divTasksOptions, document.all.cmdTasksOptionsTack)" alt='Cerrar'>
					</TD>
				</TR>
				<TR>
					<TD nowrap colspan=3>
						<A href="../transactions/pages/voucher_wizard.asp">Crear póliza</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../transactions/pages/voucher_explorer.asp">Explorador de pólizas</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../reports/balances.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;						
						<A href="../reports/other_reports.asp">Reportes</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Consultar saldos
					</TD>
					<TD nowrap align=left>
						<A href="" onclick='return(notAvailable());'>Cargar consulta</A>
						&nbsp; | &nbsp
						<A href='' onclick='return(notAvailable());'>Guardar consulta</A>
						&nbsp; &nbsp &nbsp; &nbsp &nbsp; &nbsp
						<A href='' onclick='return(getBalances());'>Ejecutar consulta</A>
					</TD>						
					<TD nowrap align=right>
						<A href="" onclick="return(showAdvancedSearch());"><span id=divAdvSearchLabel>Consulta avanzada</span></A>
						<img align=absbottom src='/empiria/images/invisible4.gif'>
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>
						<img align=absbottom src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>En el grupo de contabilidades:</b></TD>
					<TD colspan=2 nowrap width=100%>
						<SELECT name=cboGralLedgerGroups style="width:100%" onchange="return updateGralLedgers(0);">
							<%=gsCboGralLedgerGroups%>
						</SELECT>
					</TD>
				</TR>
				<TR>
					<TD nowrap valign=top><b>De la contabilidad:</b><br><br><br></TD>					
					<TD colspan=2 nowrap width=100% valign=top>
						<span id=divCboGeneralLedgers>
							<SELECT name="cboGralLedgers" width=100%>

							</SELECT>
						</span>
						<br>
						Mostrar los saldos de las contabilidades seleccionadas <b>en cascada</b> (sin consolidar)
						&nbsp;&nbsp; <INPUT type="checkbox" name=chkShowInCascade value="true" <%=gsShowInCascade%>>
					</TD>
				</TR>
				<TR>
					<TD nowrap valign=top><b>En el período comprendido:</b><br><br></TD>
					<TD colspan=2 nowrap width=100% valign=top>
						Del día: &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
						<INPUT name=txtFromDate style="width:93;height:20;" value='<%=gsFromDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromDate)'>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;al día:&nbsp; 
						<INPUT name=txtToDate style="width:93;height:20;" value='<%=gsToDate%>'>						
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToDate)'>
						 &nbsp;(día / mes / año)
						 <br>
						 <img src='/empiria/images/separator.gif' width=100% height=1px>
					</TD>
				</TR>
				<TR>
					<TD nowrap valign=top><b>Filtrado por cuenta:</b><br><br><br><br><br><br></TD>
					<TD colspan=2 nowrap width=100% valign=top>
						De la cuenta: <INPUT name=txtFromAccount style="width:130;height:20;" value='<%=gsFromAccount%>' onblur='setAccountNumber(this);'>
						&nbsp; &nbsp;a la cuenta: <INPUT name=txtToAccount style="width:130;height:20;" value='<%=gsToAccount%>' onblur='setAccountNumber(this);'>
						&nbsp;(opcional)
						&nbsp; (permiten <A href="" onclick="return(showHelp('wild_chars'))" target=_blank>comodines</A>)
						<br>
					  Mostrar sólo <b>cuentas de último nivel</b> &nbsp; &nbsp;
					  <INPUT type="checkbox" name=chkShowOnlyLastLevels value="true" <%=gsShowOnlyLastLevels%>>
					  <br>
						<b>Desglosar las cuentas auxiliares</b> &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT type="checkbox" name=chkShowSubsAccounts value="true" <%=gsShowSubsAccounts%>>
						<br>
						Incluir los <b>saldos promedio</b> en el período
						<INPUT type="checkbox" name=chkShowAverageColumn value="true" <%=gsShowAverageColumn%>>					  						
						<br>
					  <img src='/empiria/images/separator.gif' width=100% height=1px>
								
					</TD>
				</TR>
				<TR>
					<TD nowrap valign=top><b>Filtrado por auxiliar:</b><br><br><br></TD>
					<TD colspan=2 nowrap width=100% valign=top>
						Del auxiliar: &nbsp; <INPUT name=txtFromSubLedgerAccount style="width:130;height:20;" value='<%=gsFromSubLedgerAccount%>' onblur='formatSubsAccount(this);'>
						&nbsp; &nbsp; &nbsp; al auxiliar: <INPUT name=txtToSubLedgerAccount style="width:130;height:20;" value='<%=gsToSubLedgerAccount%>' onblur='formatSubsAccount(this);'>
						&nbsp;(opcional)
						&nbsp; (permiten <A href="" onclick="return(showHelp('wild_chars'))" target=_blank>comodines</A>)
						<br>
						<b>Agrupar</b> los saldos <b>por cuenta auxiliar</b>
						<INPUT type="checkbox" name=chkGroupBySubLedgerAccounts value="true" <%=gsGroupBySubLedgerAccount%>>
					</TD>
				</TR>			
				<TR id=divAdvancedSearchOptions style='display:none;'>
					<TD nowrap colspan=3>
						<TABLE class="fullScrollMenu">
							<TR class="fullScrollMenuHeader">
								<TD class="fullScrollMenuTitle" nowrap>
									Consulta avanzada
								</TD>
								<TD colspan=2 nowrap align=right>									
									<img align=absbottom src='/empiria/images/invisible4.gif'>
									<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='return(resetAdvancedSearchOptions());' alt='Actualizar ventana'>
									<img align=absbottom src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>									
								</TD>
							</TR>
							<TR>
								<TD nowrap><b>Mostrar las cuentas a: &nbsp; </b><br><br><br>&nbsp;</TD>
								<TD colspan=2 nowrap>
									<b>Nivel: </b> &nbsp; 
									<SELECT name=cboAccountPatterns style="width:240">
										<%=gsCboAccountPatterns%>
									</SELECT>
									&nbsp; 
									<b>Incluyendo: &nbsp;</b>&nbsp;
									<SELECT name=cboAccountBalanceType style="width:230">
										<%=gsCboAccountBalanceType%>
									</SELECT>
									<br>
									<b>Tipo: &nbsp;</b>&nbsp;
									<SELECT name=cboAccountTypes style="width:240">
										<OPTION value=0>-- Todos los tipos de cuenta --</OPTION>
										<%=gsCboAccountTypes%>
									</SELECT>
									&nbsp;
									<b>Naturaleza:</b> &nbsp; 
									<img align=absbottom src='/empiria/images/invisible.gif'>
									<SELECT name=cboAccountNatures style="width:230">
										<OPTION value="">-- Todas las cuentas --</OPTION>
										<%=gsCboAccountNatures%>
									</SELECT>
									<br>									
									<img src='/empiria/images/separator.gif' width=100% height=1px>
								</TD>
							</TR>
							<TR>
								<TD nowrap valign=top><b>Filtrar las pólizas por:</b><br><br>&nbsp;</TD>
								<TD colspan=2 nowrap width=100% valign=top>
									<b>Tipo de transacción:</b>&nbsp;
									<SELECT name=cboTransactionTypes style='width:260'>
										<OPTION value=0 selected> -- Todas las transacciones -- </OPTION>
										<%=gsCboTransactionTypes%>
									</SELECT> &nbsp;Todas excepto las del tipo seleccionado
									<INPUT type="checkbox" name=chkTransactionTypes value="true" <%=gsExceptSelectedTransactionType%>>
									<br>
									<b>Tipo de póliza:</b> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
									<SELECT name=cboVoucherTypes style='width:260'>			
										<OPTION value=0 selected> -- Todos los tipos de póliza -- </OPTION>
											<%=gsCboVoucherTypes%>
									</SELECT> &nbsp;Todas excepto las del tipo seleccionado
									<INPUT type="checkbox" name=chkVoucherTypes value="true" <%=gsExceptSelectedVoucherType%>>
									<br>
									<img src='/empiria/images/separator.gif' width=100% height=1px>
								</TD>
							</TR>
							<TR>
								<TD valign=top><b>Valorizar los saldos a:</b><br><br>&nbsp;</TD>
								<TD nowrap valign=top>
									Tipo de cambio:
									<SELECT name=cboExchangeRateTypes style="width:179">
										<OPTION value=0>-- No valorizar --</OPTION>
										<%=gsCboExchangeRateTypes%>
									</SELECT>
									&nbsp; 
									Moneda:&nbsp; &nbsp; &nbsp;
									<SELECT name=cboExchangeRateCurrencies style="width:220">
										<OPTION value=0>-- No valorizar --</OPTION>
										<%=gsCboCurrencies%>						
									</SELECT>
									<br>
									Fecha del tipo de cambio:
									<INPUT name=txtExchangeRateDate style="width:93" value='<%=gsExchangeRateDate%>'>
									<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToApplicationDate)'>
								</TD>
							</TR>							
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR id=divSelectedBalancesOptions style='display:none;'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap colspan=2>
						¿Qué se desea hacer con los saldos obtenidos?
					</TD>
					<TD nowrap align=right>
					  <img id=cmdSelectedVouchersOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSelectedBalancesOptions, this)' alt='Fijar la ventana'>					
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='/empiria/images/invisible.gif'>
					  <img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divSelectedBalancesOptions, document.all.cmdSelectedVouchersOptionsTack)" alt='Cerrar'>
					</TD>				
				</TR>
				<TR>
					<TD nowrap>
						<A href="" onclick="return(notAvailable());">Construir reporte y enviarlo a otro participante</A> &nbsp; &nbsp; 
						<A href="" onclick="return(notAvailable());">Imprimirlos</A> &nbsp; &nbsp; 
						<A href="" onclick="return(notAvailable());">Exportarlos a Microsoft Excel®</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<THEAD>
					<%=gsExplorerResultsHeader%>
				</THEAD>
				<% If (Len(gsExplorerResultsBody) <> 0) Then %>
					<%=gsExplorerResultsBody%>
				<% Else %>
					<TBODY>
						<TR>
							<TD colspan=9><b>No encontré ninguna cuenta con el criterio de búsqueda proporcionado.</b></TD>
						</TR>
					</TBODY>
				<% End If %>
			</TABLE>
		</TD>
	</TR>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>
<INPUT TYPE=hidden name=txtDisplayAdvancedSearchOptions value='<%=gnDisplayAdvancedSearchOptions%>'>
<INPUT TYPE=hidden name=txtDetailGralLedgerId>
<INPUT TYPE=hidden name=txtDetailCurrencyId>
<INPUT TYPE=hidden name=txtDetailStdActId>
<INPUT TYPE=hidden name=txtDetailSubledgerActId>
<INPUT TYPE=hidden name=txtDetailSectorId>
</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>