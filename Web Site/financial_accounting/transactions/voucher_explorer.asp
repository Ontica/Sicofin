<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsExplorerResultsHeader, gsExplorerResultsBody, gsVoucherInboxes, gsTackedWindows
	Dim gsCboGralLedgerGroups, gsCboTransactionTypes, gsCboVoucherTypes, gsCboBalancingTypes
	Dim gsFromElaborationDate, gsToElaborationDate, gsFromApplicationDate, gsToApplicationDate
	Dim gnGralLedgersCategory, gnGralLedger, gsShowAmountsColumns, gsVoucherNumber, gsVoucherConcept	
	Dim gsAccounts, gnTransactionTypeId, gnVoucherTypeId, gnBalancingType, gbShowAmountsCols
	Dim gnSelectedColumn
	Dim gnSelectedVoucherInbox, gsExceptSelectedTransactionType, gsExceptSelectedVoucherType

	Dim oVoucherUS, nTimeOut
	
	
	Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	
	nTimeOut = Session.Timeout
	Session.Timeout = 10
	Call Main()
	Session.Timeout = nTimeOut
	
	Sub Main()
		Call SetGlobalValues()

		gsCboGralLedgerGroups	= oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), _
																															    CLng(Session("uid")), 2, _
																															    CLng(gnGralLedgersCategory))
		gsCboTransactionTypes = oVoucherUS.CboTransactionTypes(Session("sAppServer"), _
																												   Abs(CLng(gnTransactionTypeId)))
		gsCboVoucherTypes     = oVoucherUS.CboVouchersTypes(Session("sAppServer"), _
																										    Abs(CLng(gnVoucherTypeId)))
		Set oVoucherUS = Nothing
		Call GetExplorerInformation()
	End Sub
	
	Sub GetExplorerInformation()
		Dim oExplorer, sWhere, sOrderBy
		'********************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("EFATransactionsExplorer.CExplorer")
		
		gsVoucherInboxes		    = oExplorer.CboVoucherInboxes(CLng(gnSelectedVoucherInbox))
		gsCboBalancingTypes     = oExplorer.CboBalancingTypes(CLng(gnBalancingType))
		gsExplorerResultsHeader = oExplorer.Header(Session("sAppServer"), CLng(gnSelectedVoucherInbox), _
																							 CLng(gnSelectedColumn), CBool(gbShowAmountsCols))
				
		sWhere = TransactionsFilter()
	
		sOrderBy = "numero_transaccion"
		gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), CLng(gnSelectedVoucherInbox), _
																					 Session("uid"), CBool(gbShowAmountsCols), _
																					 CStr(sWhere), CLng(gnSelectedColumn), CStr(sOrderBy))
		Set oExplorer = Nothing
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		
	End Sub
	
	Function TransactionsFilter()
		Dim oTransactionPars, bIncludeWorkgroupsVouchers
		'********************
		Set oTransactionPars = Server.CreateObject("EFAParameters.CTransactions")		

		bIncludeWorkgroupsVouchers = False
		If (gnSelectedVoucherInbox = 4) Or (gnSelectedVoucherInbox = 5) Or (gnSelectedVoucherInbox = 6) Then
			bIncludeWorkgroupsVouchers = True
		End If
		TransactionsFilter = oTransactionPars.CreateFilter(Session("sAppServer"), _
																											 Session("uid"), CBool(bIncludeWorkgroupsVouchers), _
																											 CLng(gnGralLedgersCategory), CLng(gnGralLedger), _
																											 CStr(gsFromApplicationDate), CStr(gsToApplicationDate), _
																											 CStr(gsFromElaborationDate), CStr(gsToElaborationDate), _
																											 CStr(gsVoucherNumber), CStr(gsVoucherConcept), _
																											 CStr(gsAccounts), CLng(gnTransactionTypeId), _
																											 CLng(gnVoucherTypeId), CLng(gnBalancingType))
		Set oTransactionPars = Nothing
	End Function

  Sub SetGlobalValues()
		If Len(Request.Form("cboVoucherInboxes")) <> 0 Then
			gnSelectedVoucherInbox = Request.Form("cboVoucherInboxes")	
			gnGralLedgersCategory  = Request.Form("cboGralLedgerGroups")
			gnGralLedger           = Request.Form("cboGralLedgers")
			If Len(Request.Form("txtFromApplicationDate")) <> 0 Then
				gsFromApplicationDate  = oVoucherUS.FormatDate(CDate(Request.Form("txtFromApplicationDate")))
			End If
			If Len(Request.Form("txtToApplicationDate")) <> 0 Then
				gsToApplicationDate  = oVoucherUS.FormatDate(CDate(Request.Form("txtToApplicationDate")))
			End If			
			If Len(Request.Form("txtFromElaborationDate")) <> 0 Then
				gsFromElaborationDate = oVoucherUS.FormatDate(CDate(Request.Form("txtFromElaborationDate")))
			ElseIf (Request.Form("cboVoucherInboxes") <> 1) And (Request.Form("cboVoucherInboxes") <> 4) And _
					   (Len(gsFromApplicationDate & gsToApplicationDate) = 0) Then
				gsFromElaborationDate = oVoucherUS.FormatDate(Date())
			End If
			If Len(Request.Form("txtToElaborationDate")) <> 0 Then
				gsToElaborationDate = oVoucherUS.FormatDate(Request.Form("txtToElaborationDate"))
			ElseIf (Request.Form("cboVoucherInboxes") <> 1) And (Request.Form("cboVoucherInboxes") <> 4) And _
						 (Len(gsFromApplicationDate & gsToApplicationDate) = 0) Then
				gsToElaborationDate = oVoucherUS.FormatDate(Date())
			End If
			gnTransactionTypeId    = Request.Form("cboTransactionTypes")
			gnVoucherTypeId        = Request.Form("cboVoucherTypes")
			gsVoucherNumber				 = Request.Form("txtTransactionNumber")
			gsVoucherConcept			 = Request.Form("txtTransactionConcept")
			gsAccounts						 = Request.Form("txtTransactionAccounts")
			gnBalancingType				 = Request.Form("cboBalancingModes")
			gnSelectedColumn			 = Request.Form("txtSelectedColumn")
			gsTackedWindows   = Request.Form("txtTackedWindows")
			If Len(Request.Form("chkTransactionTypes")) <> 0 Then
				gnTransactionTypeId  = -1 * CLng(gnTransactionTypeId)
				gsExceptSelectedTransactionType = "checked"
			End If
			If Len(Request.Form("chkVoucherTypes")) <> 0 Then
				gnVoucherTypeId = -1 * CLng(gnVoucherTypeId)
				gsExceptSelectedVoucherType = "checked"
			End If
			If Len(Request.Form("chkShowAmountsColumns")) <> 0 Then
				gbShowAmountsCols = True
				gsShowAmountsColumns = "checked"
			Else
				gbShowAmountsCols = False
			End If
		Else
			gnSelectedVoucherInbox = 1
			gnGralLedgersCategory  = 0
			gnGralLedger           = 0
			gsFromApplicationDate  = ""
			gsToApplicationDate    = ""
			gsFromElaborationDate  = ""
			gsToElaborationDate    = ""
			gnTransactionTypeId    = 0
			gnVoucherTypeId        = 0
			gsVoucherNumber				 = ""
			gsVoucherConcept			 = ""
			gsAccounts						 = ""
			gnBalancingType				 = 0
			gnSelectedColumn			 = 1
			gsTackedWindows				 = ""
		End If
		If (Len(Request.QueryString("inbox")) <> 0) Then
			gnSelectedVoucherInbox = CLng(Request.QueryString("inbox"))
		End If		
	End Sub	
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

function isDate(sDate) {
	var obj;		
	if (sDate != '') {
		obj = RSExecute("../financial_accounting_scripts.asp","IsDate", sDate);
		return obj.return_value;
	} else {
		return true;
	}
}

function deleteVouchers() {
	var sPendingVouchers = arrayPendingVouchers('chkAllItems');		
	var nSelectedVouchers = countCheckBoxes('chkAllItems');
	var sMsg;
	
	if(nSelectedVouchers == 0) {	
		alert('Para ejecutar esta operación necesito se seleccione al menos una póliza.');
		return false;
	}
	
	if (nSelectedVouchers > 1) {
		sMsg = 'Esta operación eliminará del sistema las ' + nSelectedVouchers + ' pólizas seleccionadas,\n' + 
					 'por lo que ya no podrán ser recuperadas.\n\n' + '¿Procedo con la operación?';
	}
	if (nSelectedVouchers == 1) {
		sMsg = 'Esta operación eliminará del sistema la póliza seleccionada, por lo que ya no\n' + 
					 'podrá ser recuperada.\n' +  '¿Procedo con la operación?';
	}
	if (confirm(sMsg)) {

		if(sPendingVouchers == '') {	
			alert("Las pólizas seleccionadas ya están en el diario o pertenecen a otros usuarios.");
			return false;	
		}	
		document.frmSend.txtPendingVouchers.value = sPendingVouchers;
		document.frmSend.txtPostedVouchers.value = '';
		document.frmSend.action = 'exec/delete_vouchers.asp';
		document.frmSend.submit();
		document.frmSend.action = '';		
	}
	return false;
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
  document.frmSend.target = '_self';
  document.frmSend.action = '';
  document.frmSend.submit();
	return false;
}

function postVouchers() {
	var sPendingVouchers = arrayPendingVouchers('chkAllItems');
	var selectedVouchers = countCheckBoxes("chkAllItems");
	var sMsg;
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}
	if (selectedVouchers > 1) {
		sMsg = 'Esta operación enviará al diario las ' + selectedVouchers + ' pólizas seleccionadas.\n\n' + 
					 'Sin embargo, esto ocurrirá únicamente con las pólizas que estén debidamente balanceadas.\n\n' +
					 '¿Procedo con la operación?';
	}
	if (selectedVouchers == 1) {
		sMsg = 'Esta operación enviará la póliza seleccionada al diario.\n\n' + 
					 'Sin embargo, esto ocurrirá únicamente si esta se encuentra debidamente balanceada.\n\n' + 
					 '¿Procedo con la operación?';
	}		
	if (confirm(sMsg)) {
		if(sPendingVouchers == '') {	
			alert("Las pólizas seleccionadas ya están en el diario o pertenecen a otros usuarios.");
			return false;	
		}	
		document.frmSend.txtPendingVouchers.value = sPendingVouchers;
		document.frmSend.target = "_self"
		document.frmSend.action = "exec/post_vouchers.asp";
		document.frmSend.submit();
		document.frmSend.action = '';
	}
	return false;
}


function arrayPendingVouchers(sCheckBoxName) {
	var i = 0, sTemp = '';
	
	if (typeof(document.all[sCheckBoxName]) == 'undefined') {
		return '';
	}	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if ((document.all[sCheckBoxName](i).checked) && (document.all[sCheckBoxName](i).value < 0)) {
				if (sTemp.length != 0) {					
					sTemp += ',' + Math.abs(document.all[sCheckBoxName](i).value);
				} else {					
					sTemp = Math.abs(document.all[sCheckBoxName](i).value);
				}
			}
		}		
	} else {
		if ((document.all[sCheckBoxName].checked) && (document.all[sCheckBoxName].value < 0)) {
			sTemp = Math.abs(document.all[sCheckBoxName].value);
		}
	}
	return sTemp;
}


function arrayPostedVouchers(sCheckBoxName) {
	var i = 0, sTemp = '';
	
	if (typeof(document.all[sCheckBoxName]) == 'undefined') {
		return '';
	}	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if ((document.all[sCheckBoxName](i).checked) && (document.all[sCheckBoxName](i).value > 0)) {
				if (sTemp.length != 0) {
					sTemp += ',' + document.all[sCheckBoxName](i).value;
				} else {
					sTemp = document.all[sCheckBoxName](i).value;
				}
			}
		}		
	} else {
		if ((document.all[sCheckBoxName].checked) && (document.all[sCheckBoxName].value > 0)) {
			sTemp = document.all[sCheckBoxName].value;
		}
	}
	return sTemp;
}

function printVouchers() {
	var sPendingVouchers = arrayPendingVouchers('chkAllItems');
	var sPostedVouchers  = arrayPostedVouchers('chkAllItems');
	var sMsg;
	
	document.frmSend.txtPendingVouchers.value = sPendingVouchers;
	document.frmSend.txtPostedVouchers.value  = sPostedVouchers;
	
	if ((sPendingVouchers == '') && (sPostedVouchers == '')) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}
	
	document.frmSend.target = "_blank";	
	if (sPendingVouchers != '') {		
		document.frmSend.action = "pending_voucher_viewer.asp";
		document.frmSend.submit();
	}
	
	if (sPostedVouchers != '') {	
		document.frmSend.action = "voucher_viewer.asp";
		document.frmSend.submit();
	}
	return false;
}

function reassignVouchers() {
	var sPendingVouchers = arrayPendingVouchers('chkAllItems');
	
	if(sPendingVouchers == '') {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;	
	}
	document.frmSend.txtPendingVouchers.value = sPendingVouchers;
			
	document.frmSend.target = '_self';
	document.frmSend.action = "exec/reassign_vouchers.asp";
	document.frmSend.submit();
	return false;
}

function resetSearchOptions() {
	var sTemp = document.all.cboVoucherInboxes.value;
	
	document.frmSend.reset();
	document.all.cboVoucherInboxes.value =(sTemp);
	return false;
}

function searchVouchers() {
	if (!isDate(document.all.txtFromApplicationDate.value)) { 
		alert("No reconozco la fecha de afectación inicial de las pólizas a buscar.");
		document.all.txtFromApplicationDate.focus();
		return false;
	}
	if (!isDate(document.all.txtToApplicationDate.value)) { 
		alert("No reconozco la fecha de afectación final de las pólizas a buscar.");
		document.all.txtToApplicationDate.focus();
		return false;
	}	
	if (!isDate(document.all.txtFromElaborationDate.value)) { 
		alert("No reconozco la fecha de elaboración inicial de las pólizas a buscar.");
		document.all.txtFromElaborationDate.focus();
		return false;
	}		
	if (!isDate(document.all.txtToElaborationDate.value)) { 
		alert("No reconozco la fecha de elaboración final de las pólizas a buscar.");
		document.all.txtToElaborationDate.focus();	
		return false;
	}
	document.frmSend.target = '_self';
	document.frmSend.action = '';
	document.frmSend.submit();
	return false;
}

function setDynamicItems() {	
	var sTemp = document.all.cboVoucherInboxes.value;
	
	if (sTemp == '2' || sTemp == '5') {
		document.all.rowBalancesModes.style.display = 'none';
		document.all.cboBalancingModes.value = 0;
		document.all.rowTransactionNumber.style.display = 'inline';
	} else {
		document.all.rowBalancesModes.style.display = 'inline';
		if (sTemp == '1' || sTemp == '3' || sTemp == '4' || sTemp == '6') {
			document.all.rowTransactionNumber.style.display = 'none';
			document.all.txtTransactionNumber.value = '';
		} else { 
			document.all.rowTransactionNumber.style.display = 'inline';
		}
	}
}

function insertTableInformation(oTableSection, nObjectId) {
  var oRow, i, j;
	var obj, aRows, aCols, sTemp;
	
  setCursor('wait');
  obj = RSExecute("../financial_accounting_scripts.asp", "Vouchers", nObjectId);
  
  sTemp = obj.return_value;
	aRows = sTemp.split("¦");
  for (i = 0; i < aRows.length; i++) {
		aCols = aRows[i].split("|");
		oRow = oTableSection.insertRow();
		for (j = 0; j < aCols.length; j++) {
			oRow.insertCell().innerHTML = aCols[j];
		}
  }
	setCursor('auto');
}

function updateGralLedgers(nSelectedItem) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGralLedgerGroups.value, nSelectedItem);
	document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="updateGralLedgers(<%=gnGralLedger%>);setDynamicItems();showTackedWindows(Array(<%=gsTackedWindows%>));">
<FORM name=frmSend action='' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Explorador de pólizas&nbsp;
		</TD>
		<TD colspan=3 align=right nowrap>						<A href="voucher_wizard.asp">Crear póliza</A>						
			<img align=absbottom src='/empiria/images/invisible4.gif'>			<img align=absbottom src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absbottom src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absbottom src='/empiria/images/invisible.gif'>
			<img align=absbottom src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
	</TR>
	<TR>
		<TD id=divSearchOptions colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Consulta de pólizas
					</TD>
					<TD nowrap align=left>
						<A href="" onclick="return(notAvailable());">Cargar consulta</A>
						&nbsp; | &nbsp
						<A href="" onclick="return(notAvailable());">Guardar consulta</A>
						<img align=absbottom src='/empiria/images/invisible8.gif'>
						<A href="" onclick="return(searchVouchers());">Ejecutar consulta</A>
					</TD>
					<TD nowrap align=right>
						<A href='' onclick="return(showOptionsWindow(document.all.divAdvancedSearchOptions));">Más opciones</A>
						<img align=absbottom src='/empiria/images/invisible4.gif'>						
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>
					  <img align=absbottom id=cmdSearchOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSearchOptions, this)' alt='Fijar la ventana'>
						<img align=absbottom src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
				</TR>
				<TR>
					<TD nowrap><b>Mostrar:</b></TD>
					<TD colspan=2 nowrap width=100%>
						<SELECT name=cboVoucherInboxes onchange="setDynamicItems();">
							<%=gsVoucherInboxes%>
						</SELECT>&nbsp;
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>En el grupo de contabilidades:</b></TD>
					<TD colspan=2 nowrap width=100%>
						<SELECT name=cboGralLedgerGroups style="width:100%" onchange="return updateGralLedgers(0);">
							<OPTION value=0 selected> -- Todos los grupos de contabilidades -- </OPTION>
							<%=gsCboGralLedgerGroups%>
						</SELECT>
					</TD>
				</TR>
				<TR>
					<TD nowrap valign=top><b>De la contabilidad:</b><br>&nbsp;</TD>					
					<TD colspan=2 nowrap width=100% valign=top>
						<span id=divCboGeneralLedgers>
							<SELECT name=cboGralLedgers" width=100%>

							</SELECT>
						</span>						
					</TD>
				</TR>				
				<TR>
					<TD nowrap><b>Con fecha de afectación:</b></TD>
					<TD colspan=2 nowrap width=100%>
						Del día:
						<INPUT name=txtFromApplicationDate style="width:100;height:20;" value='<%=gsFromApplicationDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromApplicationDate)'>
						&nbsp; &nbsp; &nbsp; al día:
						<INPUT name=txtToApplicationDate style="width:100;height:20;" value='<%=gsToApplicationDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToApplicationDate)'>&nbsp;&nbsp;
						(día / mes / año)
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Con fecha de elaboración:</b></TD>
					<TD colspan=2 nowrap width=100%>
						Del día:
						<INPUT name=txtFromElaborationDate style="width:100;height:20;" value='<%=gsFromElaborationDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromElaborationDate)'>
						&nbsp; &nbsp; &nbsp; al día:
						<INPUT name=txtToElaborationDate style="width:100;height:20;"  value='<%=gsToElaborationDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToElaborationDate)'>&nbsp;&nbsp;
						(día / mes / año)
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
							<TR id=rowTransactionNumber style='display:none;'>
								<TD nowrap><b>Cuyo número de póliza sea:</b>&nbsp;</TD>
								<TD colspan=2 nowrap>
									<INPUT name=txtTransactionNumber style="width:161;height:20;" value='<%=gsVoucherNumber%>'> 
										&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'))" target=_blank>comodines</A>)
								</TD>
							</TR>				
							<TR>
								<TD nowrap><b>Cuyo concepto sea:</b></TD>
								<TD colspan=2 nowrap width=100% valign=middle>
									<INPUT name=txtTransactionConcept style="width:400;height:20;" value='<%=gsVoucherConcept%>'>
										&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'));" target=_blank>comodines</A>)
								</TD>
							</TR>
							<TR>
								<TD nowrap><b>Con tipo de transacción:</b></TD>
								<TD colspan=2 nowrap width=100%>
									<SELECT name=cboTransactionTypes style='width:300'>
										<OPTION value=0 selected> -- Todas las transacciones -- </OPTION>
										<%=gsCboTransactionTypes%>
									</SELECT>&nbsp;Todas excepto las del tipo seleccionado
									<INPUT type="checkbox" name=chkTransactionTypes value="true" <%=gsExceptSelectedTransactionType%>>
								</TD>
							</TR>
							<TR>
								<TD nowrap><b>Con tipo de póliza:</b></TD>
								<TD colspan=2 nowrap>
									<SELECT name=cboVoucherTypes style='width:300'>			
										<OPTION value=0 selected> -- Todos los tipos de póliza -- </OPTION>
											<%=gsCboVoucherTypes%>
									</SELECT>&nbsp;Todas excepto las del tipo seleccionado
									<INPUT type="checkbox" name=chkVoucherTypes value="true" <%=gsExceptSelectedVoucherType%>>
								</TD>
							</TR>
							<TR>
								<TD nowrap><b>Que contengan las cuentas:</b></TD>
								<TD colspan=2 nowrap width=100% valign=middle>
									<INPUT name=txtTransactionAccounts style="width:400;height:20;" value='<%=gsAccounts%>'> &nbsp;&nbsp;(lista separada por comas)
								</TD>
							</TR>
							<TR id=rowBalancesModes>
								<TD nowrap><b>Filtar por balance:</b></TD>
								<TD colspan=2 nowrap>
									<SELECT name=cboBalancingModes style='width:300'>
										<%=gsCboBalancingTypes%>
									</SELECT>
								</TD>
							</TR>														
							<TR>
								<TD nowrap><b>Columnas a desplegar:</b></TD>
								<TD colspan=2 nowrap>
									Mostrar las columnas con las sumas de cargos y abonos &nbsp; &nbsp;
									<INPUT type="checkbox" name=chkShowAmountsColumns value="true" <%=gsShowAmountsColumns%>>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR id=divSelectedVouchersOptions style='display:none;'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap colspan=2>
						¿Qué se desea hacer con las pólizas seleccionadas?
					</TD>
					<TD nowrap align=right>
					  <img id=cmdSelectedVouchersOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSelectedVouchersOptions, this)' alt='Fijar la ventana'>					
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='/empiria/images/invisible.gif'>
					  <img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divSelectedVouchersOptions, document.all.cmdSelectedVouchersOptionsTack)" alt='Cerrar'>
					</TD>				
				</TR>
				<TR>
					<TD nowrap>
						<A href="" onclick="return(printVouchers());">Imprimirlas</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(postVouchers());">Enviarlas al diario</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(reassignVouchers());">Enviarlas a otro participante</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(reassignVouchers());">Exportarlas a Microsoft Excel<sup>®</sup></A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(deleteVouchers());">Eliminarlas</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
	<TR id=divTasksOptions style='display:none'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Tareas
					</TD>
					<TD nowrap align=left>
						<A href="" onclick="return(notAvailable());">Lista de tareas</A>
						&nbsp; | &nbsp
						<A href="" onclick="return(notAvailable());">Mi lista de tareas pendientes</A>
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
						<A href="../../balances/balance_explorer.asp">Explorador de saldos</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../reports/balances.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../reports/other_reports.asp">Reportes</A>
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
		<TD colspan=8><b>No encontré ninguna póliza con el criterio de búsqueda proporcionado.</b></TD>
		</TR>
	</TBODY>
<% End If %>
</TABLE>
<INPUT TYPE=hidden name=txtPostedVouchers>
<INPUT TYPE=hidden name=txtPendingVouchers>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>
</TD>
</TR>
</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>