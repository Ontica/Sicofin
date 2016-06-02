<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If

	Dim oVoucherUS
		
	Dim bOperation, bSource, nGralLedgerGroupType, bGralLedgers, bShowCboApplicationDate
	Dim bVoucherNumber, bFirstInputAccount, bUseWizard, bUseNextStep
	Dim gsVoucherTypeDescription, sLblApplicationDate, gnTransactionType
	Dim gsCboApplicationDates, gsOutOfPeriodDate,	gsCboGLCategories, gsCboGeneralLedgers
	Dim bShowFromAccount, bShowCurrencyConvertion
	Dim gsCboSources, gsCboOperations, gsCboVoucherTypes, gsCboSectors, gsCboCurrencies, gsTackedWindows
	
	Call Main()
			 
	Sub Main()
		Dim oRecordset, nSelectedVoucherType
		'***********************************
		'On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		If (Len(Request.QueryString("id")) <> 0) Then
			nSelectedVoucherType = CLng(Request.QueryString("id"))
		Else
			nSelectedVoucherType = 25
		End If
		gsCboVoucherTypes = oVoucherUS.CboVouchersTypesForAppend(Session("sAppServer"), CLng(Session("uid")), _
																														 CLng(nSelectedVoucherType))
		If nSelectedVoucherType < 0 Then
			Set oRecordset = oVoucherUS.VoucherTypeRS(Session("sAppServer"), CLng(Abs(nSelectedVoucherType)))
			gsVoucherTypeDescription  = oRecordset("object_description")
		Else
			gsVoucherTypeDescription  = "No intervendrá el asistente"
		End If
		Set oRecordset = Nothing
		
		gsTackedWindows = Request.Form("txtTackedWindows")			
		Call SetParameters(Abs(nSelectedVoucherType))
		Call RetriveParametersInformation()
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		  
	End Sub		

	Sub SetParameters(gnVoucherTypeId)
		bGralLedgers            = True
		bShowCboApplicationDate = True
		bUseWizard              = True
		bSource                 = True
		bOperation              = False
		bShowFromAccount        = False
		bShowCurrencyConvertion				= False
		gnTransactionType = 23
	  Select Case gnVoucherTypeId
			Case 26
				bUseNextStep = True
				bOperation = True				
				nGralLedgerGroupType = 1
				bFirstInputAccount = True
				sLblApplicationDate = "Fecha de afectación"				
			Case 28
				nGralLedgerGroupType = 0
				bGralLedgers = False
				bShowCboApplicationDate = False
				sLblApplicationDate = "Concentrar movimientos del día"		
			Case 29
				nGralLedgerGroupType = 2
				bShowCboApplicationDate = False
				sLblApplicationDate = "Traspasar deficientes o remanentes al día"
			Case 30
				nGralLedgerGroupType = 2
				bShowCboApplicationDate = False
				sLblApplicationDate = "Cancelar cuentas de resultados al día"
			Case 64
				nGralLedgerGroupType = 1
				bVoucherNumber = True
				sLblApplicationDate = "Fecha de afectación"
			Case 120
				nGralLedgerGroupType = 2
				bVoucherNumber = False
				bShowCboApplicationDate = False
				bShowFromAccount = True
				bShowCurrencyConvertion = True
				sLblApplicationDate = "Convertir los saldos al día"
			Case Else
				bUseWizard = False	
				nGralLedgerGroupType = 1
				sLblApplicationDate = "Fecha de afectación"
				gnTransactionType = 24
		End Select
	End Sub	
	
	Sub RetriveParametersInformation()
		If bOperation Then
			gsCboOperations = oVoucherUS.CboVoucherTypeOperations(Session("sAppServer"), CLng(Request.QueryString("id")))
		End If
		If bSource Then
			gsCboSources = oVoucherUS.CboSources(Session("sAppServer"))
		End If
		If bShowCboApplicationDate Then
			gsCboApplicationDates = "<span id=divCboApplicationDates> </span>"
		End If
		If bShowFromAccount Then
			gsCboSectors = oVoucherUS.CboSectors(Session("sAppServer"))
		End If
		If bShowCurrencyConvertion Then				
			gsCboCurrencies = oVoucherUS.CboCurrencies(Session("sAppServer"), 1)
		End If
		If Not bOperation Then
			If nGralLedgerGroupType = 1 Then
				gsCboGLCategories = oVoucherUS.CboGeneralLedgerFilledCategories(Session("sAppServer"),  CLng(Session("uid")), 1)
			ElseIf nGralLedgerGroupType = 2 Then
				gsCboGLCategories = oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), CLng(Session("uid")), 2)
			End If
		End If
	End Sub
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var gbSended = false;

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsNumeric", sNumber, nDecimals);
	return obj.return_value;
}

function transactionType(nVoucherId) {
	var obj;	
	
	obj = RSExecute("../financial_accounting_scripts.asp", "TransactionType", nVoucherId);	
	return obj.return_value;
}

function voucherId(nGralLedgerId, sVoucherNumber) {
	var obj;	
	
	obj = RSExecute("../financial_accounting_scripts.asp", "VoucherId", nGralLedgerId, sVoucherNumber);	
	return obj.return_value;	
}


function updateCboSources() {
	var obj;
	/*
	if (document.all.cboGralLedgers.value <= 32) {	
	   obj = RSExecute("../financial_accounting_scripts.asp", "CboGLSources", document.all.cboGralLedgers.value, 1);
	} else {
		 obj = RSExecute("../financial_accounting_scripts.asp", "CboGLSources", document.all.cboGralLedgers.value, 434);
  }
	document.all.divCboSources.innerHTML = obj.return_value;	
  */
}

function updateInfo() {
	<% If bShowCboApplicationDate Then %>
		updateCboApplicationDates();
	<% End If %>
	<% If bSource Then %>
		updateCboSources();	
	<% End If %>
}

function updateCboGLCategories() {
	var obj, sTemp

	if (document.all.cboOperations.value != 0) {
		<% If nGralLedgerGroupType = 1 Then %>					
		  obj = RSExecute("../financial_accounting_scripts.asp", "CboGeneralLedgerFilledCategories", 2, 0, -1 * document.all.cboOperations.value);			
		<% Else %>			
			obj = RSExecute("../financial_accounting_scripts.asp", "CboGeneralCategories", 1, 0, -1 * document.all.cboOperations.value);
		<% End If %>
		document.all.divCboGLCategories.innerHTML = obj.return_value;	
	} else {
		sTemp  = '<SELECT name=cboGLCategories style="WIDTH: 520px">';
		sTemp	+= '<OPTION value=0><< Primero debe seleccionarse la operación >></OPTION>';
    sTemp += '</SELECT>';
		document.all.divCboGLCategories.innerHTML = sTemp;
	}
	updateCboGralLedgers();
}

function updateCboGralLedgers() {
	var obj, sTemp
  <% If bOperation Then %>
		if (document.all.cboOperations.value != 0) {
			<% If nGralLedgerGroupType = 1 Then %>		
				obj = RSExecute("../financial_accounting_scripts.asp", "CboGLInCategory", document.all.cboGLCategories.value);
			<% Else %>
				obj = RSExecute("../financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGLCategories.value);	
			<% End If %>
			document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
		} else {
			sTemp  = '<SELECT name=cboGralLedgers style="WIDTH: 520px">';
			sTemp	+= '<OPTION value=0><< Primero debe seleccionarse la operación >></OPTION>';
			sTemp += '</SELECT>';
			document.all.divCboGeneralLedgers.innerHTML = sTemp;
		}
	<% Else %>
		<% If nGralLedgerGroupType = 1 Then %>
			obj = RSExecute("../financial_accounting_scripts.asp", "CboGLInCategory", document.all.cboGLCategories.value);
		<% Else %>
			obj = RSExecute("../financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGLCategories.value);	
		<% End If %>
		document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
	<% End If %>
	<% If bShowCboApplicationDate Then %>
		updateCboApplicationDates();
	<% End If %>
}

function updateCboApplicationDates() {	
	var obj;
	// <% If Request.QueryString("id") <> 28 Then %>	
	obj = RSExecute("../financial_accounting_scripts.asp", "CboOpenPeriodsDates", document.all.cboGralLedgers.value);
	//<% Else %>
	//obj = RSExecute("../financial_accounting_scripts.asp", "CboOpenPeriodsDates", 9);
	//<% End If %>
	document.all.divCboApplicationDates.innerHTML = obj.return_value;	
}

function validateDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function window_onload() {
  <% If bGralLedgers AND (NOT bOperation) Then %>
	  updateCboGralLedgers();
  <% End If %>
  document.all.txtDescription.focus();
  showTackedWindows(Array(<%=gsTackedWindows%>));
}

function cmdGenerateConcept_onclick() {
	alert("Por el momento esta opción no está disponible.\n\nGracias.");
}

function cmdCheckSpelling_onclick() {
	alert("Por el momento esta opción no está disponible.\n\nGracias.");
}

function showApplicationDatePicker() {	
	var sDate = '';
	var sOptions = 'dialogHeight:250px;dialogWidth:350px;resizable:no;scroll:no;status:no;help:no;';
		
	if (document.all.txtOutOfPeriodDate.value == '') {
		sDate = window.showModalDialog('voucher_date_picker.asp', "" , sOptions);	
		if (sDate != '') {
			document.all.txtOutOfPeriodDate.value = sDate;
			document.all.cboApplicationDates.disabled = true;
			document.all.divOutOfPeriodDate.innerText = document.all.txtOutOfPeriodDate.value;
			document.all.divOutOfPeriodDateText.innerText = 'Haga clic en esta liga para anular la fecha valor o adelantada.';
		} else {
			document.all.txtOutOfPeriodDate.value = '';
			document.all.cboApplicationDates.disabled = false;
			document.all.divOutOfPeriodDate.innerText = 'Ninguna';
			document.all.divOutOfPeriodDateText.innerText = 'Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.';
		}
		return false;
	}
	if (document.all.txtOutOfPeriodDate.value != '') {
		document.all.txtOutOfPeriodDate.value = '';
		document.all.cboApplicationDates.disabled = false;
		document.all.divOutOfPeriodDate.innerText = 'Ninguna';
		document.all.divOutOfPeriodDateText.innerText = 'Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.';	
		return false;
	}
}

function doSubmit() {
	var oTxtDescription = document.all.frmSend.txtDescription;
	var oTxtOutOfPeriodDate = document.all.frmSend.txtOutOfPeriodDate;
	var sMsg, nVoucherId;
	
	if (gbSended) {
		return false;
	}
	if (oTxtDescription.value == "") {
		alert("Necesito el concepto de la póliza.");
		oTxtDescription.focus();
		return false;
	}	
	<% If Len(gsCboVoucherTypes) <> 0 Then %>		
	 document.all.txtVoucherType.value = Math.abs(document.all.cboVoucherTypes.value)
	<% End If %>	
	<% If bOperation Then %>
		if (document.all.cboOperations.value == 0) {
			alert("Necesito se seleccione la operación que se aplicará.");
			document.all.cboOperations.focus();
			return false;
		}
	<% End If %>	
	<% If bGralLedgers AND (nGralLedgerGroupType = 1) Then %>
		if (document.all.cboGralLedgers.value == 0) {
			alert("Necesito se seleccione la contabilidad.");
			document.all.cboGralLedgers.focus();
			return false;
		}
	<% End If %>
	<% If bGralLedgers AND (nGralLedgerGroupType = 2) Then %>
		if (document.all.cboGralLedgers.value == -1) {
			alert("Necesito se seleccione la contabilidad.");
			document.all.cboGralLedgers.focus();
			return false;
		}
	<% End If %>
	<% If bShowFromAccount Then %>
		if (document.all.txtFromAccount.value == '') {
			alert("Necesito el número de la cuenta origen.");
			document.all.txtFromAccount.focus();
			return false;
		}
	<% End If %>
	<% If bShowCurrencyConvertion Then %>
		if (document.all.txtExchangeRate.value == '') {
			alert("Necesito el tipo de cambio con el que se hará la conversión de los saldos.");
			document.all.txtExchangeRate.focus();
			return false;
		}
		if (!isNumeric(document.all.txtExchangeRate.value, 6)) {
			alert("No reconozco el tipo de cambio proporcionado.");
			document.all.txtExchangeRate.focus();
			return false;
		}
	<% End If %>
	<% If bShowCboApplicationDate Then %>
		if (document.all.cboApplicationDates.value == '' && oTxtOutOfPeriodDate.value == '') {
			alert("No hay períodos abiertos y la póliza no tiene una fecha valor o adelantada.");
			return false;
		}
		if (oTxtOutOfPeriodDate.value == '') {
			if (!validateDate(document.all.cboApplicationDates.value)) {
				alert("No reconzco la fecha de la póliza. Debe existir un problema con el manejo de períodos.");
				return false;
			}
		}
		if (oTxtOutOfPeriodDate.value != '') {
			if (validateDate(oTxtOutOfPeriodDate.value)) {
				sMsg  = "La póliza será registrada con fecha valor del día " + oTxtOutOfPeriodDate.value + ".\n\n";
				sMsg += "Estas pólizas sólo pueden ser enviadas al diario por el supervisor.\n\n";
				sMsg += "¿Procedo con la creación de la póliza?";
				if (!confirm(sMsg)) {
					return false;
				}
			} else {
				alert("No reconozco la fecha valor o adelantada de la póliza.");
			}
		}
	<% Else %>
		if (oTxtOutOfPeriodDate.value == '') {
			alert("Requiero la fecha de aplicación");
			return false;
		}
		if (!validateDate(oTxtOutOfPeriodDate.value)) {			
				alert("No reconozco la fecha valor proporcionada.");
				oTxtOutOfPeriodDate.focus();
				return false;
		} else {
			sMsg  = "La póliza será registrada y elaborada con la información al día " + oTxtOutOfPeriodDate.value + ".\n\n";
			sMsg += "¿Procedo con la generación de la póliza?";
			if (!confirm(sMsg)) {
				return false;
			}		
		}	
	<% End If %>
	<% If bVoucherNumber Then  %>
		if (document.all.txtVoucherNumber.value == '') {			
			alert("Se necesita el número de la póliza que se requiere cancelar.");
			document.all.txtVoucherNumber.focus();
			return false;
		}
		nVoucherId = voucherId(document.all.cboGralLedgers.value, document.all.txtVoucherNumber.value)
		if (nVoucherId == 0) {
			alert("No encontré ninguna póliza con ese número en la contabilidad especificada.");
			document.all.txtVoucherNumber.focus();
			return false;
		}
		if (transactionType(nVoucherId) == 21) {
			alert("La póliza seleccionada no perimite revertir los movimientos por ser póliza de conversión de cuentas.");
			document.all.txtVoucherNumber.focus();
			return false;
		}
	<% End If %>
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function getVoucherWizard() {
	document.all.frmSend.action = 'voucher_wizard.asp?id=' + document.all.cboVoucherTypes.value;
	document.all.frmSend.submit();
	return false;
}

function txtFromAccount_onblur() {
	var obj;
	
	if (document.all.txtFromAccount.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountWithGLId", 1, document.all.txtFromAccount.value);
		if (obj.return_value != '') {
			document.all.txtFromAccount.value = obj.return_value;			
		} else {
			alert("No entiendo el formato de la cuenta origen.");						
		}
	}	
	return true;
}

function txtExchangeRate_onblur() {
	var obj;	
	if (document.all.txtExchangeRate.value != '' && isNumeric(document.all.txtExchangeRate.value, 6)) {
		obj = RSExecute("../financial_accounting_scripts.asp","FormatCurrency", document.all.txtExchangeRate.value, 6);
		document.all.txtExchangeRate.value = obj.return_value;
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Creación de pólizas
		</TD>
	  <TD align=right nowrap>
			<img align=middle src='/empiria/images/invisible8.gif'>
			<img align=middle src='/empiria/images/invisible8.gif'>			<img align=middle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=middle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=middle src='/empiria/images/invisible.gif'>
			<img align=middle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">
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
						<A href="voucher_explorer.asp">Explorador de pólizas</A>
						&nbsp;&nbsp;&nbsp;&nbsp;						
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
			<FORM name=frmSend action='exec/create_transaction.asp' method=post>
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD colspan=4 class=fullScrollMenuTitle>Información general de la póliza</TD>
				</TR>
			  <TR>
					<TD valign=top nowrap>Tipo de póliza:</TD>
			    <TD colspan=3 width=100%>
					<SELECT name=cboVoucherTypes onchange='return(getVoucherWizard());'>
						<%=gsCboVoucherTypes%>
					</SELECT>
			    </TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>¿Qué hará el asistente?:</TD>
					<TD colspan=3 bgcolor=LightSteelBlue>
						<%=gsVoucherTypeDescription%>
					</TD>
			  </TR>
			  <% If bOperation Then %>
			  <TR>
					<TD valign=top nowrap>Operación a ejecutar:</TD>
			    <TD colspan=3>
							<SELECT name=cboOperations style="WIDTH: 520px" onchange='return updateCboGLCategories()'>
								<OPTION value=0><< Seleccionar la operación >></OPTION>
								<%=gsCboOperations%>
							</SELECT>
			    </TD>
			  </TR>
			  <% End If %>
			  <TR>
			    <TD valign=top nowrap>Concepto de la póliza:</TD>
			    <TD colspan=3>
						<TEXTAREA name=txtDescription ROWS=3 style="WIDTH: 520px"></TEXTAREA><br>
						<INPUT type=button class=cmdSubmit name=cmdGenerateConcept value="Sugerir el concepto" onclick="return cmdGenerateConcept_onclick()">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<INPUT type=button class=cmdSubmit name=cmdCheckSpelling value="Revisar ortografía" onclick="return cmdCheckSpelling_onclick()">
			    </TD>
			  </TR>
			  <% If (nGralLedgerGroupType <> 0) Then %>
			  <TR>
			    <TD nowrap>Tipo de contabilidad:</TD>
			    <TD>
						<div id=divCboGLCategories>
							<SELECT name=cboGLCategories style="WIDTH: 520px" onchange='return updateCboGralLedgers()'>
							<% If bOperation Then %>
								<OPTION value=-1><< Primero debe seleccionarse la operación >></OPTION>				
							<% Else %>				
								<%=gsCboGLCategories%>
							<% End If %>
							</SELECT>
						</div>
			    </TD>
			  </TR>
			  <% End If %>
			  <% If bGralLedgers Then %>
			  <TR>
			    <TD>Contabilidad:</TD>
			    <TD>
						<div id=divCboGeneralLedgers>
						<% If bOperation Then %>
							<SELECT name=cboGralLedgers style="WIDTH: 520px"> 
								<OPTION value=-1><< Primero debe seleccionarse la operación >></OPTION>
			        </SELECT>				
						<% End If %>
						</div>
			    </TD>
			  </TR>
			  <% End If %>
				<% If bSource Then %>
			  <TR>
			    <TD nowrap>Origen de la transacción: &nbsp; &nbsp; &nbsp;</TD>        
			    <TD>
						<span id=divCboSources>
							<SELECT name=cboSources style="WIDTH: 520px"> 
								<%=gsCboSources%>
							</SELECT>
			     </span>
			    </TD>
			  </TR>
			  <% End If %>  
			  <% If bVoucherNumber Then %>
			  <TR>
			    <TD>Número de póliza a cancelar:</TD>
			    <TD colspan=3> 
						<INPUT name=txtVoucherNumber value="" style="WIDTH: 130px">
					</TD>
			  </TR>    
			  <% End If %>	  
			  <% If bShowFromAccount Then %>
			  <TR>
			    <TD nowrap>Cuenta origen: &nbsp; &nbsp; &nbsp;</TD>        
			    <TD nowrap>
						<INPUT name=txtFromAccount value="" style="WIDTH: 130px" onblur="return txtFromAccount_onblur()">&nbsp;&nbsp;
						Auxiliar: <INPUT name=txtFromSubsidiaryAccount value="" style="WIDTH: 100px">&nbsp;&nbsp;
						Sector:
						<SELECT name=cboFromSector style="WIDTH: 192px">
							<OPTION value=0 selected>--Todos los sectores--</OPTION>
							<%=gsCboSectors%>
						</SELECT>
						<br>
						(*) Tip: La cuenta origen y su auxiliar permiten comodines [* ?]
					</TD>
			  </TR>
			  <% End If %>
			  <% If bShowCurrencyConvertion Then %>
			  <TR>
			    <TD nowrap>De la moneda: &nbsp; &nbsp; &nbsp;</TD>   
			    <TD nowrap>
			    	<SELECT name=cboFromCurrency style="WIDTH: 155px">							
							<%=gsCboCurrencies%>
						</SELECT>&nbsp;
						A la moneda:
			    	<SELECT name=cboToCurrency style="WIDTH: 155px">							
							<%=gsCboCurrencies%>
						</SELECT>&nbsp;
						T. de cambio:
						<INPUT name=txtExchangeRate value="" style="WIDTH: 64px" onblur="return txtExchangeRate_onblur()">&nbsp;&nbsp;
			    </TD>
			  </TR>
			  <% End If %>			  
			  <TR>
			    <TD valign=top><%=sLblApplicationDate%>:</TD>
			  <% If bShowCboApplicationDate Then %>
			    <TD colspan=3> 			
						<%=gsCboApplicationDates%>
						&nbsp;
						<a href='' onclick='showApplicationDatePicker();return false;'>
							<span id=divOutOfPeriodDateText>Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.
							</span>
						</a>
						<br><br>
						Fecha valor o adelantada:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><span id=divOutOfPeriodDate>Ninguna</span></b>
						<INPUT type="hidden" name=txtOutOfPeriodDate>
					</TD>
				<% Else %>
					<TD colspan=3>
						<INPUT type="text" name=txtOutOfPeriodDate style="WIDTH: 100px"> (día / mes / año)
					</TD>
				<% End If %>
			  </TR>
			  <% If bFirstInputAccount Then %>
			  <TR>
			    <TD>Movimiento base:</TD>
			    <TD colspan=3> 			
						Número de cuenta:<INPUT name=txtFromAccount value="" style="WIDTH: 130px">&nbsp;&nbsp;
						Auxiliar:<INPUT name=txtFromSubsidiaryAccount value="" style="WIDTH: 130px">&nbsp;&nbsp;
						Sector:<INPUT name=txtFromSector value="" style="WIDTH: 130px">
						<br>
						Moneda:<INPUT name=txtFromCurrency value="" style="WIDTH: 130px">&nbsp;&nbsp;
						Monto:<INPUT name=txtFromAmount value="" style="WIDTH: 130px">&nbsp;&nbsp;
						Tipo de cambio:<INPUT name=txtFromAmount value="" style="WIDTH: 130px">&nbsp;&nbsp;
						Monto moneda base:<INPUT name=txtFromBaseAmount value="" style="WIDTH: 130px">&nbsp;&nbsp;
					</TD>
			  </TR>    
			  <% End If %>
			  <TR>
					<td>&nbsp;</td>
			    <td colspan=2 nowrap align=right>		 
					 <INPUT type="hidden" name=txtTransactionType value="<%=gnTransactionType%>">
					 <INPUT type="hidden" name=txtVoucherType value="<%=Request.QueryString("id")%>">
					 <INPUT TYPE=hidden name=txtTackedWindows>
					 <% If bUseWizard Then %>
							<% If bUseNextStep Then %>
								<INPUT class=cmdSubmit type=button name=cmdNext value="Siguiente >>" onclick="doSubmit();">					
							<% Else %>				  
								<INPUT class=cmdSubmit type=button name=cmdSend value="Crear póliza" onclick="doSubmit();">
							<% End If %>
			     <% Else %>
			        <INPUT class=cmdSubmit type=button name=cmdSend value="Crear póliza" onclick="doSubmit();">
			     <% End If %>
			     &nbsp;&nbsp;&nbsp;&nbsp;
			    </td>
			  </TR>
			</TABLE>
			</FORM>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>