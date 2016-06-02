<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect "../../exit_msg.html"
	End If
	
	Dim gnTransactionId, gnPostingId 
	
	Dim gsPostingDate, gsDescription, gnAccountId, gsAccount, gsSector, gsSubsidiaryAccount, gsSubsidiaryAccountName
	Dim gnPostingReferenceId, gspostingReference, gsResponsibilityArea
	Dim gsBudgetKey, gsDisponibilityKey, gsVerificationNumber, gsCboCurrencies, gnCurrency
	Dim gsAmount, gsBaseAmount, gsExchangeRate, gbSelectDebit, gbSelectCredit
	Dim gnTransactionPostingsCount, gsRefreshURL, gsSubsidiaryAccountPrefix
		
	Dim gnGralLedgerId, gsGralLedgerNumber, gsGralLedgerSubNumber
	Dim gsSendURL, gsTitle, gsGLBaseCurrency
	
	gnTransactionId = Request.QueryString("transactionId")
	gsRefreshURL = "posting_editor.asp?transactionId=" & gnTransactionId
	gsSendURL = "exec/save_posting.asp?transactionId=" & gnTransactionId

	If (Len(Request.QueryString("getLast")) <> 0) Then
		gsTitle = "Agregar movimiento"
		gnPostingId = 0
		Call SetValues(gnTransactionId, -1)
	ElseIf (Len(Request.QueryString("clone")) <> 0) Then
		gsTitle = "Agregar movimiento"
		gnPostingId = CLng(Request.QueryString("id"))
		gsRefreshURL = gsRefreshURL & "&id=" & gnPostingId		
		Call SetValues(gnTransactionId, gnPostingId)
		gnPostingId = 0
	ElseIf (Len(Request.QueryString("id")) = 0) Then
		gsTitle = "Agregar movimiento"
		gnPostingId = 0
		gnAccountId = 0		
		Call SetValues(gnTransactionId, 0)
	Else
		gsTitle = "Editar movimiento"
		gnPostingId = CLng(Request.QueryString("id"))
		gsRefreshURL = gsRefreshURL & "&id=" & gnPostingId		
		gsSendURL = gsSendURL & "&postingId=" & gnPostingId
		Call SetValues(gnTransactionId, gnPostingId)
	End If
				
	Sub SetValues(nTransactionId, nPostingId)
		Dim oVoucherUS, oRecordset
		'*************************
		On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		gnTransactionPostingsCount = oVoucherUS.CountPostings(Session("sAppServer"), CLng(nTransactionId))
		Set oRecordset				= oVoucherUS.GetTransactionRS(Session("sAppServer"), CLng(nTransactionId))
		gsGLBaseCurrency			= oRecordset("id_moneda_base") 
		gsPostingDate					= oRecordset("fecha_afectacion")
		gnGralLedgerId				= CLng(oRecordset("id_mayor"))		
		gsGralLedgerNumber    = oRecordset("numero_mayor")
		If IsNull(oRecordset("sub_numero_mayor")) Then
			gsGralLedgerSubNumber	= ""
		Else
			gsGralLedgerSubNumber	= CLng(oRecordset("sub_numero_mayor"))
		End If
		gnPostingReferenceId = 0
		oRecordset.Close
		gnCurrency = 0		
		gsSubsidiaryAccountPrefix = oVoucherUS.GetGLSubsidiaryLedgerPrefix(Session("sAppServer"), CLng(gnGralLedgerId))
		gsPostingReference = "Ninguno"
		If (nPostingId = -1) Then			'Recuperar último movimiento
			nPostingId = oVoucherUS.GetLastPostingId(Session("sAppServer"), CLng(nTransactionId))
		End If
		If (nPostingId <> 0) Then
			Set oRecordset = oVoucherUS.GetPendingPostingRS(Session("sAppServer"), CLng(nPostingId))
			If Not (oRecordset Is Nothing) Then
				If oRecordset("tipo_movimiento") = "D" Then
					gbSelectDebit = "selected"
				Else
					gbSelectCredit = "selected"
				End If
				gnAccountId					 = oRecordset("id_cuenta")
				gsAccount						 = oRecordset("numero_cuenta_estandar")			
				gnCurrency					 = oRecordset("id_moneda")
				gsCboCurrencies      = oVoucherUS.CboAccountCurrencies(Session("sAppServer"), CLng(gnAccountId), CLng(gnCurrency))			
				gsSubsidiaryAccount  = oRecordset("numero_cuenta_auxiliar")
				gsSubsidiaryAccount  = Right(gsSubsidiaryAccount, 16)
				gsSubsidiaryAccountName = Left(oRecordset("nombre_cuenta_auxiliar"), 40)
				gsSector						 = oRecordset("clave_sector")
				gsResponsibilityArea = oRecordset("Clave_Area_Responsabilidad")
				gsBudgetKey					 = oRecordset("clave_presupuestal")
				gsDisponibilityKey   = oRecordset("clave_disponibilidad")
				gsVerificationNumber = oRecordset("numero_verificacion")		
				gsAmount						 = oVoucherUS.FormatCurrency(oRecordset("Monto"), 2)
				gsBaseAmount				 = oVoucherUS.FormatCurrency(oRecordset("Monto_Moneda_Base"), 6)
				gsExchangeRate			 = oVoucherUS.FormatCurrency(oRecordset("Tipo_Cambio"), 6)				
				gsDescription				 = oRecordset("concepto_movimiento")
				gnPostingReferenceId = oRecordset("id_movimiento_referencia")
				If CLng(gnPostingReferenceId) = -1 Then
					gsPostingReference = "Movimiento de iniciativa"
				ElseIf CLng(gnPostingReferenceId) = 0 Then
					gsPostingReference = "Ninguno"
				Else
					gsPostingReference = "Movimiento de conformidad"
				End If
				oRecordset.Close
			Else
				gsTitle = "El movimiento no existe"
			End If
		End If
		Set oRecordset = Nothing
		Set oVoucherUS = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			If (Len(Request.QueryString) <> 0) Then
				Session("sErrPage") = Request.ServerVariables("URL") & "?" & Request.QueryString
      Else
				Session("sErrPage") = Request.ServerVariables("URL")
		  End If
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Editor de movimientos</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var gnAccountId = <%=gnAccountId%>;
var gbIsAccountDirty = true;
var gbIsSectorDirty  = true;
var gbSended = false;

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsNumeric", sNumber, nDecimals);
	return obj.return_value;
}

function deletePosting() {
	var obj;	
	if (confirm('¿Elimino este movimiento de la póliza?')) {
		window.location.href = "exec/delete_posting.asp?id=<%=gnPostingId%>";
	}
	return false;	
}

function addSubsidiaryAccount(sSubsidiaryAccount, nAccountId, sSectorKey) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","AddSubsidiaryAccount", sSubsidiaryAccount, nAccountId, sSectorKey);
	return obj.return_value;
}

function assignSubsidiaryAccount(sSubsidiaryAccount, nAccountId, sSectorKey) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","AssignSubsidiaryAccount", sSubsidiaryAccount, nAccountId, sSectorKey);
	return obj.return_value;
}

function setAmountFields() {
	if (document.all.cboCurrencies.value == '<%=gsGLBaseCurrency%>') {
		document.all.txtExchangeRate.value = '1.000000';
		document.all.txtBaseAmount.value = document.all.txtAmount.value;
	} else {
		document.all.txtExchangeRate.value = '';
		document.all.txtBaseAmount.value = '';	
	}
	return true;	
}

function setSubsidiaryAccountName() {
	var obj, sSubsidiaryAccount;		

	if (document.all.txtSubsidiaryAccount.value != '') {
		sSubsidiaryAccount = document.all.txtSubsidiaryAccountPrefix.value + document.all.txtSubsidiaryAccount.value;
	  obj = RSExecute("../financial_accounting_scripts.asp", "SubsidiaryAccountName", <%=gnGralLedgerId%>, sSubsidiaryAccount);
	  if (obj.return_value != '') {
			document.all.divSubsidiaryAccountName.innerHTML = obj.return_value.substr(0, 36);
		} else {
			document.all.divSubsidiaryAccountName.innerHTML = "Auxiliar no registrado";
		}
	} else {
		
		document.all.divSubsidiaryAccountName.innerHTML = '';
	}	
}

function enableAmountFields() {	
	if (document.all.cboCurrencies.value == '<%=gsGLBaseCurrency%>') {		
		document.all.txtExchangeRate.disabled = true;
		document.all.txtExchangeRate.style.backgroundColor = 'beige';		
		document.all.txtBaseAmount.disabled = true;
		document.all.txtBaseAmount.style.backgroundColor = 'beige';		
		return true;
	} else {		
		document.all.txtExchangeRate.disabled = false;
		document.all.txtExchangeRate.style.backgroundColor = 'white';		
		document.all.txtBaseAmount.disabled = false;
		document.all.txtBaseAmount.style.backgroundColor = 'white';		
		return true;	
	}
}

function calculateBaseAmount(nExchangeRate, nAmount) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","CalculateBaseAmount", nExchangeRate, nAmount);
	return (obj.return_value);
}

function checkPostingValues() {
	var obj, sMsg, sSubsidiaryAccount;
	if (document.all.txtAccount.value == '') {
		alert('Requiero el número de cuenta.');
		if (!document.all.txtAccount.disabled) { document.all.txtAccount.focus(); }		
		return 0;
	}
	if (gnAccountId == 0 && document.all.txtAccount.value != '') {
		alert('No entiendo el número de cuenta proporcionado.');
		if (!document.all.txtAccount.disabled) { document.all.txtAccount.focus(); }
		return 0;	
	}
	if (document.all.txtSubsidiaryAccount.value != '') {
		sSubsidiaryAccount = document.all.txtSubsidiaryAccountPrefix.value + document.all.txtSubsidiaryAccount.value;
	} else {
		sSubsidiaryAccount = '';	
	}	
	obj = RSExecute("../financial_accounting_scripts.asp","CheckPostingValues", '<%=gsPostingDate%>', gnAccountId, document.all.txtSector.value, sSubsidiaryAccount, document.all.txtBudgetKey.value, document.all.txtResponsibilityArea.value);
	switch (obj.return_value) {
	 case -1:
		alert('La cuenta proporcionada no ha sido asignada al mayor.');
		if (!document.all.txtAccount.disabled) { document.all.txtAccount.focus(); }
		return 0;
	 case -2:
		alert('Requiero el número de sector.');
		if (!document.all.txtSector.disabled) { document.all.txtSector.focus(); }
		return 0;
	 case -3:
		alert('La cuenta no maneja sectores.');
		if (!document.all.txtSector.disabled) { document.all.txtSector.focus(); }		
		return 0;
	case -4:
		alert('La cuenta no tiene registrado el sector proporcionado.');
		if (!document.all.txtSector.disabled) { document.all.txtSector.focus(); }		
		return 0;
	case -5:
	  alert('Se requiere el auxiliar.');	  
	  if (!document.all.txtSubsidiaryAccount.disabled) { document.all.txtSubsidiaryAccount.focus(); }
	  return 0;
	case -6:
	  alert('La cuenta no maneja auxiliares.');
	  if (!document.all.txtSubsidiaryAccount.disabled) { document.all.txtSubsidiaryAccount.focus(); }
	  return 0;
	case -7:
		sMsg = 'El auxiliar proporcionado no existe.\n\n' +  
					 '¿Agrego el auxiliar ' + sSubsidiaryAccount + ' y lo asigno\n' + 
					 ' a la cuenta "' + document.all.txtAccount.value + '"?';
		if (confirm(sMsg)) {
			addSubsidiaryAccount(sSubsidiaryAccount, gnAccountId, document.all.txtSector.value);
			return (-1);
		} else {			 
	    if (!document.all.txtSubsidiaryAccount.disabled) { document.all.txtSubsidiaryAccount.focus(); }	     
	    return 0;
	  }
	case -8:
		sMsg = 'El auxiliar proporcionado no ha sido asignado a la cuenta "' + document.all.txtAccount.value + '".\n\n' +  
					 '¿Asigno el auxiliar ' + sSubsidiaryAccount + ' a la cuenta "' + document.all.txtAccount.value + '"?';
		if (confirm(sMsg)) {
			assignSubsidiaryAccount(sSubsidiaryAccount, gnAccountId, document.all.txtSector.value);
			return (-1);
		} else {			 
	    if (!document.all.txtSubsidiaryAccount.disabled) { document.all.txtSubsidiaryAccount.focus(); }	     
	    return 0;
	  }		
	case -9:
	  alert('La cuenta proporcionada no maneja área de responsabilidad.');
	  if (!document.all.txtResponsibilityArea.disabled) { document.all.txtResponsibilityArea.focus(); }
	  return 0;
	case -10:
	  alert('Se requiere el área de responsabilidad.');
	  if (!document.all.txtResponsibilityArea.disabled) { document.all.txtResponsibilityArea.focus(); }
	  return 0;	  
	case -11:
	  alert('La cuenta no maneja el área de responsabilidad proporcionada.');
	  if (!document.all.txtResponsibilityArea.disabled) { document.all.txtResponsibilityArea.focus(); }
	  return 0;
	case -12:
	  alert('No reconozco el área de responsabilidad proporcionada.');
	  if (!document.all.txtResponsibilityArea.disabled) { document.all.txtResponsibilityArea.focus(); }
	  return 0;	  
	}	
  return 1;
}

function setAccountId(sAccount) {
	var obj, sStdAccountRole, sMsg;
		
	obj = RSExecute("../financial_accounting_scripts.asp", "GLAccountRole" , <%=gnGralLedgerId%>, sAccount);
	sStdAccountRole = obj.return_value;

	if (sStdAccountRole == '') {
		alert("La cuenta proporcionada no existe en el catálogo de cuentas estándar.");
		gnAccountId = 0;
		gbIsAccountDirty = true;
		return false;
	}
		
	if (sStdAccountRole == 'S') {
		alert("La cuenta proporcionada no acepta movimientos debido a que es sumaria.");
		gnAccountId = 0;
		gbIsAccountDirty = true;
		return false;			
	}

	obj = RSExecute("../financial_accounting_scripts.asp","GetAccountId", <%=gnGralLedgerId%>, sAccount);
	gnAccountId = obj.return_value;
		
	if (gnAccountId > 0) {
		gbIsAccountDirty = false;
		return true;
	}
	gnAccountId = Math.abs(gnAccountId);	
	if (gnAccountId == 0) {
		alert("La cuenta proporcionada no existe en el catálogo de cuentas estándar.");
		gnAccountId = 0;
		gbIsAccountDirty = true;
		return false;
	}
	
	sMsg  = "La cuenta proporcionada no está registrada en esta contabilidad,\n"
	sMsg +=	"pero existe en el catálogo de cuentas estándar.\n\n"
	sMsg += "¿Anexo la cuenta " + sAccount + " a la contabilidad?"
	if (confirm(sMsg)) {
		obj = RSExecute("../financial_accounting_scripts.asp", "AssignStdAccount", <%=gnGralLedgerId%>, gnAccountId);
		obj = RSExecute("../financial_accounting_scripts.asp","GetAccountId", <%=gnGralLedgerId%>, sAccount);
		gnAccountId = obj.return_value;
		if (gnAccountId > 0) {
			gbIsAccountDirty = false;
			return true;
		} else {
		  alert("Ocurrió un problema, por lo que no puede agregar la cuenta a la contabilidad.");
			gnAccountId = 0;
			gbIsAccountDirty = true;			
			return false;
		}
	}
} 

function updateCombo(sName) {
	var obj;
	
	if (sName == 'cboCurrencies') {
		<% If (gnPostingId = 0) Then %>
		obj = RSExecute("../financial_accounting_scripts.asp", "CboAccountCurrencies", gnAccountId, <%=gsGLBaseCurrency%>);
		<% Else %>
		obj = RSExecute("../financial_accounting_scripts.asp", "CboAccountCurrencies", gnAccountId, <%=gnCurrency%>);
		<% End If %>
		document.all.divCboCurrencies.innerHTML = obj.return_value;
		return true;
	}
}

function checkAmounts() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CheckAmounts", document.all.txtAmount.value, document.all.txtExchangeRate.value, document.all.txtBaseAmount.value);
	if (obj.return_value == false) {
		alert("El monto en la moneda base no coincide con el tipo de cambio y el monto proporcionados.");
		return false;
	}
	return true;	
}


function doSubmit() {
  var nResult;
	if (gbSended) {
		return false;
	}
/*
	if (document.all.txtPostingDate.value != '') {
		if (!isDate(document.all.txtPostingDate.value)) {
			alert("No reconozco la fecha de movimiento proporcionada.");
			document.all.txtPostingDate.focus();
			return false;
	  }
	}	
*/

	if (document.all.cboCurrencies.value == '') {
		alert("La cuenta proporcionada no tiene asignada ninguna moneda."); 
		return false;
	}
	controlAndSetAmounts();
	
	if (document.all.txtAmount.value == '') {
		alert("Necesito el monto del movimiento.");
		if (!document.all.txtAmount.disabled) { document.all.txtAmount.focus(); }
		return false;	
	}
	if (!isNumeric(document.all.txtAmount.value, 2)) {
		alert("No reconozo el monto del movimiento.");
		if (!document.all.txtAmount.disabled) { document.all.txtAmount.focus(); }		
		return false;
	}
	if (document.all.txtExchangeRate.value == '') {
		alert("Necesito el tipo de cambio.");
		if (!document.all.txtExchangeRate.disabled) { document.all.txtExchangeRate.focus(); }
		return false;
	}
	if (!isNumeric(document.all.txtExchangeRate.value, 6)) {
		alert("No reconozco el tipo de cambio proporcionado.");
		if (!document.all.txtExchangeRate.disabled) { document.all.txtExchangeRate.focus(); }
		return false;
	}
	if (document.all.txtBaseAmount.value == '') {
		alert("Necesito el monto en la moneda base.");
		if (!document.all.txtBaseAmount.disabled) { document.all.txtBaseAmount.focus(); }
		return false;
	}	
	if (!isNumeric(document.all.txtBaseAmount.value, 6)) {
		alert("No reconozo el monto en la moneda base.");
		if (!document.all.txtBaseAmount.disabled) { document.all.txtBaseAmount.focus(); }		
		return false;
	}
	if (!checkAmounts()) {
		return false;
	}
	while (true) {
		nResult = checkPostingValues();
	  if (nResult == 0) {			
	     return false;
	  }
	  if (nResult == 1) {
			gbSended = true;
			document.all.frmPostingEditor.submit();			
			return true;
	  }
	}
	gbSended = false;	
	return false;
}

function subsidiaryAccountNumber(nSubsidiaryAccountId, bShowComplete) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "SubsidiaryAccountNumber", nSubsidiaryAccountId, bShowComplete);	
	return obj.return_value;
}

function oDateConcept() {
	var postingDate;
	var description;
}

function oPostingRef() {
	var gralLedgerId;
	var postingId;
}

function doSearch() {
    myDialog.str = "";
    myDialog.caseSensitive = false;
    if (showModalDialog("search.htm", myDialog)==false)
        return;    // user canceled search
    else {
        // search for the string
    }
}

function displayPicker(sPickerName, oTarget) {
	var sURL = "/empiria/financial_accounting", sPars = "resizable:0;status:0;";
	var retValue;
	
	switch (sPickerName) {
		case 'account':
			sURL  += "/subs_accounts/account_picker.asp";
			sPars += "dialogHeight:320px;dialogWidth:320px;";
			alert("Por el momento esta operación está en construcción. Gracias.");						
			return false;
		case 'dateConcept':
			sURL  += "/transactions/date_concept_picker.htm";
			sPars  = "dialogHeight:200px;dialogWidth:330px;resizable:no;scroll:no;status:no;help:no;";
			oDateConcept.postingDate = document.all.txtPostingDate.value;
			oDateConcept.description = document.all.txtDescription.value;
			if (window.showModalDialog(sURL, oDateConcept, sPars)) {
				document.all.txtPostingDate.value = oDateConcept.postingDate;
				document.all.txtDescription.value =oDateConcept.description;
			};
			return false;
		case 'sector':
			sURL  += "/general/sector_picker.asp";
			sPars += "dialogHeight:320px;dialogWidth:320px;";
			alert("Por el momento esta operación está en construcción. Gracias.");
			return false;		
		case 'subsidiaryAccount':
			if (document.all.txtSubsidiaryAccount.disabled) {
				alert("No puedo ejecutar la opración solicitada, debido a que la cuenta proporcionada no maneja auxiliares.");
				return false;
			}
			sURL  += "/subs_accounts/subsidiary_account_picker.asp?gralLedgerId=<%=gnGralLedgerId%>"
			if ((document.all.txtSubsidiaryAccountPrefix.value + document.all.txtSubsidiaryAccount.value) != '') {
				sURL +="&subsAccount=" + document.all.txtSubsidiaryAccountPrefix.value + document.all.txtSubsidiaryAccount.value;
			}
			sPars += "dialogHeight:420px;dialogWidth:400px;";			
			retValue = window.showModalDialog(sURL, "" , sPars);
			if (retValue != 0) {
				document.all.txtSubsidiaryAccount.value = subsidiaryAccountNumber(retValue, false);
				document.all.txtSubsidiaryAccountPrefix.value = document.all.txtSubsidiaryAccountP.value;
				setSubsidiaryAccountName();
			}
			return false;
		case 'postingReference':
			sURL  = "/transactions/posting_reference_picker.asp";
			sPars = "dialogHeight:330px;dialogWidth:510px;resizable:no;scroll:no;status:no;help:no;";			
			oPostingRef.gralLedgerId  = <%=gnGralLedgerId%>;
			oPostingRef.postingId     = document.all.txtPostingReferenceId.value;
			if (window.showModalDialog(sURL, oPostingRef, sPars)) {
				document.all.txtPostingReferenceId.value = oPostingRef.postingId;
				setPostingReferenceDesc();
			};
			return false;
		case 'responsibilityArea':
			sURL  += "responsibility_area_picker.asp";
			sPars += "dialogHeight:320px;dialogWidth:420px;";		
			alert("Por el momento esta operación está en construcción. Gracias.");
			return false;
		case 'exchangeRates':
			sURL  += "exchange_rate_picker.asp";
			sPars += "dialogHeight:320px;dialogWidth:320px;";
			alert("Por el momento esta operación está en construcción. Gracias.");
			return false;
	}	
	//oTarget.value = window.showModalDialog(sURL, "" , sPars);
	return false;	
}

function setPostingReferenceDesc() {
	switch (Number(document.all.txtPostingReferenceId.value)) {
		case -1:
			document.all.txtPostingReference.value = "Movimiento de iniciativa"
			break;
		case 0:
			document.all.txtPostingReference.value = "Ninguno"
			break;
		default:
			document.all.txtPostingReference.value = "Movimiento de conformidad"
			break;
	}
}

function computeBaseAmount() {	
	var obj;
	
	if (gbIsAccountDirty == true) {
		alert("No reconozco la moneda proporcionada.");
		return false;
	}	
	if (document.all.txtAmount.value != '') {
		if(!isNumeric(document.all.txtAmount.value, 2)) {
			alert("No reconozco el monto de la operación.");
			if (!document.all.txtAmount.disabled) { document.all.txtAmount.focus(); }
			return false;
		}
  } else {
		alert("Para efectuar el cálculo requiero el monto de la operación.");
		if (!document.all.txtAmount.disabled) { document.all.txtAmount.focus(); }		
		return false;
	}
	if (document.all.txtExchangeRate.value != '') {
		if (document.all.cboCurrencies.value == '<%=gsGLBaseCurrency%>') {		
			document.all.txtExchangeRate.value = '1.000000';
		}
		if(!isNumeric(document.all.txtExchangeRate.value, 6)) {
			alert("No reconozco el tipo de cambio proporcionado.");
			if (!document.all.txtExchangeRate.disabled) { document.all.txtExchangeRate.focus(); }			
			return false;			
		}
	} else {
		if (document.all.cboCurrencies.value != '<%=gsGLBaseCurrency%>') {
		  obj = RSExecute("../financial_accounting_scripts.asp", "ExchangeRate" , <%=gsGLBaseCurrency%>, document.all.cboCurrencies.value, <%=gsPostingDate%>);
			document.all.txtExchangeRate.value = obj.return_value;
		} else {
			document.all.txtExchangeRate.value = '1.000000';
		}
	}	
	document.all.txtBaseAmount.value = calculateBaseAmount(document.all.txtExchangeRate.value, document.all.txtAmount.value);
	return false;
}

function deshabilitarCamposBancaria()	{
	document.all.txtSector.disabled = true;
	document.all.txtSector.style.backgroundColor = 'beige';
	document.all.txtResponsibilityArea.disabled = true;
	document.all.txtResponsibilityArea.style.backgroundColor = 'beige';		
	document.all.txtBudgetKey.disabled = true;
	document.all.txtBudgetKey.style.backgroundColor = 'beige';
	document.all.txtResponsibilityArea.disabled = true;
	document.all.txtResponsibilityArea.style.backgroundColor = 'beige';	
	document.all.txtDisponibilityKey.disabled = true;
	document.all.txtDisponibilityKey.style.backgroundColor = 'beige';
	document.all.txtVerificationNumber.disabled = true;
	document.all.txtVerificationNumber.style.backgroundColor = 'beige';			
	return true;	
}

function controlAndSetAmounts() {	
	var obj;
	
	if (document.all.txtAmount.value != '' && isNumeric(document.all.txtAmount.value, 2)) {
		obj = RSExecute("../financial_accounting_scripts.asp","FormatCurrency", document.all.txtAmount.value, 2);
		document.all.txtAmount.value = obj.return_value;
	} else {		
	  return true;
	}
	if (gbIsAccountDirty) {
	  return true;
	}	
	if (document.all.cboCurrencies.value == '<%=gsGLBaseCurrency%>') {
		document.all.txtExchangeRate.value = '1.000000';
		document.all.txtBaseAmount.value = document.all.txtAmount.value;
	} else {
	 if (document.all.txtExchangeRate.value == '') {
	  obj = RSExecute("../financial_accounting_scripts.asp", "CurrentExchangeRate" , <%=gsGLBaseCurrency%>, document.all.cboCurrencies.value);
		document.all.txtExchangeRate.value = obj.return_value;
		document.all.txtBaseAmount.value = calculateBaseAmount(document.all.txtExchangeRate.value, document.all.txtAmount.value);
	 } else {
	    if (document.all.txtExchangeRate.value != '' && isNumeric(document.all.txtExchangeRate.value, 6)) {
			  document.all.txtBaseAmount.value = calculateBaseAmount(document.all.txtExchangeRate.value, document.all.txtAmount.value);
			  if (!document.all.txtBaseAmount.disabled) { document.all.txtBaseAmount.focus(); }			  
		  }
	 }
	}
}

function txtAmount_onblur() {
	controlAndSetAmounts();
}

function txtBaseAmount_onblur() {
	var obj;	
	if (document.all.txtBaseAmount.value != '' && isNumeric(document.all.txtBaseAmount.value, 6)) {
		obj = RSExecute("../financial_accounting_scripts.asp","FormatCurrency", document.all.txtBaseAmount.value, 6);
		document.all.txtBaseAmount.value = obj.return_value;
	}
}

function txtAccount_onblur() {
	var obj;
	if (!gbIsAccountDirty) {
		return true;
	}
	gnAccountId = 0;
	if (document.all.txtAccount.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountWithGLId", <%=gnGralLedgerId%>, document.all.txtAccount.value);
		if (obj.return_value != '') {
			document.all.txtAccount.value = obj.return_value;
			setAccountId(document.all.txtAccount.value);
			updateCombo('cboCurrencies');
			if (gnAccountId != 0) {
				setAmountFields();
				enableAmountFields();				
			  enableTextboxes();
			  gbIsAccountDirty = false;  
			}
		} else {
			alert("No entiendo el formato de la cuenta proporcionada.");
			updateCombo('cboCurrencies');
			gbIsAccountDirty = true;			
		}
	}
	setSubsidiaryAccountName();
	return true;
}

function txtSubsidiaryAccount_onblur() {
	var obj;

	document.all.txtSubsidiaryAccountPrefix.value = document.all.txtSubsidiaryAccountP.value;
	if (document.all.txtSubsidiaryAccount.value != '') {
	  obj = RSExecute("../financial_accounting_scripts.asp", "FormatSubsidiaryAccount", "<%=gsGralLedgerSubNumber%>", document.all.txtSubsidiaryAccount.value);
		if (obj.return_value != '') {
			document.all.txtSubsidiaryAccount.value = obj.return_value;
		} else {
			alert("No entiendo el formato del auxiliar proporcionado.");
			setSubsidiaryAccountName();
			return true;
		}	
	}
	setSubsidiaryAccountName();
	return true;
}

function txtExchangeRate_onblur() {
	var obj;	
	if (document.all.txtExchangeRate.value != '' && isNumeric(document.all.txtExchangeRate.value, 6)) {
		obj = RSExecute("../financial_accounting_scripts.asp","FormatCurrency", document.all.txtExchangeRate.value, 6);
		document.all.txtExchangeRate.value = obj.return_value;
		if (document.all.txtAmount.value != '' && isNumeric(document.all.txtAmount.value, 6)) {
			document.all.txtBaseAmount.value = calculateBaseAmount(document.all.txtExchangeRate.value, document.all.txtAmount.value);
		}
	}
}

function cboCurrencies_onchange() {
	setAmountFields();
	enableAmountFields();
}

function txtAccount_onchange() {
	gbIsAccountDirty = true;
}

function txtSector_onblur() {
	if (gnAccountId != 0 && gbIsSectorDirty & document.all.txtSector.value != '') {
		enableTextboxes();
	}
}

function txtSector_onchange() {
  gbIsSectorDirty = true;
}
			  
function tagGetLastPosting_onclick() {
	<% If (gnTransactionPostingsCount <> 0) Then %>
		window.navigate("<%=gsRefreshURL%>" + "&getLast=true");
	<% Else %>
		alert("La póliza no tiene movimientos.")
	<% End If %>		
  return false;
}

function tagClonePosting_onclick() {
	window.navigate("<%=gsRefreshURL%>" + "&clone=true");
	return false; 	
}

function enableTextboxes() {
	var obj, status = 0, nSectorId = 0;
	
	if (document.all.txtSector.value != '' && gbIsSectorDirty) {
		obj = RSExecute("../financial_accounting_scripts.asp","GetSectorId", document.all.txtSector.value);		
		nSectorId = obj.return_value;
		if (nSectorId == 0) {
		   alert("La cuenta no maneja el sector proporcionado");
		   gbIsSectorDirty = true;
		   return false
		}
	}	
	if (gnAccountId != 0) {
	  obj = RSExecute("../financial_accounting_scripts.asp","GetAccountProps", '<%=gsPostingDate%>', gnAccountId, nSectorId);
	  status = obj.return_value;
	} else {
		return true;
	}
	if (status == -1) {
		alert("La cuenta no existe en el catálogo estándar.");
		return false;
	}
	if (status == -2) {
		alert("La cuenta proporcionada es sumaria, por lo que no acepta movimientos.");
		return false;
	}	
	if (status == -3) {
		alert("El sector proporcionado no está asignado a la cuenta.");
		gbIsSectorDirty = true;
		return false;
	}
	gbIsSectorDirty = false;
	if (status == 0) {
		document.all.txtSector.disabled = true;
		document.all.txtSector.style.backgroundColor = 'beige';
		document.all.txtSector.value = '';
		document.all.txtSubsidiaryAccount.disabled = true;
		document.all.txtSubsidiaryAccount.style.backgroundColor = 'beige';
		document.all.txtSubsidiaryAccount.value = '';
		document.all.txtSubsidiaryAccountPrefix.value = '';
		//document.all.txtResponsibilityArea.disabled = true;		
		//document.all.txtResponsibilityArea.style.backgroundColor = 'beige';
		//document.all.txtResponsibilityArea.value = '';		
		if (!document.all.txtResponsibilityArea.disabled) { document.all.txtResponsibilityArea.focus(); }
		return true;
	}
	if (status == 1) {
		document.all.txtSector.disabled = false;
		document.all.txtSector.style.backgroundColor = 'white';
		document.all.txtSubsidiaryAccount.disabled = true;
		document.all.txtSubsidiaryAccount.style.backgroundColor = 'beige';
		document.all.txtSubsidiaryAccount.value = '';
		document.all.txtSubsidiaryAccountPrefix.value = '';
		if (!document.all.txtSector.disabled && nSectorId == 0) { document.all.txtSector.focus(); }		
		return true;
	}	
	if (status == 2) {						
		document.all.txtSector.disabled = true;
		document.all.txtSector.style.backgroundColor = 'beige';
		document.all.txtSector.value = '';
		document.all.txtSubsidiaryAccount.disabled = false;
		document.all.txtSubsidiaryAccount.style.backgroundColor = 'white';
		document.all.txtSubsidiaryAccountPrefix.value = document.all.txtSubsidiaryAccountP.value;
		if (!document.all.txtSubsidiaryAccount.disabled) { document.all.txtSubsidiaryAccount.focus(); }		
		return true;
	}	
	if (status == 3) {
		document.all.txtSector.disabled = false;
		document.all.txtSector.style.backgroundColor = 'white';
		document.all.txtSubsidiaryAccount.disabled = false;
		document.all.txtSubsidiaryAccount.style.backgroundColor = 'white';
		document.all.txtSubsidiaryAccountPrefix.value = document.all.txtSubsidiaryAccountP.value;
		if (!document.all.txtSector.disabled && nSectorId == 0) { document.all.txtSector.focus(); }
		return true;
	}	
}

function refreshVoucher() {
	window.opener.document.all.ancRefreshPost.click();
	return false;
}

function window_onload() {
	document.all.cboPostingType.focus();
	document.all.txtSubsidiaryAccountPrefix.disabled = true;
	document.all.txtSubsidiaryAccountPrefix.value = document.all.txtSubsidiaryAccountP.value;
	document.all.txtSubsidiaryAccountPrefix.style.backgroundColor = 'beige';
	document.all.txtPostingReference.disabled = true;
	document.all.txtPostingReference.style.backgroundColor = 'beige';		
	<% If (gnGralLedgerId > 32) Then %>
		deshabilitarCamposBancaria();
	<% End If %>
	<% If (gnAccountId <> "") Then %>
		gbIsAccountDirty = false;
		gbIsSectorDirty = false;
		gnAccountId	= <%=gnAccountId%>;
		enableTextboxes();
		enableAmountFields();	
	<% End If %>
	
	<% If (gnPostingId <> 0) And (gnAccountId = "") Then %>
		alert("El movimiento que se desea editar fue eliminado.")
		window.close();
	<% End If %> 	
}


//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			<%=gsTitle%>
		</TD>
		<TD colspan=3 align=right nowrap>
			<A href='' onclick="return(refreshVoucher());" tabindex=-1>Refrescar póliza</A>
			<img align=absmiddle src='/empiria/images/invisible.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">								</TD>
	</TR>
  <TR>
		<TD colspan=4 nowrap>
			<FORM name=frmPostingEditor action="<%=gsSendURL%>" method=post>
				<TABLE class=applicationTable cellpadding=1>
				<TR>
					<TD colspan=4 align=right>
						<A href='' onclick="return(displayPicker('dateConcept', this));" tabindex=-1>Fecha y concepto</A>
						&nbsp; | &nbsp;
						<% If (gnPostingId = 0) Then %>
						<A href='' onclick="return(tagGetLastPosting_onclick());" tabindex=-1>Copiar datos del último movimiento</A>
						<% Else %>
						<A href='' onclick="return(tagClonePosting_onclick());" tabindex=-1>Clonar este movimiento</A>
						<% End If %>
						&nbsp; &nbsp; 
					</TD>		  
				</TR>
				<TR>
				  <TD>Movimiento:</TD>
				  <TD colspan=3>
						<SELECT name=cboPostingType style="width:210">
							<OPTION <%=gbSelectDebit%> value='D'>Cargo</OPTION>
							<OPTION <%=gbSelectCredit%> value='H'>Abono</OPTION>
						</SELECT>
					</TD>
				</TR>
				<TR>
				  <TD><a href=''onclick="return(displayPicker('account', this));" tabindex=-1>Cuenta:</a></TD>
				  <TD colspan=3>
						<INPUT name=txtAccount value="<%=gsAccount%>" style="width:140" onblur="return txtAccount_onblur()" onchange="return txtAccount_onchange()">
						&nbsp;<a href=''onclick="displayPicker('sector', this);return false;" tabindex=-1>Sector:</a>
						<INPUT name=txtSector value="<%=gsSector%>" maxlength=2 style="width:25" onblur="return txtSector_onblur()" onchange="return txtSector_onchange()">
					</TD>
				</TR>
				<TR>
				  <TD><a href='' onclick="return(displayPicker('subsidiaryAccount', this));" tabindex=-1>Número de auxiliar:</a></TD>
				  <TD colSpan=3>
						<INPUT name=txtSubsidiaryAccountPrefix value="" style="width:40px">
						<INPUT name=txtSubsidiaryAccount maxlength=16 value="<%=gsSubsidiaryAccount%>" style="width:136px" onblur="return txtSubsidiaryAccount_onblur()">
					</TD>
				</TR>
				<TR>
				  <TD>&nbsp;</TD>  
				  <TD colspan=3><span id=divSubsidiaryAccountName><%=gsSubsidiaryAccountName%></span></TD>
				</TR>
				<TR>
				<TR>
				  <TD><a href=''onclick="return(displayPicker('postingReference', this));" tabindex=-1>Mov. de referencia:</a></TD>
				  <TD colSpan=3>
						<INPUT name=txtPostingReference readonly tabindex=-1 maxlength=20 value="<%=gsPostingReference%>" style="width:210">
					</TD>
				</TR>
				<TR>
				  <TD nowrap>
						<a href=''onclick="return(displayPicker('responsibilityArea', document.all.txtResponsibilityArea.value));" tabindex=-1>Area de responsabilidad:</a>
					</TD>
				  <TD colspan=3>
						<INPUT name=txtResponsibilityArea maxlength=6 value="<%=gsResponsibilityArea%>" style="width:210">
					</TD>
				</TR>
				<TR>
				  <TD>Concepto presupuestal:</TD>
				  <TD colSpan=3>
						<INPUT name=txtBudgetKey maxlength=6 value="<%=gsBudgetKey%>" style="width:210">
					</TD>
				</TR>
				<TR>
				  <TD>Clave de disponibilidad:</TD>
				  <TD colSpan=3>
						<INPUT name=txtDisponibilityKey maxlength=1 value="<%=gsDisponibilityKey%>" style="width:210">
					</TD>
				</TR>
				<TR>
				  <TD>Número de verificación:</TD>
				  <TD colSpan=3>
						<INPUT name=txtVerificationNumber maxlength=6 value="<%=gsVerificationNumber%>" style="width:210">
					</TD>
				</TR>
				<TR>
				  <TD>Moneda:</TD>
				  <TD colspan=3>
						<div id=divCboCurrencies>
							<SELECT name=cboCurrencies style="width:210" LANGUAGE=javascript onchange="return cboCurrencies_onchange()">
							<%=gsCboCurrencies%>
							</SELECT>
						</div>
				  </TD>
				</TR>
				<TR>
				  <TD>Monto de la operación:</TD>
				  <TD colspan=3>
						<INPUT name=txtAmount value="<%=gsAmount%>" style="width:210" onblur="return txtAmount_onblur()">
					</TD>
				</TR>
				<TR>
				  <TD>
						<a href=''onclick="return(displayPicker('exchangeRates', document.all.txtExchangeRate.value));" tabindex=-1>Tipo de cambio:</a>
					</TD>
				  <TD colspan=3>
						<INPUT name=txtExchangeRate value="<%=gsExchangeRate%>" style="width:210" LANGUAGE=javascript onblur="return txtExchangeRate_onblur()">
					</TD>  
				</TR>
				<TR>
				  <TD>
						Monto en moneda base:<br><A href=''onclick="computeBaseAmount();return false;" tabindex=-1>Calcular monto</A>
				  </TD>
				  <TD colSpan=3>
						<INPUT name=txtBaseAmount value="<%=gsBaseAmount%>" style="width:210" LANGUAGE=javascript onblur="return txtBaseAmount_onblur()">
					</TD>
				</TR>
				<TR>
				  <TD>&nbsp;</TD>
				  <TD colspan=2 align=right>
						<INPUT type=hidden name=txtPostingReferenceId value="<%=gnPostingReferenceId%>">
						<INPUT type=hidden name=txtSubsidiaryAccountP value="<%=gsSubsidiaryAccountPrefix%>">
						<INPUT type=hidden name=txtPostingDate value="<%=gsPostingDate%>" style="width:210">
						<INPUT type=hidden name=txtDescription style="width:210" value="<%=gsDescription%>">
				    <% If (gnPostingId = 0) Then %>
						<INPUT class=cmdSubmit name=cmdSavePosting type=button value="Agregar" style="width:60" onclick='doSubmit();'>&nbsp;&nbsp;&nbsp;
						<% Else %>
						<INPUT class=cmdSubmit name=cmdSavePosting type=button value="Guardar" style="width:60" onclick='doSubmit();'>&nbsp;&nbsp;
						<INPUT class=cmdSubmit name=cmdDelPosting  type=button value="Eliminar" style="width:60" onclick="deletePosting();">&nbsp;&nbsp;
						<% End If %>		
						<INPUT class=cmdSubmit name=cmdCancelPosting type=button value="Cerrar" style="width:60" onclick="window.close();">
						&nbsp; &nbsp; &nbsp;
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</TD>
</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
