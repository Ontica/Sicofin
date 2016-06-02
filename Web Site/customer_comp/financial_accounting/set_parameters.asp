<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1	

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsReportId, gsReportTitle, gsReportsTable, gsCboGLGroups, gsCboCurrencies, gsCboExchangeRateTypes
	Dim bSelectGralLedger, bUniqueGralLedger, bUseInitialDate, bUseFinalDate, bUseInitialPeriod, bUseFinalPeriod, bUseSigners, bAccount, bRangeAccount, bRangeSubsAccount, bGLRange 
	Dim bUseAccountPattern, bUseExchangeRate, bUseParticipantType, bUseParticipantStatus, bRoundValues, bSelectAccounts, bChkAfectationDate, bChkUpdatedVouchers, bChkOptionToDisplay
	Dim bTittle101, bParticipantOrder, bAccountOrder, bStdAccountType, bTotalByGroups, bConFidOptions
	
	gsReportId = Request.QueryString("id")
	
	Call SetParameters()
	
	Sub SetParameters()
		bUniqueGralLedger = False		
	  Select Case gsReportId				
			Case 1
				bSelectGralLedger = True
				bUniqueGralLedger = True
				bUseInitialDate = True				
				bUseFinalDate = True
	      bUseSigners = True
				bUseExchangeRate = True
				gsReportTitle = "Cuaderno de delegaciones"
			Case 3			
				bSelectGralLedger = True
				bUseAccountPattern = False				
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False				
				bChkAfectationDate = True
				bChkUpdatedVouchers = True				
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Pólizas actualizadas o pendientes de actualizar"											
			Case 9
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True			
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Actualización de Movimientos"
			Case 18
				bSelectGralLedger = False
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Balanza de consolidación de sucursales"
			Case 51			
				bSelectGralLedger = True
				bUseAccountPattern = True
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = True
				bSelectAccounts = False
				gsReportTitle = "Resumen estado analítico de cuentas (SIF)"
      Case 53
				bSelectGralLedger = True
				bUseAccountPattern = True
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = True
				bSelectAccounts = False
				gsReportTitle = "Comparación de grupos (SIF)"
			Case 55
				bSelectGralLedger = True
				bUseAccountPattern = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = True
				bSelectAccounts = False
				gsReportTitle = "Analítico de cuentas"
			Case 58
				bSelectGralLedger = True
				bUseAccountPattern = True				
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Analítico de cuentas (saldos promedio)"																			
			Case 59
				bSelectGralLedger = True
				bUniqueGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "De cifras (SIFRRBCO)"
			Case 60
				bSelectGralLedger = False
				bUseAccountPattern = False				
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Resumen mensual de cuentas 6205 y 6206"
			Case 61
				bSelectGralLedger = true
				bUniqueGralLedger = True
				bUseAccountPattern = false				
				bUseInitialDate = false
				bUseFinalDate	= True	
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Cuentas acreedoras y deudoras"						
			Case 62
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True			
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = True
				gsReportTitle = "Movimientos de cuentas seleccionadas"
			Case 64
				bSelectGralLedger = True
				bGLRange = False
				bRangeAccount = False
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Mayor General"				
			Case 65
				bSelectGralLedger = True
				bGLRange = True
				bUseAccountPattern = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bTotalByGroups = True
				bSelectAccounts = False
				gsReportTitle = "Sumas y promedios de saldos diarios"
			Case 68
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Mensual consolidado de cuentas de cheques"
			Case 69
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Relación mensual de honorarios fiduciarios"
				
			Case 70
				bSelectGralLedger = False
				bUseAccountPattern = False				
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Disponibilidades financieras de fideicomisos"	
			Case 73
				bSelectGralLedger = True
				bUseAccountPattern = False				
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Saldos sobregirados de fideicomisos"	
				
			Case 74
				bSelectGralLedger = True
				bUseAccountPattern = false			
				bUseInitialDate = False
				bUseFinalDate	= False	
				bUseInitialPeriod = True
				bUseFinalPeriod	= True
				bUseExchangeRate = True
				bRoundValues = false
				bSelectAccounts = False
				gsReportTitle = "Comparativo Mensual de Fideicomisos"
	    Case 96
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Auxiliares cuenta 1103"
	    Case 99
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Valorización de monedas extranjeras consolidadas"
			Case 101
				bSelectGralLedger = False
				bUseAccountPattern = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = True
				bTittle101 = True
				gsReportTitle = "Análisis de saldos diarios promedio"
			Case 102
				bSelectGralLedger = True
				bUseAccountPattern = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = True
				gsReportTitle = "Resumen de saldos diarios promedio"
			Case 106
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliares de cartera"				
			Case 107
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliares administrativos"

			Case 108
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliar de activos"
			Case 109
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliares"
			Case 110
				bSelectGralLedger = True
				bGLRange = True
				bRangeAccount = True
				bUseAccountPattern = True
				bRangeSubsAccount = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Auxiliar de movimientos"
			Case 111
				bSelectGralLedger = True
				bGLRange = True
				bRangeAccount = True
				bRangeSubsAccount = True
				bUseAccountPattern = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Movimientos de Auxiliares"
			Case 113
				bSelectGralLedger = True
				bUseAccountPattern = False
				bRangeAccount = True
				bRangeSubsAccount = True
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Auxiliares por cuenta"
	    Case 115
				bSelectGralLedger = False
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= True	
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				bUseParticipantType = True
				bUseParticipantStatus = True
				bParticipantOrder = True
				gsReportTitle = "Areas operativas"	
	    Case 116	      
				bSelectGralLedger = false
				bStdAccountType = True
				bUseAccountPattern = false				
				bUseInitialDate = false
				bUseFinalDate	= True	
				bUseExchangeRate = false
				bRoundValues = false
				bSelectAccounts = False
				gsReportTitle = "Catálogo de cuentas"
	    Case 117
				bSelectGralLedger = false
				bUseAccountPattern = false				
				bUseInitialDate = false
				bUseFinalDate	= false	
				bUseExchangeRate = false
				bRoundValues = false
				bSelectAccounts = False
				bAccountOrder = True
				gsReportTitle = "Catálogo de fideicomisos"			
			Case 120
				bSelectGralLedger = False
				bUniqueGralLedger = False
				bConFidOptions = True
				bUseInitialDate = False				
				bUseFinalDate = True
			  bUseSigners = False
			  bUseExchangeRate = False
				gsReportTitle = "Agrupación de cuentas de la póliza de concentración fiduciaria"				
	    Case 126
				bSelectGralLedger = True
				bUseAccountPattern = False				
				bUseInitialDate = False
				bUseFinalDate	= True	
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Declaración anual por retención de ISR"
	    Case 127
				bSelectGralLedger = True
				bUseAccountPattern = False				
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				bChkOptionToDisplay = True
				gsReportTitle = "Determinación de la base del IVA"				
	    Case 128
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Cálculo del resultado por posición monetaria"	
			Case 144
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliares de fideicomisos"				
			Case 149
				bSelectGralLedger = True
				bGLRange = True
				bAccount = True
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= False
				bUseInitialPeriod = True
				bUseFinalPeriod	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Comparativo de movimientos netos de auxiliares de fideicomisos por cuenta y auxiliar"
			Case 150
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= False
				bUseInitialPeriod = True
				bUseFinalPeriod	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Comparativo de movimientos netos de auxiliares de fideicomisos"
			Case 151
				bSelectGralLedger = True
				bAccount = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Comparativo de cuentas de compra-venta a nivel auxiliar"

			Case 152
				bSelectGralLedger = True
				bAccount = False
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = True
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Análisis de saldos para verificar la valorización mensual de divisas de los fid. y mos."

			Case 153
				bSelectGralLedger = True
				bAccount = False
				bUseAccountPattern = False
				bUseInitialDate = False
				bUseFinalDate	= False
				bUseInitialPeriod = True
				bUseFinalPeriod	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Comparativo del gasto de operación"

			Case 156
				bSelectGralLedger = True
				bUseAccountPattern = False
				bUseInitialDate = True
				bUseFinalDate	= True
				bUseExchangeRate = False
				bRoundValues = False
				bSelectAccounts = False
				gsReportTitle = "Directorio alfabético por auxiliares de proveedores"
					
			'Case 70
			'	bSelectGralLedger = False
			'	bUseAccountPattern = False				
			'	bUseInitialDate = False
			'	bUseFinalDate	= True
			'	bUseExchangeRate = False
			'	bRoundValues = False
			'	bSelectAccounts = False
			'	gsReportTitle = "Disponibilidades financieras de fideicomisos"
			'Case 73
		  '	bSelectGralLedger = True
			'	bUseAccountPattern = True
			'	bUseInitialDate = True
			'	bUseFinalDate	= True
			'	bUseExchangeRate = True
			'	bRoundValues = False
			'	bSelectAccounts = False
			'	gsReportTitle = "Saldos sobregirados"
			'Case 61
			'	bSelectGralLedger = true
			'	bUseAccountPattern = false				
			'	bUseInitialDate = false
			'	bUseFinalDate	= True	
			'	bUseExchangeRate = true
			'	bRoundValues = false
			'	bSelectAccounts = False
			'	gsReportTitle = "Cuentas acreedoras y deudoras"
			'Case 74
			'	bSelectGralLedger = false
			'	bUseAccountPattern = false				
			'	bUseInitialDate = true
			'	bUseFinalDate	= True	
			'	bUseExchangeRate = false
			'	bRoundValues = false
			'	bSelectAccounts = False
			'	gsReportTitle = "Comparativo Mensual de Fideicomisos"
'	    Case 115
'				bSelectGralLedger = false
'				bUseAccountPattern = false				
'				bUseInitialDate = false
'				bUseFinalDate	= false	
'				bUseExchangeRate = false
'				bRoundValues = false
'				bSelectAccounts = False
'				gsReportTitle = "Areas operativas"
'	    Case 116
'				bSelectGralLedger = false
'				bUseAccountPattern = false				
'				bUseInitialDate = false
'				bUseFinalDate	= false	
'				bUseExchangeRate = false
'				bRoundValues = false
'				bSelectAccounts = False
'				gsReportTitle = "Catálogo de cuentas"

			Case Else
				gsReportTitle = "Reporte no disponible"
		End Select
		If bSelectGralLedger Then
			Call SetComboGLGroups()
		End If
		If bUseExchangeRate Then
			Call SetComboCurrencies(46, 1)
		End If
	End Sub
	
	Sub SetComboGLGroups()
		Dim oGLVoucherUS
		On Error Resume Next

		Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		gsCboGLGroups = oGLVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), CLng(Session("uid")), 2)
		Set oGLVoucherUS = Nothing
	End Sub
	
	Sub SetComboCurrencies(nExhangeRateTypeId, nCurrencyId)
		Dim oGLVoucherUS
		Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     		
		gsCboExchangeRateTypes = oGLVoucherUS.CboExchangeRateTypes(Session("sAppServer"), CLng(nExhangeRateTypeId))
		gsCboCurrencies = oGLVoucherUS.CboCurrencies(Session("sAppServer"), CLng(nCurrencyId))
		Set oGLVoucherUS = Nothing
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

var gbSended = false;

function checkAvailableExchangeRates() {
	var obj;
	var nExchRateType, dExchRateDate, dFromDate, dToDate;

  <% If (bUseInitialDate) Then %>
		dFromDate = document.all.txtInitialDate.value;
	<% Else %>
		dFromDate = '1/1/1990';
	<% End If %>
	
  <% If (bUseFinalDate) Then %>
		dToDate = document.all.txtFinalDate.value;
	<% Else %>
		dToDate = document.all.txtExchangeRateDate.value;
	<% End If %>	
	
	nExchRateType = document.all.cboExchangeRateTypes.value;
	dExchRateDate = document.all.txtExchangeRateDate.value;

	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","CheckAvailableExchangeRates", dFromDate, dToDate, nExchRateType, dExchRateDate);

	if (obj.return_value != '') {
	  sMsg  = 'No encontré los tipos de cambio necesarios para la valorización de las monedas:\n\n';
	  sMsg += obj.return_value;
	  alert (sMsg);
	  return false;
	 }
	//document.all.cboExchangeRateCurrencies.value
	return true;
}

function formatAccount(oControl) {
	var obj;
	
	if (oControl.value == '') {
		return false;
	}
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","FormatStdAccountNumber", 1, oControl.value);
	oControl.value = obj.return_value;
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

function isDate(date) {
	var obj;
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function updateGralLedgers() {
	var obj;
	<% If (Not bUniqueGralLedger) Then %>
		obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGLGroups.value);
	<% Else %>
		obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp", "CboGralLedgersInGroup2", document.all.cboGLGroups.value);
	<% End If %>
	document.all.divCboGeneralLedgers.innerHTML = obj.return_value;		
}

function frmSend_onsubmit() {
	if (gbSended) {
		return false;
	}
<% If (bUniqueGralLedger) Then %> 
	if (document.all.cboGralLedgers.value == 0) {
		alert("Requiero se seleccione una contabilidad en forma única,\nya que este reporte no acepta contabilidades consolidadas.");
		window.document.all.cboGralLedgers.focus();
		return false;		
	}
<% End If %>
<% If (bUseInitialDate) Then %>
  if (window.document.all.txtInitialDate.value == '') {
		alert("Necesito la fecha del saldo inicial.");
		window.document.all.txtInitialDate.focus();
		return false;
	}
  if (!isDate(window.document.all.txtInitialDate.value)) {
		alert("No reconozco la fecha proporcionada para el saldo inicial.");
		window.document.all.txtInitialDate.focus();
		return false;
	}
<% End If %>
<% If (bUseFinalDate) Then %>	
  if (window.document.all.txtFinalDate.value == '') {
		alert("Necesito la fecha del saldo final.");
		window.document.all.txtFinalDate.focus();
		return false;
	}
  if (!isDate(window.document.all.txtFinalDate.value)) {
		alert("No reconozco la fecha proporcionada para el saldo final.")
		window.document.all.txtFinalDate.focus();
		return false;
  }
<% End If %>
<% If (bUseInitialPeriod) Then %>
  if (window.document.all.txtInitialDate.value == '') {
		alert("Necesito la fecha del saldo inicial del primer periodo.");
		window.document.all.txtInitialDate.focus();
		return false;
	}
  if (!isDate(window.document.all.txtInitialDate.value)) {
		alert("No reconozco la fecha proporcionada para el saldo inicial del primer periodo.");
		window.document.all.txtInitialDate.focus();
		return false;
	}
  if (window.document.all.txtFinalDate.value == '') {
		alert("Necesito la fecha del saldo final del primer periodo.");
		window.document.all.txtFinalDate.focus();
		return false;
	}
  if (!isDate(window.document.all.txtFinalDate.value)) {
		alert("No reconozco la fecha proporcionada para el saldo final del primer periodo.");
		window.document.all.txtFinalDate.focus();
		return false;
	}
<% End If %>
<% If (bUseFinalPeriod) Then %>	
  if (window.document.all.txtInitialDate2.value == '') {
		alert("Necesito la fecha del saldo inicial del segundo periodo.");
		window.document.all.txtInitialDate2.focus();
		return false;
	}
  if (!isDate(window.document.all.txtInitialDate2.value)) {
		alert("No reconozco la fecha proporcionada para el saldo inicial del segundo periodo.");
		window.document.all.txtInitialDate2.focus();
		return false;
	}
  if (window.document.all.txtFinalDate2.value == '') {
		alert("Necesito la fecha del saldo final del segundo periodo.");
		window.document.all.txtFinalDate2.focus();
		return false;
	}
  if (!isDate(window.document.all.txtFinalDate2.value)) {
		alert("No reconozco la fecha proporcionada para el saldo final del segundo periodo.");
		window.document.all.txtFinalDate2.focus();
		return false;
	}
<% End If %>
<% If (bUseExchangeRate) Then %>	
	if (document.all.cboExchangeRateTypes.value != 0 && document.all.cboExchangeRateCurrencies.value == 0) {
			alert("Requiero se seleccione la moneda a la que se valorizará la balanza.");			
			document.all.cboExchangeRateCurrencies.focus();
			return false;		
	}
	if (document.all.cboExchangeRateTypes.value == 0 && document.all.cboExchangeRateCurrencies.value != 0) {
			alert("Requiero se seleccione el tipo de cambio correspondiente a la moneda a la que se valorizará la balanza.");
			document.all.cboExchangeRateTypes.focus();
			return false;
	}
	if (document.all.cboExchangeRateTypes.value != 0 && document.all.txtExchangeRateDate.value == '') {
		if (confirm('¿La fecha para tomar los tipos de cambio para efectuar la valorización es el día ' + document.all.txtFinalDate.value + '?')) {
			document.all.txtExchangeRateDate.value = document.all.txtFinalDate.value;
		} else {
			document.all.txtExchangeRateDate.focus();
			return false;
		}
	}
	if (!checkAvailableExchangeRates()) {
	  return false;
	}
<% End If %>
<% If gsReportId = 1 Then %>
	document.all.frmSend.action = './exec/create_notebook.asp';
<% End If %>
	gbSended = true;
	return true;	
}

function window_onload() {
  <% If (bSelectGralLedger) Then %>
	updateGralLedgers();
	<% End If %>
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Reportes contables fijos
		</TD>
	  <TD align=right nowrap>
	  	<A align=absmiddle href="other_reports.asp">Regresar a la lista de reportes</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<FORM name=frmSend action="./exec/create_report.asp?id=<%=gsReportId%>" method="post" onsubmit="return frmSend_onsubmit()">
				<TABLE class=applicationTable>
					<TR class=fullScrollMenuHeader>
						<TD class=fullScrollMenuTitle colspan=2><%=gsReportTitle%></TD>
					</TR>
				 <% If (bSelectGralLedger) Then %>
				  <TR>
				    <TD>Grupo de contabilidades:</TD>
				    <TD nowrap>
							<SELECT name=cboGLGroups style="width:100%" onchange="return updateGralLedgers();">
								<%=gsCboGLGroups%>
							</SELECT>
						</TD>
					</TR>
				  <TR>
				    <TD nowrap>Obtener el reporte para:<BR></TD>
				    <TD nowrap>
							<div id=divCboGeneralLedgers>
								<SELECT name=cboGralLedgers">

								</SELECT>
							</div>
							<% If (bGLRange) Then %>
							  <BR>
							  &nbsp;En el siguiente rango: &nbsp;&nbsp;&nbsp;Desde:&nbsp;<INPUT name=txtFromGL style="width:70">&nbsp;&nbsp;&nbsp;&nbsp;
							  Hasta:&nbsp;<INPUT name=txtToGL style="width:70">&nbsp;(números de mayor)<br>
							  <INPUT type="checkbox" name=chkPrintInCascade value="true">&nbsp;&nbsp;
							  Imprimir en cascada (sin consolidar), las balanzas de las contabilidades seleccionadas  
							<% End If %>  
				    </TD>    
				  </TR>
				 <% End If %>

				 <% If (bConFidOptions) Then %>
				  <TR>
				    <TD nowrap>Obtener el reporte para:<BR></TD>
				    <TD nowrap width=80%>
							<SELECT name=cboConFidOptions style="width:100%">
								<OPTION selected value=0>Todos los grupos</OPTION>
								<OPTION value=-16>2.1.1 Fideicomisos de administración [9992]</OPTION>
								<OPTION value=-18>2.1.2 Fideicomisos de garantía [9991]</OPTION>
								<OPTION value=-19>2.1.3 Fideicomisos de inversión [9993]</OPTION>
								<OPTION value=-21>2.2 Contabilidad de mandatos [9997]</OPTION>
								<OPTION value=269>(8017) Reestructuración de Cartera.-UDIS a 5 años</OPTION>
								<OPTION value=270>(8020) Reestructuración de Cartera.-UDIS a 8 años</OPTION>
								<OPTION value=267>(8034) Reestructuración de Cartera.-UDIS a 12 años</OPTION>
								<OPTION value=268>(8048) Reestructuración de Cartera.-UDIS a 15 años</OPTION>
								<OPTION value=264>(8114) Reestructuración de Cartera.-UDIS a 15 años</OPTION>
								<OPTION value=265>(8128) Reestructuración de Cartera.-UDIS a 18 años</OPTION>
								<OPTION value=266>(8131) Reestructuración de Cartera.-UDIS a 20 años</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>

				 <% If (bUseInitialDate) Then %>
					<TR>
					  <TD>Desde el día:</TD>
					  <TD nowrap width=80%><INPUT name=txtInitialDate style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
				 <% End If %>
				 <% If (bUseFinalDate) Then %>	
					<TR>
					  <TD nowrap>Hasta el día:</TD>
					  <TD nowrap width=50%><INPUT name=txtFinalDate style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
				 <% End If %>
				 <% If (bUseInitialPeriod) Then %>
					<TR>
					  <TD>Desde el día (Primer período) :</TD>
					  <TD nowrap width=50%><INPUT name=txtInitialDate style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
					<TR>  
					  <TD>Hasta el día (Primer período) :</TD>
					  <TD nowrap width=50%><INPUT name=txtFinalDate style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
				 <% End If %>
				 <% If (bUseFinalPeriod) Then %>	
					<TR>
					  <TD>Desde el día (Segundo período):</TD>
					  <TD nowrap width=50%><INPUT name=txtInitialDate2 style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
					<TR>  
					  <TD>Hasta el día (Segundo período) :</TD>
					  <TD nowrap width=50%><INPUT name=txtFinalDate2 style="width:100">&nbsp;(día / mes / año)</TD>
					</TR>
				 <% End If %>

				 <% If (bStdAccountType) Then %>
				  <TR>
				    <TD nowrap>Tipo de Contabilidad:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboStdAccountType style="width:190">
								<OPTION selected value="1">Bancaria</OPTION>
								<OPTION value="2">Fiduciaria</OPTION><BR>
							</SELECT>
				    </TD>
				  </TR>
				  <TR>
				    <TD nowrap>Rango de cuentas:<BR></TD>
				    <TD nowrap>
							De la cuenta:&nbsp;<INPUT name=txtFromAccount style="width:190">&nbsp;&nbsp;
							A la cuenta:&nbsp;<INPUT name=txtToAccount style="width:190"><br>
				    </TD>
				  </TR>  
				 <% End If %>

				 <% If (bAccount) Then %>
				  <TR>
				   <% If (gsReportId <> 151) Then %>
				    <TD nowrap>Cuenta a considerar en el reporte:<BR></TD>
				   <% Else %>
				    <TD>Cuenta en moneda extranjera a considerar en el reporte:<BR></TD>
				   <% End If %>
				    <TD nowrap>
							<INPUT name=txtFromAccount style="width:190">&nbsp;&nbsp;
				    </TD>
				  </TR>  
				 <% End If %>

				 <% If (bUseAccountPattern) Then %>
				  <TR>
				    <TD nowrap>Presentar hasta el nivel de cuenta:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboPatterns style="width:190">
								<OPTION value="&&&&">N0: 1234</OPTION>
								<OPTION value="&&&&-&&">N1: 1234-01</OPTION>
								<OPTION value="&&&&-&&-&&">N2: 1234-01-02</OPTION>
								<OPTION value="&&&&-&&-&&-&&">N3: 1234-01-02-03</OPTION>
								<OPTION value="&&&&-&&-&&-&&-&&">N4: 1234-01-02-03-04</OPTION>
								<OPTION value="&&&&-&&-&&-&&-&&-&&">N5: 1234-01-02-03-04-05</OPTION>
								<OPTION selected value="&&&&-&&-&&-&&-&&-&&-&&">N6: 1234-01-02-03-04-05-06</OPTION>		
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>
				 
				 <% If (bRangeAccount) Then %>
				  <TR>
				    <TD nowrap>Rango de cuentas:<BR></TD>
				    <TD nowrap>
							De la cuenta:&nbsp;<INPUT name=txtFromAccount style="width:130" onblur='formatAccount(this);'>&nbsp;&nbsp;
							A la cuenta:&nbsp;<INPUT name=txtToAccount style="width:130" onblur='formatAccount(this);'> &nbsp;
							(Permiten comodines [*, ?])<br>
				    </TD>
				  </TR>  
				 <% End If %>
				 
				 <% If (bRangeSubsAccount) Then %>
				  <TR>
				    <TD nowrap>Rango de auxiliares:<BR></TD>
				    <TD nowrap>
							Del auxiliar: &nbsp; &nbsp;<INPUT name=txtFromSubsAccount style="width:130" onblur='formatSubsAccount(this);'>&nbsp;&nbsp;
							Al auxiliar: &nbsp; <INPUT name=txtToSubsAccount style="width:130" onblur='formatSubsAccount(this);'> &nbsp;
							(Permiten comodines [*, ?])<br>
				    </TD>
				  </TR>  
				 <% End If %>

				 <% If (bSelectAccounts) Then %>
				  <TR>
				    <TD nowrap valign=top>
							Obtener las siguientes cuentas:<br>(separadas por comas (,))<BR>
							(e.g. 1101-01,2302-02-01,)
				    </TD>
				    <TD align=right nowrap>
							<TEXTAREA rows=3 cols=20 name=txtAccountList style="width:100%"></TEXTAREA>
				    </TD>
				  </TR>
				  <% End If %>
				  <% If (bUseExchangeRate) Then %>
					<TR>
						<TD valign=top nowrap>Valorización de saldos:</TD>		
						<TD nowrap>
							Tipo de cambio:
							<SELECT name=cboExchangeRateTypes style="width:180">
								<OPTION value=0 selected>-- No valorizar --</OPTION>
								<%=gsCboExchangeRateTypes%>
							</SELECT>
							&nbsp;Moneda:&nbsp;
							<SELECT name=cboExchangeRateCurrencies style="width:200">
								<OPTION value=0 selected>-- No valorizar --</OPTION>
								<%=gsCboCurrencies%>						
							</SELECT>
							<br>
							Fecha para el tipo de cambio:
							<INPUT name=txtExchangeRateDate style="width:115"> (día / mes / año)
						</TD>
				  </TR>
				  <% End If %>
				  <% If (bRoundValues) Then %>
				  <TR>
				    <TD nowrap>¿Obtener los saldos redondeados?:<BR></TD>
				    <TD>
							<INPUT type="checkbox" name=chkRounded value=1>			
				    </TD>
				  </TR>
				  <% End If %>
				  <% If (bTotalByGroups) Then %>
				  <TR>
				    <TD nowrap>¿Suprimir los cortes totalizadores?:<BR></TD>
				    <TD>
							<INPUT type="checkbox" name=chkTotal value="True">			
				    </TD>
				  </TR>
				  <% End If %>
				  <% If (bChkAfectationDate) Then %>
				  <TR>
				    <TD nowrap>Las fechas anteriores corresponden a:<BR></TD>
				    <TD>
							<SELECT name=cboVoucherDatesMode style="width:250">
								<OPTION value=true>Fecha de afectación</OPTION>
								<OPTION value=false>Fecha de elaboración</OPTION>				
							</SELECT>				
				    </TD>
				  </TR>
				 <% End If %>
				 <% If (bChkUpdatedVouchers) Then %>
				  <TR>
				    <TD nowrap>Mostar las siguientes pólizas:<BR></TD>
				    <TD>
							<SELECT name=cboVoucherStatus style="width:250">
								<OPTION value=true>Pólizas actualizadas</OPTION>
								<OPTION value=false>Pólizas pendientes de actualizar</OPTION>				
							</SELECT>			
				    </TD>
				  </TR>
				 <% End If %>
				<% If (bChkOptionToDisplay) Then %>
				  <TR>
				    <TD nowrap>Obtener el reporte:<BR></TD>
				    <TD>
							<SELECT name=cboOptionToDisplay style="width:250">
								<OPTION value="M">Movimientos</OPTION>
								<OPTION value="S">Saldos</OPTION>				
							</SELECT>			
				    </TD>
				  </TR>
				 <% End If %> 
				 <% If (bUseParticipantType) Then %>
				  <TR>
				    <TD nowrap>Tipo de participantes a presentar:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboParticipants style="width:190">
								<OPTION selected value="U">Usuarios</OPTION>
								<OPTION value="O">Organizaciones</OPTION>
								<OPTION value="S">Sistemas</OPTION>
								<OPTION value="R">Roles</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>
				 <% If (bUseParticipantStatus) Then %>
				  <TR>
				    <TD nowrap>Mostrar los siguientes participantes:<BR></TD>
				    <TD nowrap width=80%>
							<SELECT name=cboParticipantStatus style="width:190">
								<OPTION selected value=" ">Todos</OPTION>
								<OPTION value="A">Activo</OPTION>
								<OPTION value="S">Suspendido</OPTION>
								<OPTION value="D">Eliminado</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>

				 <% If (bParticipantOrder) Then %>
				  <TR>
				    <TD nowrap>Ordenar los participantes por:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboParticipantOrder style="width:190">
								<OPTION selected value="ParticipantKey">Clave</OPTION>
								<OPTION value="ParticipantName">Nombre</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>

				 <% If (bAccountOrder) Then %>
				  <TR>
				    <TD nowrap>Ordenar el catálogo por:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboAccountOrder style="width:190">
								<OPTION selected value="">Por tipo de contabilidad</OPTION>
								<OPTION value="Fideicomiso">Por número de fideicomiso</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>

				 <% If (bTittle101) Then %>
				  <TR>
				    <TD nowrap>Saldos acumulados para:<BR></TD>
				    <TD nowrap>
							<SELECT name=cboTittles style="width:190">
								<!--<OPTION selected value="CI">Indice inflacionario</OPTION> !-->
								<OPTION selected value="PM">Posición monetaria</OPTION>
							</SELECT>
				    </TD>
				  </TR>
				 <% End If %>
				  
				 <% If bUseSigners Then %>
					<TR>
					  <TD valign=top>Funcionario 1:</TD>
					  <TD nowrap width=90%>
							Nombre:&nbsp;&nbsp;<INPUT name=txtSigner1Name style="width:250px"><br>
							Puesto:&nbsp;&nbsp;&nbsp;<INPUT name=txtSigner1Title style="width:250px">
					  </TD>
					</TR>
					<TR>
					  <TD valign=top>Funcionario 2:</TD>
					  <TD nowrap width=90%>
							Nombre:&nbsp;&nbsp;<INPUT name=txtSigner2Name style="width:250px"><br>
							Puesto:&nbsp;&nbsp;&nbsp;<INPUT name=txtSigner2Title style="width:250px">
					  </TD>
					</TR>
				<% End If %>
					<TR align=right>
						<TD colspan=2>
							<INPUT class=cmdSubmit name=cmdBuild type=submit style='width:100;' value="Generar reporte">
							&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
							<INPUT class=cmdSubmit name=cmdCancel type=button style='width:75;' value="Cancelar" onclick='window.location.href="<%=Session("main_page")%>"'>
							&nbsp;
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


