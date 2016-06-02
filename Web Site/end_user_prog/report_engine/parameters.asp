<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnReportId, gsCboGroups, gsReportName, gsCboCurrencies, gsCurrenciesList, gsCboTransactionTypes, gsCboVoucherTypes
	Dim gsCboExchangeRateTypes, gsCboAccountGroupLevels, gnTechnology
	
	
	gnReportId = Request.QueryString("id")
	Call Main()
	
	Sub Main()
		Dim oReportsEngine, oRecordset, oGLVoucherUS
		
		Dim nRuleId
		'*******************************************
		
		Set oReportsEngine = Server.CreateObject("EUPReportBuilder.CBuilder")
		Set oRecordset = oReportsEngine.Report(Session("sAppServer"), CLng(gnReportId))				
		gsReportName   = oRecordset("reportName")
		nRuleId			   = oRecordset("reportDataSubClassId")
		gnTechnology   = oRecordset("reportTechnology")
		gsCboGroups = oReportsEngine.CboGLGroupsForRule(Session("sAppServer"), CLng(nRuleId), _
																										CLng(Session("uid")), 2, 0)
		oRecordset.Close
		Set oRecordset = Nothing
		Set oReportsEngine = Nothing
		
		Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
	  gsCboCurrencies = oGLVoucherUS.CboCurrencies(Session("sAppServer"), 1)
	  gsCboExchangeRateTypes = oGLVoucherUS.CboExchangeRateTypes(Session("sAppServer"), 49)
	  gsCboAccountGroupLevels = oGLVoucherUS.CboRuleLevels(Session("sAppServer"), 7, 2)
	  gsCurrenciesList = oGLVoucherUS.CurrenciesList(Session("sAppServer"))
		gsCboTransactionTypes = oGLVoucherUS.CboTransactionTypes(Session("sAppServer"))
		gsCboVoucherTypes = oGLVoucherUS.CboVouchersTypes(Session("sAppServer"), Session("uid"))
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
var gnBalanceRows = 1;

function updateCboGLGroups() {
	var obj;			
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp", "CboGLGroupsForRule", document.all.cboRules.value, 2, 0);
	document.all.divCboGLGroups.innerHTML = obj.return_value;	
}

function updateGralLedgers() {
	var obj;
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGLGroups.value);	
	document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
}

function updateCboLevels() {
	var obj;
	//obj = RSExecute("../../end_user_prog_scripts.asp", "CboLevels", document.all.cboRules.value);	
	//document.all.divCboLevels.innerHTML = obj.return_value;	
}

function selectedOption(oControl) {	
	return (oControl.options[oControl.selectedIndex].value);	
}

function isDate(date) {
	var obj;
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function frmSend_onsubmit() {
	var i;
	gbSended = false;
	if (gbSended) {
		if (!confirm("Actualmente se está procesando otro reporte.\n\n.¿Desea continuar?")) {			
			return false;
		}
	}
	if ((document.all.txtInitialDate1.value == '') && (!document.all.txtInitialDate1.disabled)) {		
		alert("Necesito la fecha inicial del 'período único'.");
		document.all.txtInitialDate1.focus();
		return false;
	}	
	if ((!isDate(document.all.txtInitialDate1.value)) && (!document.all.txtInitialDate1.disabled)) {
		alert("No reconozco la fecha inicial proporcionada.");
		document.all.txtInitialDate1.focus();
		return false;		
	}	
	if (document.all.txtFinalDate1.value == '') {
		alert("Necesito la fecha final del 'período único'.");
		document.all.txtFinalDate1.focus();
		return false;
	}
	if (!isDate(document.all.txtFinalDate1.value)) {
		alert("No reconozco la fecha final del 'período único'.")
		document.all.txtFinalDate1.focus();
		return false;
	}
	if (document.all.cboExchangeRateTypes.value != 0 && document.all.cboExchangeRateCurrencies.value == 0) {
		alert("Requiero se seleccione la moneda a la que se valorizarán los saldos.");
		document.all.cboExchangeRateCurrencies.focus();
		return false;
	}
	if (document.all.cboExchangeRateTypes.value == 0 && document.all.cboExchangeRateCurrencies.value != 0) {
		alert("Requiero se seleccione el tipo de cambio al que se valorizarán los saldos.");
		document.all.cboExchangeRateTypes.focus();
		return false;
	}
	gbSended = true;	  
  window.open("./exec/generator.asp", "oResultsWindow", "height=210,width=420,status=no,toolbar=no,menubar=no,location=no");
	return true;
}

function createNewBalanceRow() {
	var sRow;
  var s = new String();
  
	gnBalanceRows++;
	
	sRow  = "<TR><TD>Período " + gnBalanceRows + ":</TD>";
	sRow += "<TD><INPUT name=txtInitialDate" + gnBalanceRows + " disabled style='width:90;background:beige;'></TD>";
	sRow += "<TD nowrap><INPUT name=txtFinalDate" + gnBalanceRows + " style='width:90'>&nbsp;(día / mes / año)&nbsp;&nbsp;&nbsp;</TD>";
	sRow += "<TD nowrap><INPUT type=radio name=optBalanceType" + gnBalanceRows + " checked value=1 onclick='enabledPeriod(document.all.txtInitialDate" + gnBalanceRows + ", false);'>Saldos &nbsp;";
	sRow +=	"<INPUT type=radio name=optBalanceType" + gnBalanceRows + " value=2 onclick='enabledPeriod(document.all.txtInitialDate" + gnBalanceRows + ", true);'>Saldos promedio &nbsp;";
  sRow +=	"<INPUT type=radio name=optBalanceType" + gnBalanceRows + " value=3 onclick='enabledPeriod(document.all.txtInitialDate" + gnBalanceRows + ", true);'>Ambas columnas</TD></TR>";
  s = document.all.newBalanceRow.innerHTML;
	document.all.newBalanceRow.innerHTML = s.substring(0, s.length - 8) + sRow + "</TABLE>";
}

function enabledPeriod(oTextBox, bEnabled) {
	if (bEnabled) {
		oTextBox.style.backgroundColor = 'white';
		oTextBox.disabled = false;
	} else {
		oTextBox.style.backgroundColor = 'beige';
		oTextBox.disabled = true;
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Generador de reportes
		</TD>
	  <TD align=right nowrap>
			<A href='selector.asp'>Regresar a la lista de reportes</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Application("main_page")%>';" alt="Cerrar y regresar a la página principal">
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
						<A href="../../contabilidad/reports/balances.asp">Balanzas de comprobación</A>
						&nbsp; &nbsp; &nbsp;
						<A href="../../contabilidad/reports/other_reports.asp">Reportes contables fijos</A>
						&nbsp; &nbsp; &nbsp;
						<A href="../../contabilidad/balances/balance_explorer.asp">Explorador de saldos</A>
					</TD>
				</TR>	
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>  
			<FORM name=frmSend action="./exec/generator.asp" method="post" target="oResultsWindow" onsubmit='return(frmSend_onsubmit());'>
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle colspan=2>
						<%=gsReportName%>
					</TD>
				</TR>				
				<TR>
				  <TD>Grupo de contabilidades:</TD>
				  <TD nowrap>
						<span id=divCboGLGroups>
						<SELECT name=cboGLGroups style="width:100%" onchange="return updateGralLedgers();">
							<%=gsCboGroups%>
						</SELECT>
						</span>
					</TD>
				</TR>
				<TR>
				  <TD valign=top>Obtener el reporte para:<BR></TD>
				  <TD nowrap width='80%'>
						<div id=divCboGeneralLedgers>
							<SELECT name=cboGralLedgers style="width:100%">
								<OPTION value=0>--Todas las contabilidades en el grupo seleccionado--</OPTION>
							</SELECT>
						</div>
						<BR>
						<% If (gnTechnology <> "E") AND (gnTechnology <> "W") Then %>
						<INPUT type="checkbox" name=chkPrintInCascade value="true">&nbsp;&nbsp;
						Ejecutar el reporte en cascada para cada una de las contabilidades seleccionadas (sin consolidar)<br>
						&nbsp;En el siguiente rango: &nbsp; &nbsp;Desde:&nbsp;<INPUT name=txtFromGL style="width:70">&nbsp; &nbsp;&nbsp;
						Hasta:&nbsp;<INPUT name=txtToGL style="width:70">&nbsp;(números de mayor)
						<% End If %>
				  </TD>    
				</TR>
				<TR>
				  <TD valign=top>Fechas:</TD>
				  <TD nowrap>
				  	<SPAN id=divPeriods>
				  	<TABLE>
							<TR>
								<TD nowrap><b>Períodos</b> &nbsp;&nbsp;&nbsp;</TD>
								<TD nowrap><b>Fecha inicial</b></TD>
								<TD nowrap><b>Fecha final</b>&nbsp;&nbsp;</TD>
							</TR>
							<TR>
								<TD>Período único:</TD>
								<TD><INPUT name=txtInitialDate1 style="width:90"></TD>
								<TD nowrap><INPUT name=txtFinalDate1 style="width:90">&nbsp;(día / mes / año)&nbsp;&nbsp;&nbsp;</TD>
							</TR>
						</TABLE>
					</SPAN>	
					</TD>
				</TR>
				<TR>
				  <TD>Incluir las siguientes transacciones:</TD>
				  <TD nowrap>
						Tipo de transacción:<SELECT name=cboTransactionTypes style="width:40%">
																	<OPTION value=0 selected> -- Todas las transacciones -- </OPTION>														
																	<%=gsCboTransactionTypes%>
																</SELECT>
																&nbsp;Todas excepto las del tipo seleccionado<INPUT type="checkbox" name=chkTransactionTypes value="true">
					</TD>
				</TR>
				<TR>
				  <TD>Incluir las siguientes pólizas:<BR></TD>    
				  <TD nowrap>
						Tipo de póliza:<SELECT name=cboVoucherTypes style="width:40%">			
														<OPTION value=0 selected> -- Todos los tipos de póliza -- </OPTION>
															<%=gsCboVoucherTypes%>
														</SELECT>
														&nbsp;Todas excepto las del tipo seleccionado<INPUT type="checkbox" name=chkVoucherTypes value="true">
					</TD>
				</TR>
				<TR>
					<TD valign=top>Valorización de saldos:</TD>		
					<TD nowrap>
						Tipo de cambio<sup>*</sup>:
						<SELECT name=cboExchangeRateTypes style="width:200">
							<%=gsCboExchangeRateTypes%>
						</SELECT>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						Moneda:&nbsp;
						<SELECT name=cboExchangeRateCurrencies style="width:200">
							<%=gsCboCurrencies%>						
						</SELECT><br><br>
						<sup>*</sup>Las fechas para los tipos de cambio serán las correspondientes a los períodos seleccionados.
					</TD>
				</TR>
				<TR>
					<TD valign=top>Redondear los saldos a:</TD>		
					<TD nowrap>						
						<SELECT name=cboRoundBalancesTo>
							<OPTION value=0 selected>No redondear</OPTION>
							<OPTION value=1>Cifras en pesos (sin centavos)</OPTION>
							<OPTION value=2>Cifras en miles de pesos</OPTION>
							<OPTION value=3>Cifras en millones de pesos</OPTION>				
						</SELECT>						
					</TD>
				</TR>
				
				
				
				<TR>
					<TD>&nbsp;</TD>
				  <TD nowrap align=right>
						<INPUT type="hidden" name=txtReportId value=<%=gnReportId%>>
						<INPUT type=submit class=cmdSubmit name=cmdBuild  value="Generar el reporte...">
						&nbsp; &nbsp;
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