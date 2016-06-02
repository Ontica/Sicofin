<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReportId, gsCboGLGroups, gsCboCurrencies, gsCboExchangeRateTypes, gsCboTransactionTypes, gsCboVoucherTypes
	
	Call Main()
	
	Sub Main()
		Dim oGLVoucherUS
		'***************
		gsReportId = 16
		Set oGLVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")     		
		gsCboGLGroups		 = oGLVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), CLng(Session("uid")), 2)
		gsCboCurrencies  = oGLVoucherUS.CboCurrencies(Session("sAppServer"))
		gsCboExchangeRateTypes = oGLVoucherUS.CboExchangeRateTypes(Session("sAppServer"))
		gsCboTransactionTypes = oGLVoucherUS.CboTransactionTypes(Session("sAppServer"))
		gsCboVoucherTypes = oGLVoucherUS.CboVouchersTypes(Session("sAppServer"))
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

function isDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function updateGralLedgers() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CboGralLedgersInGroup", document.all.cboGLGroups.value);	
	document.all.divCboGeneralLedgers.innerHTML = obj.return_value;	
}

function createBalance() {
	if (gbSended) {
		return false;
	}
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
  if (document.all.cboBalanceFormat.value == 2) {
		if (document.all.txtInitialDate2.value == '') {		
				alert("Requiero la fecha del saldo inicial del segundo período.");
				window.document.all.txtInitialDate2.focus();
				return false;
		}		
		if (document.all.txtFinalDate2.value == '') {		
				alert("Requiero la fecha del saldo final del segundo período.");
				window.document.all.txtFinalDate2.focus();
				return false;
		}  
		if (!isDate(window.document.all.txtInitialDate2.value)) {
			alert("No reconozco la fecha proporcionada para el saldo inicial del segundo período.");
			window.document.all.txtInitialDate2.focus();
			return false;
		}		
		if (!isDate(window.document.all.txtFinalDate2.value)) {
			alert("No reconozco la fecha proporcionada para el saldo final del segundo período.");
			window.document.all.txtFinalDate2.focus();
			return false;
		}
	}
	if (document.all.cboExchangeRateTypes.value != 0 && document.all.cboExchangeRateCurrencies.value == 0) {
			alert("Requiero se seleccione la moneda a la que se valorizará la balanza.");			
			document.all.cboExchangeRateCurrencies.focus();
			return false;		
	}
	if (document.all.cboExchangeRateTypes.value == 0 && document.all.cboExchangeRateCurrencies.value != 0) {
			alert("Requiero se seleccione el tipo de cambio correspondiete a la moneda a la que se valorizará la balanza.");
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
	document.all.frmSend.submit();
}

function window_onload() {	
	updateGralLedgers();
	document.all.txtInitialDate2.disabled = true;
	document.all.txtInitialDate2.style.backgroundColor = 'beige';
	document.all.txtFinalDate2.disabled = true;
	document.all.txtFinalDate2.style.backgroundColor = 'beige';	
}

function cmdOptions_onclick() {
	alert("Por el momento esta opción no está disponible");
} 

function cboBalanceFormat_onchange() {
	if(document.all.cboBalanceFormat.value != 2) {
		document.all.txtInitialDate2.disabled = true;
		document.all.txtInitialDate2.style.backgroundColor = 'beige';
		document.all.txtFinalDate2.disabled = true;		
		document.all.txtFinalDate2.style.backgroundColor = 'beige';
	} else {
		document.all.txtInitialDate2.disabled = false;
		document.all.txtInitialDate2.style.backgroundColor = 'white';
		document.all.txtFinalDate2.disabled = false;
		document.all.txtFinalDate2.style.backgroundColor = 'white';
	}
}

//--> 
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Balanzas de comprobación
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
						<A href="../transactions/pages/voucher_wizard.asp">Crear póliza</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../transactions/pages/voucher_explorer.asp">Explorador de pólizas</A>
						&nbsp;&nbsp;&nbsp;&nbsp;						
						<A href="../balances/balance_explorer.asp">Explorador de saldos</A>
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
			<FORM name=frmSend action="./exec/build_balance.asp" method="post">
				<TABLE class=applicationTable>
					<TR class=fullScrollMenuHeader>
						<TD class=fullScrollMenuTitle colspan=2>Parametrización</TD>
					</TR>
					<TR>
						<TD nowrap>Tipo de balanza a obtener:</TD>
						<TD nowrap>
							<SELECT name=cboBalanceFormat LANGUAGE=javascript onchange="return cboBalanceFormat_onchange()">
								<OPTION selected value=1>Balanza tradicional</OPTION>
								<OPTION value=2>Balanza de comparación entre períodos</OPTION>
								<OPTION value=3>Balanza con saldos a nivel de auxiliar</OPTION>
								<OPTION value=4>Balanza con columna de saldos promedio</OPTION>
								<OPTION value=5>Balanza consolidada con cuentas en cascada</OPTION>
							</SELECT>
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<INPUT class=cmdSubmit name=cmdOptions style="width:175px" type=button value="Otras opciones de configuración..."  style="width:210" LANGUAGE=javascript onclick="return cmdOptions_onclick()">
						</TD>
					</TR>
				  <TR>
				    <TD nowrap>Grupo de contabilidades:</TD>
				    <TD nowrap>
							<SELECT name=cboGLGroups style="width:100%" onchange="return updateGralLedgers();">
								<%=gsCboGLGroups%>
							</SELECT>
						</TD>
					</TR>	
				  <TR>
				    <TD valign=top>Obtener balanza para:</TD>
				    <TD nowrap>
							<div id=divCboGeneralLedgers>
								<SELECT name=cboGralLedgers">

								</SELECT>
							</div><BR>
							&nbsp;En el siguiente rango: &nbsp;&nbsp;&nbsp;Desde:&nbsp;<INPUT name=txtFromGL style="width:70">&nbsp;&nbsp;&nbsp;&nbsp;
							Hasta:&nbsp;<INPUT name=txtToGL style="width:70">&nbsp;(números de mayor)<br>
							<INPUT type="checkbox" name=chkPrintInCascade value="true">&nbsp;&nbsp;
							Imprimir en cascada (sin consolidar), las balanzas de las contabilidades seleccionadas
				    </TD> 
				  </TR>
					<TR>
					  <TD valign=top>Saldos a incluir en la balanza:</TD>	  
					  <TD nowrap width=100%>
							<TABLE>
								<TR>
									<TD>&nbsp;</TD>
									<TD nowrap>Fecha del saldo inicial</TD>
									<TD>&nbsp;&nbsp;&nbsp;</TD>
									<TD nowrap>Fecha del saldo final</TD>
								</TR>
								<TR>
									<TD>Período inicial:</TD>
									<TD><INPUT name=txtInitialDate style="width:100"></TD>
									<TD>&nbsp;&nbsp;&nbsp;</TD>
									<TD><INPUT name=txtFinalDate style="width:100">&nbsp;&nbsp;(día / mes / año)</TD>
									<TD>&nbsp;&nbsp;&nbsp;</TD><br>
									<TD>
								  <INPUT type="checkbox" name=chkCascadeDates value="true">&nbsp;&nbsp;
					 		        Imprimir fechas en cascada					
					 		    </TD>    
								</TR>
								<TR>					
									<TD nowrap>Período final:</TD>
									<TD><INPUT name=txtInitialDate2 style="width:100"></TD>
									<TD>&nbsp;&nbsp;&nbsp;</TD>
									<TD><INPUT name=txtFinalDate2 style="width:100">&nbsp;&nbsp;(día / mes / año)</TD>					
								</TR>
							</TABLE>
						</TD>
					</TR>
				  <TR>
				    <TD valign=top>Rango y nivel de cuentas a obtener:<BR></TD>
				    <TD nowrap>
							De la cuenta:&nbsp;<INPUT name=txtFromAccount style="width:190">&nbsp;&nbsp;
							A la cuenta:&nbsp;<INPUT name=txtToAccount style="width:190"><br>
							Nivel:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<SELECT name=cboPatterns style="width:190">
									<OPTION value="&&&&">N0: 1234</OPTION>
									<OPTION value="&&&&-&&">N1: 1234-01</OPTION>
									<OPTION value="&&&&-&&-&&">N2: 1234-01-02</OPTION>
									<OPTION value="&&&&-&&-&&-&&">N3: 1234-01-02-03</OPTION>
									<OPTION value="&&&&-&&-&&-&&-&&">N4: 1234-01-02-03-04</OPTION>
									<OPTION value="&&&&-&&-&&-&&-&&-&&">N5: 1234-01-02-03-04-05</OPTION>
									<OPTION selected value="&&&&-&&-&&-&&-&&-&&-&&">N6: 1234-01-02-03-04-05-06</OPTION>		
								</SELECT>&nbsp;
							Incluir:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<SELECT name=cboBalanceType>
								<OPTION value=1>Sólo cuentas con saldo actual</OPTION>
								<OPTION value=2>Sólo cuentas con movimientos</OPTION>					
								<OPTION selected value=3>Cuentas con saldo actual o movimientos</OPTION>
								<OPTION value=5>Sólo cuentas con saldos sobregirados</OPTION>
								<OPTION value=4>Todas las cuentas sin importar su saldo</OPTION>					
							</SELECT>			
						</TD>		
				  </TR>
					<TR>
						<TD valign=top>Valorización de saldos:</TD>		
						<TD nowrap>
							Tipo de cambio:
							<SELECT name=cboExchangeRateTypes style="width:180">
								<OPTION value=0 selected>-- No valorizar --</OPTION>
								<%=gsCboExchangeRateTypes%>
							</SELECT>
							&nbsp;Moneda:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							<SELECT name=cboExchangeRateCurrencies style="width:200">
								<OPTION value=0 selected>-- No valorizar --</OPTION>
								<%=gsCboCurrencies%>
							</SELECT>
							<br>
							Fecha para el tipo de cambio:
							<INPUT name=txtExchangeRateDate style="width:115"> (día / mes / año)<br>
							Consolidar saldos en la moneda del tipo de cambio
							<INPUT type="checkbox" name=chkConsolidateExchangeRateCurrency>
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
				    <TD>Incluir saldos de las siguientes monedas:<BR></TD>    
				    <TD nowrap>
							Moneda:<SELECT name=cboCurrencies style="width:40%">			
												<OPTION value=0 selected> -- Todas las monedas -- </OPTION>
												<%=gsCboCurrencies%>
											</SELECT>
											&nbsp;Todas excepto la moneda seleccionada<INPUT type="checkbox" name=chkCurrencies value="true">
						</TD>		
				  </TR>    
				  <TR>
				    <TD>Obtener la balanza en el siguiente formato:<BR></TD>
				    <TD>
							<SELECT name=cboReportMode style="width:40%">
								<OPTION value="Text" selected>Archivo de texto</OPTION>
								<OPTION value="Word">Microsoft® Word</OPTION>				
								<OPTION value="Excel">Microsoft® Excel </OPTION>
							</SELECT>
				    </TD>
				  </TR>
					<TR>
					  <TD colspan=2 nowrap align=right>
							<INPUT class=cmdSubmit name=cmdSend style="width:100px" type=button value="Crear balanza" onclick="createBalance();">
							&nbsp; &nbsp;
							<INPUT class=cmdSubmit name=cmdCancel style="width:100px" type=button value="Cancelar" onclick="window.location.href = '<%=Application("main_page")%>';">
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