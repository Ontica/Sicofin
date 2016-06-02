<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsCboCurrencies, gsCboExchangeRatesTypes, gsExchangeRatesTable
	Dim gsFromDate, gsToDate
	
	Call Main()
	
	Sub Main()
		Dim oCurrenciesUS
		Dim gnCurrency, gnExchangeRate
		'****************
		'On Error Resume Next
						
		Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")					
		If (Request.Form.Count <> 0) Then
			gnCurrency     = Request.Form("cboCurrencies")
			gnExchangeRate = Request.Form("cboExchangeRateTypes")
			gsFromDate     = Request.Form("txtFromDate")
			gsToDate       = Request.Form("txtToDate")
		Else		
			gnCurrency     = 0
			gnExchangeRate = 0
			gsFromDate     = Date()
			gsToDate       = ""
		End If
		gsCboCurrencies         = oCurrenciesUS.CboCurrenciesWithException(Session("sAppServer"), 1, CLng(gnCurrency))
		gsCboExchangeRatesTypes = oCurrenciesUS.CboExchangeRatesTypes(Session("sAppServer"), CLng(gnExchangeRate))				
		If Len(gsToDate) = 0 Then
			gsExchangeRatesTable = oCurrenciesUS.ExchangeRatesTable(Session("sAppServer"), 1, CDate(gsFromDate), CDate(gsFromDate))
		Else
			gsExchangeRatesTable = oCurrenciesUS.ExchangeRatesTable(Session("sAppServer"), 1, CDate(gsFromDate), CDate(gsToDate))
		End If
		Set oCurrenciesUS = Nothing		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("./exec/exception.asp")
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
		obj = RSExecute("../general_scripts.asp","IsDate", sDate);
		return obj.return_value;
	} else {
		return true;
	}
}

function callEditor(nOperation) {	
  switch (nOperation) {  
    case 1:		//Add
			sURL = 'exchange_rates_editor.asp?date=' + arguments[1] + '&typeId=' + arguments[2];
			window.open(sURL, null, "height=440,width=332,location=0,resizable=0");
			return false;
    case 2:		//Edit
			sURL = 'exchange_rate_editor.asp?id=' + arguments[1];
			window.open(sURL, null, "height=180,width=380,location=0,resizable=0");
			return false;
		case 3:
			sURL = 'exchange_rates_editor.asp';
			window.open(sURL, null, "height=440,width=332,location=0,resizable=0");
			return false;
	}
	return false;
}

function refreshPage(nOrderId) {
	var sURL = "";	
	sURL = "exchange_rates.asp?id=" + window.document.all.cboCurrencies.value;
	sURL += "&month=" + window.document.all.cboMonths.value;
  if (nOrderId == 0) {
		window.location.href = sURL;
	} else {	
	  window.location.href = sURL + '&order=' + nOrderId;
	}
	return false;
}

function searchExchangeRates() {
	if (document.all.txtFromDate.value == '') { 	
		alert("Necesito al menos la fecha inicial del tipo de cambio.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (!isDate(document.all.txtFromDate.value)) { 
		alert("No reconozco la fecha inicial proporcionada.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (!isDate(document.all.txtToDate.value)) { 
		alert("No reconozco la fecha final proporcionada.");
		document.all.txtToDate.focus();
		return false;
	}
	document.frmSend.target = '_self';
	document.frmSend.action = '';
	document.frmSend.submit();
	return false;
}


//-->
</SCRIPT>
</HEAD>
<BODY>
<FORM name=frmSend method=post action=''>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Explorador de tipos de cambio
					</TD>
					<TD nowrap align=right>
						<A href='' onclick='return(searchExchangeRates());'>Ejecutar búsqueda</A>
						&nbsp; &nbsp; | &nbsp; &nbsp;
						<A href='' onclick='return(callEditor(3));'>Agregar tipos de cambio</A>						
						<img align=absmiddle src='/empiria/images/invisible4.gif'>						
						<img align=absmiddle src='/empiria/images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>												
						<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
						<img align=absmiddle src='/empiria/images/invisible.gif'>						
						<img align=absmiddle src='/empiria/images/close_white.gif' onclick='closeOptionsWindow(document.all.divSearchOptions, document.all.cmdSearchOptionsTack)' alt='Cerrar'>						
					</TD>
				</TR>
				<TR>
					<TD nowrap colspan=2>
						Del día: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; 
						<INPUT name=txtFromDate style="width:120;height:20;" value='<%=gsFromDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromDate)'>&nbsp; &nbsp;
						&nbsp; &nbsp; Al día: &nbsp; &nbsp;
						<INPUT name=txtToDate style="width:120;height:20;" value='<%=gsToDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToDate)'>&nbsp; &nbsp;(día / mes / año)
					</TD>
				</TR>
				<TR>
					<TD colspan=2>
						Tipo de cambio:						<SELECT name=cboExchangeRateTypes style="width:158;height:20;">
							<OPTION value=0>-- Todos los tipos --</OPTION>
							<%=gsCboExchangeRatesTypes%>
						</SELECT>
						&nbsp; &nbsp; &nbsp; Moneda:&nbsp;						<SELECT name=cboCurrencies style="width:158;height:20;">
							<OPTION value=0>-- Todas las monedas --</OPTION>
							<%=gsCboCurrencies%>
						</SELECT>
					</TD>					
				</TR>				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<THEAD>
					<TR class=fullScrollMenuHeader valign=center>
						<TD class=fullScrollMenuTitle colspan=5>
							Tipos de cambio encontrados
						</TD>
					</TR>					
					<TR class=applicationTableHeader valign=center>
						<TD nowrap><img src='/empiria/images/collapsed.gif' onclick='outline();' alt='Contraer todo'></TD>						
						<TD nowrap align="center">Moneda</TD>
						<TD nowrap align=right>Tipo de cambio</TD>
						<TD nowrap align=right>Valor</TD>
            <TD width=90%>&nbsp;</TD>
					</TR>
				</THEAD>
				<% If (Len(gsExchangeRatesTable) <> 0) Then %>
				   <%=gsExchangeRatesTable%>
				<% Else %>
					<TBODY>
						<TR>
							<TD colspan=5>
								<b>No encontré ningún tipo de cambio con las opciones de búsqueda proporcionadas.</b>
							</TD>
						</TR>
					</TBODY>
				<% End If %>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>