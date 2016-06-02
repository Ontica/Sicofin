<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsCboExchangeRatesTypes, gsExchangeRatesTable
	Dim gnExchangeRateType, gsDate
	
	Call Main()
	
	Sub Main()
		Dim oCurrenciesUS
		Dim gnCurrency, gnExchangeRate
		'****************
		'On Error Resume Next						
		Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")					
		If (Len(Request.QueryString) <> 0) Then
			gnExchangeRateType = Request.QueryString("typeId")
			gsDate						 = Request.QueryString("date")		
		ElseIf (Request.Form.Count <> 0) Then
			gnExchangeRateType = Request.Form("cboExchangeRateTypes")
			gsDate						 = Request.Form("txtDate")				
		Else
			gnExchangeRateType = 46
			gsDate = Date			
		End If
		gsCboExchangeRatesTypes = oCurrenciesUS.CboExchangeRatesTypes(Session("sAppServer"), CLng(gnExchangeRateType))
		gsExchangeRatesTable = oCurrenciesUS.ExchangeRatesEditorTable(Session("sAppServer"), 1, CLng(gnExchangeRateType), CDate(gsDate))
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

function currencyName(nCurrencyId) {
	var obj;	
	obj = RSExecute("../general_scripts.asp", "CurrencyName", nCurrencyId);
	return obj.return_value;	
}

function isDate(sDate) {
	var obj;		
	if (sDate != '') {
		obj = RSExecute("../general_scripts.asp", "IsDate", sDate);
		return obj.return_value;
	} else {
		return true;
	}
}

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../general_scripts.asp", "IsNumeric", sNumber, nDecimals);
	return obj.return_value;
}

function createArray() {
	var oObject, i;
	var sTemp = '';
	
	for (i = 0; i < document.all.txtExchangeRate.length; i++) {
		oObject = document.all.txtExchangeRate[i];
		if (oObject.value != '') {		
			if (sTemp != '') {
				sTemp += '|' + oObject.id + ';' + oObject.value;
			} else {
				sTemp = oObject.id + ';' + oObject.value;
			}
		}
	}
	return sTemp;
}

function saveItems() {
	var oObject, i;
	var sMsg;
	for (i = 0; i < document.all.txtExchangeRate.length; i++) {
		oObject = document.all.txtExchangeRate[i];
		if (oObject.value != '') {
			if (!isNumeric(oObject.value, 6)) {
				sMsg  = 'No reconozco el valor del tipo de cambio para la moneda\n';
				sMsg += '"' + currencyName(oObject.id) + '".';
				alert(sMsg);
				oObject.focus();
				return false;
			}
			if (Number(oObject.value) <= 0) {
				sMsg  = 'El valor del tipo de cambio para la moneda "' + currencyName(oObject.id) + '"\n';
				sMsg += 'debe ser positivo.'				
				alert(sMsg);
				oObject.focus();
				return false;
			}
		}
	}	
	document.frmSend.txtExchangeRatesArray.value = createArray();
	document.frmSend.target = '_self';
	document.frmSend.action = './exec/save_exchange_rates.asp';
	document.frmSend.submit();
	return false;
}

function getExchangeRates() {
	if (document.all.txtDate.value == '') {
		alert("Requiero la fecha para los tipos de cambio.");
		document.all.txtDate.focus();
		return false;
	}
	if (!isDate(document.all.txtDate.value)) {
		alert("No reconozco la fecha para los tipos de cambio.");
		document.all.txtDate.focus();
		return false;
	}
	document.frmSend.target = '_self';
	document.frmSend.action = 'exchange_rates_editor.asp';
	document.frmSend.submit();
	return false;	
}


function cboExchangeRateTypes_onchange() {
		if (document.all.txtDate.value == '') {
			return false;
		}
		if (!isDate(document.all.txtDate.value)) {
			return false;
		}
		getExchangeRates();
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmSend method=post action=''>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Edición de tipos de cambio
					</TD>
					<TD nowrap align=right>
						<A href='' onclick='return(getExchangeRates());'>Refrescar</A>
						<img align=absmiddle src='/empiria/images/invisible4.gif'>						
						<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
						<img align=absmiddle src='/empiria/images/invisible.gif'>						
						<img align=absmiddle src='/empiria/images/close_white.gif' onclick='window.close();' alt='Cerrar'>						
					</TD>
				</TR>
				<TR>
					<TD nowrap colspan=2>
						Fecha: &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						<INPUT name=txtDate style="width:112;height:20;" value='<%=gsDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtDate)'>
						&nbsp; (día / mes / año)
					</TD>
				</TR>
				<TR>
					<TD nowrap colspan=2>						
						Tipo de cambio:						<SELECT name=cboExchangeRateTypes style="width:150;height:20;" LANGUAGE=javascript onchange="return cboExchangeRateTypes_onchange()">
							<%=gsCboExchangeRatesTypes%>
						</SELECT>						
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
		  <DIV STYLE="overflow:auto; float:bottom; width=100%; height=325px">
			<TABLE class=applicationTable>
				<THEAD>
					<TR class=fullScrollMenuHeader valign=center>
						<TD class=fullScrollMenuTitle colspan=3>Valores para cada una de las monedas</TD>
					</TR>				
					<TR class=applicationTableHeader valign=center>
						<TD nowrap align="center">Moneda</TD>
						<TD nowrap align=right>Valor</TD>
            <TD width=90%>&nbsp;</TD>    
					</TR>
				</THEAD>				
				<%=gsExchangeRatesTable%>				
			</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap align=right>			
			<INPUT type="hidden" name=txtExchangeRatesArray>
			<INPUT class=cmdSubmit name=cmdSave type=button value='Guardar' style='width:85;' onclick='return(saveItems());'>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		</TD>
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>