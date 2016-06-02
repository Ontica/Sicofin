<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnItemId, gsCurrencyName, gsExchangeRateName, gsDate, gsExchangeRate, gbEdit
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oCurrenciesUS
		gbEdit = False
		gnItemId = 0
		Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")		
		gsCurrencyName = oCurrenciesUS.CboCurrenciesWithException(Session("sAppServer"), 1)
		gsExchangeRateName = oCurrenciesUS.CboExchangeRatesTypes(Session("sAppServer"))
		Set oCurrenciesUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oCurrenciesUS, oRecordset
		'****************************
		gbEdit = True		
		gnItemId = CLng(nItemId)		
		Set oCurrenciesUS	 = Server.CreateObject("AOCurrenciesUS.CServer")	
		Set oRecordset		 = oCurrenciesUS.GetExchangeRateRS(Session("sAppServer"), CLng(nItemId))		
		gsCurrencyName		 = oCurrenciesUS.CurrencyName(Session("sAppServer"), CLng(oRecordset("to_currency_id")), True)		
		gsExchangeRateName = oCurrenciesUS.ExchangeRateName(Session("sAppServer"), CLng(oRecordset("exchange_rate_type_id")))
		gsDate						 = oCurrenciesUS.FormatDate(oRecordset("from_date"))
		gsExchangeRate		 = oCurrenciesUS.FormatCurrency(oRecordset("exchange_rate"), 6)
		oRecordset.Close
		Set oCurrenciesUS = Nothing						
		Set oRecordset = Nothing		
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Tipos de cambio</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function isDate(sDate) {
	var obj;
	obj = RSExecute("../general_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../general_scripts.asp", "IsNumeric" , sNumber, nDecimals);
	return obj.return_value;
}

function validate() {
	var dDocument = window.document.all;
	
	if (!isNumeric(dDocument.txtExchangeRate.value, 6)) {
		alert("No reconozco el valor del tipo de cambio proporcionado.");
		dDocument.txtExchangeRate.focus();
		return false;
	}	
	if (Number(dDocument.txtExchangeRate.value) <= 0) {
		alert("El valor del tipo de cambio debe ser positivo.");
		dDocument.txtExchangeRate.focus();
		return false;
	}		
	return true;
}

function deleteItem() {
	if (confirm('¿Elimino el tipo de cambio?')) {		
		window.document.frmEditor.action = "./exec/delete_exchange_rate.asp?id=<%=gnItemId%>";		
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "./exec/save_exchange_rate.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();
	}
	return false;
}

function txtFromDate_onblur() {
	var dDocument = window.document.all;
	
	if ((dDocument.txtFromDate.value != '') && (dDocument.txtToDate == '')) {
		if (isDate(dDocument.txtFromDate.value)) {
			dDocument.txtToDate.value = dDocument.txtFromDate.value;
		}
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="window.document.all.txtExchangeRate.focus();">
<FORM name=frmEditor method="post">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Edición de tipos de cambio
		</TD>
		<TD align=right nowrap>
			<img align=absMiddle src="/empiria/images/help_red.gif" onclick='notAvailable();' alt="Ayuda">	
			<img align=absMiddle src="/empiria/images/invisible.gif">
			<img align=absMiddle src="/empiria/images/close_red.gif" onclick='window.close();' alt="Cerrar">
		</TD>
	</TR>
  <TR>
		<TD colspan=2 nowrap>
			<TABLE class=applicationTable cellpadding=1>
				<TR>
				  <TD valign=middle nowrap>Moneda:</TD>
				  <TD>
						&nbsp; <b><%=gsCurrencyName%></b>
					</TD>  
				</TR>
				<TR>
				  <TD valign=middle nowrap>Tipo de cambio:</TD>
				  <TD>
						&nbsp; <b><%=gsExchangeRateName%></b>
					</TD>
				</TR>
				<TR>
				  <TD valign=middle nowrap>Fecha:</TD>
				  <TD>
						&nbsp;<b><%=gsDate%></b>
					</TD>
				</TR>
				<TR>
				  <TD valign=middle nowrap>Valor:</TD>
				  <TD>
						$<INPUT name=txtExchangeRate value="<%=gsExchangeRate%>" style="width:120;height:20;"> (Moneda nacional)
					</TD>
				</TR>
				<TR>
				  <TD colspan=2 align=right nowrap>
						<% If gbEdit Then %>
						<INPUT class=cmdSubmit name=cmdDelete type=button value='Eliminar' style='width:65;' onclick="return(deleteItem());">&nbsp; &nbsp; &nbsp; &nbsp;				  
						<% End If %>
						<INPUT class=cmdSubmit name=cmdSave type=button value='Aceptar' style='width:65;' onclick='return(saveItem());'>&nbsp; &nbsp;
						<INPUT class=cmdSubmit name=cmdCancel type=button value='Cancelar' style='width:65;' onclick='window.close();'>
						&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
