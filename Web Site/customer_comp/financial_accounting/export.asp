<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim sTargetAppServer
	Dim dBalancesLastExpDate, dTransactionsLastExpDate
	Dim sDaysSinceLastBalancesExpDate, sDaysSinceLastTransactionsExpDate
	
	sTargetAppServer = "GemPyC"
	
	'Call Main()
	
	Sub Main()
		Dim oObject, dTemp
		'On Error Resume Next
		
		Set oObject = Server.CreateObject("SCFIGemPyC.CInterface")
		
		dBalancesLastExpDate = oObject.BalancesLastExportationDate(CStr(sTargetAppServer))
		If IsDate(dBalancesLastExpDate) Then
			dTemp	= Now() - CDate(dBalancesLastExpDate)
			sDaysSinceLastBalancesExpDate = FormatSinceDate(dTemp)
		Else
			sDaysSinceLastBalancesExpDate = "Indeterminado"
		End If		

		dTransactionsLastExpDate = oObject.TransactionsLastExportationDate(CStr(sTargetAppServer))
		If IsDate(dTransactionsLastExpDate) Then
			dTemp	= Now() - CDate(dTransactionsLastExpDate)
			sDaysSinceLastTransactionsExpDate = FormatSinceDate(dTemp)
		Else
			sDaysSinceLastTransactionsExpDate = "Indeterminado"
		End If
		
		Set oObject = Nothing
	End Sub
	
	Function FormatSinceDate(dDate)
		Dim sTemp
		If (Int(dDate) = 1) Then
			sTemp = "1 d�a, "
		Else
			sTemp = Int(dDate) & " d�as, "
		End If
		If (Hour(dDate) = 1) Then
			sTemp = sTemp & "1 hora, "
		Else
			sTemp = sTemp & Hour(dDate) & " horas, "
		End If
		If (Minute(dDate) = 1) Then
			sTemp = sTemp & "1 minuto. "
		Else
			sTemp = sTemp & Minute(dDate) & " minutos."
		End If		
		FormatSinceDate = sTemp
	End Function
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var gbSended = false;

function isDate(sDate) {
	var obj;
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function cmdExportOrabanks_onclick() {
  if (gbSended) {
		return false;
  }
	if (document.all.txtOrabanksDate.value == '') {
		alert("Necesito la fecha de elaboraci�n de las p�lizas a exportar.");
		document.all.txtOrabanksDate.focus();
		return false;
	}
	if (!isDate(document.all.txtOrabanksDate.value)) {
		alert("No reconozco la fecha de elaboraci�n proporcionada.");
		document.all.txtOrabanksDate.focus();
		return false;
	}
	if (document.all.txtFromHour.value == '') {
		alert("Necesito la hora de inicio de elaboraci�n de las p�lizas a exportar.");
		document.all.txtFromHour.focus();
		return false;
	}		
	if (!isDate(document.all.txtFromHour.value)) {
		alert("No reconozco la hora de inicio.");
		document.all.txtFromHour.focus();
		return false;
	}
	if (document.all.txtToHour.value == '') {
		alert("Necesito la hora de t�rmino de elaboraci�n de las p�lizas a exportar.");
		document.all.txtToHour.focus();
		return false;
	}	
	if (!isDate(document.all.txtToHour.value)) {
		alert("No reconozco la hora de t�rmino.");
		document.all.txtToHour.focus();
		return false;
	}
  if (!confirm('�Contin�o con la exportaci�n de las p�lizas al sistema Ora*banks?')) {
		return false;
	}	
	gbSended = true;
	document.frmSend.action = './exec/export_orabanks.asp';
	document.frmSend.submit();
	return true;	
}

function cmdExportPyCTransactions_onclick() {
  if (gbSended) {
		return false;
  }
  if (!confirm('�Contin�o con la exportaci�n de las p�lizas al sistema PyC?')) {
		return false;
	}
  gbSended = true;
	document.frmSend.action = './exec/export_others.asp?id=1';
	document.frmSend.submit();
	return true;	
}

function cmdExportPyCBalances_onclick() {
  if (gbSended) {
		return false;
  }
  if (!confirm('�Contin�o con la exportaci�n de los saldos al sistema PyC?')) {
		return false;
	}
  gbSended = true;
	document.frmSend.action = './exec/export_others.asp?id=2';
	document.frmSend.submit();
	return true;	
}

function cmdExportSigro_onclick() {
  if (gbSended) {
		return false;
  }
	if (document.all.txtSigroDate.value == '') {
		alert("Necesito la fecha del saldo para los reportes regulatorios.");
		document.all.txtSigroDate.focus();
		return false;
	}
	if (!isDate(document.all.txtSigroDate.value)) {
		alert("No reconozco la fecha del saldo para los reportes regulatorios.");
		document.all.txtSigroDate.focus();
		return false;
	}
  if (gbSended) {
		return false;
  }
  if (!confirm('�Contin�o con la exportaci�n de los saldos al Sigro?')) {
		return false;
	}	
  gbSended = true;
	document.frmSend.action = './exec/export_others.asp?id=3';
	document.frmSend.submit();
	return true;	
}

function cmdExportBalances_onclick() {
  if (gbSended) {
		return false;
  }
	if (document.all.txtFromDate.value == '') {
		alert("Necesito la fecha de inicio para la exportaci�n de saldos.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (!isDate(document.all.txtFromDate.value)) {
		alert("No reconozco la fecha de inicio para la exportaci�n de saldos.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (document.all.txtToDate.value == '') {
		alert("Necesito la fecha de t�rmino para la exportaci�n de saldos.");
		document.all.txtToDate.focus();
		return false;
	}
	if (!isDate(document.all.txtToDate.value)) {
		alert("No reconozco la fecha de t�rmino para la exportaci�n de saldos.");
		document.all.txtToDate.focus();
		return false;
	}
  if (!confirm('�Contin�o con la exportaci�n de los saldos a la tabla de Oracle�?')) {
		return false;
	}		
  gbSended = true;
	document.frmSend.action = './exec/export_others.asp?id=4';
	document.frmSend.submit();
	return true;	
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<FORM name=frmSend method="post">
<TABLE align=center border=0 cellPadding=3 cellSpacing=3 width="95%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>a) Exportaci�n de p�lizas en formato Ora*banks</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD>
			<FONT color=maroon>
				Crea los archivos de exportaci�n con las p�lizas de un d�a para cada una de las contabilidades.<br>
				Estos archivos pueden ser le�dos por el sistema Ora*banks.
			</FONT>
		</TD>
	</TR>	
	<TR bgcolor=khaki>
	  <TD nowrap>
			Crear los archivos de exportaci�n de todas las p�lizas <b>elaboradas</b> el d�a: &nbsp;&nbsp;
	  <INPUT name="txtOrabanksDate" style="width:95px" value="<%=Date()%>">&nbsp;&nbsp;(d�a / mes / a�o)
	  <br><br>
			<b>De las:</b>&nbsp;&nbsp;<INPUT name="txtFromHour" maxlength=8 style="width:90px" value="00:00:00"> &nbsp;&nbsp;(hr:min:seg)
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<b>a las:</b>&nbsp;&nbsp;<INPUT name="txtToHour" maxlength=8 style="width:90px" value="23:59:59"> &nbsp;&nbsp;(hr:min:seg)
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdExportOrabanks value="Exportar p�lizas" style="WIDTH: 140px" LANGUAGE=javascript onclick="return cmdExportOrabanks_onclick()">
	  </TD>
	</TR>
	<TR>
		<TD nowrap><br><IMG src='/empiria/images/pleca.gif' width=100% height=1px></IMG></TD>
	</TR>
	<TR> 
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>b) Exportaci�n de p�lizas hacia el sistema PyC</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD>
			<FONT color=maroon>
				Exporta todas las p�lizas que se encuentran en el diario del sistema de contabilidad financiera y que a�n no han sido
				exportadas<br>
				al Sistema de Presupuestos y Costos.<br>
			</FONT>
		</TD>
	</TR>
	<TR bgcolor=khaki>
	  <TD nowrap>
	  Fecha de elaboraci�n de la �ltima p�liza exportada: &nbsp;&nbsp;<b><%=dTransactionsLastExpDate%></b><br>
	  Tiempo transcurrido desde la �ltima exportaci�n: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=sDaysSinceLastTransactionsExpDate%></b>
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <INPUT type="button" name=cmdExportPyCTransactions value="Exportar p�lizas PyC" style="WIDTH: 140px" LANGUAGE=javascript onclick="return cmdExportPyCTransactions_onclick()">
	  </TD>
	</TR>
	<TR>
		<TD nowrap><br><IMG src='/empiria/images/pleca.gif' width=100% height=1px></IMG></TD>
	</TR>
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>c) Exportaci�n de saldos hacia el sistema PyC</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD>
			<FONT color=maroon>
				Actualiza la tabla de saldos del Sistema de Presupuestos y Costos con los saldos del sistema de contabilidad financiera.<br>
			</FONT>
		</TD>
	</TR>
	<TR bgcolor=khaki>
	  <TD nowrap>
	  Fecha de la �ltima exportaci�n de saldos hacia el PyC: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=dBalancesLastExpDate%></b><br>
	  Tiempo transcurrido desde la �ltima exportaci�n de saldos: &nbsp;<b><%=sDaysSinceLastBalancesExpDate%></b>
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <INPUT type="button" name=cmdExportPyCBalances value="Exportar saldos PyC" style="WIDTH: 140px" LANGUAGE=javascript onclick="return cmdExportPyCBalances_onclick()">
	  </TD>
	</TR>
	<TR>
		<TD nowrap><br><IMG src='/empiria/images/pleca.gif' width=100% height=1px></IMG></TD>
	</TR>	
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>d) Exportaci�n de saldos para el Sigro (reportes regulatorios)</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD>
			<FONT color=maroon>
				Crea la tabla de saldos empleada por el sistema de reportes regulatorios.<br>
			</FONT>
		</TD>
	</TR>
	<TR bgcolor=khaki>
	  <TD nowrap>
		Crear la tabla de saldos para el Sigro <b>al d�a</b>: &nbsp;
	  <INPUT name="txtSigroDate" style="width:95px">&nbsp;&nbsp;(d�a / mes / a�o)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	  <INPUT type="button" name=cmdExportSigro value="Exportar a Sigro" style="WIDTH: 140px" LANGUAGE=javascript onclick="return cmdExportSigro_onclick()">
	  </TD>
	</TR>
	<TR>
		<TD nowrap><br><IMG src='/empiria/images/pleca.gif' width=100% height=1px></IMG></TD>
	</TR>
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>e) Exportaci�n de saldos. (Para uso exclusivo de la Subdirecci�n de Inform�tica)</STRONG></FONT></TD>
	</TR>
	<TR>
		<TD>
			<FONT color=maroon>
				Crea la tabla "ZSaldos" en Oracle con las balanzas seleccionadas y cuyos saldos est�n entre las fechas proporcionadas.<br>
			</FONT>
		</TD>
	</TR>	
	<TR bgcolor=khaki>
	  <TD nowrap>
		Generar los saldos <b>para</b>:&nbsp;&nbsp;&nbsp;
		<SELECT name=cboGralLedgers style="WIDTH: 240px">
			<OPTION value=0>Todas las contabilidades</OPTION>
			<OPTION value=1>Contabilidad bancaria</OPTION>
			<OPTION value=2>Contabilidad fiduciaria</OPTION>		</SELECT><br><br>
		<b>Del d�a:</b> &nbsp;
	  <INPUT name="txtFromDate" style="width:95px">&nbsp;&nbsp;(d�a / mes / a�o)&nbsp;&nbsp;&nbsp;&nbsp;
	  <b>al d�a:</b>&nbsp;
	  <INPUT name="txtToDate" style="width:95px">&nbsp;&nbsp;(d�a / mes / a�o) &nbsp;&nbsp;&nbsp;&nbsp;
	  <INPUT type="button" name=cmdExportBalances value="Exportar saldos" style="WIDTH: 140px" LANGUAGE=javascript onclick="return cmdExportBalances_onclick()">
	  </TD>
	  <INPUT type=hidden name=txtTargetAppServer value="<%=sTargetAppServer%>">
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>