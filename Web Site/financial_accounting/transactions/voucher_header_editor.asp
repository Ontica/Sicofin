<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsDescription, gsCboVoucherTypes, gsCboApplicationDates, gsCboSources, gsActionPage
	Dim gsGralLedgerName, gsAlternativeDate
	
	Call Main()
			 
	Sub Main()
		Dim oVoucherUS, oRecordset
		'*************************
		On Error Resume Next
		Set oVoucherUS   = Server.CreateObject("AOGLVoucherUS.CVoucher")				
		Set oRecordset   = oVoucherUS.GetTransactionRS(Session("sAppServer"), CLng(Request.QueryString("id")))
		gsDescription    = oRecordset("concepto_transaccion")
		gsGralLedgerName = oRecordset("nombre_mayor")
		
		gsCboVoucherTypes = oVoucherUS.CboVouchersTypes(Session("sAppServer"), CLng(oRecordset("id_tipo_poliza")))
		gsCboSources = oVoucherUS.CboSources(Session("sAppServer"), CLng(oRecordset("id_fuente")))
		gsCboApplicationDates = oVoucherUS.CboOpenPeriodsDates(Session("sAppServer"), CLng(oRecordset("id_mayor")), oRecordset("fecha_afectacion"))
		If oVoucherUS.IsDateInGLPeriod(Session("sAppServer"), CLng(oRecordset("id_mayor")), CDate(oRecordset("fecha_afectacion"))) Then
			gsAlternativeDate = ""
		Else
			gsAlternativeDate = oVoucherUS.FormatDate(CDate(oRecordset("fecha_afectacion")))
		End If	  
		Set oRecordset = Nothing
		gsActionPage = "exec/save_transaction.asp?id=" & Request.QueryString("id")
			
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		  
	End Sub
			
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Encabezado de la póliza</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function validateDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function doSubmit() {
	var oTxtDescription = document.all.txtDescription;
	var sMsg;
	if (oTxtDescription.value == "") {
		alert("Necesito el concepto de la póliza.");
		oTxtDescription.focus();
		return false;
	}	
	if (document.all.cboApplicationDates.value == '' && document.all.txtAlternativeDate.value == '') {
		alert("No hay períodos abiertos y la póliza no tiene una fecha valor o adelantada.");
		return false;
	}
	if (!validateDate(document.all.cboApplicationDates.value)) {
		alert("No reconzco la fecha de la póliza. Debe existir un problema con el manejo de períodos.");
		return false;
	}
	document.all.frmSend.submit();
	return true;
}

function window_onload() {
	if (document.all.txtAlternativeDate.value == '') {
		document.all.cboApplicationDates.disabled = false;
		document.all.divAlternativeDate.innerText = 'Ninguna';
		document.all.divAlternativeDateText.innerText = 'Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.';	
	} else {
		document.all.cboApplicationDates.disabled = true;
		document.all.divAlternativeDate.innerText = document.all.txtAlternativeDate.value;
		document.all.divAlternativeDateText.innerText = 'Haga clic en esta liga para anular la fecha valor o adelantada.';		
	}
	document.all.txtDescription.focus();
}


function cmdCheckSpelling_onclick() {
	alert("Por el momento esta opción no está disponible.\n\nGracias.");
}

function cmdGenerateConcept_onclick() {
	alert("Por el momento esta opción no está disponible.\n\nGracias.");
}

function showApplicationDatePicker() {	
	var sDate = '';
	var sOptions = 'dialogHeight:250px;dialogWidth:350px;resizable:no;scroll:no;status:no;help:no;';
		
	if (document.all.txtAlternativeDate.value == '') {
		sDate = window.showModalDialog('voucher_date_picker.asp', "" , sOptions);	
		if (sDate != '') {
			document.all.txtAlternativeDate.value = sDate;
			document.all.cboApplicationDates.disabled = true;
			document.all.divAlternativeDate.innerText = document.all.txtAlternativeDate.value;
			document.all.divAlternativeDateText.innerText = 'Haga clic en esta liga para anular la fecha valor o adelantada.';
		} else {
			document.all.txtAlternativeDate.value = '';
			document.all.cboApplicationDates.disabled = false;
			document.all.divAlternativeDate.innerText = 'Ninguna';
			document.all.divAlternativeDateText.innerText = 'Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.';
		}
		return false;
	}
	if (document.all.txtAlternativeDate.value != '') {
		document.all.txtAlternativeDate.value = '';
		document.all.cboApplicationDates.disabled = false;
		document.all.divAlternativeDate.innerText = 'Ninguna';
		document.all.divAlternativeDateText.innerText = 'Haga clic en esta liga si la póliza tiene fecha valor o es adelantada.';	
		return false;
	}
}

function cmdCancel_onclick() {
	window.location.href = 'voucher_editor.asp?id=<%=Request.QueryString("id")%>';
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload="return window_onload()">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Encabezado de la póliza
		</TD>
		<TD colspan=3 align=right nowrap>
			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.close();" alt="Cerrar">								</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<FORM name=frmSend action='<%=gsActionPage%>' method=post>
			<TABLE class=applicationTable>	
				<TR>
					<TD valign=top nowrap>Tipo de póliza:</TD>
				  <TD colspan=3>			
						<SELECT name=cboVoucherTypes style="WIDTH: 380px">
							<%=gsCboVoucherTypes%>
						</SELECT>			
				  </TD>
				</TR>
				<TR>
				  <TD valign=top nowrap>Concepto:</TD>
				  <TD colspan=3>
						<TEXTAREA name=txtDescription ROWS=3 style="WIDTH: 380px"><%=gsDescription%></TEXTAREA><br>
						<INPUT type=button class=cmdSubmit name=cmdGenerateConcept value="Sugerir el concepto" onclick="return cmdGenerateConcept_onclick()">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<INPUT type=button class=cmdSubmit name=cmdCheckSpelling value="Revisar ortografía" onclick="return cmdCheckSpelling_onclick()">
				  </TD>
				</TR>
				<TR>
				  <TD>Contabilidad:</TD>
				  <TD colspan=3>
						<b><%=gsGralLedgerName%></b>
				  </TD>
				</TR>  	
				<TR>
				  <TD>Origen de la transacción:</TD>
				  <TD colspan=3>
						<span id=divCboSources>
							<SELECT name=cboSources style="WIDTH: 380px"> 
								<%=gsCboSources%>
							</SELECT>
				   </span>
				  </TD>
				</TR>  
				<TR>
				  <TD valign=top>Fecha de afectación:</TD>
				  <TD colspan=3>
						<SELECT name='cboApplicationDates' style='WIDTH: 130px'>
							<%=gsCboApplicationDates%>
						</SELECT><br>
						<a href=''onclick='showApplicationDatePicker();return false;'>
							<span id=divAlternativeDateText>
								Haga clic en esta liga si la póliza tiene fecha valor o es adelantada
							</span>
						</a>		
					</TD>
				</TR>
				<TR>
				  <TD valign=top>
						Fecha valor o adelantada:
					</TD>
					<TD colspan=3>
						<b><span id=divAlternativeDate>Ninguna</span></b>
					</TD>
				</TR>  			
				<TR>
				  <td colspan=4 align=right>
				  <INPUT type="hidden" name=txtAlternativeDate value='<%=gsAlternativeDate%>'>
				   <INPUT class=cmdSubmit name=cmdSend type=button value="Guardar cambios" style='width:100' onclick='doSubmit();'>
				   &nbsp;&nbsp;&nbsp;&nbsp;
				   <INPUT class=cmdSubmit name=cmdCancel type=button value="Cancelar" style='width:100' onclick="window.close();">
				  </td>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
