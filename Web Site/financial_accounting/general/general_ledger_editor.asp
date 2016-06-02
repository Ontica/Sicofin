<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
		
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
				
	Dim gnItemId, gsTitle, gsGralLedgerName, gnStdAccountTypeId, gsGralLedgerNumber, gsSubsAccountsPrefix
	Dim gsCboVouchersGroup, gsCboReportsGroup, gsCboCalendars, gsCboCurrencies
	
  If CLng(Request.QueryString("id")) = 0 Then
		Call AddItem()
	Else
		Call EditItem(CLng(Request.QueryString("id")))
	End If

	Sub AddItem()
		Dim oVoucherUS,  oGralLedgerUS
		'****************
		gsTitle = "Nueva contabilidad"
		gnItemId = 0		
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")		
		gsCboVouchersGroup = oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), 0, 1)
		gsCboReportsGroup = oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), 0, 2)
		Set oVoucherUS = Nothing
		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		gsCboCurrencies = oGralLedgerUS.CboCurrencies(Session("sAppServer"), 1)		
		gsCboCalendars = oGralLedgerUS.CboCalendars(Session("sAppServer"))
		Set oGralLedgerUS = Nothing
	End Sub
	
	Sub EditItem(nItemId)
		Dim oVoucherUS, oGralLedgerUS, oRecordset, nGroupId
		'**************************************************
		gsTitle = "Edición de la contabilidad"
		gnItemId = CLng(nItemId)
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")
		Set oRecordset = oGralLedgerUS.GetGralLedgerRS(Session("sAppServer"), CLng(gnItemId))			
		gsCboCurrencies = oGralLedgerUS.CboCurrencies(Session("sAppServer"), CLng(oRecordset("id_moneda_base")))
		gsCboCalendars = oGralLedgerUS.CboCalendars(Session("sAppServer"), CLng(oRecordset("id_calendario")))
		gsSubsAccountsPrefix = 	oRecordset("prefijo_cuentas_auxiliares")
		gnStdAccountTypeId = oRecordset("id_tipo_cuentas_std")
		gsGralLedgerName   = oRecordset("nombre_mayor")
		gsGralLedgerNumber = oRecordset("numero_mayor")		
		oRecordset.Close		
		Set oGralLedgerUS = Nothing
		
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		nGroupId = oVoucherUS.GralLedgerGroupId(Session("sAppServer"), Clng(nItemId), 1)
		gsCboVouchersGroup = oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), 0, 1, CLng(nGroupId))
		nGroupId = oVoucherUS.GralLedgerGroupId(Session("sAppServer"), Clng(nItemId), 2)
		gsCboReportsGroup = oVoucherUS.CboGeneralLedgerCategories(Session("sAppServer"), 0, 2, CLng(nGroupId))
		Set oVoucherUS = Nothing
	End Sub
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function canDelete() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "CanDeleteGL", <%=gnItemId%>);
	return obj.return_value;	
}

function isRootGroup(nGroup) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "GLGroupIsRoot", nGroup);
	return obj.return_value;	
}

function GLGroupStdAccountId(nGroup) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "GLGroupStdAccountId", nGroup);
	return obj.return_value;	
}


function cmdDelete_onclick() {
	if (!canDelete()) {
		alert("La contabilidad no puede ser eliminada debido a que ya tiene movimientos.");
		return false;
	}
	if (confirm('¿Elimino la contabilidad (<%=gsGralLedgerNumber%>) <%=gsGralLedgerName%>?')) {
		window.document.frmEditor.action = "exec/delete_general_ledger.asp?id=<%=gnItemId%>";
		window.document.frmEditor.submit();
	}
}

function saveItem() {
	if (validate()) {
		window.document.frmEditor.action = "exec/save_general_ledger.asp";
		window.document.frmEditor.submit();
	}
	return false;
}

function validate() {
	var dDocument = window.document.all;
	var sTemp;
	
	if (dDocument.txtGralLedgerName.value == '') {
		alert("Requiero el nombre de la contabilidad.");
		dDocument.txtGralLedgerName.focus();
		return false;
	}	
	if (dDocument.txtGralLedgerNumber.value == '') {
		alert("Requiero el número de contabilidad.");
		dDocument.txtGralLedgerNumber.focus();
		return false;
	}	
  if (dDocument.txtSubsAccountsPrefix.value == '') {
		alert("Requiero el prefijo para las cuentas auxiliares.");
		dDocument.txtSubsAccountsPrefix.focus();
		return false;
	}
  if (dDocument.cboVouchersGroups.value == 0) {
		alert("Requiero la selección del grupo de contabilidades para la captura de pólizas.");
		dDocument.cboVouchersGroups.focus();
		return false;
	}
  if (dDocument.cboReportGroups.value == 0) {
		alert("Requiero la selección del grupo de contabilidades para la generación de reportes.");
		dDocument.cboReportGroups.focus();
		return false;
	}
	if (isRootGroup(dDocument.cboVouchersGroups.value)) {
		alert("El grupo seleccionado para la captura de pólizas no permite la incorporación de contabilidades.");
		dDocument.cboVouchersGroups.focus();
		return false;
	}
	if (isRootGroup(dDocument.cboReportGroups.value)) {
		alert("El grupo seleccionado para la generación de reportes no permite la incorporación de contabilidades.");		
		dDocument.cboReportGroups.focus();
		return false;				
	}	
	nBaseStdAccountId = GLGroupStdAccountId(dDocument.cboVouchersGroups.value);
	if (nBaseStdAccountId != GLGroupStdAccountId(dDocument.cboReportGroups.value)) {
		alert("Los grupos de contabilidades seleccionados manejan catálogos de cuenta distintos entre sí.");
		dDocument.cboVouchersGroups.focus();
		return false;
	}
	<% If (gnItemId <> 0) Then %>
		if(nBaseStdAccountId != <%=gnStdAccountTypeId%>) {		
			sTemp = "Los grupos de contabilidades seleccionados manejan un catálogo de cuentas distinto al de esta contabilidad.";
			alert(sTemp);	
			dDocument.cboVouchersGroups.focus();
			return false;
		}
	<% End If %>
	return true;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmEditor action="" method="post">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			<%=gsTitle%>
		</TD>
		<TD colspan=2 align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>						<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.close();" alt="Cerrar">		</TD>
	</TR>
	<TR>
		<TD colspan=3> 
			<TABLE class=applicationTable>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2><b>Información general</b></TD>
				</TR>
				<TR>
					<TD>Número de contabilidad:</TD>
					<TD><INPUT name=txtGralLedgerNumber value="<%=gsGralLedgerNumber%>" maxlength=6 style="HEIGHT: 22px; WIDTH: 30%"></TD>  
				</TR>		
				<TR>
					<TD valign=top>Nombre:</TD>
					<TD valign=top width=100%><TEXTAREA name=txtGralLedgerName rows=3 style="WIDTH: 100%"><%=gsGralLedgerName%></TEXTAREA></TD>
				</TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2><b>Parametrización referente a su funcionamiento</b></TD>
				</TR>
				<TR>
					<TD valign=top>Prefijo para sus auxiliares:</TD>
					<TD><INPUT name=txtSubsAccountsPrefix value="<%=gsSubsAccountsPrefix%>" maxlength=4 style="HEIGHT: 22px; WIDTH: 30%"></TD>					
				</TR>				
				<TR>
				  <TD>Moneda base para sus movimientos:</TD>
					<TD>
						<SELECT name=cboCurrencies style="HEIGHT: 22px; WIDTH: 100%">							
							<%=gsCboCurrencies%>
						</SELECT>
					</TD>
				</TR>
				<TR>
				  <TD nowrap>Calendario que utiliza:</TD>
					<TD>
						<SELECT name=cboCalendars style="HEIGHT: 22px; WIDTH: 100%">
							<%=gsCboCalendars%>
						</SELECT>
					</TD>
				</TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2><b>Grupos a los que pertenece<b></TD>
				</TR>
				<TR>
					<TD nowrap>Para la captura de pólizas:</TD>
					<TD>		
						<SELECT name=cboVouchersGroups style="HEIGHT: 22px; WIDTH: 100%">
							<OPTION value=0>(Seleccionar grupo de contabilidades)</OPTION>
							<%=gsCboVouchersGroup%>
						</SELECT>
						&nbsp;
					</TD>
				</TR>
				<TR>
					<TD nowrap>Para la generación de reportes:</TD>
					<TD>
						<SELECT name=cboReportGroups style="HEIGHT: 22px; WIDTH: 100%">
							<OPTION value=0>(Seleccionar grupo de contabilidades)</OPTION>
							<%=gsCboReportsGroup%>
						</SELECT>
						<br>&nbsp;
					</TD>		
				</TR>
				<!--
				<TR class=applicationTableRowDivisor>
					<TD colspan=2><b>Mayores auxiliares que emplea<b></TD>
				</TR>
				<TR>
					<TD nowrap>Para la generación de reportes:</TD>
					<TD>
						<SELECT name=cboCategories style="HEIGHT: 22px; WIDTH: 100%">
							<OPTION value=0>(Seleccionar grupo de contabilidades)</OPTION>
							<%=gsCboVouchersGroup%>
						</SELECT>
					</TD>		
				</TR>
				-->				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD align=right colspan=3>
			<INPUT name=txtItemId type="hidden" value="<%=gnItemId%>">
			<% If (gnItemId <> 0) Then %>
			<INPUT class=cmdSubmit name=cmdDelete type=button value='Eliminar' onclick='return cmdDelete_onclick()' style='width:75;'>
			&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
			<% End If %>
			<INPUT class=cmdSubmit name=cmdEditItem type=button value='Aceptar' onclick='return saveItem()' style='width:75;'>
			&nbsp; &nbsp;
			<INPUT class=cmdSubmit name=cmdCancel type=button value='Cancelar' onclick='window.close();' style='width:75;'>
			&nbsp; &nbsp;
		</TD>
	</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
