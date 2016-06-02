<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gnItemId, gnSubsidiaryLedgerId, gsSubsidiaryAccountNumber, gsSubsidiaryAccountName, gsSubsidiaryAccountDescription
	Dim gsCboStdAccountTypes, gsCboStdAccountNature, gsSubsidiaryAccountPrefix, gnSubsLedgerType, gsCboSubsidiaryLedgers
	Dim gsSubsidiaryAccountExtAttrs
	
  If (Len(Request.QueryString("id")) = 0) Then
		Call AddItem(CLng(Request.QueryString("gralLedger")), CLng(Request.QueryString("subsLedgerId")))
	Else
		Call EditItem(CLng(Request.QueryString("gralLedger")), CLng(Request.QueryString("subsLedgerId")), CLng(Request.QueryString("id")))		
	End If

	Sub AddItem(nGralLedger, nSubsidiaryLedgerId)
		Dim oGralLedger, oRecordset, sTemp
		'*********************************
		gbEdit = False
		gsTitle = "Agregar cuenta auxiliar"
		gnItemId = 0
		gnSubsidiaryLedgerId = nSubsidiaryLedgerId		
		Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
		Call SetSubsidiaryLedgerType(nGralLedger, gnSubsidiaryLedgerId)
		sTemp = oGralLedger.GetNextSubsidiaryAccountNumber(Session("sAppServer"), CLng(gnSubsidiaryLedgerId))
		Set oRecordset = oGralLedger.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(gnSubsidiaryLedgerId))		
		Call SetExtendedAttributes(gnSubsLedgerType, 0)
		gsSubsidiaryAccountPrefix = oRecordset("prefijo_cuentas_auxiliares")
		gsSubsidiaryAccountNumber = Right(sTemp, 16)
		oRecordset.Close		
		Set oRecordset = Nothing				
		Set oGralLedger = Nothing
	End Sub
	
	Sub EditItem(nGralLedger, nSubsidiaryLedgerId, nItemId)
		Dim oGralLedger, oRecordset
		'***************************************
		gbEdit = True
		gsTitle = "Editar cuenta auxiliar"
		gnItemId = CLng(nItemId)
		gnSubsidiaryLedgerId = nSubsidiaryLedgerId
		Call SetSubsidiaryLedgerType(nGralLedger, gnSubsidiaryLedgerId)
		Set oGralLedger = Server.CreateObject("AOGralLedgerUS.CServer")
		Set oRecordset = oGralLedger.GetSubsidiaryAccountRS(Session("sAppServer"), CLng(nItemId))				
		Call SetExtendedAttributes(gnSubsLedgerType, nItemId)
		Set oGralLedger = Nothing
		gsSubsidiaryAccountNumber			 = oRecordset("numero_cuenta_auxiliar")
		gsSubsidiaryAccountName				 = oRecordset("nombre_cuenta_auxiliar")
		gsSubsidiaryAccountDescription = oRecordset("descripcion")		
		oRecordset.Close
		Set oRecordset = Nothing		
	End Sub
	
	Sub SetSubsidiaryLedgerType(nGralLedgerId, nSubsidiaryLedgerId)
		Dim oGralLedgerUS, oRecordset
		'************************************************************
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")		
		gsCboSubsidiaryLedgers  = oGralLedgerUS.CboSubsidiaryLedgers(Session("sAppServer"), CLng(nGralLedgerId), CLng(nSubsidiaryLedgerId))
		Set oRecordset = oGralLedgerUS.GetSubsidiaryLedgerRS(Session("sAppServer"), CLng(nSubsidiaryLedgerId))
		gnSubsLedgerType = oRecordset("id_tipo_mayor_auxiliar")
		Set oGralLedgerUS = Nothing		
	End Sub

	Sub SetExtendedAttributes(nSubsidiaryLedgerType, nSubsAccountId)
		Dim oGralLedger
		'*************************************************************
		Set oGralLedger = Server.CreateObject("AOGralLedgerUS.CServer")
		gsSubsidiaryAccountExtAttrs = oGralLedger.SubsidiaryAccountExtendedAttrs(Session("sAppServer"), _
																																						 CLng(nSubsidiaryLedgerType), _
																																						 CLng(nSubsAccountId))		
		Set oGralLedger = Nothing	
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

var bSubsLedgerTypeChanged = false;

function isDate(sDate) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

function isNumeric(sNumber, nDecimals) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "IsNumeric" , sNumber, nDecimals);
	return obj.return_value;
}

function isSubsidiaryAccountNumberValid(sSubsAccountNumber) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "IsSubsAccountNumberValid", sSubsAccountNumber);
	return obj.return_value;
}

function existsSubsidiaryAccountNumber(sSubsidiaryAccountNumber) {
	var obj, nSubsidiaryLedgerId;
	
	nSubsidiaryLedgerId = document.all.cboSubsidiaryLedgers.value;	
	obj = RSExecute("../financial_accounting_scripts.asp","ExistsSubsidiaryAccountNumber", nSubsidiaryLedgerId, sSubsidiaryAccountNumber);	
	return obj.return_value;
}

function getExtendedAttributes() {
	var obj, nSubsLedgerId;
	  
	nSubsLedgerId = document.all.cboSubsidiaryLedgers.value;	
	obj = RSExecute("../financial_accounting_scripts.asp","SubsAccountExtendedAttrs", nSubsLedgerId, 0);
	return obj.return_value;
}

function getNextNumber() {
	var obj, nSubsidiaryLedgerId;
	
	nSubsidiaryLedgerId = document.all.cboSubsidiaryLedgers.value;	
	obj = RSExecute("../financial_accounting_scripts.asp", "GetNextSubsAccountNumber", nSubsidiaryLedgerId);
	if (obj.return_value != '') {
		document.all.txtSubsidiaryAccountNumber.value = obj.return_value;
	} else {
		alert("Debido a algún problema, no pude ejecutar la operación solicitada.");
	}	
}

function validate() {
	var dDocument = window.document.all;
	<% If Not gbEdit Then %>
	if (dDocument.txtSubsidiaryAccountNumber.value == '') {
		alert("Requiero el número de auxiliar.");
		dDocument.txtSubsidiaryAccountNumber.focus();
		return false;
	}	
	if (!isSubsidiaryAccountNumberValid(dDocument.txtSubsidiaryAccountNumber.value)) {
		alert("No reconozco el formato del auxiliar proporcionado.");
		dDocument.txtSubsidiaryAccountNumber.focus();
		return false;
	}	
	if (existsSubsidiaryAccountNumber('<%=gsSubsidiaryAccountPrefix%>' + dDocument.txtSubsidiaryAccountNumber.value)) {
		alert("El auxiliar proporcionado ya fue dado de alta.");
		dDocument.txtSubsidiaryAccountNumber.focus();
		return false;
	}	
	<% End If %>
	if (dDocument.txtSubsidiaryAccountName.value == '') {
		alert("Requiero el nombre del auxiliar.");
		dDocument.txtSubsidiaryAccountName.focus();
		return false;
	}	
  return true;
}

function validateDate(oControl) {
	if (oControl.value != '') {
		if (!isDate(oControl.value)) {
			alert("No reconozco la fecha proporcionada.");
			oControl.value = '';
			oControl.focus();
			return false;
		}
	}
	return true;
}

function validateNumeric(oControl,nDecimals) {
	if (oControl.value != '') {
		if (!isNumeric(oControl.value, nDecimals)) {
			alert("No reconozco el valor proporcionado.");
			oControl.value = '';
			oControl.focus();			
			return false;
		}
	}
	return true;
}


function txtSubsidiaryAccountNumber_onblur() {
	var obj, nSubsidiaryLedgerId;
	gnAccountId = 0;
	nSubsidiaryLedgerId = document.all.cboSubsidiaryLedgers.value;	
	if (document.all.txtSubsidiaryAccountNumber.value != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatSubsidiaryAccountWithSLId", nSubsidiaryLedgerId, document.all.txtSubsidiaryAccountNumber.value);
		if (obj.return_value != '') {
			document.all.txtSubsidiaryAccountNumber.value = obj.return_value;			
		} else {
			alert("No reconozco el formato del auxiliar proporcionado.");
		}
	}
	return true;
}

function saveItem() {
	if (!validate()) {
		return false;
	}
	document.all.frmEditor.action = "exec/save_subsidiary_account.asp";
	document.all.frmEditor.submit();
	return false;
}

function deleteItem() {
	if (confirm('¿Elimino el auxiliar "<%=gsSubsidiaryAccountName%>"?')) {
		document.all.frmEditor.action = "./exec/delete_subsidiary_account.asp?id=<%=gnItemId%>";
		document.all.frmEditor.submit();		
	}
}

function cboSubsidiaryLedgers_onchange() {
	var sMsg;	
	<% If (gnItemId <> 0) Then %>	
	if (!bSubsLedgerTypeChanged) {
		sMsg  = 'Al modificar el tipo de auxiliar se perderá toda la información complementaria del auxiliar.\n\n';
		sMsg += '¿Continúo con el cambio del tipo de auxiliar?';
		if (!confirm(sMsg)) {
			document.all.cboSubsidiaryLedgers.value = <%=gnSubsidiaryLedgerId%>;	
			return false;
		}	
		bSubsLedgerTypeChanged = true;
	}
	<% End If %>
	sMsg  = '<TABLE class=applicationTable height=101%>' + getExtendedAttributes();
	sMsg +=	'<TR><TD colspan=4 valign=top nowrap width=100% height=100%></TD></TR></TABLE>';	
	document.all.divExtendedAttrs.innerHTML = sMsg;
	
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
		</TD>		<TD colspan=3 align=right nowrap>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">								</TD>
	</TR>
	<TR>
	  <TD colspan=4>
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD colspan=4>
						Información general
					</TD>
				</TR>
				<TR>
					<TD nowrap>Tipo cuenta auxiliar:</TD>
					<TD>
						<SELECT name=cboSubsidiaryLedgers style="width:280" onchange="return cboSubsidiaryLedgers_onchange();">
							<%=gsCboSubsidiaryLedgers%>						</SELECT>
					</TD>
				</TR>
				<TR>
					<TD nowrap>Cuenta auxiliar:</TD>
					<% If gbEdit Then %>
					<TD colspan=3 width=100%><b><%=gsSubsidiaryAccountNumber%></b></TD>
					<% Else %>
					<TD colspan=3 width=100% nowrap>
						<INPUT name=txtSubsidiaryAccountPrefix value="<%=gsSubsidiaryAccountPrefix%>" style="HEIGHT: 22px; WIDTH:40px" readonly>
						<INPUT name=txtSubsidiaryAccountNumber value="<%=gsSubsidiaryAccountNumber%>" maxlength=16 style="HEIGHT: 22px;" onblur="txtSubsidiaryAccountNumber_onblur()">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit name=cmdGetNextNumber type=button value="Sugerir" onclick="getNextNumber()">&nbsp; &nbsp;
					</TD>
				  <% End If %>
				</TR>
				<TR>
					<TD valign=top nowrap>Nombre:</TD>
					<TD colspan=3>
						<TEXTAREA name=txtSubsidiaryAccountName rows=2 style="width:280"><%=gsSubsidiaryAccountName%></TEXTAREA>
				</TR>
				<TR>
				  <TD valign=top nowrap>Descripción:</TD>
				  <TD colspan=3>
						<TEXTAREA name=txtSubsidiaryAccountDescription rows=4 style="width:280"><%=gsSubsidiaryAccountDescription%></TEXTAREA>						
					</TD>
				</TR>
				<TR class=fullScrollMenuHeader>
					<TD colspan=4>
						Información específica (según el tipo de auxiliar)
					</TD>
				</TR>				
			</TABLE>									
			<DIV id=divExtendedAttrs STYLE="overflow:auto;float:bottom;width=100%; height=130px">
				<TABLE class=applicationTable height=101%>
					<%=gsSubsidiaryAccountExtAttrs%>
					<TR><TD colspan=4 valign=top nowrap width=100% height=100%></TD></TR>			
				</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 align=right>				  
	  <INPUT name=txtItemId type="hidden" value="<%=gnItemId%>">
		<% If gbEdit Then %>
		<INPUT class=cmdSubmit tabindex=-1 name=cmdDeleteItem type=button value="Eliminar" style="width:75" onclick="return deleteItem()"> 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
		&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
		<% End If %>
	  <INPUT class=cmdSubmit name=cmdAddItem type=button value="Aceptar" style="width:75" onclick="return saveItem()">&nbsp; &nbsp;
	  <INPUT class=cmdSubmit name=cmdCancel type=button value="Cancelar" style="width:75" onclick="window.close();"> &nbsp; &nbsp; &nbsp; &nbsp;
	</TD>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
