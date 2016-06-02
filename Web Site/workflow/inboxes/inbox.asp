<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsExplorerResultsHeader, gsExplorerResultsBody, gsVoucherInboxes, gsTackedWindows
	Dim gsCboPriorities
	Dim gsFromDate, gsToDate
	Dim gnGralLedgersCategory, gsVoucherNumber, gsItemSubject, gsSendedBy
	Dim gnTransactionTypeId, gnVoucherTypeId, gnItemsType, gnSelectedColumn
	Dim gnSelectedInbox, gsExceptSelectedTransactionType, gsExceptSelectedVoucherType

	Dim oVoucherUS
	Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		
	Call Main()

	Sub Main()
		Call SetGlobalValues()
		Set oVoucherUS = Nothing
		Call GetExplorerInformation()
	End Sub
	
	Sub GetExplorerInformation()
		Dim oExplorer, sWhere, sOrderBy
		'********************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("MHInboxExplorer.CExplorer")
		
		gsVoucherInboxes		    = oExplorer.CboVoucherInboxes(CLng(gnSelectedInbox))
		gsCboPriorities         = oExplorer.CboPriorities(CLng(gnItemsType))		
		gsExplorerResultsHeader = oExplorer.Header(Session("sAppServer"), CLng(gnSelectedInbox), _
																							 CLng(gnSelectedColumn))
		gsExplorerResultsHeader = Replace(gsExplorerResultsHeader, "../images/", "/empiria/images/workflow/")
		gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), CLng(gnSelectedInbox), Session("uid"), _
																					 CStr(sWhere), CLng(gnSelectedColumn), CStr(sOrderBy))
		gsExplorerResultsBody = Replace(gsExplorerResultsBody, "../images/", "/empiria/images/workflow/")
		Set oExplorer = Nothing
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  'Response.Redirect("./exec/exception.asp")
		End If		
	End Sub

  Sub SetGlobalValues()
		If Len(Request.Form("cboVoucherInboxes")) <> 0 Then
			gnSelectedInbox = Request.Form("cboVoucherInboxes")	
			gnGralLedgersCategory  = Request.Form("cboGralLedgerGroups")
			If Len(Request.Form("txtFromDate")) <> 0 Then
				gsFromDate  = oVoucherUS.FormatDate(CDate(Request.Form("txtFromDate")))
			End If
			If Len(Request.Form("txtToDate")) <> 0 Then
				gsToDate  = oVoucherUS.FormatDate(CDate(Request.Form("txtToDate")))
			End If			
			gnTransactionTypeId    = Request.Form("cboTransactionTypes")
			gnVoucherTypeId        = Request.Form("cboVoucherTypes")
			gsVoucherNumber				 = Request.Form("txtTransactionNumber")
			gsItemSubject			     = Request.Form("txtTransactionConcept")			
			gsSendedBy						 = Request.Form("txtSendedBy")
			gnItemsType				     = Request.Form("cboBalancingModes")
			gnSelectedColumn			 = Request.Form("txtSelectedColumn")
			gsTackedWindows        = Request.Form("txtTackedWindows")
			If Len(Request.Form("chkTransactionTypes")) <> 0 Then
				gnTransactionTypeId  = -1 * CLng(gnTransactionTypeId)
				gsExceptSelectedTransactionType = "checked"
			End If
			If Len(Request.Form("chkVoucherTypes")) <> 0 Then
				gnVoucherTypeId = -1 * CLng(gnVoucherTypeId)
				gsExceptSelectedVoucherType = "checked"
			End If
		Else
			gnSelectedInbox = 1
			gnGralLedgersCategory  = 0			
			gsFromDate  = ""
			gsToDate    = ""
			gnTransactionTypeId    = 0
			gnVoucherTypeId        = 0
			gsVoucherNumber				 = ""
			gsItemSubject			     = ""
			gsSendedBy						 = ""
			gnItemsType				 = 0
			gnSelectedColumn			 = 1
			gsTackedWindows				 = ""
		End If
		If (Len(Request.QueryString("inbox")) <> 0) Then
			gnSelectedInbox = CLng(Request.QueryString("inbox"))
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
		obj = RSExecute("../workflow_scripts.asp","IsDate", sDate);
		return obj.return_value;
	} else {
		return true;
	}
}

function viewAttachments(nItemId) {
	var sURL;
	
	sURL = 'attachments.asp?id=' + nItemId;
	window.open(sURL, "_blank")
	return false;
}


function deleteVouchers() {	
	var nSelectedVouchers = countCheckBoxes('chkAllItems');
	var sMsg;
	
	if(nSelectedVouchers == 0) {	
		alert('Para ejecutar esta operación necesito se seleccione al menos una póliza.');
		return false;
	}
	
	if (nSelectedVouchers > 1) {
		sMsg = 'Esta operación eliminará del sistema los ' + nSelectedVouchers + ' elementos seleccionados,\n' + 
					 'por lo que ya no podrán ser recuperados.\n\n' + '¿Procedo con la operación?';
	}
	if (nSelectedVouchers == 1) {
		sMsg = 'Esta operación eliminará del sistema el elemento seleccionado, por lo que ya no\n' + 
					 'podrá ser recuperado.\n' +  '¿Procedo con la operación?';
	}
	if (confirm(sMsg)) {
		if(sPendingVouchers == '') {	
			alert("Las pólizas seleccionadas no se pueden eliminar debido a que ya están en el diario.");
			return false;	
		}	
		document.frmSend.txtPendingTasks.value = sPendingVouchers;
		document.frmSend.txtPostedVouchers.value = '';
		document.frmSend.action = './exec/delete_vouchers.asp';
		document.frmSend.submit();
		document.frmSend.action = '';		
	}
	return false;
}

function orderBy(nSelectedColumn) {
  var nTemp = document.all.txtSelectedColumn.value;
  
  if (nTemp == '') {
		document.all.txtSelectedColumn.value = nSelectedColumn;
	} else {
		if ((nTemp == nSelectedColumn) || (Math.abs(nSelectedColumn) == nSelectedColumn)) {
			if (nTemp == nSelectedColumn) {
				document.all.txtSelectedColumn.value = (-1 * nSelectedColumn);
			} else {
				document.all.txtSelectedColumn.value = nSelectedColumn;
			}
		} else {
			document.all.txtSelectedColumn.value = nSelectedColumn;
		}
  }
  document.frmSend.target = '_self';
  document.frmSend.action = '';
  document.frmSend.submit();
	return false;
}

function assignMeTasks() {
	var sPendingVouchers = arrayPendingTasks('chkAllItems');
	var selectedItems = countCheckBoxes("chkAllItems");
	var sMsg;
	if(selectedItems == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una tarea.");
		return false;
	}
	if (selectedItems > 1) {
		sMsg = 'Esta operación le asignará las ' + selectedItems + ' tareas seleccionadas.\n\n' + 					 
					 '¿Procedo con la operación?';
	}
	if (selectedItems == 1) {
		sMsg = 'Esta operación le asignará la tarea seleccionada.\n\n' + 					 
					 '¿Procedo con la operación?';
	}		
	if (confirm(sMsg)) {
		if(sPendingVouchers == '') {	
			alert("Las tareas seleccionadas ya están en el diario");
			return false;	
		}	
		document.frmSend.txtPendingTasks.value = sPendingVouchers;
		document.frmSend.target = "_self"
		document.frmSend.action = "./exec/assign_me_tasks.asp";
		document.frmSend.submit();
		document.frmSend.action = '';
	}
	return false;
}


function arrayPendingTasks(sCheckBoxName) {
	var i = 0, sTemp = '';
	
	if (typeof(document.all[sCheckBoxName]) == 'undefined') {
		return '';
	}	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				if (sTemp.length != 0) {					
					sTemp += ',' + Math.abs(document.all[sCheckBoxName](i).value);
				} else {					
					sTemp = Math.abs(document.all[sCheckBoxName](i).value);
				}
			}
		}		
	} else {
		if (document.all[sCheckBoxName].checked) {
			sTemp = Math.abs(document.all[sCheckBoxName].value);
		}
	}
	return sTemp;
}

function reassignVouchers() {
	var sPendingVouchers = arrayPendingTasks('chkAllItems');
	
	if(sPendingVouchers == '') {	
		alert("Para ejecutar esta operación necesito se seleccione al menos un elemento.");
		return false;	
	}
	document.frmSend.txtPendingTasks.value = sPendingVouchers;
			
	document.frmSend.target = '_self';
	document.frmSend.action = "./exec/reassign_vouchers.asp";
	document.frmSend.submit();
	return false;
}

function resetSearchOptions() {
	var sTemp = document.all.cboVoucherInboxes.value;
	
	document.frmSend.reset();
	document.all.cboVoucherInboxes.value =(sTemp);
	return false;
}

function searchItems() {
	if (!isDate(document.all.txtFromDate.value)) { 
		alert("No reconozco la fecha inicial.");
		document.all.txtFromDate.focus();
		return false;
	}
	if (!isDate(document.all.txtToDate.value)) { 
		alert("No reconozco la fecha final.");
		document.all.txtToDate.focus();
		return false;
	}	
	document.frmSend.target = '_self';
	document.frmSend.action = '';
	document.frmSend.submit();
	return false;
}

function setDynamicItems() {	
	var sTemp = document.all.cboVoucherInboxes.value;
	/*	
	if (sTemp == '4' || sTemp == '9') {
		document.all.rowBalancesModes.style.display = 'none';
		document.all.cboBalancingModes.value = 0;
		document.all.rowTransactionNumber.style.display = 'inline';
	} else {
		document.all.rowBalancesModes.style.display = 'inline';
		if (sTemp == '1' || sTemp == '2' || sTemp == '3' || sTemp == '6' || sTemp == '7' || sTemp == '8') {
			document.all.rowTransactionNumber.style.display = 'none';
			document.all.txtTransactionNumber.value = '';
		} else { 
			document.all.rowTransactionNumber.style.display = 'inline';
		}
	}
	*/
}

function insertTableInformation(oTableSection, nObjectId) {
  var oRow, i, j;
	var obj, aRows, aCols, sTemp;
	
  setCursor('wait');
  obj = RSExecute("../workflow_scripts.asp", "Vouchers", nObjectId);
  
  sTemp = obj.return_value;
	aRows = sTemp.split("¦");
  for (i = 0; i < aRows.length; i++) {
		aCols = aRows[i].split("|");
		oRow = oTableSection.insertRow();
		for (j = 0; j < aCols.length; j++) {
			oRow.insertCell().innerHTML = aCols[j];
		}
  }
	setCursor('auto');
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="setDynamicItems();showTackedWindows(Array(<%=gsTackedWindows%>));">
<FORM name=frmSend action='' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Bandeja de entrada&nbsp;
		</TD>
		<TD colspan=3 align=right nowrap>
			Ver:<img align=middle src='/empiria/images/invisible.gif'>
			<SELECT name=cboVoucherInboxes  LANGUAGE=javascript onchange="setDynamicItems()">
				<%=gsVoucherInboxes%>
			</SELECT>
			<img align=middle src='/empiria/images/invisible.gif'>		  <A href="" onclick='return(searchItems());'>Ejecutar</A>
			<img align=middle src='/empiria/images/invisible.gif'>			<A href='' onclick="return(showOptionsWindow(document.all.divSearchOptions));">Más opciones</A>
			&nbsp;|&nbsp;			<A href='' onclick="return(showOptionsWindow(document.all.divSelectedVouchersOptions));">Selección</A>
			<img align=middle src='/empiria/images/invisible.gif'>			<img align=middle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=middle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=middle src='/empiria/images/invisible.gif'>
			<img align=middle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
	</TR>
	<TR id=divSearchOptions style='display:none;'>
		<TD colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Buscar elemento
					</TD>
					<TD nowrap align=left>
						<A href="" onclick="return(notAvailable());">Cargar búsqueda</A>
						&nbsp; | &nbsp
						<A href="" onclick="return(notAvailable());">Guardar búsqueda</A>
						<img align=top src='/empiria/images/invisible8.gif'>
						<img align=top src='/empiria/images/invisible8.gif'>
						<img align=top src='/empiria/images/invisible8.gif'>
						<A href="" onclick="return(searchItems());">Ejecutar búsqueda</A>
					</TD>
					<TD nowrap align=right>
						<img src='/empiria/images/invisible4.gif'>
						<img src='/empiria/images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>
					  <img id=cmdSearchOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSearchOptions, this)' alt='Fijar la ventana'>
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
						<img src='/empiria/images/invisible.gif'>						
						<img src='/empiria/images/close_white.gif' onclick='closeOptionsWindow(document.all.divSearchOptions, document.all.cmdSearchOptionsTack)' alt='Cerrar'>
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Recibido:</b></TD>
					<TD colspan=2 nowrap width=100%>
						Entre el día:
						<INPUT name=txtFromDate style="width:100;height:20;" value='<%=gsFromDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromDate)'>
						&nbsp; &nbsp; &nbsp; y el día:
						<INPUT name=txtToDate style="width:100;height:20;" value='<%=gsToDate%>'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtToDate)'>&nbsp;&nbsp;
						(día / mes / año)
					</TD>
				</TR>
				<TR id=rowTransactionNumber style='display:none;'>
					<TD nowrap><b>Cuyo número de póliza sea:</b>&nbsp;</TD>
					<TD colspan=2 nowrap>
						<INPUT name=txtTransactionNumber style="width:161;height:20;" value='<%=gsVoucherNumber%>'> 
							&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'))" target=_blank>comodines</A>)
					</TD>
				</TR>				
				<TR>
					<TD nowrap><b>Cuyo asunto sea:</b></TD>
					<TD colspan=2 nowrap width=100% valign=middle>
						<INPUT name=txtTransactionConcept style="width:400;height:20;" value='<%=gsItemSubject%>'>
							&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'));" target=_blank>comodines</A>)
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Enviado por:</b></TD>
					<TD colspan=2 nowrap width=100% valign=middle>
						<INPUT name=txtSendedBy style="width:400;height:20;" value='<%=gsSendedBy%>'>
							&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'));" target=_blank>comodines</A>)
					</TD>
				</TR>				
				<TR>
					<TD nowrap><b>Cuya proridad sea: &nbsp; &nbsp; &nbsp;</b></TD>
					<TD colspan=2 nowrap>
						<SELECT name=cboPriorities style='width:300'>
							<%=gsCboPriorities%>
						</SELECT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR id=divSelectedVouchersOptions style='display:none;'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap colspan=2>
						¿Qué se desea hacer con los elementos seleccionados?
					</TD>
					<TD nowrap align=right>					
					  <img id=cmdSelectedVouchersOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSelectedVouchersOptions, this)' alt='Fijar la ventana'>
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='/empiria/images/invisible.gif'>
					  <img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divSelectedVouchersOptions, document.all.cmdSelectedVouchersOptionsTack)" alt='Cerrar'>
					</TD>				
				</TR>
				<TR>
					<TD nowrap>						
						<A href="" onclick="return(assignMeTasks());">Asignarme las tareas</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(printVouchers());">Generar un reporte</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(reassignVouchers());">Enviarlos a otro participante</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(reassignVouchers());">Exportarlos a Microsoft Excel®</A>&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="" onclick="return(deleteVouchers());">Eliminarlos</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
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
						<A href="../../balances/balance_explorer.asp">Explorador de saldos</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../reports/balances.asp">Balanzas de comprobación</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../../reports/other_reports.asp">Reportes</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
<TR><TD colspan=4 nowrap>
<TABLE class=applicationTable> 
<THEAD>
	<TR class=fullScrollMenuHeader valign=center>
		<TD colspan=9 class=fullScrollMenuTitle>
			Elementos de la bandeja de entrada
		</TD>
	</TR>	
	<TR class=applicationTableHeader valign=center>
		<%=gsExplorerResultsHeader%>
	</TR>
</THEAD>
<% If (Len(gsExplorerResultsBody) <> 0) Then %>
	<%=gsExplorerResultsBody%>
<% Else %>
	<TBODY>
		<TR>
		<TD colspan=9><b>No encontré ningún elemento con el criterio de búsqueda proporcionado.</b></TD>
		</TR>
	</TBODY>
<% End If %>
</TABLE>
<INPUT TYPE=hidden name=txtPostedVouchers>
<INPUT TYPE=hidden name=txtPendingTasks>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>
</TD>
</TR>
</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>