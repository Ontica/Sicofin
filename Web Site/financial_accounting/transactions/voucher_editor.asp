<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gnTransactionId, gsVoucherHeader, gsVoucherPostings, gsTransactionStatus, gnDateStatus, gsTackedWindows
	Dim bUserIsSupervisor, bCanDelete
	
	Call Main()
	
	Sub Main()
		Dim oVoucher
		gnTransactionId     = Request.QueryString("id")
		Set oVoucher        = Server.CreateObject("AOGLVoucherUS.CVoucher")
		gsVoucherHeader     = oVoucher.Header(Session("sAppServer"), CLng(gnTransactionId))
		If (Request.QueryString("analize") = "true") Then
			gsVoucherPostings   = oVoucher.GetPostings(Session("sAppServer"), CLng(gnTransactionId), True, True)
		Else
			gsVoucherPostings   = oVoucher.GetPostings(Session("sAppServer"), CLng(gnTransactionId), True, False)
		End If
		gsTransactionStatus = oVoucher.TransactionStatus(Session("sAppServer"), CLng(gnTransactionId))
		gnDateStatus				= oVoucher.DateStatus(Session("sAppServer"), CLng(gnTransactionId))
		bUserIsSupervisor   = oVoucher.UserIsSupervisor(Session("sAppServer"), CLng(gnTransactionId), CLng(Session("uid")))
		gsTackedWindows     = Request.Form("txtTackedWindows")
		Set oVoucher = Nothing
		Set oVoucher = Server.CreateObject("AOGLVoucher.CServer")
		bCanDelete   = oVoucher.CanDelete(Session("sAppServer"), CLng(gnTransactionId))
		Set oVoucher =  Nothing
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
<meta http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oPostingWindow = null;

function refreshPostings() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "PostingsTable", <%=gnTransactionId%>, false);
	document.all.divPostingsTable.innerHTML = obj.return_value;
	
	obj = RSExecute("../financial_accounting_scripts.asp", "TransactionStatus", <%=gnTransactionId%>);
	document.all.tdTransactionStatus.innerHTML = obj.return_value;
	return false;
}

function refreshAll() {
	window.document.location.href = window.document.location.href;
	return false;
}

function deleteVoucher() {
	var sMsg;
	
	<% If bCanDelete Then %>
		sMsg  = 'Esta operaci�n eliminar� la p�liza con todos sus movimientos.\n\n';
		sMsg += '�Procedo con la eliminaci�n?';
		return (confirm(sMsg));
	<% Else %>
		sMsg  = 'No se puede eliminar la p�liza debido a que tiene movimientos protegidos.\n\n';
		alert(sMsg);
		return false;
	<% End If %>
}

function checkVoucher() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","ValidateTransaction", <%=gnTransactionId%>);
	return obj.return_value;
}

function msgValidateVoucher() {	
	var nResult;
	nResult = checkVoucher();
	switch (nResult) {
		case 0:
			alert('�La p�liza est� balanceada correctamente!');
			return false;
		case -1:
			alert('La p�liza a�n no tiene movimientos.');
			return false;
		case -2:
			alert('La p�liza tiene movimientos con informaci�n incompleta o incongruente.');
			return false;
		case -3:
			alert('La p�liza a�n no est� balanceada.');
			return false;
	}
}


function postVoucher() {
	var sMsg, nResult;
	
	nResult = checkVoucher();
	if (nResult != 0) {
		<% If (gnDateStatus < 0) AND (Not bUserIsSupervisor) Then %>
			switch(nResult) {
				case -1:
					sMsg = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al supervisor debido a que a�n no tiene movimientos.';
					break;
				case -2:
					sMsg = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al supervisor debido a que tiene movimientos con informaci�n incompleta o incongruente.';
					break;
				case -3:
					sMsg = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al supervisor debido a que a�n no est� balanceada.';
					break;
			}
		<% ElseIf (gnDateStatus < 0) AND (bUserIsSupervisor) Then %>
			switch(nResult) {
				case -1:
					sMsg  = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al diario \n';
					sMsg += 'debido a que a�n no tiene movimientos.';
					break;
				case -2:
					sMsg  = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al diario \n';
					sMsg += 'debido a que tiene movimientos con informaci�n incompleta o incongruente.';
					break;
				case -3:
					sMsg = 'Esta p�liza, con fecha de afectaci�n atrasada, no puede enviarse al diario \n';
					sMsg = 'debido a que a�n no est� balanceada.';
					break;
			}
		<% ElseIf (gnDateStatus = 0) Then %>
			switch(nResult) {
				case -1:
					sMsg = 'La p�liza no puede incorporarse al diario debido a que a�n no tiene movimientos.';
					break;
				case -2:
					sMsg  = 'La p�liza no puede incorporarse al diario debido a que tiene movimientos con \n';
					sMsg += 'informaci�n incompleta o incongruente.';
					break;
				case -3:
					sMsg = 'La p�liza no puede incorporarse al diario debido a que a�n no est� balanceada.';
					break;
			}		
	  <% ElseIf (gnDateStatus > 0) Then %>
			switch(nResult) {
				case -1:
					sMsg = 'Esta p�liza, con fecha de afectaci�n adelantada, a�n no tiene movimientos.';
					break;
				case -2:
					sMsg  = 'Esta p�liza, con fecha de afectaci�n adelantada, tiene movimientos con \n';
					sMsg += 'informaci�n incompleta o incongruente.';
					break;
				case -3:
					sMsg = 'Esta p�liza, con fecha de afectaci�n adelantada, a�n no est� balanceada.';
					break;
			}	
	  <% End If %>
		alert(sMsg);
		return false;
	} else {			
		<% If (gnDateStatus < 0) AND (Not bUserIsSupervisor) Then %>
	     sMsg  = 'Esta p�liza tiene fecha de afectaci�n atrasada, por lo que esta operaci�n la enviar�\n';
	     sMsg += 'al supervisor para que �ste la incorpore al diario.\n\n';
	     sMsg += '�Procedo con el env�o de la p�liza al supervisor?';
	     return (confirm(sMsg));
		<% ElseIf (gnDateStatus < 0) AND (bUserIsSupervisor) Then %>
	     sMsg  = 'Esta p�liza tiene fecha de afectaci�n atrasada y ser� incorporada al diario,\n';
	     sMsg += 'por lo cual ya no podr� ser modificada.\n\n'
	     sMsg += '�Procedo con la contabilizaci�n de la p�liza con fecha de afectaci�n atrasada?';
	     return (confirm(sMsg));		
	  <% ElseIf (gnDateStatus = 0) Then %>
	     sMsg  = 'Esta operaci�n guardar� la p�liza en el diario, por lo cual ya no podr� modificarse.\n\n';
	     sMsg += '�Procedo con la operaci�n?';
	     return (confirm(sMsg));
	  <% ElseIf (gnDateStatus > 0) Then %>
	     sMsg  = 'Debido a que la p�liza tiene una fecha de afectaci�n adelantada, esta operaci�n\n';
	     sMsg += 'estar� disponible cuando sea abierto el per�odo correspondiente.\n\n';
	     sMsg += 'Gracias.';	     
	     alert (sMsg);
	     return (false);
		<% End If %>		
	}
}

function callEditor(nOperation, nItemId) {
	var sURL, sOpt;
  switch (nOperation) {  
    case 1:		//Add			
			sURL = 'posting_editor.asp?transactionId=<%=gnTransactionId%>';
			sOpt = 'height=465px,width=370px,resizable=no,scrollbars=no,status=no,location=no';
			if (oPostingWindow == null || oPostingWindow.closed) {
				oPostingWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oPostingWindow.focus();
				oPostingWindow.navigate(sURL);
			}
			return false;
    case 2:		//Edit
			sURL = 'posting_editor.asp?transactionId=<%=gnTransactionId%>&id=' + nItemId;
			sOpt = 'height=465px,width=370px,resizable=no,scrollbars=no,status=no,location=no';
			if (oPostingWindow == null || oPostingWindow.closed) {
				oPostingWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oPostingWindow.focus();
				oPostingWindow.navigate(sURL);
			}
			return false;			
			//oPostingWindow = window.open(sURL, null, sOpt)
			//oPostingWindow.focus();
			//window. showModelessDialog(sURL, null, sOpt);
			return false;		
		case 3:   //Edit voucher header			
			sURL = 'voucher_header_editor.asp?id=<%=gnTransactionId%>';
			sOpt = 'height=300px,width=520px,resizable=no,scrollbars=no,status=no,location=no';
			//sOpt = 'dialogHeight:330px;dialogWidth:520px;resizable:no;scroll:no;status:no;help:no;'
			window.open(sURL, '_blank', sOpt);
			//window.showModalDialog(sURL, null, sOpt);			
			return false;
	}
	return false;
}

function window_onunload() {
	if (oPostingWindow != null && !oPostingWindow.closed) {
		oPostingWindow.close();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload="window_onunload();" topmargin=0>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Editor de p�lizas
		</TD>
		<TD colspan=3 align=right nowrap>
			<A href="voucher_wizard.asp">Crear otra p�liza</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>
			<A href="voucher_explorer.asp">Explorador de p�lizas</A>
			<img align=absmiddle src='/empiria/images/invisible.gif'>			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la p�gina principal">								</TD>
	</TR>
	<TR id=divTasksOptions style='display:none'>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class='fullScrollMenuHeader'>
					<TD class='fullScrollMenuTitle' nowrap>
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
						<A href="transaction_selector.asp">Asignar esta p�liza a otro participante</A>
						&nbsp;&nbsp;&nbsp;&nbsp;					
						<A href="transaction_selector.asp">Exportar a MS Excel<sup>�</sup></A>						
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="transaction_selector.asp">Explorador de saldos</A>			
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="transaction_selector.asp">Balanzas de comprobaci�n</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="transaction_selector.asp">Reportes</A>
						<img src='/empiria/images/invisible.gif'>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=fullScrollMenu>			
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle nowrap width=315>
						<% If (gnDateStatus = 0) Then %>
							P�liza
						<% ElseIf (gnDateStatus < 0) Then %>
							<img align=absmiddle src='/empiria/images/exclamation.gif' alt='P�liza con fecha atrasada'>
							&nbsp; 
							P�liza con fecha atrasada
						<% ElseIf (gnDateStatus > 0) Then %>
							<img align=absmiddle src='/empiria/images/exclamation.gif' alt='P�liza con fecha adelantada'>
							&nbsp; 
							P�liza con fecha adelantada
						<% End If %>
					</TD>
					<TD nowrap>
						<% If (gnDateStatus < 0) AND (Not bUserIsSupervisor) Then %>
						<A href="exec/send_voucher_to_check.asp?id=<%=gnTransactionId%>" onclick="return postVoucher()">Enviar al supervisor (fecha valor)</A>
						<% ElseIf (gnDateStatus < 0) AND (bUserIsSupervisor) Then %>
						<A href="exec/post_voucher.asp?id=<%=gnTransactionId%>" onclick="return postVoucher()">Enviar al diario (fecha valor)</A>
						<% ElseIf gnDateStatus = 0 Then %>
						<A href="exec/post_voucher.asp?id=<%=gnTransactionId%>" onclick="return postVoucher()">Enviar al diario</A>
						<% ElseIf gnDateStatus > 0 Then %>
						<A href="" onclick="return msgValidateVoucher()">Revisar (p�liza adelantada)</A>
						<% End If %>	
						&nbsp; &nbsp; | &nbsp			
						<A href='' onclick='return(callEditor(3,0));'>Editar encabezado</A>
						&nbsp; | &nbsp
						<A href="exec/delete_voucher.asp?id=<%=gnTransactionId%>" onclick="return deleteVoucher()">Eliminar</A>
						&nbsp; | &nbsp
						<A href="pending_voucher_viewer.asp?id=<%=gnTransactionId%>" target="_blank">Imprimir</A>
					</TD>
					<TD align=right>
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
				<TR>
				<%=gsVoucherHeader%>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=fullScrollMenu>			
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle nowrap width=315>
						Movimientos
					</TD>
					<TD nowrap>
						<A id=ancAddPost href="" onclick="return(callEditor(1, 0))">Agregar</A>
						&nbsp; &nbsp; | &nbsp
						<A id=ancRefreshPost href='' onclick='return(refreshPostings());'>Refrescar</A>
						&nbsp; | &nbsp
						<A href="voucher_editor.asp?id=<%=gnTransactionId%>&analize=true">Analizar</A>
						&nbsp; | &nbsp
						<A href='' onclick='return(notAvailable())'>Balancear</A>
						<A id=ancRefreshAll href='' onclick='return(refreshAll());'></A>
					</TD>					
					<TD colspan=1 align=right>
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					</TD>
				</TR>
			</TABLE>
			<span id=divPostingsTable>
				<TABLE class=applicationTable>
					<TR class=applicationTableHeader>
					  <TD nowrap width=120><b>N�m. de cuenta</b></TD>
					  <TD><b>Sec</b></TD>
					  <TD width=40%><b>Descripci�n</b></TD>
					  <TD><b>Verif</b></TD>
					  <TD><b>Area</b></TD>
						<TD align=center><b>Moneda</b></TD>
					  <TD align=center nowrap><b>T. de cambio</b></TD>
					  <TD colspan=3 align=center width=30%><b>Importes</b></TD>
					</TR>
					<TR class=applicationTableHeader>
					  <TD><b><i>Auxiliar</i></b></TD>
					  <TD>&nbsp;</TD>
					  <TD><b><i>Concepto</i></b></TD>
					  <TD colspan=3>&nbsp;</TD>
					  <TD align=center>&nbsp;</TD>
					  <TD align=center><b>Parcial</b></TD>
					  <TD align=center><b>Debe</b></TD>
					  <TD align=center><b>Haber</b></TD>
					</TR>
					<%=gsVoucherPostings%>
				</TABLE>
			</span>
			<TABLE class=fullScrollMenu>
				<TR class=fullScrollMenuHeader valign=top>
	        <% If Len(gsTransactionStatus) <> 0 Then %>
					<TD valign=top class=fullScrollMenuTitle width=200>
						<img align=top src='/empiria/images/exclamation.gif' alt='P�liza con fecha atrasada'>
						Estado
					</TD>
					<TD id=tdTransactionStatus valign=top align=right>
						<%=gsTransactionStatus%>
					</TD>
			    <% Else %>
					<TD valign=top class=fullScrollMenuTitle width=200>
						<img align=top src='/empiria/images/exclamation.gif' alt='P�liza balanceada correctamente'>
						Estado
					</TD>
					<TD id=tdTransactionStatus valign=top align=right>
						La p�liza est� balanceada correctamente.
					</TD>
	        <% End If %>
				</TR>
	    </TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden name=txtTackedWindows>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>