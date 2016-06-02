<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsPendingVouchersTable
	
	Call Main()
	
	Sub Main()
		Dim oVoucherUS
		'*************
		On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		Select Case Request.QueryString("order")
			Case ""
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "")
			Case "1"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "fecha_afectacion")
			Case "2"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "fecha_registro")
			Case "3"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "tipo_transaccion, fecha_afectacion")
			Case "4"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "tipo_poliza, fecha_afectacion")
			Case "5"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "concepto_transaccion, fecha_afectacion")
			Case "6"
				gsPendingVouchersTable = oVoucherUS.PendingVouchers(Session("sAppServer"), CLng(Session("uid")), "concepto_transaccion, fecha_afectacion, nombre_autorizada_por")
		End Select		
		Set oVoucherUS = Nothing
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
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function countCheckBoxes(sCheckBoxName) {
	var i= 0, counter = 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				counter++;
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			counter++;
		}
	}
	return counter;
}

function selectCheckBoxes(sCheckBoxName, bCheck) {
	var i= 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			document.all[sCheckBoxName](i).checked = bCheck;
		}		
	} else {
		document.all[sCheckBoxName].checked = bCheck;		
	}
	return true;	
}

function deleteVouchers() {
	var selectedVouchers = countCheckBoxes("chkVoucher");
	var sMsg;
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}
	if (selectedVouchers > 1) {
		sMsg = 'Esta operación eliminará del sistema las ' + selectedVouchers + ' pólizas seleccionadas,\n' + 
		'por lo que ya no podrán ser recuperadas.\n\n' + '¿Procedo con la operación?';
	}
	if (selectedVouchers == 1) {
		sMsg = 'Esta operación eliminará del sistema la póliza seleccionada, por lo que ya no\n' + 
		'podrá ser recuperada.\n' +  '¿Procedo con la operación?';
	}	
	if (confirm(sMsg)) {
		window.document.frmSend.action = "exec/delete_vouchers.asp";
		window.document.frmSend.submit();
	}	
	return false;
}

function postVouchers() {
	var selectedVouchers = countCheckBoxes("chkVoucher");
	var sMsg;
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}
	if (selectedVouchers > 1) {
		sMsg = 'Esta operación enviará al diario las ' + selectedVouchers + ' pólizas seleccionadas.\n\n' + 
					 'Sin embargo, esto ocurrirá únicamente con las pólizas que estén debidamente balanceadas.\n\n' +
					 '¿Procedo con la operación?';
	}
	if (selectedVouchers == 1) {
		sMsg = 'Esta operación enviará la póliza seleccionada al diario.\n\n' + 
					 'Sin embargo, esto ocurrirá únicamente si esta se encuentra debidamente balanceada.\n\n' + 
					 '¿Procedo con la operación?';
	}	
	if (confirm(sMsg)) {
		window.document.frmSend.action = "exec/post_vouchers.asp";
		window.document.frmSend.submit();
	}	
	return false;
}

function printVouchers() {
	var selectedVouchers = countCheckBoxes("chkVoucher");
	var sMsg;
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}	
	window.document.frmSend.target = "_blank";
	window.document.frmSend.action = "pending_voucher_viewer.asp";
	window.document.frmSend.submit();	
	return false;
}

function reassignVouchers() {
	var selectedVouchers = countCheckBoxes("chkVoucher");
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;	
	}	
	window.document.frmSend.action = "exec/reassign_vouchers.asp";
	window.document.frmSend.submit();	
	return false;
}

function refreshPage(nOrderId) {
  if (nOrderId == 0) {
		window.location.href = "pending_vouchers.asp";
	} else {	
		window.location.href = "pending_vouchers.asp" + '?order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<!--<DIV STYLE="overflow:auto; float:bottom; width=98%; height=70px">!-->
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="98%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Mis pólizas pendientes</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>
			<A href="" onclick="window.location.href=window.location.href;return false;">Refrescar página</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="voucher_wizard.asp">Crear póliza</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="transaction_selector.asp">Asistente para crear pólizas</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="posted_vouchers.asp">Mis pólizas actualizadas</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp
			<A href="" onclick="window.location.href = '<%=Application("main_page")%>';return false;">Cerrar</A>
		</TD>
	</TR>
	<TR>
		<TD></TD>		
		<TD colspan=3 align=right nowrap>
			<% If Len(gsPendingVouchersTable) <> 0 Then %>
			<font color=maroon><b>Pólizas seleccionadas:</b></font>&nbsp;&nbsp;
			<A href="" onclick="printVouchers(); return false">Imprimir</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="" onclick="postVouchers(); return false">Enviar al diario</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="" onclick="deleteVouchers();return false;">Eliminar</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="" onclick="reassignVouchers(); return false">Enviar a otro participante</A>
			<% End If %>	
		</TD>
	</TR>
</TABLE>
<!--</DIV>!-->
<!--<DIV STYLE="overflow:auto; float:bottom; width=98%; height=85%">!-->
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="98%">
	<FORM name=frmSend action="" method="post">
<% If Len(gsPendingVouchersTable) <> 0 Then %>
	<TR>
		<TD><INPUT type=checkbox name=chkAllItems onclick="selectCheckBoxes('chkVoucher', this.checked);"></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Afectación</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Elaboración</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Tipo de transacción</b></FONT></A></TD>
	  <TD nowrap width=140 align="center"><A href="" onclick="return refreshPage(4);"><FONT color=maroon><b>Tipo de póliza</b></FONT></A></TD>	  
	  <TD nowrap width=60% align="center"><A href="" onclick="return refreshPage(5);"><FONT color=maroon><b>Concepto</b></FONT></A></TD>
	  <TD nowrap width=100 align="center"><A href="" onclick="return refreshPage(6);"><FONT color=maroon><b>Estado</b></FONT></A></TD>
	</TR>	
	<%=gsPendingVouchersTable%>
	</FORM>
<% Else %>
	<TR><TD colspan=8 align=center><b>No hay pólizas pendientes.</b></TD></TR>
<% End If %>
</TABLE>
<BR>&nbsp;
<!--</DIV>!-->
</BODY>
</HTML>