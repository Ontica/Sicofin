<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsPostedVouchersTable, gsCurrentPage, gsCurrentDate
	
	Call Main()
	
	Sub Main()		
		gsCurrentPage = "posted_vouchers.asp?order=" & Request.QueryString("order")
		If Len(Request.Form("txtUseElaborationDate")) <> 0 Then
			If CBool(Request.Form("txtUseElaborationDate")) Then
				gsCurrentDate = Request.Form("txtDate")
				Call SetTable("", gsCurrentDate)
			Else
				gsCurrentDate = Request.Form("txtDate")
				Call SetTable(gsCurrentDate, "")
			End If		
		Else
			Dim oVoucherUS
			Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
			gsCurrentDate = Date
			gsCurrentDate = oVoucherUS.FormatDate(CDate(gsCurrentDate))
			Set oVoucherUS = Nothing
		End If		
	End Sub
	
	Sub SetTable(sApplicationDate, sElaborationDate)
		Dim oVoucherUS
		'*************
		On Error Resume Next	
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")		
		Select Case Request.QueryString("order")
			Case ""				
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "")
				gsCurrentPage = "posted_vouchers.asp"
			Case "1"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "numero_transaccion")
			Case "2"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "fecha_afectacion, numero_transaccion")
			Case "3"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "fecha_registro, numero_transaccion")
			Case "4"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "tipo_transaccion, numero_transaccion")
			Case "5"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "tipo_poliza, numero_transaccion")
			Case "6"
				gsPostedVouchersTable = oVoucherUS.PostedVouchers(Session("sAppServer"), CLng(Session("uid")), sApplicationDate, sElaborationDate, "concepto_transaccion, numero_transaccion")
		End Select
		gsCurrentDate = oVoucherUS.FormatDate(CDate(gsCurrentDate))
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
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function isDate(sDate) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", sDate);
	return obj.return_value;
}

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

function printVouchers() {
	var selectedVouchers = countCheckBoxes("chkVoucher");
	var sMsg;
	if(selectedVouchers == 0) {	
		alert("Para ejecutar esta operación necesito se seleccione al menos una póliza.");
		return false;
	}	
	window.document.frmSend.target = "_blank";
	window.document.frmSend.action = "voucher_viewer.asp";
	window.document.frmSend.submit();	
	return false;
}

function refreshPage(nOrderId) {
  if (nOrderId == 0) {
		window.document.all.frmSend.action = "posted_vouchers.asp";
		window.document.all.frmSend.submit();		
	} else {	
		window.document.all.frmSend.action = "posted_vouchers.asp" + '?order=' + nOrderId;
		window.document.all.frmSend.submit();		
	}
	return false;
}

function searchPostedVoucher(bSearchForApplicationDate) {
	if (window.document.all.txtDate.value == '') {
		alert("Necesito la fecha para efectuar la búsqueda.");
		return false;	
	}
	if (!isDate(window.document.all.txtDate.value)) {
		alert("No reconozco la fecha proporcionada");
		return false;			
	}
	if (bSearchForApplicationDate) {
		window.document.all.txtUseElaborationDate.value = 0;
	} else {
		window.document.all.txtUseElaborationDate.value = -1;
	}	
	window.document.frmSend.target = "";
	window.document.frmSend.action = "<%=gsCurrentPage%>";
	window.document.all.frmSend.submit();
	return true;
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<!--<DIV STYLE="overflow:auto; float:bottom; width=100%; height=78px">-->
<FORM name=frmSend action="" method="post">
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="98%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Mis pólizas actualizadas</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>
	    <% If Len(gsPostedVouchersTable) <> 0 Then %>		
			<A href="" onclick="printVouchers(); return false">Imprimir seleccionadas</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<% End If %>
			<A href="voucher_wizard.asp">Crear póliza</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="transaction_selector.asp">Asistente para crear pólizas</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="pending_vouchers.asp">Mis pólizas pendientes</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp
			<A href="" onclick="window.location.href = '<%=Application("main_page")%>';return false;">Cerrar</A>
		</TD>
	</TR>	
	<TR>
		<TD></TD>		
		<TD colspan=3 align=right valign=middle nowrap>			
			<font color=maroon><b>Buscar:</b></font>&nbsp;&nbsp;
			<INPUT type="text" name=txtDate  value="<%=gsCurrentDate%>" style="width:105;height:22;">&nbsp;&nbsp;(día/mes/año)&nbsp;
			<INPUT type="hidden" name=txtUseElaborationDate  value="<%=Request.Form("txtUseElaborationDate")%>">
			<font color=maroon><b>, por fecha de:</b></font>&nbsp;&nbsp;
			<A href="" onclick="searchPostedVoucher(false);return false;">Elaboración</A>&nbsp;&nbsp;			
			<A href="" onclick="searchPostedVoucher(true);return false;">Afectación</A>&nbsp;&nbsp;			
		</TD>
	</TR>
</TABLE>
<!--</DIV>!-->
<!--<DIV STYLE="overflow:auto; float:bottom; width=100%; height=85%">!-->
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="98%">
<% If Len(gsPostedVouchersTable) <> 0 Then %>  
	<TR>
	  <TD><INPUT type=checkbox name=chkAllItems onclick="selectCheckBoxes('chkVoucher', this.checked);"></TD>
		<TD nowrap align="center"><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Número de póliza</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Afectación</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Elaboración</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(4);"><FONT color=maroon><b>Tipo de transacción</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(5);"><FONT color=maroon><b>Tipo de póliza</b></FONT></A></TD>
	  <TD nowrap width=70% align="center"><A href="" onclick="return refreshPage(6);"><FONT color=maroon><b>Concepto</b></FONT></A></TD>
	</TR>
	<%=gsPostedVouchersTable%>
<% Else %>
	<% If Len(Request.Form("txtUseElaborationDate")) <> 0 Then %>
		<% If CBool(Request.Form("txtUseElaborationDate")) Then %>
			<TR><TD colspan=10 align=center><b>No encontré pólizas elaboradas el día <%=gsCurrentDate%>.</b></TD></TR>
		<% Else %>
			<TR><TD colspan=10 align=center><b>No encontré pólizas con fecha de afectación del día <%=gsCurrentDate%>.</b></TD></TR>
		<% End If %>
	<% Else %>
		<TR><TD colspan=10 align=center><b>Para buscar las pólizas actualizadas se debe elegir una fecha y luego hacer clic en las ligas 'Elaboración' o 'Afectación', según se requiera.</b></TD></TR>
	<% End If %>
<% End If %>
</TABLE>
</FORM>
<!--</DIV>!-->
</BODY>
<BR>&nbsp;
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>