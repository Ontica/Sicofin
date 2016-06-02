<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim nScriptTimeout, gStdAccountCatalogues, gsStandardAccountsTable
	Dim gnStdAccountCatalogId, gsHistoricDate, gsFromStdAccount, gsToStdAccount, gnOrder
	
	nScriptTimeout = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
	
	Sub Main()
		Dim oStdAccountMgr, dDate, sFilter, sOrder
		'****************************************
		'On Error Resume Next
		Set oStdAccountMgr = Server.CreateObject("EFAStdActUS.CServer")
		gnStdAccountCatalogId = Request.Form("cboStdAccountCatalogues")
		gnOrder = Request.QueryString("order")
		If (Len(gnOrder) = 0) Then
			gnOrder = 0
		End If
		If (Len(gnStdAccountCatalogId) <> 0) Then						
			gStdAccountCatalogues = oStdAccountMgr.CboStdAccountCatalogues(Session("sAppServer"), CLng(gnStdAccountCatalogId))
			gsFromStdAccount = Request.Form("txtFromStdAccount")
			gsToStdAccount = Request.Form("txtToStdAccount")			
			gsHistoricDate = Request.Form("txtHistoricDate")			
			If (Len(gsHistoricDate) <> 0) Then
				dDate = gsHistoricDate
			Else
				dDate = Date
			End If			
			sFilter = GetSQLFilter()
			Select Case gnOrder
				Case "0"
					sOrder = ""
				Case "1"
					sOrder = "numero_cuenta_estandar"
				Case "2"
					sOrder = "nombre_cuenta_estandar, numero_cuenta_estandar"
				Case "3"
					sOrder = "rol_cuenta, numero_cuenta_estandar"
				Case "4"
					sOrder = "object_name, numero_cuenta_estandar"					
				Case "5"
					sOrder = "naturaleza DESC, numero_cuenta_estandar"
				Case Else
					sOrder = ""
			End Select
			gsStandardAccountsTable = oStdAccountMgr.TblStdAccounts(Session("sAppServer"), CLng(gnStdAccountCatalogId), dDate, CStr(sFilter), CStr(sOrder))
		Else			
			gnStdAccountCatalogId = 0
			gStdAccountCatalogues = oStdAccountMgr.CboStdAccountCatalogues(Session("sAppServer"))			
			gsStandardAccountsTable = ""
		End If
		Set oStdAccountMgr = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If
	End Sub
	
	Function GetSQLFilter()
		Dim sTemp
		'********************		
		If (Len(gsFromStdAccount) <> 0) Then
			gsFromStdAccount = Replace(gsFromStdAccount, "*", "%")
			gsFromStdAccount = Replace(gsFromStdAccount, "?", "_")
		End If
		If (Len(gsToStdAccount) <> 0) Then
			gsToStdAccount = Replace(gsToStdAccount, "*", "%")
			gsToStdAccount = Replace(gsToStdAccount, "?", "_")
		End If		
		If (Len(gsFromStdAccount) <> 0) AND (Len(gsToStdAccount) <> 0) Then					
			sTemp = "numero_cuenta_estandar BETWEEN '" & gsFromStdAccount & "' AND '" & gsToStdAccount & "'"
		ElseIf (Len(gsFromStdAccount) <> 0) AND (Len(gsToStdAccount) = 0) Then
			sTemp = "numero_cuenta_estandar LIKE '" & gsFromStdAccount & "'"
		ElseIf (Len(gsFromStdAccount) = 0) AND (Len(gsToStdAccount) <> 0) Then
			sTemp = "numero_cuenta_estandar LIKE '" & gsToStdAccount & "'"
		ElseIf (Len(gsFromStdAccount) = 0) AND (Len(gsToStdAccount) = 0) Then
			sTemp = ""
		End If
		GetSQLFilter = sTemp
		If (Len(gsFromStdAccount) <> 0) Then
			gsFromStdAccount = Replace(gsFromStdAccount, "%", "*")
			gsFromStdAccount = Replace(gsFromStdAccount, "_", "?")
		End If
		If (Len(gsToStdAccount) <> 0) Then
			gsToStdAccount = Replace(gsToStdAccount, "%", "*")
			gsToStdAccount = Replace(gsToStdAccount, "_", "?")
		End If		
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

function formatAccount(sAccount) {
	var obj;
	if (sAccount != '') {
		obj = RSExecute("../financial_accounting_scripts.asp", "FormatStdAccountNumber", document.frmSend.cboStdAccountCatalogues.value, sAccount);
		if (obj.return_value != '') {
			return (obj.return_value);
		} else {
			alert("No entiendo el formato de la cuenta por la que se desea hacer el filtrado.");
			return '';
		}
	} else {
		return '';
	}
}

function callEditor(nOperation, nItemId) {
	var sURL, sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
    switch (nOperation) {  
		case 1:		//Add
		   sURL = 'standard_account_editor.asp?type_id=<%=gnStdAccountCatalogId%>';
		   sOptions = "height=400,width=600,location=0,resizable=0";
		   window.open(sURL, null, sOptions);
		   return false;
		case 2:		//Edit
		   sURL = 'standard_account_editor.asp?type_id=<%=gnStdAccountCatalogId%>&id=' + nItemId;
		   sOptions = "height=400,width=600,location=0,resizable=0";
		   window.open(sURL, null, sOptions);
		   return false;
		case 3:		//Monedas
			sURL = 'standard_account_currencies.asp&id=' + nItemId;
			sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
			window.showModalDialog(sURL, "" , sOptions);
			return false;
		case 4:		// Sectores
			sURL = 'standard_account_sectors.asp&id=' + nItemId;
			sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
			window.showModalDialog(sURL, "" , sOptions);
			return false;
		case 5:   // Saldos
			sURL = 'standard_account_balances.asp&id=' + nItemId;
			sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
			window.showModalDialog(sURL, "" , sOptions);
			return false;
		case 6:	  // Mayores
			sURL = 'standard_account_gral_ledgers.asp&id=' + nItemId;
			sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
			window.showModalDialog(sURL, "" , sOptions);
			return false;
		case 7:   // Historia
			sURL = 'standard_account_history.asp&id=' + nItemId;		
			sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
			window.showModalDialog(sURL, "" , sOptions);
			return false;
	}
	return false;
}

function showAccounts(bShowDatePicker) {
	var sDate, sOptions = 'dialogHeight:280px;dialogWidth:330px;resizable:0;status:0;';
	if (bShowDatePicker) {
		sDate = window.showModalDialog('voucher_date_picker.asp', "" , sOptions);	
		if (sDate != '') {
			document.frmSend.txtHistoricDate.value = sDate;
		} else {
			document.frmSend.txtHistoricDate.value = '';
		}		
	} else {
		document.frmSend.txtHistoricDate.value = '';
	}
	refreshPage(<%=gnOrder%>);
	return false;
}

function formatAccountRanges() {
	var sTemp;
	
	sTemp = document.frmSend.txtFromStdAccount.value;
	if (sTemp != '') {
		sTemp = formatAccount(sTemp);
		if (sTemp != '') {
			document.frmSend.txtFromStdAccount.value = sTemp;		
		} else { 
			document.frmSend.txtFromStdAccount.focus();
			return false
		}
	}
	sTemp = document.frmSend.txtToStdAccount.value;
	if (sTemp != '') {
		sTemp = formatAccount(sTemp);
		if (sTemp != '') {
			document.frmSend.txtToStdAccount.value = sTemp;		
		} else { 
			document.frmSend.txtToStdAccount.focus();
			return false
		}
	}
	return true;
}

function refreshPage(nOrderId) {	
	if (window.document.all.cboStdAccountCatalogues.value == 0) {
		alert("Requiero se seleccione el catálogo de cuentas estándar.");
		document.frmSend.cboStdAccountCatalogues.focus();
		return false;
	}
	if (!formatAccountRanges()) {
		return false;
	}
  if (nOrderId == 0) {
		document.frmSend.action = 'ledger_accounts_viewer.asp';
	} else {
		document.frmSend.action = 'ledger_accounts_viewer.asp?order=' + nOrderId;		  
	}
	document.frmSend.submit();
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="100%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Catálogos de cuentas</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>
			<A href="" onclick="refreshPage(<%=gnOrder%>);return false;">Refrescar página</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="std_account_categories.asp">Definición de catálogos</A>&nbsp;&nbsp;&nbsp;&nbsp;
			<A href="" onclick="window.location.href = '<%=Application("main_page")%>';return false;">Cerrar</A>
		</TD>
	</TR>
	<TR>
		<FORM name=frmSend action="" method=post>		
		<INPUT type=hidden name=txtHistoricDate style="width:85;height:22;" value="<%=gsHistoricDate%>">		
		<TD colspan=4 align=right nowrap>
			<b>Buscar en:</b>&nbsp;&nbsp;
			<SELECT name=cboStdAccountCatalogues>
				<%=gStdAccountCatalogues%>
			</SELECT>
			&nbsp;&nbsp;
			<b>De la cuenta:</b>
			<INPUT type=text name=txtFromStdAccount style="width:85;height:22;" value="<%=gsFromStdAccount%>">			
			<b>A la cuenta:</b>
			<INPUT type=text name=txtToStdAccount style="width:85;height:22;" value="<%=gsToStdAccount%>">
			&nbsp;&nbsp;
			<A href='' onclick='showAccounts(false);return false;'>Al día de hoy</A>
			&nbsp;&nbsp;
			<A href='' onclick='showAccounts(true);return false;'>Otra fecha</A>
		</TD>
		</FORM>
	</TR>	
</TABLE>
<% If (gnStdAccountCatalogId > 0) Then %>
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="100%">
<% If Len(gsStandardAccountsTable) <> 0 Then %>
	<TR>
	  <TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Número de cuenta</b></FONT></A></TD>
	  <TD align=center width=60%><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Nombre</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Rol</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(4);"><FONT color=maroon><b>Tipo de cuenta</b></FONT></A></TD>
	  <TD nowrap align=center><A href="" onclick="return refreshPage(5);"><FONT color=maroon><b>Naturaleza</b></FONT></A></TD>
	  <TD nowrap align=center><FONT color=maroon><b>Más información</b></FONT></TD>
	</TR>
	<%=gsStandardAccountsTable%>
<% Else %>
	<TR>
		<TD colspan=6 align=center>
			El catálogo de cuentas seleccionado está vacío.&nbsp;&nbsp;&nbsp;&nbsp;
			<a href='' onclick='callEditor(1,<%=gnStdAccountCatalogId%>);return false;'>Agregar cuenta</a>
		</TD>		
	</TR>
	<TR><TD colspan=6 align=center></TD></TR>
<% End If %>
</TABLE>
<% End If %>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>