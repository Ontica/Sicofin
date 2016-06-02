<%@ Language=VBScript %>
<%
	Option Explicit

	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If

  Dim gsElaborationDate
	Dim gsFTPServer, gsFTPDirectory, gsCboVoucherTypes
	
	Call Main()
	
	Sub Main()		
		Dim iVouchersTextFile, oVoucherUS
		
		Set iVouchersTextFile = Server.CreateObject("SCFIVouchersTextFile.CServer")
		gsFTPServer = iVouchersTextFile.FTPServer
		gsFTPDirectory = iVouchersTextFile.GetUserDirectory(Session("sAppServer"), Session("uid"))
		Set iVouchersTextFile = Nothing
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		gsCboVoucherTypes = oVoucherUS.CboVouchersTypes(Session("sAppServer"), 25)
		gsElaborationDate = oVoucherUS.FormatDate(Now)
		Set oVoucherUS = Nothing
	End Sub
	
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function isDate(sDate) {
	var obj;
	obj = RSExecute("/empiria/financial_accounting/financial_accounting_scripts.asp", "IsDate", sDate);
	return obj.return_value;
}

function frmUpload_onsubmit() {
	if(document.frmUpload.txtFileName.value == '') {
		alert("Requiero se exporte el archivo con las pólizas (Paso 1).")
		return false;
	}
	if(document.frmUpload.txtElaborationDate.value == '') {
		alert("Requiero la fecha de elaboración de las pólizas (Paso 2).")
		return false;
	}
	if(!isDate(document.frmUpload.txtElaborationDate.value)) {
		alert("No reconozco la fecha de elaboración de las pólizas (Paso 2).")
		return false;
	}
	return true;
}

function aoxUploader_TransferSuccess(sFileName) {
  document.frmUpload.txtFileName.value = sFileName;
}

function aoxUploader_TransferFail(oError) {
	alert(oError.Description);
}

function aoxUploader_Error(oError) {
	alert(oError.Description);
}

//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=aoxUploader EVENT=TransferSuccess(sFileName)>
<!--
 aoxUploader_TransferSuccess(sFileName)
//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=aoxUploader EVENT=TransferFail(oError)>
<!--
 aoxUploader_TransferFail(oError)
//-->
</SCRIPT>
<SCRIPT LANGUAGE=javascript FOR=aoxUploader EVENT=Error(oError)>
<!--
 aoxUploader_Error(oError)
//-->
</SCRIPT>
</HEAD>
<BODY topmargin=0>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Importador de pólizas
		</TD>
		<TD colspan=3 align=right nowrap>						<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD colspan=2>						
				    <P>Este programa permite anexar las pólizas contenidas en un archivo al sistema de contabilidad financiera.<br>
						   Estos archivos se obtienen de aplicaciones como el SIF, GEM, Mesa de dinero y otras.
				    </P>     
				  </TD>
				</TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2 valign=top>
						<b>Paso 1. Transferir el archivo con las pólizas al servidor</b>
					</TD>
				</TR>
				<TR>
					<TD colspan=2 valign=top>		
						Haga clic en el botón 'Examinar...' para seleccionar el archivo con las pólizas que desea incorporar.<br><br>
						Despúes haga clic en 'Enviar' para transferir el archivo de su equipo hacia el equipo servidor.						
					</TD>
				</TR>
				<TR>
					<TD colspan=2 nowrap valign=top align=center>
						<OBJECT classid="CLSID:443B33BF-659C-440A-AEE1-7CA5287CFBE4" codeBase="/empiria/bin/aoupload.cab#version=-1,-1,-1,-1" height=58 name=aoxUploader style="LEFT: 0px; TOP: 0px" title="" width=464 VIEWASTEXT>
							<PARAM NAME="_ExtentX" VALUE="12277">
							<PARAM NAME="_ExtentY" VALUE="1535">
							<PARAM NAME="AutoReplaceFile" VALUE="-1">
							<PARAM NAME="BackColor" VALUE="16777215">
							<PARAM NAME="BackStyle" VALUE="0">
							<PARAM NAME="Enabled" VALUE="-1">
							<PARAM NAME="FileSelectorFilter" VALUE="Todos los archivos (*.*)|*.*">
							<PARAM NAME="ForeColor" VALUE="-2147483640">
							<PARAM NAME="FTPServer" VALUE="<%=gsFTPServer%>">
							<PARAM NAME="ServerDirectory" VALUE="<%=gsFTPDirectory%>">
							<PARAM NAME="ShowSuccessTransferMsg" VALUE="-1"></OBJECT>
					</TD>
				</TR>
				<FORM name="frmUpload" action="./exec/import_transactions.asp" method=post LANGUAGE=javascript onsubmit="return frmUpload_onsubmit()">
				<TR class=applicationTableRowDivisor>
					<TD colspan=2 valign=top>
						<b>Paso 2. Información sobre las pólizas y parámetros de importación</b>
					</TD>
				</TR>	
				<TR>
					<TD nowrap>
						Fecha de elaboración para las pólizas:
					</TD>
					<TD nowrap>
						<b><%=gsElaborationDate%></b>
					</TD>
				</TR>	
				<TR>
				  <TD nowrap>
						Tipo de pólizas que contiene el archivo:
					</TD>		
				  <TD nowrap>
						<SELECT name="cboStdAccountTypes" style="WIDTH: 280px">
						 <OPTION value=1 
				      selected>Contabilidad bancaria</OPTION>
						 <OPTION value=2>Contabilidad fiduciaria</OPTION>
						</SELECT>
					</TD>
				</TR>	
				<TR>
				  <TD nowrap>
						Tipo de transacciones que contiene el archivo:
					</TD>	
				  <TD nowrap>
						<SELECT name=cboVoucherTypes style="WIDTH: 280px">
							<%=gsCboVoucherTypes%>
						</SELECT>
					</TD>
				</TR>
				<TR>
				  <TD nowrap>
						¿Distribuir las pólizas importadas hacia los <br>grupos de trabajo correspondientes?
					</TD>
					<TD>
						<input type="checkbox" name=chkForwardToUsers value=1>
					</TD>		
				</TR>
				<TR>
				  <TD nowrap>
						¿Generar en forma automática las cuentas auxiliares que <br>aún no hayan sido dadas de alta?
					</TD>
					<TD>
						<input type="checkbox" name=chkAutoGenerateSubsidiaryAccounts value=1>
					</TD>		
				</TR>
				<TR>
				  <TD nowrap>
						¿Los movimientos de las pólizas importadas podrán <br>modificarse con el editor de pólizas?
					</TD>
					<TD>
						<input type="checkbox" name=chkProtectPostings value=1 checked>
					</TD>		
				</TR>
				<TR class=applicationTableRowDivisor>
					<TD colspan=2 valign=top>
						<b>Paso 3. Iniciar el proceso de importación</b>
					</TD>
				</TR>				
				<TR>
				  <TD>						
						Una vez que el archivo se envió y que toda la información solicitada <br>esté correcta,
						haga clic en el botón 'Importar' para iniciar el proceso. <br><br>¡Gracias!						
					</TD>
					<TD>
						<INPUT type=hidden name="txtElaborationDate" width="80%" style="HEIGHT: 22px; WIDTH: 150px" value="<%=Now%>">
						<INPUT type="hidden" name="txtFileName">
						<INPUT type="submit" class=cmdSubmit name=cmdRead value="Importar pólizas" style="WIDTH: 140px">
					</TD>
				</TR>
				</FORM>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>