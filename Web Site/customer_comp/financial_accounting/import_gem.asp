<%@ Language=VBScript %>
<%
	Option Explicit
		
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If


  Dim sSourceAppServer
  sSourceAppServer = "GemPyC"
  
  Dim gsTackedWindows
  Dim gnGemPendingTransactions, gnGemErrorTransactions, gsTransactionsLastImpDate, gsDaysSinceLastTransactionsImpDate
  Dim gsImportTimer
  
  Call Main()
  
  Sub Main()
		Dim iGemPyC, dTemp
		'*****************

		Set iGemPyC = Server.CreateObject("SCFIGemPyC.CInterface")
		
		gnGemPendingTransactions  =	iGemPyC.GEMPendingTransactionsCount(CStr(sSourceAppServer))
		gnGemErrorTransactions    = iGemPyC.GEMErrorTransactionsCount(CStr(sSourceAppServer))
		gsTransactionsLastImpDate = iGemPyC.TransactionsLastImportationDate(CStr(sSourceAppServer))
		gsImportTimer							= iGemPyC.ImportTimer(CStr(sSourceAppServer))
				
		If IsDate(gsTransactionsLastImpDate) Then
			dTemp	= Now() - CDate(gsTransactionsLastImpDate)
			gsDaysSinceLastTransactionsImpDate = FormatSinceDate(dTemp)
		Else
			gsDaysSinceLastTransactionsImpDate = "Indeterminado"
		End If
		
  End Sub
    
	Function FormatSinceDate(dDate)
		Dim sTemp
		If (Int(dDate) = 1) Then
			sTemp = "1 día, "
		Else
			sTemp = Int(dDate) & " días, "
		End If
		If (Hour(dDate) = 1) Then
			sTemp = sTemp & "1 hora, "
		Else
			sTemp = sTemp & Hour(dDate) & " horas, "
		End If
		If (Minute(dDate) = 1) Then
			sTemp = sTemp & "1 minuto, "
		Else
			sTemp = sTemp & Minute(dDate) & " minutos, "
		End If
		If (Second(dDate) = 1) Then
			sTemp = sTemp & "1 segundo."
		Else
			sTemp = sTemp & Second(dDate) & " segundos."
		End If		
		FormatSinceDate = sTemp
	End Function
			
%>

<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oImportWindow = null, oFilesWindow = null;

function importGEM() {	
	var sURL = './exec/import_gem.asp';
	var sOpt = 'height=200,width=420,status=no,toolbar=no,menubar=no,location=no';
	
  oImportWindow = window.open(sURL, null, sOpt);
	return false;
}

function retriveFiles() {
	var sURL = 'imported_gem_files.asp';
	var sOpt = 'height=300,width=420,status=no,toolbar=no,menubar=no,location=no';
	
  oImportWindow = window.open(sURL, null, sOpt);
	return false;	
}

function window_onunload() {
	if (oImportWindow != null && !oImportWindow.closed) {
		oImportWindow.close();
	}
	if (oFilesWindow != null && !oFilesWindow.closed) {
		oFilesWindow.close();
	}	
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload="window_onunload();" topmargin=0>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Importador de pólizas del sistema GEM
		</TD>
		<TD colspan=3 align=right nowrap>
			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">								</TD>
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
						<A href="import.asp">Importador de pólizas mediante archivo</A>
						&nbsp; | &nbsp;
						<A href="voucher_explorer.asp">Explorador de pólizas</A>
						&nbsp; | &nbsp;
						<A href="../../balances/balance_explorer.asp">Explorador de saldos</A>
						&nbsp; | &nbsp;
						<A href="../../reports/balances.asp">Balanzas de comprobación</A>
						&nbsp; | &nbsp;
						<A href="../../reports/other_reports.asp">Reportes contables</A>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<TR class=fullScrollMenuHeader>
					<TD colspan=2>						
				    <P>Este programa permite traspasar las pólizas del sistema GEM (Gobiernos, estados y municipios) al 
				    sistema de contabilidad financiera.
				    </P>
				    <P>El administrador de flujo de trabajo realiza esta tarea en forma automática cada <b><%=gsImportTimer%></b>.<br>
				    Sin embargo, con esta herramienta es posible importar las pólizas en forma manual.
				    </P>      
				  </TD>
				</TR>
				<TR>
				  <TD colspan=2 valign=top nowrap>
						<br>
						Fecha de la última importación:
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<b><%=gsTransactionsLastImpDate%></b>
						<br><br>
						Tiempo transcurrido desde esa fecha: 
						&nbsp; &nbsp;&nbsp; &nbsp; &nbsp;
						<b><%=gsDaysSinceLastTransactionsImpDate%></b>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type="button" name=cmdRetriveFiles value="Recuperar los archivos con los resultados de las importaciones realizadas ..." style="WIDTH: 380px" onclick='retriveFiles();'>
					</TD>
				</TR>				
				<TR>
					<TD>
						<br>
						Número de pólizas pendientes de importar: 
						<% If (gnGemPendingTransactions > 1) Then %>
							<b><%=gnGemPendingTransactions%> pólizas</b>
						<% ElseIf (gnGemPendingTransactions = 1) Then %>
							<b>Una póliza</b>
						<% ElseIf (gnGemPendingTransactions = 0) Then %>
							<b>Ninguna</b>
						<% Else  %>
							<b>No pude determinar el número de pólizas pendientes</b>
						<% End If %>
						<% If (gnGemPendingTransactions <> 0) Then %>
						<br><br>						
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type="button" name=cmdPendingReport value="Obtener informe" style="WIDTH: 120px" onclick='notAvailable();'>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type="button" name=cmdImport value="Importar pólizas" style="WIDTH: 120px" onclick="return importGEM();">						
						<br>&nbsp;
						<% End If %>
					</TD>
				</TR>
				<TR>
				  <TD colspan=2 valign=top>
						<br>
						Número de pólizas que se han intentado<br>
						importar pero que contienen errores: 
						&nbsp; &nbsp; &nbsp; &nbsp;						
						<% If (gnGemErrorTransactions > 1) Then %>
							<b><%=gnGemErrorTransactions%> pólizas</b>
						<% ElseIf (gnGemErrorTransactions = 1) Then %>
							<b>Una póliza</b>
						<% ElseIf (gnGemErrorTransactions = 0) Then %>
							<b>Ninguna</b>
						<% Else  %>
							<b>No pude determinar el número de pólizas con error</b>
						<% End If %>
						<% If (gnGemErrorTransactions <> 0) Then %>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						<INPUT class=cmdSubmit type="button" name=cmdErrorReport value="Obtener informe" style="WIDTH: 120px" onclick='notAvailable();'>
						<br>&nbsp;
						<% End If %>
					</TD>
				</TR>								
			</TABLE>
		</TD>
	</TR>	
</TABLE>
</BODY>
</HTML>