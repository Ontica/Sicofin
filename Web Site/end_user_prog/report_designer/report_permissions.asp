<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTemplateFile, gsTitle
	Dim gnReportId, gsReportName, gsReportDescription, gsCboDataObjects, gsTackedWindows
	Dim gsHelpFile, gsIconFile, gsCboReportStatus
		
	Call Main()
	
	Sub Main()
		Dim oReportMgr, oRecordset, nSelectedVoucherType
		'***********************************************
		'On Error Resume Next
		
		Call FileServerSettings()
		
		Set oReportMgr = Server.CreateObject("AOReportsMgr.CManager")
		If (Len(Request.QueryString("id")) <> 0) Then
			gnReportId = CLng(Request.QueryString("id"))
			Set oRecordset = oReportMgr.GetReport(Session("sAppServer"), CLng(gnReportId))
			gsReportName				= oRecordset("reportName")
			gsReportDescription = oRecordset("reportDescription")
			gsCboDataObjects		= oReportMgr.CboEntityDataObjects(Session("sAppServer"), CLng(oRecordset("reportEntityId")), _
																														CLng(oRecordset("reportObjectId")))
			gsTemplateFile			= oRecordset("reportTemplate")
			gsHelpFile					= oRecordset("reportHelp")
			gsIconFile					= oRecordset("reportIcon")
			gsCboReportStatus   = oReportMgr.CboReportStatus(Session("sAppServer"), CStr(oRecordset("status")))
			gsTitle						  = "Edición de permisos"
		Else
			gnReportId = 0
			gsCboDataObjects    = oReportMgr.CboEntityDataObjects(Session("sAppServer"), 0)
			gsCboReportStatus   = oReportMgr.CboReportStatus(Session("sAppServer"))
			gsTitle						  = "Edición de permisos"
		End If	
		gsTackedWindows = Request.Form("txtTackedWindows")
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("./exec/exception.asp")
		End If		  
	End Sub		

	Sub FileServerSettings()
		Dim oFileServer
		'*********************
		Set oFileServer = Server.CreateObject("EGEFileManager.CDirectory")
		gsFTPServer    = oFileServer.FTPServer
		gsFTPDirectory = "./"
		Set oFileServer = Nothing
	End Sub
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Diseñador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var gbSended = false;

function oFile() {
	var ftpServer;
	var	ftpDirectory;
	var	ftpTransferTo;
  var fileName;
}

function pickFile(sFileToPick, oTarget) {
	var sURL, sPars;
	var retValue;


	sURL  = "../pickers/file_uploader.htm";
	sPars = "dialogHeight:195px;dialogWidth:490px;resizable:no;scroll:no;status:no;help:no;";
		
	oFile.ftpServer		 = '<%=gsFTPServer%>';
	oFile.ftpDirectory = '';
	
	oFile.ftpTransferTo = '/banobras/contabilidad/reports';
	switch (sFileToPick) {
		case 'templateFile':
			oFile.ftpTransferTo += '/templates/';
			break;
		case 'iconFile':
			oFile.ftpTransferTo += '/icons/';
			break;
		case 'helpFile':
			oFile.ftpTransferTo += '/help/';
			break;
	}
  
	if (window.showModalDialog(sURL, oFile, sPars)) {
		switch (sFileToPick) {
			case 'templateFile':
				document.all.divTemplateFile.innerHTML = '<b>' +  oFile.fileName + '</b>';
				document.all.txtTemplateFile.value = oFile.fileName;
				return false;
			case 'iconFile':
				document.all.divIconFile.innerHTML = '<b>' +  oFile.fileName + '</b>';
				document.all.txtIconFile.value = oFile.fileName;
				return false;
			case 'helpFile':
				document.all.divHelpFile.innerHTML = '<b>' +  oFile.fileName + '</b>';
				document.all.txtHelpFile.value = oFile.fileName;
				return false;
		}
	}
}

function doSubmit() {
	var sMsg, nVoucherId;

	if (document.all.txtName.value == '') {
		alert('Requiero el nombre del reporte.');
		document.all.txtName.focus();
		return false;
	}	
	if (document.all.cboReportTypes.value != 'E') {
		alert('La creación de reportes con el tipo seleccionado aún no está disponible.');
		document.all.cboReportTypes.focus();
		return false;
	}	
	if (document.all.txtTemplateFile.value == '') {
		alert('Requiero se seleccione el archivo con el patrón de diseño del reporte.');
		return false;
	}	
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			<%=gsTitle%>
		</TD>
	  <TD align=right nowrap>
	    <% If CLng(gnReportId) <> 0 Then %>			<A href="" onclick="return(notAvailable());">Definir columnas</A>			&nbsp; | &nbsp;			<A href="" onclick="return(notAvailable());">Permisos</A>			<% End If %>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<FORM name=frmSend action='./exec/save_report.asp' method=post>
			<TABLE class=applicationTable>
			  <TR>
					<TD valign=top nowrap>Nombre del reporte:</TD>
			    <TD colspan=3 width=100%>
						<TEXTAREA name=txtName ROWS=2 style="width:400px"><%=gsReportName%></TEXTAREA><br>
						<INPUT type=button class=cmdSubmit name=cmdCheckSpelling value="Revisar ortografía" onclick="return cmdCheckSpelling_onclick()">
			    </TD>
			  </TR>
			  <TR>
					<TD valign=top nowrap>Descripción:</TD>
			    <TD colspan=3 width=100%>
						<TEXTAREA name=txtDescription ROWS=3 style="width:400px"><%=gsReportDescription%></TEXTAREA><br>
						<INPUT type=button class=cmdSubmit name=cmdCheckSpelling value="Revisar ortografía" onclick="return cmdCheckSpelling_onclick()">
			    </TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Sistema al que pertenece:</TD>
					<TD colspan=3>
						<SELECT name=cboReportSystems style="width:300">
							<OPTION value=1>Sistema de contabilidad financiera</OPTION>
						</SELECT>
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Entidad de la información:</TD>
					<TD colspan=3>
						<SELECT name=cboDataEntities style="width:300">
							<OPTION value=19>Base de conocimiento contable</OPTION>
						</SELECT>
					</TD>
			  </TR>			  
			  <TR nowrap>
					<TD valign=top nowrap>Origen de la información:</TD>
					<TD colspan=3>
						<SELECT name=cboDataObjects style="width:300">
							<OPTION value=0>-- Mixta (Proviene de diferentes orígenes) --</OPTION>
							<%=gsCboDataObjects%>
						</SELECT>
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Tipo del reporte:</TD>
					<TD colspan=3>
						<SELECT name=cboReportTypes style="width:300">
							<OPTION value="E" selected>Microsoft® Excel</OPTION>
							<OPTION value="W">Microsoft® Word</OPTION>							
							<OPTION value="X">Documento XML</OPTION>
							<OPTION value="H">Documento HTML con hoja de estilo</OPTION>
							<OPTION value="S">Documento HTML simple</OPTION>
							<OPTION value="T">Archivo de texto</OPTION>							
						</SELECT>
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Archivo con el patrón de diseño:</TD>
					<TD valign=top><span id=divTemplateFile><b><%=gsTemplateFile%></b></span></TD>
					<TD colspan=2 align=right>
						<INPUT class=cmdSubmit type=button name=cmdFileTransfer value="Archivo ..." onclick="return(pickFile('templateFile', this));">
						&nbsp; 
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Archivo con la ayuda del reporte:</TD>
					<TD valign=top><span id=divHelpFile><b><%=gsHelpFile%></b></span></TD>
					<TD colspan=2 align=right>
						<INPUT class=cmdSubmit type=button name=cmdFileTransfer value="Archivo ..." onclick="return(pickFile('helpFile', this));">
						&nbsp; 
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Archivo con el icono del reporte:</TD>
					<TD valign=top><span id=divIconFile><b><%=gsIconFile%></b></span></TD>
					<TD colspan=2 align=right>
						<INPUT class=cmdSubmit type=button name=cmdFileTransfer value="Archivo ..." onclick="return(pickFile('iconFile', this));">
						&nbsp; 
					</TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top nowrap>Estado del reporte:</TD>
					<TD colspan=3>
						<SELECT name=cboReportStatus style="width:300">
							<%=gsCboReportStatus%>
						</SELECT>
					</TD>
			  </TR>		  	  
			  <TR>
			    <td colspan=4 nowrap align=right>
			     <br>
					 <INPUT type="hidden" name=txtReportId value="<%=gnReportId%>">
					 <INPUT type="hidden" name=txtTemplateFile value="<%=gsTemplateFile%>">
					 <INPUT type="hidden" name=txtHelpFile value="<%=gsHelpFile%>">			
					 <INPUT type="hidden" name=txtIconFile value="<%=gsIconFile%>">					 
					 <INPUT TYPE=hidden name=txtTackedWindows>
						<% If CLng(gnReportId) = 0 Then %>
						<INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Crear" onclick="doSubmit();">
						<% Else %>
						<INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Guardar" onclick="doSubmit();">
						<% End If %>
						&nbsp; &nbsp;
			     <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
			     &nbsp;
			    </td>
			  </TR>
			</TABLE>
			</FORM>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>