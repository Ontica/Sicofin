<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gsFTPServer, gsFTPDirectory, gsTransferToPath, gsTemplateFile, gsTitle
	Dim gnReportId, gsReportName, gsReportDescription, gsReportKeywords, gsReportCategories, gsTackedWindows
	Dim gsHelpFile, gsIconFile, gsStatus, gsCboTechnologies, gsCboReportStructures, gsCboClasses
	Dim gsCboSubClasses, gsCboDataItemOrders, gsReportLayout
		
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oDictionary, oRecordset, nSelectedVoucherType
		'****************************************************
		'On Error Resume Next
		
		Call FileServerSettings()
		
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
		Set oDictionary     = Server.CreateObject("AOReportsDesigner.CDictionary")
		gsTransferToPath = oReportDesigner.FilesPath		
		If (Len(Request.QueryString("id")) <> 0) Then
			gnReportId = CLng(Request.QueryString("id"))
			Set oRecordset       = oReportDesigner.GetReport(Session("sAppServer"), CLng(gnReportId))
			gsReportName				 = oRecordset("reportName")
			gsReportDescription  = oRecordset("reportDescription")
			gsReportKeywords     = oRecordset("reportKeywords")
			gsReportCategories   = oReportDesigner.CboReportCategories(Session("sAppServer"), CLng(oRecordset("reportCategoryId")))
			gsCboClasses         = oDictionary.CboClasses(Session("sAppServer"), CLng(oRecordset("reportDataClassId")))
			gsCboSubClasses      = oDictionary.CboSubClasses(Session("sAppServer"), CLng(oRecordset("reportDataClassId")), _
																											 CLng(oRecordset("reportDataSubClassId")))
			gsCboDataItemOrders  = oDictionary.CboDataItemOrders(Session("sAppServer"), CLng(oRecordset("reportDataClassId")), _
																												   CLng(Session("uid")), CLng(oRecordset("reportDataOrderId")))
			gsCboTechnologies		 = oReportDesigner.CboReportTechnologies(oRecordset("reportTechnology"))
			gsTemplateFile			 = oRecordset("reportTemplateFile")
			gsHelpFile					 = oRecordset("reportHelpFile")
			gsIconFile					 = oRecordset("reportIconFile")
			gsStatus						 = oRecordset("reportStatus")
			If Len(gsReportName) > 76 Then
				gsTitle	= Left(gsReportName, 76) & "..."
			Else
				gsTitle	= gsReportName
			End If
			If gsStatus = "S" Then
				gsTitle = gsTitle & " (suspendido)"
			End If
		Else
			gnReportId = 0
			gsReportCategories = oReportDesigner.CboReportCategories(Session("sAppServer"))
			gsCboClasses       = oDictionary.CboClasses(Session("sAppServer"))
			gsCboClasses       = "<OPTION value=0>-- Seleccionar una estructura de información --</OPTION>" & gsCboClasses
			gsCboTechnologies	 = oReportDesigner.CboReportTechnologies()			
			gsTitle						 = "Nuevo reporte"
		End If		
		
		If Len(gsCboSubClasses) = 0 Then
			gsCboSubClasses = "<OPTION value=0>-- No existen subestructuras para la estructura seleccionada --</OPTION>"
		End If
		If Len(gsCboDataItemOrders) = 0 Then
			gsCboDataItemOrders  = "<OPTION value=0>-- No existen ordenamientos para la estructura seleccionada --</OPTION>"
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

	sURL  = "/empiria/general/file_uploader/file_uploader.htm";
	sPars = "dialogHeight:195px;dialogWidth:490px;resizable:no;scroll:no;status:no;help:no;";
		
	oFile.ftpServer		 = '<%=gsFTPServer%>';
	oFile.ftpDirectory = '';	
	
	oFile.ftpTransferTo = '<%=gsTransferToPath%>';	
	switch (sFileToPick) {
		case 'templateFile':
			oFile.ftpTransferTo += 'templates/';
			break;
		case 'iconFile':
			oFile.ftpTransferTo += 'icons/';
			break;
		case 'helpFile':
			oFile.ftpTransferTo += 'help/';
			break;
	}
  
	if (window.showModalDialog(sURL, oFile, sPars)) {
		switch (sFileToPick) {
			case 'templateFile':				
				document.all.txtTemplateFile.value = oFile.fileName;
				return false;
			case 'iconFile':				
				document.all.txtIconFile.value = oFile.fileName;
				return false;
			case 'helpFile':				
				document.all.txtHelpFile.value = oFile.fileName;
				return false;
		}
	}
}

function updateCboSubClasses() {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "CboSubClasses", document.all.cboClasses.value, 0);
	document.all.divSubClasses.innerHTML = obj.return_value;
}

function updateCboDataItemOrders() {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "CboDataItemOrders", document.all.cboClasses.value, 0);
	document.all.divDataOrders.innerHTML = obj.return_value;
}

function refreshDataItemCombos() {
	updateCboSubClasses();
	updateCboDataItemOrders();
}

function pickData(sData, oTarget) {
	switch(sData) {
		case 'reportLayout':
			alert("Report Layout");
			return;
	}	
}

function doSubmit() {
	var sMsg, nVoucherId;

	if (document.all.txtName.value == '') {
		alert('Requiero el nombre del reporte.');
		showStep(1);
		document.all.txtName.focus();
		return false;
	}		
	if (document.all.cboReportCategories.value == 0) {
		alert('Necesito se seleccione la categoría a la que pertenece el reporte.');
		showStep(2);
		document.all.cboReportCategories.focus();
		return false;
	}		
	if (document.all.cboClasses.value == 0) {
		alert('Necesito se seleccione el origen de la información del reporte.');
		showStep(2);
		document.all.cboClasses.focus();
		return false;
	}
	if (document.all.cboReportTechnologies.value == 0) {
		alert('Se necesita seleccionar la tecnología con la que se construirá el reporte.');
		showStep(3);
		document.all.cboReportTechnologies.focus();
		return false;
	}	
	if (document.all.cboReportTechnologies.value != 'E' && document.all.cboReportTechnologies.value != 'T') {
		alert('Por el momento sólo se pueden generar reportes en MS Excel® o en archivos de texto.');
		showStep(3);
		document.all.cboReportTechnologies.focus();
		return false;
	}
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function deleteReport() {
	var obj;		
	if (confirm("¿Elimino el reporte <%=gsReportName%>?")) {
		obj = RSExecute("../end_user_prog_scripts.asp", "DeleteReport", <%=gnReportId%>);
		window.opener.location.href = window.opener.location.href;
		window.close();
	}
}

function suspendReport() {
	var sMsg;
	
	sMsg  = 'Al suspender el reporte no será posible crearlo con el "Generador de reportes" sino hasta que\n';
	sMsg += 'vuelva a ser activado.\n\n';
	sMsg += "¿Suspendo la generación de este reporte?";
	if (confirm(sMsg)) {
		obj = RSExecute("../end_user_prog_scripts.asp", "SuspendReport", <%=gnReportId%>);
		window.opener.location.href = window.opener.location.href;
		window.close();
	}
}

function activateReport() {
	var sMsg;
	
	sMsg  = 'Los reportes activos pueden ser invocados con el "Generador de reportes" por los usuarios\n';
	sMsg += 'que tengan derechos sobre el mismo.\n\n';
	sMsg += "¿Activo la generación de este reporte?";	
	if (confirm(sMsg)) {		
		obj = RSExecute("../end_user_prog_scripts.asp", "ActivateReport", <%=gnReportId%>);
		window.opener.location.href = window.opener.location.href;
		window.close();
	}
}

function showStep(nStep) {
	switch(nStep) {	
		case 1:
			document.all.step1.style.display = 'inline';
			document.all.stepTitle1.style.fontWeight = 'bold';
			document.all.step2.style.display = 'none';
			document.all.stepTitle2.style.fontWeight = 'normal';
			document.all.step3.style.display = 'none';			
			document.all.stepTitle3.style.fontWeight = 'normal';
			return false;
		case 2:
			document.all.step1.style.display = 'none';
			document.all.stepTitle1.style.fontWeight = 'normal';
			document.all.step2.style.display = 'inline';
			document.all.stepTitle2.style.fontWeight = 'bold';
			document.all.step3.style.display = 'none';
			document.all.stepTitle3.style.fontWeight = 'normal';
			return false;
		case 3:
			document.all.step1.style.display = 'none';
			document.all.stepTitle1.style.fontWeight = 'normal';
			document.all.step2.style.display = 'none';
			document.all.stepTitle2.style.fontWeight = 'normal';
			document.all.step3.style.display = 'inline';
			document.all.stepTitle3.style.fontWeight = 'bold';
			return false;
	}		
}

function cboReportTechnologies_onchange() {
	if (document.all.cboReportTechnologies.value == 'E') {
		document.all.cmdTemplateFile.disabled = false;
	} else {
		document.all.cmdTemplateFile.disabled = true;
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onload='showStep(1);'>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Diseñador de reportes
		</TD>
	  <TD align=right nowrap>					<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;			
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<FORM name=frmSend action='./exec/save_report.asp' method=post>
			<TABLE class='fullScrollMenu'>				
				<TR class="fullScrollMenuHeader">
					<TD class=fullScrollMenuTitle nowrap>
						<%=gsTitle%>						
					</TD>
					<TD align=right nowrap>
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='document.all.frmSend.reset();' alt="Refrescar">						
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
			  <TR class=applicationTableRowDivisor>
					<TD valign=top nowrap colspan=4>
						<a id=stepTitle1 href='' onclick='return(showStep(1));'>Datos generales</a>
						&nbsp; &nbsp; &nbsp; | &nbsp; &nbsp; &nbsp;
						<a id=stepTitle2 href='' onclick='return(showStep(2));'>Origen de la información</a>						
						&nbsp; &nbsp; &nbsp; | &nbsp; &nbsp; &nbsp;
						<a id=stepTitle3 href='' onclick='return(showStep(3));'>Formato del reporte</a>
						</TD>
			  </TR>
			</TABLE>
			<TABLE id=step1 class=applicationTable style='display:none;' height=240>	
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Identificación del reporte</b></TD>
				</TR>			
				<TR>
					<TD valign=top nowrap>Nombre del reporte:</TD>
				  <TD colspan=3 width=100%>
						<TEXTAREA name=txtName rows=2 style="width:430px"><%=gsReportName%></TEXTAREA>
				  </TD>
				</TR>
				<TR>
					<TD valign=top nowrap>Descripción:</TD>
				  <TD colspan=3 width=100%>
						<TEXTAREA name=txtDescription rows=4 style="width:430px"><%=gsReportDescription%></TEXTAREA>
				  </TD>
				</TR>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Información de apoyo para los usuarios del reporte</b></TD>
				</TR>
				<TR>
					<TD valign=top nowrap>Palabras clave:<br>(separadas por espacios)</TD>
				  <TD colspan=3 width=100%>
						<TEXTAREA name=txtKeywords rows=4 style="width:430px"><%=gsReportKeywords%></TEXTAREA>
				  </TD>
				</TR>								
				<TR nowrap>
					<TD valign=top nowrap>Archivo de ayuda:</TD>
					<TD valign=top colspan=3>
						<INPUT name=txtHelpFile style="width:275px" value='<%=gsHelpFile%>' readonly>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFileTransfer style='width:65' value="Archivo ..." onclick="return(pickFile('helpFile', this));">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdClean style='width:65' value="Limpiar" onclick="document.all.txtHelpFile.value='';" >
					</TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap> Icono del reporte:</TD>
					<TD valign=top colspan=3>					
						<INPUT name=txtIconFile style="width:275px" value='<%=gsIconFile%>' readonly>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFileTransfer style='width:65' value="Archivo ..." onclick="return(pickFile('iconFile', this));">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdClean style='width:65' value="Limpiar" onclick="document.all.txtIconFile.value='';">
					</TD>
				</TR>
				<TR nowrap height=100%>
					<TD valign=top colspan=4>&nbsp;</TD>
				</TR>				
			</TABLE>			  
			<TABLE id=step2 class=applicationTable style='display:none;' height=240>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Clasificación del reporte</b></TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap>Clasificar el reporte en la categoría:</TD>
					<TD colspan=3 width=400>
						<SELECT name=cboReportCategories style="width:330">
						  <OPTION value=0>-- Seleccionar la categoría del reporte --</OPTION>
							<%=gsReportCategories%>
						</SELECT>
						<br>&nbsp;
					</TD>
				</TR>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>¿Qué información desplegará el reporte y en qué orden?</b></TD>
				</TR>								
				<TR nowrap>
					<TD valign=top>Desplegar la información contenida en la estructura de datos:</TD>
					<TD colspan=3>
						<SELECT name=cboClasses style="width:330" onchange='refreshDataItemCombos();'>
							<%=gsCboClasses%>
						</SELECT>
						<br>
					</TD>
				</TR>			  
				<TR nowrap>
					<TD valign=top>Restringir la información de la estructura de datos seleccionada empleando:</TD>
					<TD colspan=3>
						<div id=divSubClasses>
						<SELECT name=cboSubClasses style="width:330">
							<%=gsCboSubClasses%>
						</SELECT>
						</div>
						<br>
					</TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap>Ordenar la información del reporte utilizando:</TD>
					<TD colspan=3>					
						<div id=divDataOrders>
						<SELECT name=cboDataOrders style="width:330">
							<%=gsCboDataItemOrders%>
						</SELECT>
						</div>
						<br>
					</TD>
				</TR>				
				<TR nowrap height=100%>
					<TD valign=top colspan=4>&nbsp;</TD>
				</TR>				
			</TABLE>
			<TABLE id=step3 class=applicationTable style='display:none;' height=240>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Tecnología que se empleará en la construcción del reporte</b></TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap>Obtener el reporte en archivos:</TD>
					<TD colspan=3>					
						<SELECT name=cboReportTechnologies style="width:270" onchange="return cboReportTechnologies_onchange()">
							<OPTION value=0>-- Seleccionar la tecnología del reporte --</OPTION>
							<%=gsCboTechnologies%>
						</SELECT>
						<br>&nbsp;
					</TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap>Archivo con el reporte prediseñado:</TD>
					<TD valign=top colspan=3 nowrap>
						<INPUT name=txtTemplateFile style="width:270px" value='<%=gsTemplateFile%>' readonly>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;						
						<INPUT class=cmdSubmit type=button name=cmdTemplateFile style='width:130' value="Archivo prediseñado ..." onclick="return(pickFile('templateFile', this));">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdCleanTemplateFile style='width:70' value="Limpiar" onclick="document.all.txtTemplateFile.value='';">						
						<br>&nbsp;
					</TD>
				</TR>
				<TR class=applicationTableRowDivisor nowrap>
					<TD valign=top colspan=4><b>Estructura y formato de presentación del reporte</b></TD>
				</TR>
				<TR nowrap>
					<TD valign=top nowrap>Parámetros de presentación del reporte:</TD>
					<TD colspan=3>
						<TEXTAREA name=txtLayout rows=3 style="width:270px" readonly><%=gsReportLayout%></TEXTAREA>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdReportLayout style='width:70' value="Modificar ..." onclick="return(pickData('reportLayout', this));">
						<br>&nbsp;						
					</TD>
				</TR>				
				<TR nowrap height=100%>
					<TD valign=top colspan=4>&nbsp;</TD>
				</TR>
			</TABLE>
			<TABLE width=100%>
				<TR>
					<TD align=right>
					 <% If CLng(gnReportId) <> 0 Then %>
						<INPUT class=cmdSubmit type=button name=cmdDelete style='width:70' value="Eliminar" onclick='deleteReport();'>
						&nbsp; &nbsp; &nbsp; &nbsp;
							<% If gsStatus <> "S" Then %>
							<INPUT class=cmdSubmit type=button name=cmdSuspend style='width:70' value="Suspender" onclick="suspendReport();">
							&nbsp; &nbsp; &nbsp; &nbsp;
							<% Else %>
							<INPUT class=cmdSubmit type=button name=cmdActivate style='width:70' value="Activar" onclick="activateReport();">
							&nbsp; &nbsp; &nbsp; &nbsp;
							<% End If %>
						<% End If %>
						<INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Guardar" onclick="doSubmit();">						
						&nbsp;
				   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
				   &nbsp; &nbsp;
				  </TR>
				</TR>
			</TABLE>
			<INPUT type="hidden" name=txtReportId value="<%=gnReportId%>">			
			<INPUT TYPE=hidden name=txtTackedWindows>
			</FORM>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>