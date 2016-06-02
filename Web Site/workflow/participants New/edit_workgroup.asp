<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnWorkgroupId, gbEditMode, gsTitle
	Dim gsName, gsHistoricDate
	Dim gsParticipantKey, gsParticipantName, gsDescription, gsStatus
			
	Call Main()
		
	Sub Main()
		Dim oParticipant, oRecordset
		'***************************
		On Error Resume Next
		
		If (Len(Request.QueryString("id")) <> 0) Then
			gnWorkgroupId = Request.QueryString("id")
			gbEditMode = True
		Else
			gnWorkgroupId = 0
			gbEditMode = False
		End If
				
		If gbEditMode Then
			Set oParticipant   = Server.CreateObject("MHParticipantsMgr.CParticipant")
			Set oRecordset     = oParticipant.Attributes(Session("sAppServer"), CLng(gnWorkgroupId))		  
		  gsParticipantName  = oRecordset("participantName")
		  gsTitle						 = Left(gsParticipantName, 30)
		  gsDescription      = oRecordset("description")
			gsParticipantKey	 = oRecordset("participantKey")			
			gsStatus					 = oRecordset("status")
			Set oRecordset     = Nothing
			gsName = gsParticipantName
			If gsStatus = "S" Then													
				gsTitle = gsTitle & " (Suspendido)"				
			ElseIf gsStatus = "D" Then				
				gsTitle = gsTitle & " (Eliminado)"
			End If						
			Set oParticipant = Nothing						
		Else
			gsTitle			 = "Nuevo grupo de trabajo"
		End If		
		If (Err.number <> 0) Then					
			Session("errNumber") = Err.number
			Session("errDesc")   = Err.description
			Session("errSource") = Err.source
			Err.Clear
			Response.Redirect(Application("errorPage"))
		End If		
	End Sub			
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Administración de grupos de trabajo</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oPermissionsWindow = null; 

function whoHasUserKey(sKey, nUserId) {
	var obj;
	obj = RSExecute("../workflow_scripts.asp", "WhoHasUserKey", sKey, nUserId);
	return obj.return_value;
}

function save() {		
	if (document.all.frmSend.txtName.value == "") {
		alert("Necesito el nombre del grupo de trabajo o rol.");
		document.all.frmSend.txtName.focus();
		return false;
	}
	if (document.all.frmSend.txtParticipantKey.value == "") {
		alert("Necesito el identificador del grupo de trabajo o rol.");
		document.all.frmSend.txtParticipantKey.focus();
		return false;
	}	
	document.all.frmSend.submit();
	return true;
}

function reactivate() {
	var sMsg;	
	sMsg  = "Esta operación reactivará el grupo de trabajo en el sistema.\n\n";		
	sMsg += "¿Reactivo el grupo de trabajo '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnWorkgroupId%>&status=1";
	}
}

function suspend() {
	var sMsg;	
	
	sMsg  = "Esta operación suspenderá, en forma temporal, al grupo de trabajo, por lo que sus miembros\n";
	sMsg += "serán también suspendidos.\n\n";
	sMsg += "Los documentos y mensajes de sus miembros permanecerán intactos en las bandejas respectivas,\n";	
	sMsg += "por lo que les será posible recuperarlos cuando el grupo sea reactivado.\n\n";		
	sMsg += "¿Suspendo temporalmente del sistema al grupo '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnWorkgroupId%>&status=2";
	}
}

function delete_() {
	var sMsg;
	sMsg  = "Esta operación eliminará al grupo de trabajo del sistema.\n\n";
	sMsg  = "Sin embargo esta tarea no eliminará del sistema a los miembros del grupo.\n\n";		
	sMsg += "¿Elimino permanentemente del sistema al grupo '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnWorkgroupId%>&status=3";
	}
}

function callEditor(sWindow, nItemId) {
	var sURL, sOpt;
  switch (sWindow) {  
		case 'users':
			sURL = 'workgroup_users.asp?id=' + nItemId;
			window.location.href = sURL;
			return false;			
    case 'tasks':
			sURL = 'workgroup_tasks.asp?id=' + nItemId;
			//sOpt = 'height=360px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			//if (oPermissionsWindow == null || oPermissionsWindow.closed) {
			//	oPermissionsWindow = window.open(sURL, '_blank', sOpt);
			//} else {
			//	oPermissionsWindow.focus();
			//	oPermissionsWindow.navigate(sURL);
			//}
			window.location.href = sURL;
			return false;
	}
	return false;
}

function window_onunload() {
	if (oPermissionsWindow != null && !oPermissionsWindow.closed) {
		oPermissionsWindow.close();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox onunload="return window_onunload()">
<FORM name=frmSend action='./exec/save_workgroup.asp' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			<%=gsTitle%>
		</TD>
		<TD colspan=3 align=right nowrap>
			<img align=absmiddle src='/empiria/images/help_white.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_white.gif' onclick="window.close();" alt="Cerrar">								</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">					
					<TD align=right nowrap>
						<% If (gnWorkgroupId) <> 0 Then %>
					  <a href='' onclick='return(callEditor("users", <%=gnWorkgroupId%>));'>Usuarios</a>
						&nbsp; | &nbsp;
					  <a href='' onclick='return(callEditor("tasks", <%=gnWorkgroupId%>));'>Tareas</a>
						&nbsp; | &nbsp;						
						<A href="" onclick="return(notAvailable());">Imprimir</A>
						&nbsp; &nbsp;
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='document.all.frmSend.reset();' alt="Refrescar">
						<% Else %>
						<img align=absmiddle src='/empiria/images/invisible.gif'>
						<% End If %>						
					</TD>
				</TR>
			</TABLE>
			<TABLE class=applicationTable>
				<TR>
					<TD nowrap>Grupo de trabajo (rol):</TD>
				  <TD>
						<INPUT name=txtName value='<%=gsName%>' style='width:260px'>
					</TD>
				</TR>
				<TR>
					<TD nowrap>Descripción:</TD>
				  <TD>
						<TEXTAREA rows=4 name=txtDescription style='width:260px'><%=gsDescription%></TEXTAREA>
					</TD>
				</TR>				
				<TR>
					<TD>Identificador del grupo:</TD>
				  <TD width=80%>
						<INPUT name=txtParticipantKey value='<%=gsParticipantKey%>' style='width:125px'>
					</TD>
				</TR>
				<TR>
					<TD colspan=2 align=right nowrap>
						<br>
						<% If gnWorkgroupId <> 0 Then %>
							<% If gsStatus = "A" Then %>
						  <INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Suspender' onclick='suspend();' tabindex=-1>
						  <% ElseIf (gsStatus = "S") Or (gsStatus = "D") Then %>
							<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Reactivar' onclick='reactivate();' tabindex=-1>
						  <% End If %>						
						&nbsp;
							<% If (gsStatus <> "D") Then %>
							<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Eliminar' onclick='delete_();' tabindex=-1>
							<% End If %>
						&nbsp; &nbsp; &nbsp; &nbsp;
						<% End If %>
						<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Aceptar' onclick='save();'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdClose style='width:80;' value='Cancelar' onclick='window.close();'>
						<br>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT type=hidden name=txtUserId value=<%=gnWorkgroupId%>>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
