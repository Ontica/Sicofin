<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnUserId, gbEditMode, gsTitle
	Dim gsLastName1, gsLastName2, gsName, gsIsFemale, gsBornDate, gsJobEMail, gsJobPhone, gsHistoricDate
	Dim gsParticipantKey, gsParticipantName, gsStatus
			
	Call Main()
		
	Sub Main()
		Dim oParticipant, oRecordset, oExtAttributes
		'*******************************************
		On Error Resume Next
		
		If (Len(Request.QueryString("id")) <> 0) Then
			gnUserId = Request.QueryString("id")
			gbEditMode = True
		Else
			gnUserId = 0
			gbEditMode = False
		End If
				
		If gbEditMode Then
			Set oParticipant   = Server.CreateObject("MHParticipantsMgr.CParticipant")
			Set oRecordset     = oParticipant.Attributes(Session("sAppServer"), CLng(gnUserId))		  
		  gsParticipantName  = oRecordset("participantName")
		  gsTitle						 = Left(gsParticipantName, 30)
			gsParticipantKey	 = oRecordset("participantKey")
			gsStatus					 = oRecordset("status")
			Set oRecordset     = Nothing
			Set oExtAttributes = oParticipant.ExtendedAttributes(Session("sAppServer"), CLng(gnUserId))
			gsLastName1	= oExtAttributes("LastName1")
			gsLastName2 = oExtAttributes("LastName2")		
			gsName = oExtAttributes("FirstName")							
			If (CBool(oExtAttributes("IsFemale"))) Then
		    gsIsFemale = "checked"
			Else
				gsIsFemale = ""
			End If
			gsJobEMail = oExtAttributes("JobEmail")
			gsJobPhone = oExtAttributes("JobPhone")
			If gsStatus = "S" Then
				If Len(gsIsFemale) <> 0 Then
					gsTitle = gsTitle & " (Suspendida)"
				Else
					gsTitle = gsTitle & " (Suspendido)"
				End If
			ElseIf gsStatus = "D" Then
				If Len(gsIsFemale) <> 0 Then
					gsTitle = gsTitle & " (Eliminada)"
				Else
					gsTitle = gsTitle & " (Eliminado)"
				End If
			End If			
			Set oExtAttributes = Nothing
			Set oParticipant = Nothing						
		Else
			gsTitle			 = "Agregar usuario(a)"			
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
<TITLE>Administración de usuarios</TITLE>
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

function validateDate(date) {
	var obj;
	obj = RSExecute("../workflow_scripts.asp", "IsDate", date);
	return obj.return_value;
}

function save() {
	var sTemp;
	
	if (document.all.frmSend.txtName.value == "") {
		alert("Necesito el nombre del usuario(a).");
		document.all.frmSend.txtName.focus();
		return false;
	}
	if (document.all.frmSend.txtLastName1.value == "") {
		alert("Necesito el apellido paterno del usuario(a).");
		document.all.frmSend.txtLastName1.focus();
		return false;
	}	
	if (document.all.frmSend.txtParticipantKey.value == "") {
		alert("Necesito el identificador de acceso al sistema.");
		document.all.frmSend.txtParticipantKey.focus();
		return false;
	}
	sTemp = whoHasUserKey(document.all.frmSend.txtParticipantKey.value, <%=gnUserId%>);	
	if (sTemp != '') {
		sTemp = 'El identificador de acceso proporcionado está asignado a:\n\n' + sTemp;
		alert(sTemp);
		document.all.frmSend.txtParticipantKey.focus();
		return false;	
	}
	<% If gnUserId = 0 Then %>
	if(document.all.frmSend.txtHistoricDate.value == '') {
		alert("Necesito la fecha a partir de la cual se permitirá el acceso al participante.");
		document.all.frmSend.txtHistoricDate.focus();
		return false;
	}
	if(!validateDate(document.all.frmSend.txtHistoricDate.value)) {
		alert("No reconozco la fecha de acceso al sistema.");
		document.all.frmSend.txtHistoricDate.focus();
		return false;
	}
	<% End If %>
	document.all.frmSend.submit();
	return true;
}

function reactivate() {
	var sMsg;
	<% If Len(gsIsFemale) <> 0 Then %> 
		sMsg  = "Esta operación reactivará a la participante en el sistema.\n\n";
	<% Else %>
		sMsg  = "Esta operación reactivará al participante en el sistema.\n\n";
	<% End If %>
	
	sMsg += "¿Reactivo a '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnUserId%>&status=1";
	}
}

function suspend() {
	var sMsg;
	<% If Len(gsIsFemale) <> 0 Then %> 
		sMsg  = "Esta operación suspenderá, en forma temporal, a la participante, por lo que las tareas pendientes que\n";
	<% Else %>
		sMsg  = "Esta operación suspenderá, en forma temporal, al participante, por lo que las tareas pendientes que\n";
	<% End If %>	
	sMsg += "tenga en este momento serán redistribuidas entre los miembros de sus grupos de trabajo.\n\n";
	sMsg += "Sus documentos y mensajes permanecerán intactos en sus bandejas personales,\n";
	<% If Len(gsIsFemale) <> 0 Then %> 
		sMsg += "por lo que le será posible recuperarlos cuando sea reactivada en el sistema.\n\n";
	<% Else %>
		sMsg += "por lo que le será posible recuperarlos cuando sea reactivado en el sistema.\n\n";
	<% End If %>
	
	sMsg += "¿Suspendo temporalmente del sistema a '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnUserId%>&status=2";
	}
}

function delete_() {
	var sMsg;
	<% If Len(gsIsFemale) <> 0 Then %> 
		sMsg  = "Esta operación eliminará a la participante del sistema, por lo que todas sus tareas pendientes,\n";
	<% Else %>
		sMsg  = "Esta operación eliminará al participante del sistema, por lo que todas sus tareas pendientes,\n";
	<% End If %>
	sMsg += "si las tuviera, serán redistribuidas entre los miembros de sus grupos de trabajo.\n\n";
	sMsg += "Además se borrarán y perderán todos los documentos y mensajes de sus bandejas personales.\n\n";			
	sMsg += "¿Elimino permanentemente del sistema a '<%=gsParticipantName%>'?";
	if (confirm(sMsg)) {
		window.location.href = "./exec/change_status.asp?id=<%=gnUserId%>&status=3";
	}
}

function callEditor(sWindow, nItemId) {
	var sURL, sOpt;
  switch (sWindow) {   
		case 'tasks':
			 sURL = 'user_tasks.asp?id=' + nItemId;
			 window.location.href = sURL;
			 return false;
    case 'objects':
			sURL = 'user_permissions.asp?id=' + nItemId;
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
<FORM name=frmSend action='./exec/save_user.asp' method=post>
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
						<% If (gnUserId) <> 0 Then %>
					  <a href='' onclick='return(notAvailable());'>Grupos de trabajo</a>
						&nbsp; | &nbsp;
					  <a href='' onclick='return(notAvailable());'>Habilidades</a>
						&nbsp; | &nbsp;
					  <a href='' onclick='return(callEditor("tasks", <%=gnUserId%>));'>Tareas</a>
						&nbsp; | &nbsp;						
					  <a href='' onclick='return(callEditor("objects", <%=gnUserId%>));'>Objetos</a>
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
					<TD nowrap>Nombre(s):</TD>
				  <TD>
						<INPUT name=txtName value='<%=gsName%>' style='width:260px'>
					</TD>
				</TR>			
				<TR>
					<TD nowrap>Apellido paterno:</TD>
				  <TD>			
						<INPUT name=txtLastName1 value='<%=gsLastName1%>' style='width:260px'>
					</TD>
				</TR>
				<TR>
					<TD nowrap>Apellido materno:</TD>
				  <TD>
						<INPUT name=txtLastName2 value='<%=gsLastName2%>' style='width:260px'>
					</TD>
				</TR>
				<TR>
					<TD nowrap>¿Es del sexo femenino? &nbsp; </TD>
				  <TD>
						<INPUT type=checkbox name=chkIsFemale value=true <%=gsIsFemale%>><br>&nbsp;
					</TD>
				</TR>
				<!--<TR>
					<TD nowrap>Fecha de nacimiento:</TD>
				  <TD>
						<INPUT name=txtBornDate value='<%=gsBornDate%>' style='width:90px'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtBornDate);'>
						&nbsp; (día / mes / año)<br>&nbsp;	
					</TD>
				</TR>
				!-->
				<TR>
					<TD>Correo electrónico en la organización:</TD>
				  <TD>
						<INPUT name=txtJobEMail value='<%=gsJobEMail%>' style='width:260px'><br>&nbsp;
					</TD>
				</TR>
				<TR>
					<TD>Teléfono en la organización:</TD>
				  <TD>
						<INPUT name=txtJobPhone value='<%=gsJobPhone%>' style='width:260px'><br>&nbsp;
					</TD>
				</TR>
				<TR>
					<TD>Identificador de acceso al sistema:</TD>
				  <TD width=80%>
						<INPUT name=txtParticipantKey value='<%=gsParticipantKey%>' style='width:125px'>
					</TD>
				</TR>
				<% If (gnUserId = 0) Then %>
				<TR>
					<TD>Permitir el acceso a partir del:</TD>
				  <TD>
						<INPUT name=txtHistoricDate value='<%=gsHistoricDate%>' style='width:90px'>
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='showCalendar(document.all.txtFromDate);'>
						&nbsp; (día / mes / año)<br>&nbsp;
					</TD>
				</TR>
				<% End If %>
				<TR>
					<TD colspan=2 align=right nowrap>
						<br>
						<% If gnUserId <> 0 Then %>
							<% If gsStatus = "A" Then %>
						  <INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Suspender' onclick='suspend();' tabindex=-1>
						  <% ElseIf (gsStatus = "S") Or (gsStatus = "D") Then %>
							<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Reactivar' onclick='reactivate();' tabindex=-1>
						  <% End If %>						
						&nbsp;
							<% If (gsStatus <> "D") Then %>
							<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Eliminar' onclick='delete_();' tabindex=-1>
							<% End If %>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
						<% End If %>
						<INPUT class=cmdSubmit type=button name=cmdSend style='width:80;' value='Aceptar' onclick='save();'>
						&nbsp;
						<INPUT class=cmdSubmit type=button name=cmdClose style='width:80;' value='Cancelar' onclick='window.close();'>
						&nbsp;<br>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
<INPUT type=hidden name=txtUserId value=<%=gnUserId%>>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
