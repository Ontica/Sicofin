<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsTackedWindows, gnSelectedColumn
	Dim gnParticipantType, gsCboParticipantTypes, gsExplorerResultsHeader, gsExplorerResultsBody
	Dim gsUserNames, gsGroups, gsSkills, gsTasks
	
	Call Main()	

	Sub Main()
		Call SetGlobalValues()
		Call GetExplorerInformation()
	End Sub
	
	Sub GetExplorerInformation()
		Dim oExplorer, sWhere, sOrderBy
		'***************************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("EWFParticipantsUS.CExplorer")
		gsCboParticipantTypes		= oExplorer.CboParticipantTypes(CLng(gnParticipantType))
		gsExplorerResultsHeader = oExplorer.Header(CLng(gnSelectedColumn))		
		'sWhere = oExplorer.BuildSearchParametersString(CLng(gnGralLedgersCategory), _
		'																							 CStr(gsFromApplicationDate), _
		'																							 CStr(gsToApplicationDate), _
		'																							 CStr(gsFromElaborationDate), _
		'																							 CStr(gsToElaborationDate), _
		'																							 CStr(gsVoucherNumber), _
		'																							 CStr(gsVoucherConcept), _
		'																							 CStr(gsAccounts), CLng(gnTransactionTypeId), _
		'																							 CLng(gnVoucherTypeId), CLng(gnBalancingType))
		'sOrderBy = "numero_transaccion"
		
		If gnParticipantType <> -1 Then
			gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), CLng(gnParticipantType), 0, _
																						 CStr(sWhere), CStr(sOrderBy), CLng(gnSelectedColumn))
		End If
		Set oExplorer = Nothing
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("./exec/exception.asp")
		End If		
	End Sub

  Sub SetGlobalValues() 
		If (CLng(Request.Form.Count) <> 0) Then
			gnParticipantType = Request.Form("cboParticipantTypes")
			gsUserNames				= Request.Form("txtUserName")
			gsGroups					= Request.Form("txtGroups")
			gsSkills					= Request.Form("txtSkills")
			gsTasks						= Request.Form("txtTasks")	
			gnSelectedColumn	= Request.Form("txtSelectedColumn")
			gsTackedWindows   = Request.Form("txtTackedWindows")
		Else
			gnParticipantType = -1
			gsUserNames				= ""
			gsGroups					= ""
			gsSkills					= ""
			gsTasks						= ""
			gnSelectedColumn	= 1
			gsTackedWindows		= ""			
		End If
	End Sub		
%>	
<HTML>
<HEAD>
<TITLE>Empiria®: Administrador del flujo de trabajo</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var oEditWindow = null;

function orderBy(nSelectedColumn) {
  var nTemp = document.all.txtSelectedColumn.value;
  
  if (nTemp == '') {
		document.all.txtSelectedColumn.value = nSelectedColumn;
	} else {
		if ((nTemp == nSelectedColumn) || (Math.abs(nSelectedColumn) == nSelectedColumn)) {
			if (nTemp == nSelectedColumn) {
				document.all.txtSelectedColumn.value = (-1 * nSelectedColumn);
			} else {
				document.all.txtSelectedColumn.value = nSelectedColumn;
			}
		} else {
			document.all.txtSelectedColumn.value = nSelectedColumn;
		}
  }
  document.all.frmSend.action = '';
  document.all.frmSend.submit();
	return false;
}

function callEditor(nOperation, nItemId) {
	var sURL, sOpt;
		
  switch (nOperation) {  
    case 1:		//Edit user
			sURL = 'edit_user.asp?id=' + nItemId;
			sOpt = 'height=360px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			break;
    case 2:		//Edit workgroup
			sURL = 'edit_workgroup.asp?id=' + nItemId;	
			sOpt = 'height=260px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			break;
    case 3:		//Edit organization
			sURL = 'edit_system.asp?id=' + nItemId;	
			sOpt = 'height=260px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			break;			
    case 4:		//Edit system
			sURL = 'edit_system.asp?id=' + nItemId;	
			sOpt = 'height=260px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			break;
	}
	if (oEditWindow == null || oEditWindow.closed) {
		oEditWindow = window.open(sURL, '_blank', sOpt);
	} else {
		oEditWindow.focus();
		oEditWindow.navigate(sURL);
	}	
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload="unloadWindows(oEditWindow)">
<FORM name=frmSend action='' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Explorador de participantes
		</TD>
		<TD colspan=3 align=right nowrap align=top>
			Tipos de participantes: &nbsp;
			<SELECT name=cboParticipantTypes style='width:180px;'>
				<% If gnParticipantType = -1 Then %>
					<OPTION value='' selected>-- Tipo de participante --</OPTION>
				<% End If %>
					<%=gsCboParticipantTypes%>
			</SELECT>
			<A href='' onclick="document.all.frmSend.submit();return false;">Explorar</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>
			<A href='' onclick="return(showOptionsWindow(document.all.divSearchOptions));">Búsqueda avanzada</A>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			
			<img align=absmiddle src='/empiria/images/task_red.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">		</TD>
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
						<A href="balances.asp" target='_blank'>Administración del flujo de trabajo</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="financial_statements.asp">Administración de tareas</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="other_reports.asp">Visor del flujo de trabajo</A>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<A href="../balances/balance_explorer.asp">Estadísticas de desempeño</A>
						<img src='/empiria/images/invisible.gif'>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR id=divSearchOptions style='display:none;'>
		<TD colspan=4 nowrap>
			<TABLE class="fullScrollMenu">
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap>
						Búsqueda de participantes
					</TD>
					<TD nowrap align=right>
						<img src='/empiria/images/invisible4.gif'>
						<img src='/empiria/images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>
					  <img id=cmdSearchOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSearchOptions, this)' alt='Fijar la ventana'>
						<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
						<img src='/empiria/images/invisible.gif'>						
						<img src='/empiria/images/close_white.gif' onclick='closeOptionsWindow(document.all.divSearchOptions, document.all.cmdSearchOptionsTack)' alt='Cerrar'>
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Por nombre:</b>&nbsp;</TD>
					<TD colspan=2 nowrap width=90%>
						<INPUT name=txtUserName style="width:230;height:20;" value='<%=gsUserNames%>'> 
							&nbsp;&nbsp;(permite el empleo de <A href="" onclick="return(showHelp('wild_chars'))" target=_blank>comodines</A>)
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Por sus habilidades:</b>&nbsp;</TD>
					<TD colspan=2 nowrap>
						<INPUT name=txtSkills style="width:230;height:20;" value='<%=gsSkills%>'> 
							&nbsp; <INPUT class=cmdSubmit type=button name=cmdSend value="Lista de habilidades ..." style='width:120;' onclick="doSubmit();">
					</TD>
				</TR>				
				<TR>
					<TD nowrap><b>Por sus participantes relacionados:</b>&nbsp;</TD>
					<TD colspan=2 nowrap>
						<INPUT name=txtUsers style="width:230;height:20;" value='<%=gsGroups%>'>
							&nbsp; <INPUT class=cmdSubmit type=button name=cmdSend value="Lista de participantes ..." style='width:120;' onclick="doSubmit();">
					</TD>
				</TR>
				<TR>
					<TD nowrap><b>Por las tareas que efectúan:</b>&nbsp;</TD>
					<TD colspan=2 nowrap>
						<INPUT name=txtTasks style="width:230;height:20;" value='<%=gsTasks%>'> 
							&nbsp; <INPUT class=cmdSubmit type=button name=cmdSend value="Lista de tareas ..." style='width:120;' onclick="doSubmit();">	
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">
					<TD class="fullScrollMenuTitle" nowrap colspan=2>
						Resultados de la consulta
					</TD>
					<TD align=right nowrap>
						<img align=absbottom src='/empiria/images/refresh_white.gif' onclick='document.all.frmSend.submit();' alt="Refrescar">						
					</TD>
				</TR>				
				<TR id=divSelectedItemsOptions style='display:none;'>
					<TD colspan=4 nowrap>
						<TABLE class='fullScrollMenu'>
							<TR class="fullScrollMenuHeader">
								<TD class="fullScrollMenuTitle" nowrap colspan=2>
									¿Qué se desea hacer con los grupos de trabajo seleccionados?
								</TD>
								<TD nowrap align=right>
								  <img id=cmdSelectedVouchersOptionsTack src='/empiria/images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSelectedVouchersOptions, this)' alt='Fijar la ventana'>					
									<img src='/empiria/images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
								  <img src='/empiria/images/invisible.gif'>
								  <img src='/empiria/images/close_white.gif' onclick="closeOptionsWindow(document.all.divSelectedVouchersOptions, document.all.cmdSelectedVouchersOptionsTack)" alt='Cerrar'>
								</TD>				
							</TR>
							<TR>
								<TD nowrap>
									<A href="" onclick="return(notAvailable());">Incluir usuarios</A>&nbsp; &nbsp; 
									<A href="" onclick="return(notAvailable());">Ver sus listas de tareas pendientes</A>
								</TD>
							</TR>			
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD colspan=4 nowrap>
						<TABLE class=applicationTable>				
							<THEAD>
								<TR class=applicationTableHeader valign=center>
									<%=gsExplorerResultsHeader%>
								</TR>
							</THEAD>
							<% If (Len(gsExplorerResultsBody) <> 0) Then %>
								<%=gsExplorerResultsBody%>
							<% Else %>
								<TBODY>
									<TR>
										<TD colspan=5>Seleccionar de la lista superior el tipo de participante que se desea desplegar.</TD>
									</TR>
								</TBODY>
							<% End If %>
						</TABLE>
					</TD>
				</TR>
			</TD>
		</TR>
	</TD>
</TR>
<INPUT TYPE=hidden name=txtSelectedColumn value='<%=gnSelectedColumn%>'>
<INPUT TYPE=hidden name=txtTackedWindows>
</TABLE>
</FORM>
</BODY>
</HTML>