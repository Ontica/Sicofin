<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsTackedWindows, gnSelectedColumn
	Dim gsExplorerResultsHeader, gsExplorerResultsBody
	Dim gsUserNames, gsGroups, gsSkills, gsTasks
	
	Call Main()	

	Sub Main()
		Call SetGlobalValues()
		Call GetExplorerInformation()
	End Sub
	
	Sub GetExplorerInformation()
		Dim oExplorer, vGralLedgers, bShowSubsAccounts, bShowDebitCreditCols, sWhere, sOrderBy
		'********************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("MHParticipantsExplorer.CExplorer")
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

		If (CLng(Request.Form.Count) <> 0) Then
			gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), 1, 0, _
																						 CStr(sWhere), CStr(sOrderBy), CLng(gnSelectedColumn))
		Else
			gsExplorerResultsBody = oExplorer.Body(Session("sAppServer"), 1, 0, _
																						 CStr(sWhere), CStr(sOrderBy), CLng(gnSelectedColumn))			
		End If
		Set oExplorer = Nothing
		
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("../programs/exception.asp")
		End If		
	End Sub

  Sub SetGlobalValues() 
		If (CLng(Request.Form.Count) <> 0) Then
			gsUserNames				= Request.Form("txtUserName")
			gsGroups					= Request.Form("txtGroups")
			gsSkills					= Request.Form("txtSkills")
			gsTasks						= Request.Form("txtTasks")	
			gnSelectedColumn	= Request.Form("txtSelectedColumn")
			gsTackedWindows   = Request.Form("txtTackedWindows")
		Else
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
<TITLE>Aldea®: Administrador del flujo de trabajo</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="../resources/mahler.css">
<script language="JavaScript" src="../programs/client_scripts.js"></script>
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
    case 1:		//Add    		
			sURL = 'edit_user.asp'
			sOpt = 'height=360px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			if (oEditWindow == null || oEditWindow.closed) {
				oEditWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oEditWindow.focus();
				oEditWindow.navigate(sURL);
			}
			return false;
    case 2:		//Edit
			sURL = 'edit_user.asp?id=' + nItemId;
			sOpt = 'height=360px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			if (oEditWindow == null || oEditWindow.closed) {
				oEditWindow = window.open(sURL, '_blank', sOpt);
			} else {
				oEditWindow.focus();
				oEditWindow.navigate(sURL);
			}
			return false;
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
			<img align=absmiddle src='../images/users.gif' height=20> Administración de usuarios
		</TD>
		<TD colspan=3 align=right nowrap align=top>
			<A href='participants_mgr.asp'>Administración de participantes</A>
			<img align=absmiddle src='../images/invisible4.gif'>
			<img align=absbottom src='../images/refresh_white.gif' onclick='document.all.frmSend.submit();' alt="Refrescar">
			<img align=absmiddle src='../images/task_white.gif' onclick='showOptionsWindow(document.all.divTasksOptions);' alt="Tareas">			<img align=absmiddle src='../images/help_white.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='../images/invisible.gif'>
			<img align=absmiddle src='../images/close_white.gif' onclick="window.location.href = '<%=Session("main_page")%>';" alt="Cerrar y regresar a la página principal">		</TD>
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
					  <img id=cmdTasksOptionsTack src='../images/tack_white.gif' onclick='tackOptionsWindow(document.all.divTasksOptions, this)' alt='Fijar la ventana'>					
						<img src='../images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
					  <img src='../images/invisible.gif'>
						<img src='../images/close_white.gif' onclick="closeOptionsWindow(document.all.divTasksOptions, document.all.cmdTasksOptionsTack)" alt='Cerrar'>
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
						<img src='../images/invisible.gif'>
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
						Usuarios
					</TD>
					<TD align=right nowrap>
					  <a href='' onclick='return(callEditor(1, 0));'>Agregar usuario(a)</a>
						&nbsp; | &nbsp;
						<A href='' onclick="return(showOptionsWindow(document.all.divSearchOptions));">Buscar</A>
						&nbsp; | &nbsp;
						<A href="" onclick="return(notAvailable());">Imprimir</A>
						&nbsp; | &nbsp;
						<A href='' onclick="return(showOptionsWindow(document.all.divSelectedItemsOptions));">Usuarios seleccionados</A>						
						&nbsp; &nbsp;
						<img align=absbottom src='../images/refresh_white.gif' onclick='document.all.frmSend.submit();' alt="Refrescar">						<img align=absmiddle src='../images/help_white.gif' onclick='notAvailable();' alt="Ayuda">
					</TD>
				</TR>
				<TR id=divSearchOptions style='display:none;'>
					<TD colspan=4 nowrap>
						<TABLE class="fullScrollMenu">
							<TR class="fullScrollMenuHeader">
								<TD class="fullScrollMenuTitle" nowrap>
									Búsqueda de usuarios
								</TD>
								<TD nowrap align=right>
									<img src='../images/invisible4.gif'>
									<img src='../images/refresh_white.gif' onclick='return(resetSearchOptions());' alt='Actualizar ventana'>
								  <img id=cmdSearchOptionsTack src='../images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSearchOptions, this)' alt='Fijar la ventana'>
									<img src='../images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>					  
									<img src='../images/invisible.gif'>						
									<img src='../images/close_white.gif' onclick='closeOptionsWindow(document.all.divSearchOptions, document.all.cmdSearchOptionsTack)' alt='Cerrar'>
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
								<TD nowrap><b>Por los grupos a los que pertenecen:</b>&nbsp;</TD>
								<TD colspan=2 nowrap>
									<INPUT name=txtGroups style="width:230;height:20;" value='<%=gsGroups%>'>
										&nbsp; <INPUT class=cmdSubmit type=button name=cmdSend value="Grupos de trabajo ..." style='width:120;' onclick="doSubmit();">							
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
				<TR id=divSelectedItemsOptions style='display:none;'>
					<TD colspan=4 nowrap>
						<TABLE class='fullScrollMenu'>
							<TR class="fullScrollMenuHeader">
								<TD class="fullScrollMenuTitle" nowrap colspan=2>
									¿Qué se desea hacer con los usuarios seleccionados?
								</TD>
								<TD nowrap align=right>
								  <img id=cmdSelectedVouchersOptionsTack src='../images/tack_white.gif' onclick='tackOptionsWindow(document.all.divSelectedVouchersOptions, this)' alt='Fijar la ventana'>					
									<img src='../images/help_white.gif' onclick='notAvailable();' alt='Ayuda'>
								  <img src='../images/invisible.gif'>
								  <img src='../images/close_white.gif' onclick="closeOptionsWindow(document.all.divSelectedVouchersOptions, document.all.cmdSelectedVouchersOptionsTack)" alt='Cerrar'>
								</TD>				
							</TR>
							<TR>
								<TD nowrap>
									<A href="" onclick="return(notAvailable());">Incluirlos en un grupo de trabajo</A>&nbsp; &nbsp; 
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
										<TD colspan=8><b>No encontré ningún usuario con el criterio de búsqueda proporcionado.</b></TD>
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