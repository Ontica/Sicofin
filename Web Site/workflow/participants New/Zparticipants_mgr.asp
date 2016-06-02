<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsTackedWindows
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

function refreshPage(nOrderId) {
  if (nOrderId == 0) {
		window.location.href = 'transaction_selector.asp';
	} else {
		window.location.href = 'transaction_selector.asp' + '?order=' + nOrderId;
	}
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY onload="showTackedWindows(Array(<%=gsTackedWindows%>));" onunload="unloadWindows(oEditWindow)">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Administración de participantes
		</TD>
		<TD colspan=3 align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>
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
	<TR>
		<TD colspan=4 nowrap>
			<TABLE>
				<TR>					
					<TD width=50% valign=top>
						<TABLE class=blockTable>
							<TR class=fullScrollMenuHeader>
								<TD nowrap>
									<b>Grupos de trabajo</b>
								</TD>
								<TD nowrap align=right>
									<a href='' onclick='return(notAvailable());'>Crear</a> &nbsp; | &nbsp;
									<a href='participants_explorer.asp'>Consultar</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Pendientes</a>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD rowspan=3 valign=top>
									<img src='/empiria/images/workflow/participants.gif' alt='Ver grupos de trabajo'>
								</TD>
							<TR>
								<TD>
										Los grupos de trabajo o <b>roles</b>, son conjuntos de usuarios que están facultados 
										para realizar las mismas actividades o tareas.<br><br>
										Cada grupo de trabajo tiene asignados uno o más coordinadores para efectos
										administrativos o para manejar los casos excepcionales.<br><br> 
										Algunos ejemplos son: administrador del sistema, auxiliar administrativo,
										pagador de nóminas, programador, auxiliar contable, analista financiero,
										jefe de producción, cajero, etc.
								</TD>
							</TR>
							</TR>
						</TABLE>
					</TD>
					<TD>&nbsp;&nbsp;</TD>
					<TD width=50% valign=top>
						<TABLE class=blockTable>
							<TR class=fullScrollMenuHeader>
								<TD nowrap>
									<b>Usuarios</b>
								</TD>
								<TD nowrap align=right>
									<a href='' onclick='return(notAvailable());'>Crear</a> &nbsp; | &nbsp;
									<a href='users.asp'>Consultar</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Pendientes</a>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD rowspan=3 valign=top>
									<img src='/empiria/images/workflow/users.gif' alt='Consulta de usuarios'>
								</TD>
								<TR>
									<TD>
										Permite definir a los usuarios de las actividades que intervienen en 
										los flujos de trabajo.<br><br>
										Los usuarios pueden pertenecer a grupos de trabajo y normalmente
										son miembros de una o más organizaciones. <br><br>
										Así mismo, tienen asignadas diferentes actividades y objetos
										(como nóminas y contabilidades), ya sea por ellos mismos o por su 
										inclusión en dichos grupos de trabajo u organizaciones.<br><br><br><br>
									</TD>
								</TR>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR><TD>&nbsp;</TD></TR>				
				<TR>
					<TD valign=top>
						<TABLE class=blockTable>
							<TR class=fullScrollMenuHeader>
								<TD nowrap>
									<b>Organizaciones</b>
								</TD>
								<TD nowrap align=right>
									<a href='' onclick='return(notAvailable());'>Crear</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Consultar</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Pendientes</a>&nbsp;
								</TD>
							</TR>						
							<TR>
								<TD rowspan=3 valign=top>
									<img src='/empiria/images/workflow/organizations.gif' alt='Ver grupos de trabajo'>
								</TD>
								<TR>
									<TD>
										Define a las organizaciones y a las áreas o departamentos de las 
										mismas que están involucradas en el flujo de trabajo.<br><br>
										Típicamente tienen asignada una lista de los usuarios que 
										pertenecen a la misma, los cuales tienen diferentes
										puestos o roles dentro de la organización.<br><br>
										En ocasiones las tareas en el flujo de trabajo son direccionadas a 
										organizaciones en vez de a grupos de trabajo o a usuarios.
									</TD>
								</TR>
							</TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD> 
					<TD valign=top>
						<TABLE class=blockTable>
							<TR class=fullScrollMenuHeader>
								<TD nowrap>
									<b>Sistemas</b>
								</TD>
								<TD nowrap align=right>
									<a href='' onclick='return(notAvailable());'>Crear</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Consultar</a> &nbsp; | &nbsp;
									<a href='' onclick='return(notAvailable());'>Pendientes</a>&nbsp;
								</TD>
							</TR>						
							<TR>
								<TD rowspan=3 valign=top>
									<img src='/empiria/images/workflow/systems.gif' alt='Administración de sistemas'>
								</TD>
								<TR>
									<TD>
											Se trata de participantes que ejecutan las tareas del flujo de trabajo
											en forma automática (sin la intervención de usuarios).<br><br>
											Por ejemplo, escáners, agentes electrónicos, radares, monitores de seguridad, 
											máquinas automáticas, despachadores de actividades y otros.
											<br><br><br><br><br>
									</TD>
								</TR>
							</TR>
						</TABLE>					
					</TD>
				</TR>
				<TR><TD>&nbsp;</TD></TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>