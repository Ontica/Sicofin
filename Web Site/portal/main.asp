<%
	Option Explicit  
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If	
	
	Dim gsToDoQuickBar, gsTasksQuickBar
	
	Call Main()
	
	Sub Main()
		Dim oTaskBars, oInboxes
		'************		
		Set oTaskBars   = Server.CreateObject("EWMTasksUS.CTaskBars")
		gsToDoQuickBar	= oTaskBars.ToDoQuickBar(Session("sAppServer"), Session("uid"))				
		gsTasksQuickBar = oTaskBars.QuickBar(Session("sAppServer"), Session("uid"))
		Set oTaskBars   = Nothing
	End Sub
%>
<html>
<head>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/main_page.css">
<meta name="generator" content="Microsoft FrontPage 4.0">
<title>Banobras - Intranet corporativa</title>
<script src="/empiria/bin/client_scripts/clock.js"></script>
<script LANGUAGE=javascript>
<!--
   function window_onload() {     
     writeDate('oDate', true, true)
   }

//-->
</SCRIPT>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function openWindow(sWindow) {
	switch (sWindow) {
		case 'userConfig':
			
	}
	return false;
}

//-->
</SCRIPT>
</head>
<body LANGUAGE=javascript onload="return window_onload()" leftmargin=1>
  <!--BEGIN_TABLA_PRINCIPAL-->
  <table border="0" cellpadding="0" cellspacing="0" width="768">
    <tr>
			<!--BEGIN_COLUMNA_1-->
      <td rowspan=3 valign="top" width="137">
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=162>
        <TR height=7px><td></td></TR>
        <TR valign=top>
          <TD align=middle>						
            <TABLE class="taskTable" border=0 cellPadding=0 cellSpacing=0 width=100%>
              <TR align=left valign=top>
                <TD>
									<TABLE class="taskTable" border=0 cellPadding=3 cellSpacing=0 width=100%>
										<TR class=sectionDivision valign=middle height=20px>
											<TD>
							 					<a title='Lista rápida de asuntos y tareas pendientes'>Mis asuntos pendientes</a>							 					
							 				</TD>
											<TD align=right>
												<A class=ellipsisLink title='Muestra todas las tareas y asuntos pendientes por hacer y permite personalizar esta lista rápida de tareas pendientes.'
													 onclick="return(openWindow('userConfig'));">
													. . .
												</A>
											</TD>	
							 			</TR>
							 		</TABLE>
							 		<TABLE class="taskTable" border=0 cellPadding=3 cellSpacing=0 width=100%>
										<%=gsToDoQuickBar%>												
										<TR height=10px><td></td></TR>
									</TABLE>
                </TD>
              </TR>             
         <TR align=left valign=top>
					<TD>
						<TABLE class="taskTable" border=0 cellPadding=2 cellSpacing=0 width=100%>
							<TR class=sectionDivision valign=middle height=20px>
								<TD>
									<a title='Lista de las tareas que empleo frecuentemente'>Mis tareas frecuentes</a>
								</TD>
								<TD align=right>
									<A class=ellipsisLink title='Presenta la lista de tareas completa y permite configurar esta lista rápida de tareas frecuentes.'
										 onclick="return(openWindow('userConfig'));">
										. . .
									</A>
								</TD>								
							</TR>
						</TABLE>
						<TABLE class="taskTable" border=0 cellPadding=2 cellSpacing=0 width=100%>
							<%=gsTasksQuickBar%>
							<TR height=10px><td></td></TR>							
						</TABLE>									                    
					</TD>
				</TR>
				 <!--BEGIN_SECCION_FAVORITOS-->
         <TR align=left valign=top>
					<TD>
						<TABLE class="taskTable" border=0 cellPadding=2 cellSpacing=0 width=100%>
							<TR class=sectionDivision valign=middle height=20px>
								<TD>
									<a title='Lista personal con los sitios Web de mayor interés'>Mis favoritos</a>
								</TD>
								<TD align=right>
									<A class=ellipsisLink title='Permite administrar los sitios favoritos y personalizar esta lista de acceso rápido.'
										 onclick="return(openWindow('userConfig'));">
										. . .
									</A>
								</TD>								
							</TR>
						</TABLE>
						<TABLE class="taskTable" border=0 cellPadding=2 cellSpacing=0 width=100%>
							<TR><TD><a href="http://msdn.microsoft.com/" target="_blank" TITLE="Microsoft Developer´s Network">MSDN</a></TD></TR>
							<TR><TD><a href="http://hotmail.com/" target="_blank" TITLE="Servicio gratuito de correo electrónico">Hotmail</a></TD></TR>
							<TR><TD><a href="http://www.jornada.unam.com.mx/" target="_blank" TITLE="Periódico La Jornada">La Jornada</a></TD></TR>
							<TR><TD><a href="http://www.xml.com/pub/" target="_blank" TITLE="Sitio con recursos de XML">XML.com</a></TD></TR>
							<TR height=8px><td></td></TR>
						</TABLE>									                    
					</TD>
				</TR>								          
				<!--END_SECCION_FAVORITOS-->
		    </TABLE>
		  </TD>
		</TR>
    <TR vAlign=top height=8px><TD></TD></TR>
	</TABLE>
  </td>
  <!--END_COLUMNA_1-->
  <!--BEGIN_COLUMNA_2-->
  <td rowspan="3" valign="top" width="420">
     <table border="0" cellpadding="0" cellspacing="3" width="100%">
      <tr height=1px><td></td></tr>
			<TR height=20>
				<td>
					<!--BEGIN_FECHA_HORA-->
					<table border="0" cellpadding="0" cellspacing="0">
						<tr class=sectionDivision height=20 width=100%>							
							<TD id=oDate nowrap>&nbsp;</TD>
							<TD >&nbsp;</TD>
							<td width=100% align=right nowrap>
								<A title='Muestra una lista con las opciones personales de configuración del sistema.'
									 class=ellipsisLink onclick="return(openWindow('userConfig'));">
									<%=Session("user_name")%>
								</A>
							</td>
						</tr>
					</table>
					<!--END_FECHA_HORA-->
				</td>
			</TR>
			<!--BEGIN_DOCUMENTOS-->
			<TR>
				<td>
					<table class=newsAbstract border="0" cellpadding="3" cellspacing="3" width="412">
					<TR>
						<td>
							<table border="0" cellpadding="0" cellspacing="0" width="412">
								<tr>
									<td>
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td>
													<A class=newsTitle>Gracias ...<br></A>
												</td>
											</tr>
											<tr>
												<td class=newsAbstract>
													Agradecemos a todos los usuarios sus acertados comentarios y sugerencias, así 
													como su infinita paciencia, para la liberación de esta primera versión 
													operativa del sistema de contabilidad financiera.
													<br><br>
													Estamos seguros que sin su colaboración nos hubiera sido imposible realizar
													los trabajos encomendados.
													<br><br>
													Así mismo, aprovechamos para ponernos a sus órdenes en la dirección de correo
													<A href='mailto:soporte_banobras@ontica.com.mx'>soporte_banobras@ontica.com.mx</A> o en nuestro
													sitio <A href='http://www.ontica.com.mx/'>www.ontica.com.mx</A>
													<br><br>
													Muchas gracias. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Atentamente<br><br>
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>El equipo de desarrollo de</i><br>&nbsp;
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>La Vía Ontica, S.C.</b>
													<br>
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
						</td>
					</TR>			
					</table>
				</td>
			</TR>		
			<TR>
				<td>
				</td>		
			</TR>
			<!--END_DOCUMENTOS-->
			<TR height=20>
				<td>
					<!--BEGIN_BANNER_SERVICIOS-->
					<table border="0" width="100%" cellpadding="0" cellspacing="0">
						<tr class=sectionDivision height=20 >
							<TD colspan=2>Servicios</TD>
							<TD align=right>
								<A class=ellipsisLink title='Muestra la lista con todos los servicios electrónicos disponibles y permite editar los elementos que aparecen en esta lista.'
									 onclick="return(openWindow('userConfig'));">
									. . .
								</A>
							</TD>
						</tr>			
						<TR class=tip border=0 cellPadding=0 cellSpacing=0 width=412 height=18>
							<TD TITLE="Permite reservar boletos de avión en línea">
								<a href="main.asp">Reservaciones aéreas</a></TD>
							<TD TITLE="Conozca las efemérides del día de hoy">
								<a href="http://encarta.msn.com/features/onThisDay.asp" target='_blank'>En un día como hoy...</a></TD>
							<TD TITLE="Indicadores económicos nacionales y extranjeros">
								<a href="main.asp">Indicadores económicos</a></TD>
						</TR>
						<TR class=tip border=0 cellPadding=0 cellSpacing=0 width=412 height=18>
							<TD TITLE="Agente electrónico experto en consultoría financiera">
								<a href="main.asp">Mi agente financiero</a></TD>
							<TD TITLE="Servicio para consultar el estado del tiempo">
								<a href="main.asp">Estado del tiempo</a></TD>
							<TD TITLE="Periódicos y revistas nacionales y extranjeros">
								<a href="main.asp">Periódicos y revistas</a></TD>
						</TR>
						<TR class=tip border=0 cellPadding=0 cellSpacing=0 width=412 height=18>
							<TD TITLE="Nuestras recomendaciones culturales">
								<a href="main.asp">Cultura y entretenimiento</a></TD>
							<TD TITLE="Lista de restaurantes y bares cercanos">
								<a href="main.asp">Para comer</a></TD>
							<TD TITLE="Nuestro servicio de anuncios clasificados entre usuarios de la intranet">
								<a href="main.asp">El corcho</a></TD>
						</TR>
					</TABLE>
					<!--END_TABLA_SERVICIOS-->					
				</td>
			</tr>
			<tr height=8><td></td></tr>
			<TR height=20>
				<td>
					<!--BEGIN_BANNER_SERVICIOS-->
					<table border="0" width="100%" cellpadding="0" cellspacing="0">
						<tr class=sectionDivision height=20>
							<TD colspan=2>Tip del día</TD>
							<TD align=right>
								<A class=ellipsisLink title='Al hacer clic en esta liga se despliega el siguiente tip.'
									 onclick="return(openWindow('userConfig'));">
									. . .
								</A>
							</TD>						
						</tr>	
						<TR class=tip border=0 cellPadding=0 cellSpacing=0 width=412 height=18>
							<TD colspan=3>
								La mayoría de los botones y ligas tienen un breve mensaje de ayuda.<br>
								Para desplegarlos basta con colocar el cursor del ratón sobre dichos elementos.
							</TD>
						</TR>
					</TABLE>
					<!--END_TABLA_SERVICIOS-->
				</td>
			</tr>
			<tr><td colspan=2><br><IMG height=7 src="/empiria/images/pleca.gif" width=100%></td></tr>
			<TR height=20>
				<td>
					<!--BEGIN_BANNER_SERVICIOS-->
					<table class=copyright border="0" width="100%" cellpadding="0" cellspacing="0">
						<TR>
							<TD>Desarrollado especialmente para el Banco Nacional de Obras y Servicios Públicos, S.N.C.</TD>
						</TR>					
						<TR>
							<TD><A href='http://www.ontica.com.mx/'>Empiria © México 2001. La Vía Ontica, S.C. Todos los derechos reservados.</A></TD>
						</TR>						
					</TABLE>
					<!--END_TABLA_SERVICIOS-->
				</td>
			</tr>			
		</table>
	</td>
	<!--END_COLUMNA_2-->
	<!--BEGIN_COLUMNA_3-->
	<td rowspan="3" valign="top" width="145" align="left">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">		
		<!--BEGIN_TIP_DEL_DIA-->
		<tr height=7><td></td></tr>
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr class=sectionDivision height=20>
						<td>Herramientas
						</td>
						<TD align=right>
							<A class=ellipsisLink title='Muestra la lista con todas las herramientas incluidas en el sistema y permite configurar esta lista rápida de herramientas.'
								 onclick="return(openWindow('userConfig'));">
								. . .
							</A>
						</TD>
					</tr>
					<tr>
						<td class=tasktable colspan=2>
					  <TABLE class=taskTable cellPadding=2 cellSpacing=3 width=100%>
						  <TR>
							  <TD><a TITLE="Calculadora">Calculadora</a></TD>
						  </TR>
						  <TR>
							  <TD><a TITLE="Calendario">Calendario</a></TD>
						  </TR>
						  <TR>
  							<TD><a TITLE="Post-its®">Post-its®</a></TD>
	  					</TR>
							<TR>
								<TD><a TITLE="Tipos de cambio">Tipos de cambio</a></TD>
		  				</TR>
		  			</TABLE>								
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<!--END_TIP_DEL_DIA-->		
		<tr>
		<td>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr height=7px colspan=2><td></td></tr>						
				<tr class=sectionDivision height=20>
					<td>Encuesta</td>
						<TD align=right>
							<A class=ellipsisLink title='Muestra los resultados de encuestas levantadas anteriormente.'
								 onclick="return(openWindow('userConfig'));">
								. . .
							</A>
						</TD>
				</tr>						
				<!--BEGIN_ENCUESTA_DEL_DIA-->
			  <tr valign=top>
				 	<td colspan=2>						 	
				 		<table class=pollTable border="0" cellpadding="0" cellspacing="0" width="100%">
				 			<tr>
				 				<td colspan="4" height="5"></td>
				 			</tr>
				 			<tr>
				 				<td colspan="4" align="middle">						 					
				 						¿Es sencilla la operación del diseñador de reportes?						 					
				 				</td>
				 			</tr>
				 			<tr height="2"><td colspan="4" align="middle"></td></tr>
				 			<form action="main.asp" method="post" name=frmPoll>
				 			<tr>
				 				<td width="30%" align="middle">
				 					<input type="radio" name="opcion" value="si">Sí
				 				</td>
				 				<td colspan="3" align="middle">						 					
				 					<input type="radio" name="opcion" value="no">No						 					
				 				</td>
				 			</tr>
				 			<tr height="12"><td colspan="4" align="middle"></td></tr>
				 			<tr>
				 				<td colspan="4" align="middle">
				 					<input type="submit" class=pollButton name="accion" style='width:65;' value="Votar">&nbsp;&nbsp;&nbsp;
				 					<input type="submit" class=pollButton name="accion" style='width:65;' value="Resultados">
				 					<input type="hidden" name="IdPregunta" value="101">
				 				</td>
				 			</tr>
				 			<tr height="8"><td colspan="4"></td></tr>						 			
				 			</FORM>
				 		</table>
				 		<!--END_ENCUESTA_DEL_DIA-->
					</td>
				</tr>
			</table>
			<tr height=8><td></td></tr>
		  <!--BEGIN_INDICADORES_ECONOMICOS-->
		  <tr>
				<td>
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
						<TR class=sectionDivision height=20>
							<TD>Indicadores económicos</td>
							<TD align=right>
								<A class=ellipsisLink title='Presenta los indicadores económicos de fechas anteriores.'
									 onclick="return(openWindow('userConfig'));">
									. . .
								</A>
							</TD>
						</TR>						
						<tr>
							<td align=middle colspan=2><IMG alt="Estados financieros" src="/empiria/images/portal/grafica.gif" border=0></td>
						</tr>
					</table>
				</td>
			</tr>
			<!--END_INDICADORES_ECONOMICOS-->			
			</table>
			</td>
		</tr>
	</table>
	<!--END_TABLA_PRINCIPAL-->
</body>
</html>


