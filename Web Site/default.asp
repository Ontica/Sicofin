<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim gsUserName, gsMessage
	
	On Error Resume Next
		
	If IsObject(Session("rsUserInfo")) Then
		If (Session("rsUserInfo").State = 1) Then
			Session("rsUserInfo").Close
			Set Session("rsUserInfo") = Nothing
		End If		
	End If

	gsMessage = Request.Form("txtMessage")
	gsUserName = Request.Form("txtUserName")
	
	Session.Abandon
%>
<html>
<head>
<title>Banobras - Acceso a la intranet corporativa.</title>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/logon.css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function validate() {
	var sMsg;
	if (document.all.txtUserName.value == '') {
		alert('Para ingresar a la intranet se requiere la identificación de acceso.');
		document.all.txtUserName.focus();
		return false;
  }
	if (document.all.txtPassword.value == '') {
		alert('Para ingresar a la intranet se requiere la contraseña de acceso.');
		document.all.txtPassword.focus();
		return false;
  }
  document.all.frmLogon.submit();
	return true
}

function window_onload() {
<% If (Len(gsMessage) <> 0) Then %>
	alert("<%=gsMessage%>");
	document.all.txtPassword.focus();
<% Else %>
	document.all.txtUserName.focus();
<% End If %>	
}

function cmdSend_onclick() {

}

//-->
</SCRIPT>
</head>
<body onload="return window_onload()">
<br><br>
<P align=center>
<TABLE width="350" border="0" cellpadding="0" cellspacing="0">
	<TR>
		<TD nowrap colspan=2>
			<TABLE bgColor=#003333 border=0 cellPadding=0 cellSpacing=0 width="100%">
			  <TR>
			    <TD valign=top align=left height=60 rowspan=2 width=163>
						<IMG alt=Banobras border=0 height=60 src="images/banobras.gif" width=163 align=top>
					</TD>
			    <TD valign=top align=right>
						<IMG border=0 height=40 src="images/collage.jpg" width=426 align=top>
			    </TD>
				</TR>
			  <TR>
					<TD nowrap valign=middle align=right background="images/fondo.gif" height=20 width="100%">
						Bienvenidos a la intranet corporativa &nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR height=240 nowrap>
		<TD valign=top>
			<TABLE class=linksTable width="163" height=100% border="0" cellpadding="0" cellspacing="0">
				<TR height=25>					
					<TD>
						&nbsp;<A href='http://www.banobras.gob.mx/'>Sitio Internet de Banobras</A>
					</TD>
				</TR>
				<TR>
					<TD class=separator>&nbsp;Otros sitios de interés</TD>
				</TR>
				<TR>
					<TD>
						<TABLE class=linksTable width="163" height=100% border="0" cellpadding="2" cellspacing="2">
							<TR>
								<TD><A href='http://www.presidencia.gob.mx/'>Presidencia de la República</A></TD>
							</TR>
							<TR>
								<TD><A href='http://www.shcp.gob.mx/'>Secretaría de Hacienda</A></TD>
							</TR>
							<TR>
								<TD><A href='http://www.banxico.org.mx/'>Banco de México</A></TD>
							</TR>							
							<TR>
								<TD><A href='http://www.cnbv.gob.mx/'>Comisión Nacional Bancaria y de Valores</A></TD>
							</TR>
							<TR>
								<TD><A href='http://www.bancomext.gob.mx/'>Banco de Comercio Exterior</A></TD>
							</TR>
							<TR>
								<TD><A href='http://www.nafinsa.gob.mx/'>Nacional Financiera</A></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
		<TD width=440 align=center valign=top>
			<TABLE class=logonTable width="280">
				<TR>
					<TD>&nbsp; &nbsp;</TD>
					<TD valign=top nowrap>
						<FORM name="frmLogon" action="/empiria/central/logon/logon.asp" method="post">
							<TABLE class=formTable>
								<TR>
								  <TD colspan=3>Para ingresar a la intranet requiero el identificador y la <br>contraseña de acceso:<br>&nbsp;</TD>
								</TR>							
								<TR>
									<TD nowrap>&nbsp;</TD>
								  <TD><b>Identificación: &nbsp;</b></TD>
								  <TD><INPUT name=txtUserName value="<%=gsUserName%>" style='width:170;'></TD>
								</TR>
								<TR>
									<TD nowrap>&nbsp; &nbsp;</TD>
								  <TD><b>Contraseña:</b></TD>
								  <TD><INPUT type=password name=txtPassword style='width:170;'></TD>
								</TR>
								<TR>
								  <TD colspan=3 align=right>
										<INPUT type=button class=cmdButton name=cmdSend style="width:60" value=Aceptar onclick="validate();">
										&nbsp; &nbsp;
										<INPUT type=button class=cmdButton name=cmdHelp style="width:60" value=Ayuda>
									</TD>
								</TR>
							</TABLE>
						</FORM>
					</TD>
				</TR>	
				<TR>				
					<TD>&nbsp; &nbsp;</TD>
					<TD nowrap>
						<TABLE class=optionsTable>
							<TR>
								<TD nowrap>
									&nbsp;<A href=''>Olvidé mi contraseña de acceso</A>
								</TD>
							</TR>
							<TR>
								<TD nowrap>
									&nbsp;<A href=''>Solicitar acceso a la intranet</A>
								</TD>								
							</TR>
							<TR>
								<TD nowrap>
									&nbsp;<A href=''>Verificar la compatibilidad de mi navegador</A>&nbsp; &nbsp;
								</TD>								
							</TR>							
						</TABLE>
					</TD>
				</TR>
				<TR height=100%><TD colspan=2></TD></TR>
				<TR>
					<TD colspan=2 class=copyright>Desarrollado especialmente para el Banco Nacional de Obras y Servicios Públic<font color=#696969>os, </font>S.N.C.</TD>
				</TR>
				<TR>
					<TD colspan=2 class=copyright><A href='http://www.ontica.com.mx/'>© México 2001. La Vía Ontica, S.C. Todos los derechos reservados.</A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</P>
</body>
</html>
