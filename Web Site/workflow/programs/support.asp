<%
  Option Explicit     
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
	Dim gsErrPage, gsCancelPage, gsSupportPage
	Dim gsErrNumber, gsErrSource, gsErrDescription
	  	 
  Call SendToSupport()

	Sub SendToSupport()	
		'Request.Form("txtErrNumber")
		'Request.Form("txtErrSource")
		'Request.Form("txtErrDescription")
		'Request.Form("txtErrPage")
		'Request.Form("txtAdditionalDescription")
	End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="../resources/pages_style.css">
<TITLE>La Aldea� Ontica - Administrador del flujo de trabajo</TITLE>
</HEAD>
<BODY>
<H3>Recibimos la solicitud de soporte.</H3>
A la brevedad nos comunicaremos con usted v�a correo electr�nico para informarle sobre la resoluci�n de este problema.
<br><br>
Gracias.
<br><br>
Atentamente
<br><br>
El equipo de soporte t�cnico
<br><br>
<A href="../../main.asp">Regresar a la p�gina principal</A>
</BODY>
</HTML>