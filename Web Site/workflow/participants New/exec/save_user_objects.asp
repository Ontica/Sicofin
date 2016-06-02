<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnUserId, gnEntityId, gsParticipantName, gnUserObjects
	 
  Call Main()
   
  Sub Main()
		Dim oParticipant, sArray, i
		'******************
		On Error Resume Next
		gnUserId = CLng(Request.Form("txtUserId"))
		gnEntityId = CLng(Request.Form("txtEntityId"))
		gnUserObjects = Request.Form("chkItems").Count
		Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
		gsParticipantName = oParticipant.Attributes(Application("sAppServer"), CLng(gnUserId)).Fields("participantName")
		If gnUserObjects > 0 Then			
			sArray = Request.Form("chkItems")(1)
			For i = 2 To gnUserObjects
				sArray = sArray & "," & Request.Form("chkItems")(i)
			Next
			oParticipant.ObjectPermissions Application("sAppServer"), CLng(gnEntityId), CLng(gnUserId), Split(sArray, ",")
		Else
			oParticipant.ObjectPermissions Application("sAppServer"), CLng(gnEntityId), CLng(gnUserId), Null
		End If
		Set oParticipant = Nothing
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
<meta http-equiv="Pragma" content="no-cache">
<TITLE>Administración de usuarios</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="..//empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=2>
			<TABLE class=fullScrollMenu>
				<TR class=fullScrollMenuHeader>
					<TD colspan=3 class=fullScrollMenuTitle>
						La operación se efectuó correctamente
					</TD>
					<TD align=right nowrap>
						<img align=absmiddle src='../../images/close_white.gif' onclick="window.close();" alt="Cerrar">						
					</TD>
				</TR>
				<TR>
					<TD width=100% colspan=4>
						<br>
						Se modificaron los permisos de acceso de <b><%=gsParticipantName%></b>
						<img src='/empiria/images/separator.gif' width=100% height=1><br><br>
						<b>¿Qué se desea hacer?</b><br><br>
						<a href='../edit_user.asp'>Agregar un nuevo usuario(a)</a><br><br>
						<a href='../edit_user.asp?id=<%=gnUserId%>'>Regresar a editar la información de <%=gsParticipantName%></a><br><br>
						<a href='' onclick='window.close();return false;'>Cerrar esta ventana</a>
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
</HTML>