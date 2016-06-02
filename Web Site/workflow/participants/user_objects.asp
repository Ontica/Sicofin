<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnUserId, gbEditMode
	Dim gsLastName1, gsLastName2, gsName, gsIsFemale, gsBornDate, gsJobEMail, gsJobPhone, gsHistoricDate
	Dim gsParticipantKey, gsParticipantName, gsStatus

	Dim gnEntityId, gsObjectTypeName, gsObjectsTable, gsTitle
	
	Call Main()
		
	Sub Main()
		Dim oParticipant, oRecordset
		'***************************
		On Error Resume Next
		
		If (Len(Request.QueryString("id")) <> 0) Then
			gnUserId = Request.QueryString("id")
			gbEditMode = True
		Else
			gnUserId = 0
			gbEditMode = False
		End If
				
		gnEntityId = Request.QueryString("typeId")
		If gnEntityId = 9 Then
			gsObjectTypeName = "Contabilidades"
		ElseIf gnEntityId = 15 Then
			gsObjectTypeName = "Reportes contables"
		End If
		Set oParticipant   = Server.CreateObject("MHParticipantsMgr.CParticipant")
		Set oRecordset     = oParticipant.Attributes(Session("sAppServer"), CLng(gnUserId))		  		
		gsTitle						 = Left(oRecordset("participantName"), 30) & " (Objetos)"
		Set oRecordset     = Nothing
		gsObjectsTable     = oParticipant.PermissionsTable(Session("sAppServer"), CLng(gnEntityId), CLng(gnUserId))		
		Set oParticipant   = Nothing
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
<link REL="stylesheet" TYPE="text/css" HREF="../resources/mahler.css">
<script language="JavaScript" src="../programs/client_scripts.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function save() {
	document.frmSend.submit()
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmSend action='./programs/save_user_objects.asp' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			<%=gsTitle%>
		</TD>
		<TD colspan=3 align=right nowrap>
			<img align=absmiddle src='../images/help_white.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='../images/invisible.gif'>
			<img align=absmiddle src='../images/close_white.gif' onclick="window.close();" alt="Cerrar">								</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class='fullScrollMenu'>
				<TR class="fullScrollMenuHeader">					
					<TD nowrap>
						<b><%=gsObjectTypeName%></b>
					</TD>
					<TD nowrap align=right> 
					  <a href='' onclick='return(notAvailable());'>Seleccionar todo</a>
						&nbsp; | &nbsp;
						<A href="" onclick="return(save());">Guardar</A>
						&nbsp; &nbsp;
						<img align=absbottom src='../images/refresh_white.gif' onclick='document.all.frmSend.reset();' alt="Refrescar">						
					</TD>
				</TR>
			</TABLE>
			<DIV STYLE="overflow:auto; float:bottom; width=100%; height=385px">
			<TABLE class=applicationTable>
				<%=gsObjectsTable%>
			</TABLE>
			</DIV>
		</TD>
	</TR>
</TABLE>
<INPUT type=hidden name=txtUserId value=<%=gnUserId%>>
<INPUT type=hidden name=txtEntityId value=<%=gnEntityId%>>

</FORM>
</BODY>
</HTML>
