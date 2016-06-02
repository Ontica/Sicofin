<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnWorkgroupId, gbEditMode
	Dim gsLastName1, gsLastName2, gsName, gsIsFemale, gsBornDate, gsJobEMail, gsJobPhone, gsHistoricDate
	Dim gsParticipantKey, gsParticipantName, gsStatus

	Dim gsWorkgroupUsers, gsTitle
	
	Call Main()
		
	Sub Main()
		Dim oWorkgroup, oRecordset
		'***************************
		On Error Resume Next
		
		If (Len(Request.QueryString("id")) <> 0) Then
			gnWorkgroupId = Request.QueryString("id")
			gbEditMode = True
		Else
			gnWorkgroupId = 0
			gbEditMode = False
		End If
														
		Set oWorkgroup   = Server.CreateObject("MHParticipantsMgr.CParticipant")
		Set oRecordset   = oWorkgroup.Attributes(Session("sAppServer"), CLng(gnWorkgroupId))
		gsTitle					 = Left(oRecordset("participantName"), 40)
		Set oRecordset   = Nothing
		gsWorkgroupUsers = oWorkgroup.RelatedParticipantsTable(Session("sAppServer"), CLng(gnWorkgroupId), 1301)
		Set oWorkgroup   = Nothing
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
<FORM name=frmSend action='./programs/save_workgroup_users.asp' method=post>
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
						<b>Usuarios en el grupo de trabajo</b>
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
			<DIV STYLE="overflow:auto; float:bottom; width=100%; height=200px">
			<TABLE class=applicationTable>
				<%=gsWorkgroupUsers%>
			</TABLE>
			</DIV>
		</TD>
	</TR>
</TABLE>
<INPUT type=hidden name=txtWorkgroupId value=<%=gnWorkgroupId%>>
</FORM>
</BODY>
</HTML>
