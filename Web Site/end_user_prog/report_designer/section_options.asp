<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gnSectionId, gsSectionName
		
	gnSectionId = Request.QueryString("id")
	
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oRecordset
		'******************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")		
		Set oRecordset = oReportDesigner.GetSection(Session("sAppServer"), CLng(gnSectionId))
		gsSectionName  = oRecordset("sectionName")
		oRecordset.Close
		Set oReportDesigner = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("./exec/exception.asp")
		End If		  
	End Sub
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Diseñador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oWindow = null;

function deleteSection() {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "DeleteSection", <%=gnSectionId%>);	
	return(obj.return_value);	
}

function callEditor(sOperation) {
	var sURL, sPars;
	
	switch (sOperation) {
		case 'properties':
			sURL  = 'create_section.asp?id=<%=gnSectionId%>';
			sPars = 'height=420px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oWindow = createWindow(oWindow, sURL, sPars);
			window.close();
			return false;
		case 'editRows':
			sURL  = 'create_section.asp?id=<%=gnSectionId%>';
			sPars = 'height=420px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oWindow = createWindow(oWindow, sURL, sPars);
			window.close();
			return false;
		case 'delete':
			if (confirm("¿Elimino el renglón '<%=gsSectionName%>' y todos los elementos que contiene?")) {
				deleteSection();
				window.opener.location.href = window.opener.location.href;
				window.close();
			}
		case 'divideSection':
			notAvailable();
			return false;		
	}	
	return false;
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap>		
			Operaciones sobre <%=gsSectionName%>
		</TD>
	  <TD align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;
		</TD>
	</TR>	
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('properties'));">Editar las propiedades de la sección</A></TD>
			  </TR>
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>			  			
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('editRows'));">Agregar o eliminar renglones de la sección</A></TD>
			  </TR>
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>			  
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('divideSection'));">Dividir la sección en forma horizontal</A></TD>
			  </TR>			  
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('delete'));">Eliminar la sección con todos sus elementos</A></TD>
			  </TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>