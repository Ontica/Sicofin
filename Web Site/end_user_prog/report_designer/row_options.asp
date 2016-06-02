<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gnItemId, gsItemName	
		
	gnItemId = Request.QueryString("id")
		
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oRecordset
		'******************************
		'On Error Resume Next
		Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")				
		Set oRecordset = oReportDesigner.GetItem(Session("sAppServer"), CLng(gnItemId))		
		gsItemName     = oRecordset("itemName")
		oRecordset.Close
		Set oReportDesigner = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect(".exec/exception.asp")
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

function deleteRow() {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "DeleteRow", <%=gnItemId%>);	
	return(obj.return_value);	
}

function insertRow(nDirection) {
	var obj;
	
	obj = RSExecute("../end_user_prog_scripts.asp", "InsertRow", <%=gnItemId%>, nDirection);
	return(obj.return_value);	
}

function callEditor(sOperation) {
	var sURL, sPars;
	
	switch (sOperation) {
		case 'addItem':
			sURL  = 'edit_item.asp?rowId=<%=gnItemId%>';
			sPars = 'height=420px,width=560px,resizable=no,scrollbars=no,status=no,location=no';
			oWindow = createWindow(oWindow, sURL, sPars);
			window.close();
			return false;
		case 'properties':
			sURL  = 'row_properties.asp?id=<%=gnItemId%>';
			sPars = 'height=340px,width=450px,resizable=no,scrollbars=no,status=no,location=no';
			oWindow = createWindow(oWindow, sURL, sPars);
			window.close();
			return false;
		case 'insertBefore':
			if (confirm("¿Agrego un renglón antes del renglón '<%=gsItemName%>'?")) {
				insertRow(-1);
				window.opener.location.href = window.opener.location.href;
			}
			return false;
		case 'insertAfter':
			if (confirm("¿Agrego un renglón después del renglón '<%=gsItemName%>'?")) {
				insertRow(1);
				window.opener.location.href = window.opener.location.href;
			}			
			return false;
		case 'copyToAll':
			notAvailable();
			return false;
			sURL  = 'create_section.asp?id=<%=gnItemId%>';
			sPars = 'height=420px,width=580px,resizable=no,scrollbars=no,status=no,location=no';
			oWindow = createWindow(oWindow, sURL, sPars);
			window.close();
			return false;
		case 'delete':
			if (confirm("¿Elimino el renglón '<%=gsItemName%>' y todos los elementos que contiene?")) {
				deleteRow();
				window.opener.location.href = window.opener.location.href;
				window.close();
			}			
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
			<b>Operaciones sobre '<%=gsItemName%>'</b>
		</TD>
	  <TD align=right nowrap>
			<img align=absmiddle src='/empiria/images/invisible4.gif'>			<img align=absmiddle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=absmiddle src='/empiria/images/invisible.gif'>
			<img align=absmiddle src='/empiria/images/close_red.gif' onclick='window.close();' alt="Cerrar">&nbsp;
		</TD>
	</TR>
	<FORM name=frmSend action='./exec/create_section.asp' method=post>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('addItem'));">Agregar un elemento de información</A></TD>
			  </TR>
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('properties'));">Editar las propiedades de este renglón</A></TD>
			  </TR>			  			  
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>			  
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('insertBefore'));">Insertar un renglón antes de este renglón</A></TD>
			  </TR>
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>			  
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('insertAfter'));">Insertar un renglón después de este renglón</A></TD>
			  </TR>			  
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('copyToAll'));">Heredar las propiedades de este renglón a todos los de su sección</A></TD>
			  </TR>
				<TR>
					<TD nowrap>&nbsp;</TD>
			  </TR>			  			  
				<TR class=applicationTableRowDivisor>
					<TD nowrap><A href='' onclick="return(callEditor('delete'));">Eliminar el renglón con todos sus elementos</A></TD>
			  </TR>
			</TABLE>
		</TD>
	</TR>
	</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>