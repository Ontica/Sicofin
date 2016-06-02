<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	
	Dim gnItemId, gsItemName, gnReportDataId, gsPosition, gsCboFilters, gsCboPrintConditions, gsPrintLayout
	
	Call Main()
	
	Sub Main()
		Dim oReportDesigner, oRecordset, nReportId, nFilterId
		'*****************************************
		'On Error Resume Next						
		gnItemId = Request.QueryString("id")				
	  Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDesigner")
	  Set oRecordset = oReportDesigner.GetItem(Session("sAppServer"), CLng(gnItemId))
	  nReportId  = oRecordset("reportId")
	  gsPosition = oRecordset("itemRow")
	  gsItemName = oRecordset("itemName")
	  nFilterId  = oRecordset("itemFilterId")
	  oRecordset.Close
	  Set oRecordset = oReportDesigner.GetReport(Session("sAppServer"), CLng(nReportId))
	  gnReportDataId = oRecordset("reportDataSourceId")
	  If CLng(oRecordset("reportDataFilterId")) <> 0 Then
			nFilterId = oRecordset("reportDataFilterId")
		End If
	  oRecordset.Close
	  Set oReportDesigner = Nothing
	  Set oReportDesigner = Server.CreateObject("AOReportsDesigner.CDictionary")
	  gsCboFilters = oReportDesigner.CboDataItemFilters(Session("sAppServer"), CLng(gnReportDataId), CLng(Session("uid")), CLng(nFilterId))	  
	  Set oRecordset = Nothing
	  Set oReportDesigner = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Diseñador de reportes</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oBuildWindow = null;
var gbSended = false;

function doSubmit() {
	var sMsg, nVoucherId;
	
	if (gbSended) {
		return false;
	}
	if (document.all.txtName.value == '') {
		alert('Requiero el nombre del renglón.');
		document.all.txtName.focus();
		return false;
	}
	gbSended = true;
	document.all.frmSend.submit();
	return true;
}

function openWindow(sWindowName) {
	var sURL, sPars;
	
	switch (sWindowName) {	
		case 'createFilter':
			sURL = '../dictionary/edit_filter.asp?id=0&classId=<%=gnReportDataId%>';
			sPars = 'height=380px,width=400px,resizable=no,scrollbars=no,status=no,location=no';
			oBuildWindow = createWindow(oBuildWindow, sURL, sPars);
			return false;
		case 'createOperation':
			notAvailable();
/*			
			if (oReportItemWindow == null || oReportItemWindow.closed) {
				oReportItemWindow = window.open(sURL, '_blank', sPars);
			} else {
				oReportItemWindow.focus();
				oReportItemWindow.navigate(sURL);
			}
*/
			return false;
	}	
	return false;	
}

function openFile(sFileName) {
	return false;
}

function buildOperation() {
	notAvailable();	
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<FORM name=frmSend action='./exec/save_row.asp' method=post>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>		
			Propiedades del renglón
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
					<TD colspan=4><b>Identificación del renglón</b></TD>
			  </TR>			
			  <TR>
					<TD valign=top nowrap>Nombre del renglón:</TD>
			    <TD colspan=3 width=430>
						<INPUT name=txtName style="height:20px;width:320px;" value='<%=gsItemName%>'>
			    </TD>	    
			  </TR>			  
			  <TR>
					<TD valign=top nowrap>Posición del renglón:</TD>
			    <TD colspan=3 width=430>
						<b><%=gsPosition%></b>
			    </TD>	    
			  </TR>
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Filtro sobre todos los elementos del renglón que apliquen</b></TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>Filtrar la información mediante:</TD>
					<TD colspan=3>
						<SELECT name=cboFilters style="width:320">
							<OPTION value=0>-- No aplicar ningún filtro --</OPTION>
							<%=gsCboFilters%>
						</SELECT>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='width:110' value="Crear filtro ..." onclick="openWindow('createFilter');">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='width:110' value="Editar filtro ..." onclick="openWindow('editFilter');">
						<br>&nbsp;
					</TD>
			  </TR>
			  <TR class=applicationTableRowDivisor>
					<TD colspan=4><b>Condición para la impresión del renglón</b></TD>
			  </TR>
			  <TR nowrap>
					<TD valign=top>Imprimir el renglón:</TD>
					<TD colspan=3>
						<SELECT name=cboPrintConditions style="width:320">
							<OPTION value=0>-- Siempre --</OPTION>
							<OPTION value=-1>-- Nunca --</OPTION>
							<%=gsCboPrintConditions%>
						</SELECT>
						<br><br>
						&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='width:110' value="Crear condición ..." onclick="openWindow('createCondition');">
						&nbsp; &nbsp;
						<INPUT class=cmdSubmit type=button name=cmdFilter style='width:110' value="Editar condición ..." onclick="openWindow('editCondition');">						
						<br>&nbsp;
					</TD>
			  </TR>			  
			</TABLE>
		</TD>
	</TR>
	<TR>
	  <td colspan=4 nowrap align=right>
	   <INPUT type="hidden" name=txtItemId value="<%=gnItemId%>">
	   <INPUT type="hidden" name=txtPrintLayout value="<%=gsPrintLayout%>">
		 <INPUT class=cmdSubmit type=button name=cmdSend style='width:150' value="Parámetros de impresión ..." onclick="printLayout();">
		 &nbsp; &nbsp; &nbsp;
		 <INPUT class=cmdSubmit type=button name=cmdSend style='width:70' value="Aceptar" onclick="doSubmit();">						
			&nbsp; &nbsp;
	   <INPUT class=cmdSubmit type=button name=cmdCancel style='width:70' value="Cancelar" onclick='window.close();'>
	   &nbsp; &nbsp; &nbsp;
	  </td>
	</TR>	
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>