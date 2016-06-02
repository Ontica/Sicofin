<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gbEdit
	Dim gsAccountKey, gsAccountName, gnGralLedgerId, gnStdAccountId
	Dim gsSubsidiaryLedgersCheckBoxList  
	
	gsTitle = "Asignación de mayores auxiliares"

	gnGralLedgerId = Request.QueryString("gralLedgerId")
	gnStdAccountId = Request.QueryString("stdAccountId")
	Call Main(gnGralLedgerId, gnStdAccountId)

	Sub Main(nGralLedgerId, nStdAccountId)
		Dim oGralLedgerUS, oRecordset
		'****************		
		Set oGralLedgerUS = Server.CreateObject("AOGralLedgerUS.CServer")		
		gsSubsidiaryLedgersCheckBoxList = oGralLedgerUS.SubsidiaryLedgersCheckBoxList(Session("sAppServer"), CLng(nGralLedgerId))		
		Set oRecordset = oGralLedgerUS.GetStdAccountRS(Session("sAppServer"), CLng(nStdAccountId))		
		gsAccountKey = oRecordset("numero_cuenta_estandar")
		gsAccountName = oRecordset("nombre_cuenta_estandar")
		oRecordset.Close
		Set oRecordset = Nothing
		Set oGralLedgerUS = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

window.returnValue = 0;

function countCheckBoxes(sCheckBoxName) {
	var i= 0, counter = 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				counter++;
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			counter++;
		}
	}
	return counter;
}

function selectCheckBoxes(sCheckBoxName, bCheck) {
	var i= 0;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			document.all[sCheckBoxName](i).checked = bCheck;
		}		
	} else {
		document.all[sCheckBoxName].checked = bCheck;		
	}
	return true;	
}

function addSubsidiaryLedger(nSubsidiaryLedgerId, nAccountId) {
	var obj;
	obj = RSExecute("../transactions/exec/scripts_voucher.asp", "AssignSubsidiaryLedger", nSubsidiaryLedgerId, nAccountId);
}

function deleteSubsidiaryLedger(nSubsidiaryLedgerId, nAccountId) {

}

function updateSubsidiaryLedgers(sCheckBoxName) {
	var obj;
	var nAccountId = 0, i= 0;
	
	obj = RSExecute("../transactions/exec/scripts_voucher.asp", "AddAccount", <%=gnGralLedgerId%>, <%=gnStdAccountId%>);
	nAccountId = obj.return_value;
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				addSubsidiaryLedger(document.all[sCheckBoxName](i).value, nAccountId);
			} else {
				deleteSubsidiaryLedger(document.all[sCheckBoxName](i).value, nAccountId);
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			addSubsidiaryLedger(document.all[sCheckBoxName].value, nAccountId);
		} else {
			deleteSubsidiaryLedger(document.all[sCheckBoxName].value, nAccountId);
		}			
	}
	return (nAccountId);
}

function save() {
	var counter = countCheckBoxes('chkSubsidiaryLedgers');		
	var sMsg, nAccountId = 0;
	if (counter == 0) {		
		alert("Requiero la selección de al menos uno de los mayores auxiliares.");
		return false;
	}
	if (counter == 1) {
		sMsg = "Se asignará el mayor auxiliar seleccionado a la cuenta <%=gsAccountKey%>.\n\n" + 
					 "¿Procedo con la operación?";
	}
	if (counter > 1) {
		sMsg = "Se asignarán los " + counter + " mayores auxiliares seleccionados a la cuenta\n" + 
					 "<%=gsAccountKey%>.\n\n" + 
					 "¿Procedo con la operación?";
	}	
	if (confirm(sMsg)) {
		nAccountId = updateSubsidiaryLedgers('chkSubsidiaryLedgers');
		alert("Los cambios fueron efectuados satisfactoriamente");
		window.returnValue = nAccountId;
		window.close();
		return true;
	} else {
		return false;
	}
	if (counter != 0) {
		alert("Requiero la selección de al menos uno de los mayor auxiliares.");
		return false;
	}	
  return true;
}  

//-->
</SCRIPT>
</HEAD>
<BODY scroll=no>
<FORM name=frmEditor action="exec/assign_subsidiary_ledgers.asp" method="post">
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
<TR>  
  <TD colSpan=3 bgcolor=khaki><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>  
</TR>
<TR>  
  <TD colSpan=3 bgcolor=khaki><FONT face=Arial color=maroon><STRONG>Cuenta: <%=gsAccountKey%></STRONG></FONT>
  </TD>  
</TR>
<TR>  
  <TD colSpan=3 bgcolor=khaki><FONT face=Arial color=maroon><STRONG><%=gsAccountName%></STRONG></FONT>  
  </TD>  
</TR>
<TR>
  <TD colspan=3>
		<%=gsSubsidiaryLedgersCheckBoxList%>
	</TD>
</TR>
<TR>  
  <TD align=right>
		<INPUT name=cmdAddItem type=button value="Agregar" LANGUAGE=javascript onclick="return save()">&nbsp;&nbsp;&nbsp;
		<INPUT name=cmdCancel type=button value="Cancelar" onclick="window.close();">
	</TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
