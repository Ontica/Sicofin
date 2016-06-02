<HTML>
<HEAD>
<TITLE>Editor de movimientos</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var oLastItem = null;

function oPostingRef() {
	var gralLedgerId;
	var postingId;
}

function loadPostingsTable() {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp", "PendingPostingsReferences", oPostingRef.gralLedgerId, oPostingRef.postingId);
	document.all.pendingPostingsTable.innerHTML = obj.return_value;
}


function pickData() {
  if (window.dialogArguments != null) {
      window.dialogArguments.postingId = oPostingRef.postingId;
  }
  window.returnValue = true;
  window.close();
}

function loadArguments() {
	if (window.dialogArguments != null) {
		oPostingRef.gralLedgerId = window.dialogArguments.gralLedgerId;
		oPostingRef.postingId    = window.dialogArguments.postingId;
	}
	if (oPostingRef.postingId == -1) {
		document.all.optReferenceType[0].click();
		window.returnValue = false;
		return false;
	}
	if (oPostingRef.postingId == 0) {
		document.all.optReferenceType[2].click();
		window.returnValue = false;
		return false;
	}
	if (oPostingRef.postingId > 0) {
		document.all.optReferenceType[1].click();
		loadPostingsTable();	
		window.returnValue = false;
		return false;
	}	
}

function selectItem() {
	var oItem = event.srcElement;
	var oRow = getObjectParent(oItem, 'TR');
	
	if (oLastItem != null) {
		getObjectParent(oLastItem, 'TR').className = '';
	}
	oPostingRef.postingId = oItem.value;
	oRow.className = 'applicationTableSelectedRow';
	oLastItem = oItem;
}

function changeReferenceType() {
	if (document.all.optReferenceType[0].checked) {
		document.all.divPostingsTable.style.display = 'none';
		oPostingRef.postingId = -1;
		oLastItem = null;
		return true;
	}
	if (document.all.optReferenceType[1].checked) {
		loadPostingsTable();
		document.all.divPostingsTable.style.display = 'inline';
		return true;
	}
	if (document.all.optReferenceType[2].checked) {
		document.all.divPostingsTable.style.display = 'none';
		oPostingRef.postingId = 0;
		oLastItem = null;
		return true;
	}	
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox LANGUAGE=javascript onload="return loadArguments()">
<TABLE class=standardPageTable>
	<TR>
		<TD nowrap class=applicationTitle>
			Movimiento de referencia
		</TD>
		<TD align=right nowrap>
			<img align=absMiddle src="/empiria/images/help_red.gif" onclick='notAvailable();' alt="Ayuda">	
			<img align=absMiddle src="/empiria/images/invisible.gif">
			<img align=absMiddle src="/empiria/images/close_red.gif" onclick='window.close();' alt="Cerrar">
		</TD>
	</TR>
  <TR>
		<TD colspan=2 nowrap>
			<TABLE class=applicationTable cellpadding=1>
				<TR>
					<TD nowrap>Contabilidad:</TD>
					<TD width=100%><b><span id=gralLedgerName></span></b></TD>
				</TR>			
				<TR>
				  <TD nowrap>Tipo de referencia:</TD>
				  <TD nowrap>
						<INPUT type="radio" name=optReferenceType value=1 onclick='return changeReferenceType();'>Movimiento de iniciativa
						<INPUT type="radio" name=optReferenceType value=2 onclick='return changeReferenceType();'>Movimiento de conformidad						
						&nbsp; &nbsp; &nbsp;
						<INPUT type="radio" name=optReferenceType value=0 onclick='return changeReferenceType();'>Ninguno
					</TD>
				</TR>
				<TR id=divPostingsTable style='display:none;'>
				  <TD nowrap>
						Iniciativas pendientes:<br><br>
						<a href=''>Refrescar</a>
					</TD>
				  <TD nowrap>
						<SPAN id=pendingPostingsTable STYLE="overflow:auto; float:bottom; width=99%; height=180px">
						</SPAN>
					</TD>
				</TR>
				<TR>  
				  <TD colspan=2 align=right nowrap>
						<INPUT class=cmdSubmit name=cmdAddItem type=button value="Aceptar" style="WIDTH: 60px" onclick='pickData();'>&nbsp;&nbsp;&nbsp;
						<INPUT class=cmdSubmit name=cmdCancel type=button value="Cancelar" style="WIDTH: 60px" onclick='window.close();'>
						&nbsp; &nbsp; &nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>