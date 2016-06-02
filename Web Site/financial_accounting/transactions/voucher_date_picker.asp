<HTML>
<HEAD>
<TITLE>Fecha de la p�liza</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script language="JavaScript" src="/empiria/bin/client_scripts/general.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

window.returnValue = '';

function validateDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function formatDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","FormatDate", date, 'dd/mmm/yyyy');
	return obj.return_value;	
}

function pickData() {
	var sDate = document.all.txtDate.value;
	
  if (sDate == '') {
     alert('Necesito se proporcione la fecha valor o adelantada que tendr� la p�liza.');
     document.all.txtDate.focus;
     return false;
  }
  if (!validateDate(sDate)) {
     alert('La fecha proporcionada tiene un formato que no reconozco.');
     document.all.txtDate.focus;
     return false;
  }
	window.returnValue = formatDate(document.all.txtDate.value);
	window.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=3 nowrap class=applicationTitle>
			Fecha de la p�liza
		</TD>
	  <TD align=right nowrap>
			<img align=middle src='/empiria/images/help_red.gif' onclick='notAvailable();' alt="Ayuda">			<img align=middle src='/empiria/images/invisible.gif'>
			<img align=middle src='/empiria/images/close_red.gif' onclick="window.close();" alt="Cerrar">
		</TD>
	</TR>
	<TR>
		<TD colspan=4 nowrap>
			<TABLE class=applicationTable>
				<TR>
				  <TD colspan=2>
						Las p�lizas con <b>fecha de afectaci�n atrasada (fecha valor)</b> son 
						enviadas para su autorizaci�n al �rea correspondiente. El
						personal de dicha �rea es quien, de ser el caso, ingresar� la p�liza en el diario.<br><br>
						Si una p�liza tiene una <b>fecha de afectaci�n futura o adelantada</b>
						s�lo podr� ser enviada al diario una vez que se llegue a esa fecha.
						Llegado el d�a, el administrador de flujo de trabajo nos recordar� que es momento 
						de incorporar nuestra p�liza en el diario.
				  </TD>  
				</TR>
				<TR>
				  <TD nowrap>Fecha valor o adelantada:</TD>
				  <TD nowrap>
						<INPUT name=txtDate value="" style="WIDTH: 80px">
						<img align=absbottom src='/empiria/images/calendar.gif' alt='Despliega el calendario' onclick='return(showCalendar(document.all.txtDate));'>
						(d�a / mes / a�o)
				  </TD>
				</TR>
				<TR>
					<TD>&nbsp;</TD>
				  <TD align=right>
						<INPUT class=cmdSubmit name=cmdAddItem type=button value="Aceptar" LANGUAGE=javascript onclick="return pickData()">&nbsp;&nbsp;&nbsp;
						<INPUT class=cmdSubmit name=cmdCancel type=button value="Cancelar" onclick="window.close();">
						&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
