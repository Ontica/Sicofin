<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
			
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function canChangePassword(userName, password) {
	var obj;
	obj = RSExecute("../workflow_scripts.asp", "CanChangePassword", userName, password);	
	return (obj.return_value);
}

function checkValues() {
	var sMsg;
			
	if (document.all.txtUserName.value == '') {
		alert('Para modificar la contrase�a requiero la identifiaci�n de acceso.');
		document.all.txtUserName.focus();
		return false;
  }
	if (document.all.txtPassword.value == '') {
		alert('Para modificar la contrase�a requiero la contrase�a actual.');
		document.all.txtPassword.focus();
		return false;
  }  
	if ( document.all.txtNewPassword.value.length < 8 ) {
		alert('Necesito que las contrase�as tengan un m�nimo de ocho caracteres.');		
		document.all.txtNewPassword.value = '';
		document.all.txtPasswordConfirmation.value = '';
		document.all.txtNewPassword.focus();
		return false;
  }
	if ( document.all.txtNewPassword.value != document.all.txtPasswordConfirmation.value ) {
		alert('La nueva contrase�a y su confirmaci�n no son iguales.');		
		document.all.txtNewPassword.value = '';
		document.all.txtPasswordConfirmation.value = '';
		document.all.txtNewPassword.focus();
		return false;
  } 
  /*
	if ( !canChangePassword(document.all.txtUserName.value, document.all.txtPassword.value) ) {	
		sMsg  = "No puedo ejecutar la operaci�n debido a que la identificaci�n de acceso\n";
		sMsg += "o la contrase�a no coinciden con las registradas.";
		alert(sMsg);
		document.all.txtUserName.focus();
		return false;
	}	
	*/
  sMsg  = "Est� operaci�n modificar� la contrase�a de acceso al sistema.\n\n";  
  sMsg += "�Contin�o con la ejecuci�n?";
  if (!confirm(sMsg)) {		
		return false;
	}	
}
	
//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=52px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="70%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Modificar mi contrase�a</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>	  
	    <A href="" onclick="window.location.href = '<%=Session("main_page")%>';return false;">Cerrar</A>
		</TD>
	</TR>
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=90%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="70%">
<FORM name=frmSend action="./exec/upd_password.asp" method="post" onsubmit="return checkValues();">
  <TR>
    <TD nowrap>Identificaci�n de acceso al sistema:<BR></TD>
    <TD>
			<INPUT name=txtUserName style="width=200">
    </TD>
  </TR>
  <TR>
    <TD>Contrase�a actual:<BR></TD>
    <TD>
			<INPUT name=txtPassword type=password style="width=200">
    </TD>
  </TR>
  <TR>
    <TD>Nueva contrase�a:<BR></TD>
    <TD>
			<INPUT name=txtNewPassword type=password style="width=200">
    </TD>
  </TR>
  <TR>
    <TD>Confirmaci�n de la nueva contrase�a :<BR></TD>
    <TD>
			<INPUT name=txtPasswordConfirmation type=password style="width=200">
    </TD>
  </TR>    
	<TR>
	  <TD nowrap align=middle><INPUT name=cmdBuild type=submit value="Aceptar">&nbsp;</TD>
	  <TD nowrap align=middle><INPUT name=cmdCancel type=button value="Cancelar" onclick="window.location.href = '<%=Session("main_page")%>';">
	  </TD>
	</TR>
</FORM>
</TABLE>
<BR>&nbsp;
</DIV>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>