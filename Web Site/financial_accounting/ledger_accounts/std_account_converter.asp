<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsTitle, gnStdAccountId
	Dim gnStdAccountCatId, gsStdAccountNumber, gsStdAccountName
	Dim gsStdAccountParentName, gsStdAccountParentNumber
	Dim gnStdAccountCatalogName, gsCboStdAccountRoles
	

	Call Main(CLng(Request.Form("txtStdAccountCatId")), CLng(Request.Form("txtItemId")))

	Sub Main(nStdAccountCatId, nStdAccountId)
		Dim oStdAccountMgr, oRecordset, nStdAccountParentId
		'**************************************************
		gsTitle = "Conversión por eliminación"
		gnStdAccountCatId       = nStdAccountCatId
		gnStdAccountId					= nStdAccountId
		Set oStdAccountMgr      = Server.CreateObject("EFAStdActUS.CServer")		
		Set oRecordset					= oStdAccountMgr.GetStdAccount(Session("sAppServer"), CLng(nStdAccountId))
		gnStdAccountCatalogName = oRecordset("nombre_tipo_cuentas_std")
		gsStdAccountNumber      = oRecordset("numero_cuenta_estandar")
		gsStdAccountName        = oRecordset("nombre_cuenta_estandar")
		Set oRecordset          = Nothing
		nStdAccountParentId     = oStdAccountMgr.StdAccountParentId(Session("sAppServer"), CLng(nStdAccountId))
		If (nStdAccountParentId <> 0) Then
			Set oRecordset			     = oStdAccountMgr.GetStdAccount(Session("sAppServer"), CLng(nStdAccountParentId))			
			gsStdAccountParentName   = oRecordset("nombre_cuenta_estandar")
			gsStdAccountParentNumber = oRecordset("numero_cuenta_estandar")
			gsCboStdAccountRoles     = oStdAccountMgr.CboStdAccountRoles(Session("sAppServer"))
		End If		  
		Set oStdAccountMgr = Nothing		
	End Sub
	
%>
<HTML>
<HEAD>
<TITLE><%=gsTitle%></TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
function cmdNext_onclick() {
	var sMsg;
	
	sMsg = "¿Elimino la cuenta '<%=gsStdAccountNumber%>' del catálogo '<%=gnStdAccountCatalogName%>'?";
	if (confirm(sMsg)) {
		document.frmEditor.submit();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<br>
<FORM name=frmEditor action="exec/delete_std_account.asp" method="post">
<TABLE align=center border=0 cellPadding=3 cellSpacing=0 width="500px">
<TR bgcolor=khaki>  
  <TD	nowrap><FONT face=Arial size=3 color=maroon><STRONG><%=gsTitle%></STRONG></FONT></TD>
	<TD colspan=3 align=right nowrap>		
		<A href="standard_accounts.asp">Regresar al catálogo estándar</A>&nbsp;&nbsp;&nbsp;&nbsp;
		<A href="" onclick="window.location.href = '<%=Application("main_page")%>';return false;">Cerrar</A>
	</TD>	
</TR>
<TR bgcolor=LightCoral>  
  <TD	nowrap colspan=4 align=right><FONT face=Arial size=3 color=maroon><STRONG>Último paso</STRONG></FONT></TD>
</TR>
</TABLE>
<TABLE align=center border=1 cellPadding=3 cellSpacing=0 width="500px">
<TR>
  <TD colspan=4>
  <FONT face=Arial size=2 color=maroon><STRONG>Importante: </STRONG></FONT><br><br>
  Probablemente la eliminación de esta cuenta genere una o más pólizas de conversión.<br><br>
  Es decir, se generará una póliza para cada una de las contabilidades
  en donde la cuenta a eliminar tenga un saldo distinto de cero, trasladado dicho saldo
  a la cuenta madre (<%=gsStdAccountParentNumber%>) y dejando el saldo de la cuenta eliminada (<%=gsStdAccountNumber%>) en cero.<br><br>
  </TD>
</TR>
<TR>
  <TD nowrap>Catálogo de cuentas:</TD>  
  <TD colspan=3><b><%=gnStdAccountCatalogName%></b></TD>
</TR>
<TR>
	<TD nowrap valign=top>Cuenta madre:</TD>
	<TD colspan=3>
		<b><%=gsStdAccountParentNumber%></b><br>
		<%=gsStdAccountParentName%>
	</TD>
</TR>
<TR>
  <TD nowrap valign=top>Cuenta a eliminar:</TD>
  <TD colspan=3>
		<b><%=gsStdAccountNumber%></b><br>
		<%=gsStdAccountName%>
	</TD>
</TR>
<TR>
  <TD>Después de la conversión,<br>¿cuál será el nuevo rol de la cuenta madre?</TD>
  <TD colspan=3 valign=top>
		<SELECT name=cboParentNewRole style="WIDTH:250px">
			<%=gsCboStdAccountRoles%>
		</SELECT>
		<INPUT name=txtStdAccountId type="hidden" value="<%=gnStdAccountId%>">
	</TD>
</TR>
<TR>
	<TD><INPUT name=cmdCancel type=button value="Cancelar" onclick="window.location.href='standard_accounts.asp';"></TD>
	<TD colspan=3 nowrap align=right>	
		<INPUT name=cmdNext type=button value="Iniciar conversión" LANGUAGE=javascript onclick="return cmdNext_onclick()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	</TD>  
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
