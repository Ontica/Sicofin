<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleGroupId, gnRuleDefId, gbSaveAsChild, gsCboRuleGroupPosition, gsCboOperators, gnParentRuleGroupId
		
	Call Main()

	Sub Main()
		Dim oRule, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'****************************************************************
		gnRuleGroupId = Request.QueryString("id")
		gnRuleDefId   = Request.QueryString("ruleDefId") 
		gbSaveAsChild = CBool(Request.QueryString("derivated"))
		Set oRule     = Server.CreateObject("EFARulesMgrBS.CRule")
		gsCboOperators	= oRule.CboOperators(Session("sAppServer"))
		If gbSaveAsChild Then
			gsCboRuleGroupPosition = oRule.CboRuleChilds(Session("sAppServer"), CLng(gnRuleGroupId))			
		Else
			gsCboRuleGroupPosition = oRule.CboRuleSiblings(Session("sAppServer"), CLng(gnRuleDefId), CLng(gnRuleGroupId))			
		End If
		Set oRule = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Rango de cuentas</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function validate() {
	if (document.all.txtDescription.value == '') {
		alert("Requiero el nombre del grupo.");
		document.all.txtDescription.focus();
		return false;		
	}
	if (document.all.txtFactor.value == '') {
		alert("Requiero el factor que se le aplicará al rango");
		document.all.txtFactor.focus();
		return false;
	}
	return true;
}

function sendInfo() {
	if (validate()) {
		document.frmEditor.submit();
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY scroll=no>
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
<TR bgcolor=khaki height=25>
  <TD>
		<FONT face=Arial color=maroon><STRONG>Insertar grupo de cuentas</STRONG></FONT>		
  </TD>
  <TD nowrap align=right>
		<A href='' onclick='sendInfo();return false;'>Aceptar</A>&nbsp;&nbsp;
 		<A href='' onclick='window.close();return false;'>Cancelar</A>
  </TD>
</TR>
</TABLE>
<FORM name=frmEditor action="./exec/save_rule_group.asp" method="post">
<TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
<TR>
  <TD nowrap>Insertar después de:</TD>
  <TD nowrap>
		<SELECT name=cboInsertAfterItems style="WIDTH: 260px">
		  <OPTION value=0>Insertar como primer grupo</OPTION>
		  <%=gsCboRuleGroupPosition%>
		</SELECT>
	</TD>
</TR>
<TR>
  <TD>Número de grupo:</TD>
  <TD><INPUT name=txtNumber maxlength=16 style="width:100px"></TD>
</TR>
<TR>
  <TD valign=top>Nombre del grupo:</TD>
  <TD>
		<TEXTAREA name=txtDescription ROWS=4 style="width:300px"></TEXTAREA>
	</TD>
</TR>
<TR>
  <TD>Operación sobre el grupo padre:</TD>
  <TD>
		<SELECT name=cboOperators style="WIDTH: 120px">
		  <%=gsCboOperators%>
		</SELECT>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		Factor:		
		<INPUT name=txtFactor value="1" style="width:60px">
		<INPUT name=txtRuleId type=hidden value=0>
		<INPUT name=txtRuleDefId type=hidden value=<%=gnRuleDefId%>>
		<INPUT name=txtRuleGroupId type=hidden value=<%=gnRuleGroupId%>>
		<INPUT name=txtSaveAsChild type=hidden value=<%=gbSaveAsChild%>>
  </TD>  
</TR>
<TR>
  <TD>Incluir el grupo:</TD>
  <TD>
		<SELECT name=cboFilters style="WIDTH: 180px">			
			<OPTION value=0 selected>Siempre</OPTION>
			<OPTION value=1>Sólo si su saldo es positivo</OPTION>
			<OPTION value=-1>Sólo si su saldo es negativo</OPTION>
		</SELECT>
  </TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>