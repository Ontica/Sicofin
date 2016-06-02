<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleDefId, gnRuleId, gsNumber, gsName, gsCboOperators, gsCboReferencedGroups, gsFactor, gsFilter
		
	Call Main()

	Sub Main()
		Dim oRule, oRuleDef, oRecordset, nSubsidiaryLedgerId, nSubsidiaryAccountId
		Dim nGLAccountId, nSectorId
		'****************************************************************
		gnRuleId			 = Request.QueryString("id")
		gnRuleDefId		 = Request.QueryString("ruleDefId")
		Set oRule			 = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset = oRule.RuleRS(Session("sAppServer"), CLng(gnRuleId))
		gsNumber			 = oRecordset("numero_grupo_cuenta")
		gsName				 = oRecordset("nombre_grupo_cuenta")
		gsCboOperators = oRule.CboOperators(Session("sAppServer"), CStr(oRecordset("operador")))
		gsFactor			 = oRecordset("factor")
		If IsNull(oRecordset("filtro_cuentas_origen")) Then 
			gsFilter = ""
		Else
			gsFilter = oRecordset("filtro_cuentas_origen")
		End If	
		Set oRule = Nothing
		
		Set oRuleDef = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		gsCboReferencedGroups = oRuleDef.CboRules(Session("sAppServer"), _
																							CLng(gnRuleDefId), CLng(oRecordset("id_agrupador_origen")), _
																							CLng(oRecordset("id_grupo_cuenta_padre")) )
		Set oRuleDef = Nothing
		Set oRecordset  = Nothing
	End Sub	
%>
<HTML>
<HEAD>
<TITLE>Base de conocimiento contable</TITLE>
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
		<FONT face=Arial color=maroon><STRONG>Edición de la referencia</STRONG></FONT>		
  </TD>
  <TD nowrap align=right>
		<A href='' onclick='sendInfo();return false;'>Aceptar</A>&nbsp;&nbsp;
 		<A href='' onclick='window.close();return false;'>Cancelar</A>
  </TD>
</TR>
</TABLE>
<FORM name=frmEditor action="./exec/save_group_reference.asp" method="post">
<TABLE border=0 cellPadding=2 cellSpacing=1 width="100%">
<TR>
  <TD>Número de la referencia:</TD>
  <TD><INPUT name=txtNumber maxlength=16 value="<%=gsNumber%>" style="width:100px"></TD>
</TR>
<TR>
  <TD valign=top>Nombre de la referencia:</TD>
  <TD>
		<TEXTAREA name=txtDescription ROWS=4 style="width:300px"><%=gsName%></TEXTAREA>
	</TD>
</TR>
<TR>
  <TD valign=top>Grupo referenciado:</TD>
  <TD>
  	<SELECT name=cboReferencedGroups style="WIDTH: 100%">
			<%=gsCboReferencedGroups%>
		</SELECT>
	</TD>
</TR>
<TR>
  <TD>Operación sobre el grupo padre:</TD>
  <TD>
		<SELECT name=cboOperators style="WIDTH: 180px">
		  <%=gsCboOperators%>
		</SELECT>
		&nbsp;&nbsp;&nbsp;
		Factor:
		<INPUT name=txtFactor value="<%=gsFactor%>" style="width:60px">
		<INPUT name=txtRuleId type=hidden value=<%=gnRuleId%>>		
  </TD>
</TR>
<TR>
  <TD>Incluir la referencia:</TD>
  <TD>
		<SELECT name=cboFilters style="WIDTH: 180px">
			<% If (Len(gsFilter) = 0) Then %>
			<OPTION value=0 selected>Siempre</OPTION>
			<OPTION value=1>Sólo si su saldo es positivo</OPTION>
			<OPTION value=-1>Sólo si su saldo es negativo</OPTION>			
			<% ElseIf gsFilter = "(saldo_actual > 0)" Then %>
			<OPTION value=0>Siempre</OPTION>
			<OPTION value=1 selected>Sólo si su saldo es positivo</OPTION>
			<OPTION value=-1>Sólo si su saldo es negativo</OPTION>
			<% ElseIf gsFilter = "(saldo_actual < 0)" Then %>
			<OPTION value=0>Siempre</OPTION>
			<OPTION value=1>Sólo si su saldo es positivo</OPTION>
			<OPTION value=-1 selected>Sólo si su saldo es negativo</OPTION>
			<% End If %>
		</SELECT>
  </TD>
</TR>
</TABLE>
</FORM>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>