<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gsGroupName, gnRuleDefId, gnRuleGroupId, gnRuleChildsCount, gnRuleChildsType
	
	Call Main()

	Sub Main()
		Dim oRule, oRecordset
		'*************************************************************
		gnRuleDefId				= Request.QueryString("ruleDefId")
		gnRuleGroupId     = Request.QueryString("id")
		Set oRule         = Server.CreateObject("EFARulesMgrBS.CRule")
		Set oRecordset    = oRule.RuleRS(Session("sAppServer"), CLng(gnRuleGroupId))		
		gsGroupName       = oRecordset("nombre_grupo_cuenta")
		gnRuleChildsCount = CLng(oRule.RuleChildsCount(Session("sAppServer"), CLng(gnRuleGroupId)))
		gnRuleChildsType  = CLng(oRule.RuleChildsType(Session("sAppServer"), CLng(gnRuleGroupId)))
		Set oRule = Nothing
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Agrupador de reglas</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

 function deleteRuleGroup() {
   var obj;
   
   obj = RSExecute("../financial_accounting_scripts.asp", "DeleteRuleGroup", <%=gnRuleGroupId%>);
	 window.close();
   return false
 }
 
 function callEditor(nOption) {
	 var sMsg = '';
	 switch(nOption) {
		case 1:
		 	window.location = "add_group_reference.asp?id=<%=gnRuleGroupId%>&ruleDefId=<%=gnRuleDefId%>"
			break;			
	  case 2:
		 	window.location = "add_rule.asp?id=<%=gnRuleGroupId%>&ruleDefId=<%=gnRuleDefId%>"
			break;
	  case 3:
			window.location = "add_rule_group.asp?id=<%=gnRuleGroupId%>&ruleDefId=<%=gnRuleDefId%>&derivated=false"
			break;
	  case 4:
			window.location = "add_rule_group.asp?id=<%=gnRuleGroupId%>&ruleDefId=<%=gnRuleDefId%>&derivated=true"
			break;
	  case 5:
			window.location = "edit_rule_group.asp?id=<%=gnRuleGroupId%>"
			break;			
	  case 6:
			<% If (gnRuleChildsCount = 0) Then %>
			  sMsg = '¿Elimino el grupo <%=gsGroupName%>?';
			<% ElseIf (gnRuleChildsCount = 1) Then %>
				sMsg = 'El grupo <%=gsGroupName%> tiene un elemento.\n¿Procedo con la eliminación?';
			<% ElseIf (gnRuleChildsCount > 1) Then %>
				sMsg = 'El grupo <%=gsGroupName%> tiene <%=gnRuleChildsCount%> elementos.\n¿Procedo con la eliminación?';
			<% End If %>
			if (confirm(sMsg)) {
				deleteRuleGroup();
			}			
			break;
	  case 7:
			alert("la operación 'mover' aún no está disponible");
			break;
	  case 8:
			alert("La operación 'copiar' aún no está disponible");
			break;
	 }
 }
 
//-->
</SCRIPT>
</HEAD>
<BODY scroll=no>
<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
<TR bgcolor=khaki height=30>
  <TD nowrap>
		<FONT face=Arial color=maroon><STRONG>
			¿Qué se desea hacer?
		</STRONG></FONT>
  </TD>
  <TD align=right>
		<A href='' onclick='window.close();'>Cerrar</A>
  </TD>
</TR>
</TABLE>
<br>
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
<% If (gnRuleChildsType > 0) Or (gnRuleChildsCount = 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(1);return false;'>Insertar una referencia a otro grupo dentro del grupo <b><%=gsGroupName%></b></A>
		<br>&nbsp;
	</TD>
</TR>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(2);return false;'>Insertar un rango de cuentas dentro del grupo <b><%=gsGroupName%></b></A>
		<br>&nbsp;
	</TD>
</TR>
<% End If %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(3);return false;'>Insertar un grupo al mismo nivel que <b><%=gsGroupName%></b></A>
		<br>&nbsp;
  </TD>
</TR>
<% If (gnRuleChildsType <= 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(4);return false;'>Insertar un grupo descendiente de <b><%=gsGroupName%></b></A>
		<br>&nbsp;
	</TD>
</TR>
<% End If %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(5);return false;'>Editar la información del grupo <b><%=gsGroupName%></b></A>
		<br>&nbsp;
	</TD>
</TR>
<% If (gnRuleChildsCount <> 0) AND (gnRuleChildsType > 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(6);return false;'>Eliminar el grupo <b><%=gsGroupName%></b> y los elementos que contiene</A>
		<br>&nbsp;
	</TD>
</TR>
<% End If %>
<% If (gnRuleChildsCount = 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(6);return false;'>Eliminar el grupo <b><%=gsGroupName%></b></A>
		<br>&nbsp;
  </TD>
</TR>
<% End If %>
<% If (gnRuleChildsCount <> 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(7);return false;'>Mover los elementos de <b><%=gsGroupName%></b> a otro grupo</A>
		<br>&nbsp;
	</TD>
</TR>
<% End If %>
<% If (gnRuleChildsCount <> 0) Then %>
<TR>
  <TD valign=top nowrap></TD>
  <TD>
		<A href='' onclick='callEditor(8);return false;'>Copiar los elementos de <b><%=gsGroupName%></b> a otro grupo</A>
		<br>&nbsp;
	</TD>
</TR>
<% End If %>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/financial_accounting/")</script>
</HTML>