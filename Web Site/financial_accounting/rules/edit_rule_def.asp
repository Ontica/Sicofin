<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnRuleId, gsRuleName, gsDescription, gsCboGLCategories, gsCboMethods, gsVersion, gsActionPage
	Dim gsFromDate, gsToDate, gbUsesVersions
	
	Call Main()
			 
	Sub Main()
		Dim oRuleDef, oRecordset
		'*************************
		On Error Resume Next
		gnRuleId			 = CLng(Request.QueryString("id"))
		Set oRuleDef   = Server.CreateObject("EFARulesMgrBS.CRuleDefinition")
		Set oRecordset = oRuleDef.RuleDefRS(Session("sAppServer"), CLng(gnRuleId))
		gsRuleName		 = oRecordset("nombre_regla_contable")
		gsDescription  = oRecordset("descripcion")
		If CLng(oRecordset("version_regla")) <> 0 Then
			gsVersion = oRecordset("version_regla") & ".0"
			gbUsesVersions = True
		Else
			gsVersion	= "Única (depende de la fecha en que se editan las cuentas de la regla)"
			gbUsesVersions = False
		End If
		gsFromDate	      = oRecordset("fecha_inicio")				
		gsCboGLCategories = oRuleDef.CboGLCategories(Session("sAppServer"), CLng(oRecordset("id_tipo_cuentas_std")))
		gsCboMethods      = oRuleDef.CboMethods(Session("sAppServer"), CLng(oRecordset("id_metodo")))
		If oRecordset("fecha_fin") = CDate("31/Dic/9999") Then
			gsToDate = ""
		Else
			gsToDate = oRecordset("fecha_inicio")
		End If
		Set oRecordset = Nothing
		gsActionPage = "./exec/save.asp?id=" & gnRuleId
			
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("/empiria/central/exceptions/exception.asp")
		End If		  
	End Sub
			
%>
<HTML>
<HEAD>
<META http-equiv="Pragma" content="no-cache">
<TITLE>Banobras - Intranet corporativa</TITLE>
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<script language="JavaScript" src="/empiria/bin/ms_scripts/rs.htm"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function validateDate(date) {
	var obj;
	obj = RSExecute("../financial_accounting_scripts.asp","IsDate", date);
	return obj.return_value;
}

function cmdCheckSpelling_onclick() {
	alert("Por el momento esta opción no está disponible.\n\nGracias.");
}

function createNewVersion() {
	alert("Por el momento esta opción no está disponible");
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<TABLE align=center border=0 bgcolor=khaki cellPadding=3 cellSpacing=3 width="90%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Editor de reglas</STRONG></FONT></TD>
	  <TD colspan=3 align=right nowrap>	
			<A href="rules_def.asp">Base de conocimiento contable</A>&nbsp;&nbsp;&nbsp;&nbsp;
	    <A href="rules.asp?id=<%=gnRuleId%>">Editar agrupación</A>&nbsp;&nbsp;&nbsp;&nbsp;
	    <% If gbUsesVersions Then %>
	    <A href="" onclick="createNewVersion();return false;">Crear nueva versión</A>&nbsp;&nbsp;&nbsp;&nbsp;
	    <% End If %>
			<A href="<%=Application("main_page")%>">Cerrar</A>
		</TD>
	</TR>
</TABLE>
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="90%">
<FORM name=frmSend action="<%=gsActionPage%>" method=post onsubmit='return frmSend_onsubmit()'>
  <TR>
		<TD valign=top nowrap>Nombre de la regla:</TD>
    <TD colspan=3>
			<TEXTAREA name=txtDescription ROWS=2 style="WIDTH: 520px"><%=gsRuleName%></TEXTAREA>
    </TD>
  </TR>
  <TR>
    <TD valign=top>Descripción:</TD>
    <TD colspan=3><TEXTAREA name=txtDescription ROWS=4 style="WIDTH: 520px"><%=gsDescription%></TEXTAREA><br>			
    </TD>
  </TR>
  <TR>
    <TD nowrap>Contabilidad a la que se aplica:</TD>  
		<TD>
			<SELECT name=cboGLCategory style="WIDTH: 520px"> 
				<%=gsCboGLCategories%>
			</SELECT>
		</TD>
  </TR>
  <TR>
    <TD nowrap>Programa que la ejecuta:</TD> 
    <TD>
			<SELECT name=cboMethods style="WIDTH: 520px"> 
				<%=gsCboMethods%>
			</SELECT>
		</TD>
  </TR>  
  <TR>
    <TD valign=top>Versión de la regla:</TD>
    <TD colspan=3>
			<b><%=gsVersion%></b>
		</TD>		
  </TR>  
  <TR>
    <TD valign=top nowrap>Fecha de aplicación:</TD>
    <TD colspan=3 nowrap>			
			Desde: <INPUT type=text name=txtFromDate style="WIDTH: 95px" value="<%=gsFromDate%>">			
			&nbsp;&nbsp;
			Hasta: <INPUT type=text name=txtToDate style="WIDTH: 95px" value="<%=gsToDate%>">
			&nbsp;(dejar vacío si no se sabe hasta cuándo se aplicará)
		</TD>
  </TR>
  <TR>
		<td>&nbsp;</td>
    <td colspan=3 nowrap>     
     <INPUT language=javascript name=cmdSend style="HEIGHT: 25px; WIDTH: 120px" type=submit value="Guardar cambios">
     &nbsp;&nbsp;&nbsp;&nbsp;     
     <INPUT language=javascript name=cmdCancel style="HEIGHT: 25px; WIDTH: 120px" type=button value="Cancelar" onclick="window.location.href = 'rules_def.asp';">
    </td>
  </TR>
</FORM>
</TABLE>
</BODY>
<script language="JavaScript">RSEnableRemoteScripting("/empiria/bin/ms_scripts/")</script>
</HTML>
