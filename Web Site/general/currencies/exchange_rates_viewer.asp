<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsCboCurrencies, gsCboMonths, gsExchangeRatesTable, gnSelectedCurrency, gnSelectedMonth
	
	Call Main()
	
	Sub Main()
		Dim oCurrenciesUS
		'****************
		On Error Resume Next
		Set oCurrenciesUS = Server.CreateObject("AOCurrenciesUS.CServer")
		If (Len(Request.QueryString("id")) <> 0) Then			
			gnSelectedCurrency = CLng(Request.QueryString("id"))
			gnSelectedMonth = CLng(Request.QueryString("month"))
		Else
			gnSelectedCurrency = 0
			gnSelectedMonth = 0
		End IF
		gsCboCurrencies = oCurrenciesUS.CboCurrenciesWithException(Session("sAppServer"), 1, CLng(gnSelectedCurrency))
		gsCboMonths = oCurrenciesUS.CboExchangeRatesMonths(CLng(gnSelectedMonth))
		Select Case Request.QueryString("order")
			Case ""
				gsExchangeRatesTable = oCurrenciesUS.GetExchangeRatesHTMLTableRO(Session("sAppServer"), 1, CLng(gnSelectedCurrency), CLng(gnSelectedMonth), "")
			Case "1"
				gsExchangeRatesTable = oCurrenciesUS.GetExchangeRatesHTMLTableRO(Session("sAppServer"), 1, CLng(gnSelectedCurrency), CLng(gnSelectedMonth), "To_Date")
			Case "2"
				gsExchangeRatesTable = oCurrenciesUS.GetExchangeRatesHTMLTableRO(Session("sAppServer"), 1, CLng(gnSelectedCurrency), CLng(gnSelectedMonth), "To_Date, From_Date")
			Case "3"
				gsExchangeRatesTable = oCurrenciesUS.GetExchangeRatesHTMLTableRO(Session("sAppServer"), 1, CLng(gnSelectedCurrency), CLng(gnSelectedMonth), "Exchange_Rate_Type, From_Date")
			Case "4"
				gsExchangeRatesTable = oCurrenciesUS.GetExchangeRatesHTMLTableRO(Session("sAppServer"), 1, CLng(gnSelectedCurrency), CLng(gnSelectedMonth), "Exchange_Rate, From_Date")
		End Select		
		Set oCurrenciesUS = Nothing
		If (Err.number <> 0) Then
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description			
			Session("sErrPage") = Request.ServerVariables("URL")		  
		  Response.Redirect("./exec/exception.asp")
		End If
	End Sub
%>
<HTML>
<HEAD>
<TITLE>Banobras - Intranet corporativa</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function refreshPage(nOrderId) {
	var sURL = "";	
	sURL = "exchange_rates_viewer.asp?id=" + window.document.all.cboCurrencies.value;
	sURL += "&month=" + window.document.all.cboMonths.value;
  if (nOrderId == 0) {
		window.location.href = sURL;
	} else {	
	  window.location.href = sURL + '&order=' + nOrderId;
	}
	return false;
}

function cboMonths_onchange() {
	var sURL = "";	
	sURL = "exchange_rates_viewer.asp?id=" + window.document.all.cboCurrencies.value;
	sURL += "&month=" + window.document.all.cboMonths.value;
	window.location.href = sURL;
}

function cboCurrencies_onchange() {
	var sURL = "";	
	sURL = "exchange_rates_viewer.asp?id=" + window.document.all.cboCurrencies.value;
	sURL += "&month=" + window.document.all.cboMonths.value;
	window.location.href = sURL;
}

//-->
</SCRIPT>
</HEAD>
<BODY SCROLL=NO>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=83px">
<BR>
<TABLE align=center border=0 bgcolor=Khaki cellPadding=3 cellSpacing=3 width="70%">
	<TR>
		<TD nowrap><FONT face=Arial size=3 color=maroon><STRONG>Tipos de cambio</STRONG></FONT></TD>
	  <TD colspan= 2 nowrap align=right>Ver los tipos de cambio a:&nbsp;
			<SELECT name=cboCurrencies onchange="return cboCurrencies_onchange()">
				<%=gsCboCurrencies%>
			</SELECT>
		</TD>
		<TD colspan nowrap align=right>en el mes de:&nbsp;
			<SELECT name=cboMonths onchange="return cboMonths_onchange()">
				<%=gsCboMonths%>
			</SELECT>					
		</TD>
	</TR>
	<TR>
		<TD nowrap>&nbsp;</TD>
	  <TD colspan=3 align=right nowrap>			
			<INPUT type="button" name=cmdRefresh value="Actualizar vista" style="WIDTH:120px" LANGUAGE=javascript onclick="window.location.href=window.location.href;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<INPUT type="button" name=cmdReturn value="Cerrar" style="WIDTH:80px" LANGUAGE=javascript onclick="window.location.href = '<%=Session("main_page")%>';">
		</TD>
	</TR>	
</TABLE>
</DIV>
<DIV STYLE="overflow:auto; float:bottom; width=100%; height=85%">
<TABLE align=center border=1 cellPadding=3 cellSpacing=3 width="70%">
<% If Len(gsExchangeRatesTable) <> 0 Then %>
	<A href="#SCROLLABLE_DIV_TOP"></A>
	<TR>
	  <TD nowrap><A href="" onclick="return refreshPage(1);"><FONT color=maroon><b>Fecha inicial</b></FONT></A></TD>
	  <TD nowrap><A href="" onclick="return refreshPage(2);"><FONT color=maroon><b>Fecha final</b></FONT></A></TD>
	  <TD nowrap align="center"><A href="" onclick="return refreshPage(3);"><FONT color=maroon><b>Tipo</b></FONT></A></TD>
	  <TD nowrap align=right><A href="" onclick="return refreshPage(4);"><FONT color=maroon><b>Tipo de cambio</b></FONT></A></TD>
	</TR>
	<%=gsExchangeRatesTable%>
	<TR>
	  <TD nowrap colspan=4 align=right><A href="#SCROLLABLE_DIV_TOP">Subir</A></TD>
	</TR>	
<% Else %>
	<TR><TD colspan=4 align=center>No hay tipos de cambio definidos para esta moneda en el mes seleccionado.</TD></TR>
<% End If %>
</TABLE>
<BR>&nbsp;
</DIV>
</BODY>
</HTML>