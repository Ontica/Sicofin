<%
  Option Explicit
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim oReports, gsIndexFileName, gsVerticalFileName, gsHorizontalFileName, nScriptTimeout
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
					
	Sub Main()
		Dim oVoucherUS, vGralLedgers, bRounded, sTitle, sTemp, dExcRateDate, bPrintInCascade, bTotal
		
		'On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
		Set oReports = Server.CreateObject("SCFNotebook.CNotebook")
				
		If (Len(Request.Form("cboGralLedgers")) <> 0 ) Then		
			If (Len(Request.Form("txtFromGL")) = 0) Then
				If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
					sTemp = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), ",")
					vGralLedgers = Split(sTemp, ",")
				Else
					vGralLedgers = CLng(Request.Form("cboGralLedgers"))
				End If
			Else
				vGralLedgers = oVoucherUS.GetGLRangeArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), CLng(Request.Form("txtFromGL")), CLng(Request.Form("txtToGL")))
			End If
		End If
		If Len(Request.Form("chkRounded")) <> 0 Then			
			bRounded = CLng(Request.Form("chkRounded")) = 1
		End If
  	If Len(Request.Form("chkTotal")) <> 0 Then
			bTotal = False
		Else
			bTotal = True
		End If
		If Len(Request.Form("txtExchangeRateDate")) <> 0 Then
			dExcRateDate = Request.Form("txtExchangeRateDate")
		Else
			dExcRateDate = Date()
		End If
		
  	If Len(Request.Form("chkPrintInCascade")) <> 0 Then
			bPrintInCascade = True
		Else
			bPrintInCascade = False
		End If

		Call CreateNotebook(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
									      dExcRateDate, Request.Form("txtSigner1Name"), Request.Form("txtSigner1Title"), _
									      Request.Form("txtSigner2Name"), Request.Form("txtSigner2Title"), _
									      Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"), _
									      gsIndexFileName, gsVerticalFileName, gsHorizontalFileName)
					
		gsIndexFileName      = oReports.URLFilesPath & gsIndexFileName
		gsVerticalFileName   = oReports.URLFilesPath & gsVerticalFileName
		gsHorizontalFileName = oReports.URLFilesPath & gsHorizontalFileName
		Set oReports = Nothing
	End Sub  
                             
	Sub CreateNotebook(nId, dInitialDate, dFinalDate, dExcRateDate, sSigner1Name, sSigner1Title, _
										 sSigner2Name, sSigner2Title, nExchangeRateType, nExcRateCurrency, _
										 sIndexFileName, sVerticalFileName, sHorizontalFileName)
										 
		oReports.Generate Session("sAppServer"), CLng(nId), _
											CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
											CStr(sSigner1Name), CStr(sSigner1Title), _
											CStr(sSigner2Name), CStr(sSigner2Title), _
											CLng(nExchangeRateType), CLng(nExcRateCurrency), _
											sIndexFileName, sVerticalFileName, sHorizontalFileName
	End Sub
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</head>
<body>
<table>
<tr>
	<td><font size=2><b>La información solicitada está lista.</b></font></td>	
</tr>
<tr>
	<td><font size=2><b>¿Qué desea hacer?</b></font></td>	
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td>
		<a href="<%=gsIndexFileName%>" target="_blank" onclick="showReportInBrowser();return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="<%=gsIndexFileName%>" target="_blank">	
			Indice del cuaderno
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="<%=gsVerticalFileName%>" target="_blank" onclick="showReportInBrowser();return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="<%=gsVerticalFileName%>" target="_blank"> 	
			Reportes para impresión en modo vertical
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="<%=gsHorizontalFileName%>" target="_blank" onclick="showReportInBrowser();return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="<%=gsHorizontalFileName%>" target="_blank">	
			Reportes para impresión en modo horizontal
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<a href="" onclick="window.history.back();">
			Cerrar esta ventana y perder la información obtenida.
		</a>
		<br>
	</td>	
</tr>
</table>
</body>
</html>