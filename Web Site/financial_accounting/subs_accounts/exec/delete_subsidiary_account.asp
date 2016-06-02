<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Call DeleteItem(CLng(Request.QueryString("id")))    
	
  Sub DeleteItem(nItemId)
		Dim oGralLedger
		'**************************
		On Error Resume Next
		Set oGralLedger = Server.CreateObject("AOGralLedger.CSubsidiaryLedger")
		oGralLedger.DeleteSubsidiaryAccount Session("sAppServer"), CLng(nItemId)
		Set oGralLedger = Nothing
		If (Err.number = 0) Then
			Set Session("oError") = Nothing
		Else
			Set Session("oError") = Err
			Response.Redirect Application("error_page")
		End If
  End Sub 
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
</head>
<body class=bdyDialogBox>
<TABLE class=standardPageTable>
	<TR>
		<TD colspan=2>
			<TABLE class=fullScrollMenu>
				<TR class=fullScrollMenuHeader>
					<TD class=fullScrollMenuTitle>
						La eliminación se efectuó correctamente
					</TD>
					<TD colspan=3 align=right nowrap>
						<img align=absmiddle src='/empiria/images/close_white.gif' onclick="window.close();" alt="Cerrar">											</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD nowrap> &nbsp; &nbsp; </TD>
		<TD width=100%>
			<br>			
			<a href='' onclick='window.close();return false;'>Cerrar esta ventana</a>
		</TD>
	</TR>
</TABLE>
</body>
</html>