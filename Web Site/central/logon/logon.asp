<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

	Dim gsUserName, gsMessage
					
	Call Check()
	
	Sub Check()		
		Dim oLogon
		'*********
		On Error Resume Next
		Set oLogon = Server.CreateObject("ECELogonServices.CLogon")
		gsUserName = Request.Form("txtUserName")		
		Session("uid") = oLogon.UserId(Session("sAppServer"), CStr(gsUserName), CStr(Request.Form("txtPassword")), Session.SessionID)
		Set oLogon = Nothing
		If Err.number = 0 Then
			If Session("uid") <> 0 Then
				If GetUserInfo() Then
					Response.Redirect "/empiria/portal/index.asp"
				Else
					Session("uid")       = 0
					Session("user_name") = ""					
				End If
			Else
				gsMessage = "No reconozco la contraseña o el identificador proporcionados."
			End If
		Else
			gsMessage = "Ocurrió el siguiente problema: \n\n" & Replace(Replace(Replace(Err.description,  Chr(34), Chr(39)), "\", "\\"), vbCrLf, "\n")
			gsMessage = gsMessage & "\n\n" & "Fuente: " & Err.source & " / Número: "  & Err.number
		End If
  End Sub
  
  Function GetUserInfo()
	  Dim oParticipant
		'*******************
		On Error Resume Next
		Set oParticipant = Server.CreateObject("ECEParticipantsMgr.CParticipant")
		Session("user_name") = oParticipant.Name(Session("sAppServer"), Session("uid"))
		Set oParticipant = Nothing
		If (Err.number = 0) Then
			GetUserInfo = True
		Else
			GetUserInfo = False
			gsMessage = "Ocurrió el siguiente problema: \n\n" & Replace(Replace(Replace(Err.description, Chr(34), Chr(39)), "\", "\\"), vbCrLf, "\n")
			gsMessage = gsMessage & "\n\n" & "Fuente: " & Err.source & " /  Número: "  & Err.number
		End If
	End Function
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Expires" CONTENT="-1">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
<% If (Len(gsMessage) <> 0) Then %>
	document.all.frmSend.submit();	
<% End If %>
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<form name=frmSend action="/empiria/default.asp" method="post">
<INPUT type=hidden name="txtUserName" value="<%=gsUserName%>">
<INPUT type=hidden name="txtMessage" value="<%=gsMessage%>">
</form>
</BODY>
<meta http-equiv="Pragma" content="no-cache">
</HTML>