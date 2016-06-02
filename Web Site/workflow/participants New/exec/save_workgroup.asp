<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
  Call Main()
   
  Sub Main()
		Dim oParticipants, oParticipant, oExtendedAttrs, sParticipantName, nUserId
		'*************************************************************************
		'On Error Resume Next						
		nUserId = CLng(Request.Form("txtUserId"))
		If CLng(nUserId) = 0 Then
			Set oParticipants = Server.CreateObject("MHParticipantsMgr.CParticipants")
			Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")	
			nUserId = oParticipants.Append(Application("sAppServer"), CStr(Request.Form("txtName")), 3, _
																		 CStr(Request.Form("txtParticipantKey")), CStr(Request.Form("txtDescription")), _
																		 1, , Date)
			Set oParticipants = Nothing
		Else
			Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
			nUserId = oParticipant.Save(Application("sAppServer"), CLng(nUserId), CStr(Request.Form("txtName")), _
																  CStr(Request.Form("txtParticipantKey")), CStr(Request.Form("txtDescription")), , Date)			
			Set oParticipant = Nothing
		End If				
		If (Err.number <> 0) Then
			Session("errNumber") = Err.number
			Session("errDesc")   = Err.description
			Session("errSource") = Err.source
			Err.Clear
			Response.Redirect(Application("errorPage"))
		End If
  End Sub
%>
<HTML>
<HEAD>
<meta http-equiv="Pragma" content="no-cache">
</HEAD>
<BODY  onload='window.close();'>
</BODY>
</HTML>