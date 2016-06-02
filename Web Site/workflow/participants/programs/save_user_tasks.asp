<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnUserId, gnEntityId, gsParticipantName
	 
  Call Main()
   
  Sub Main()
		Dim oParticipant, gnWorkgroupTasks, sArray, i
		'******************
		'On Error Resume Next
		gnUserId = CLng(Request.Form("txtUserId"))		
		Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
		
		gnWorkgroupTasks = CLng(Request.Form("chkItems").Count)
		If (gnWorkgroupTasks > 0) Then
			sArray = Request.Form("chkItems")(1)
			For i = 2 To gnWorkgroupTasks
				sArray = sArray & "," & Request.Form("chkItems")(i)
			Next
			oParticipant.ParticipantTasks Session("sAppServer"), CLng(gnUserId), Split(sArray, ",")
		Else
			oParticipant.ParticipantTasks Session("sAppServer"), CLng(gnUserId), Null
		End If
		Set oParticipant = Nothing
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
<BODY onload='window.close();'>
</BODY>
</HTML>