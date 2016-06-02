<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim gnWorkgroupId, gnEntityId, gsParticipantName
	 
  Call Main()
   
  Sub Main()
		Dim oParticipant, gnParticipantRelations, sArray, i
		'******************
		'On Error Resume Next
		gnWorkgroupId = CLng(Request.Form("txtWorkgroupId"))		
		Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
		
		gnParticipantRelations = CLng(Request.Form("chkItems").Count)
		If (gnParticipantRelations > 0) Then
			sArray = Request.Form("chkItems")(1)
			For i = 2 To gnParticipantRelations
				sArray = sArray & "," & Request.Form("chkItems")(i)
			Next
			oParticipant.RelatedParticipants Session("sAppServer"), CLng(gnWorkgroupId), 1301, Split(sArray, ",")
		Else
			oParticipant.RelatedParticipants Session("sAppServer"), CLng(gnWorkgroupId), 1301, Null
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