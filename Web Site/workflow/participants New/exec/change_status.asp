<%
  Option Explicit     
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
		
	Dim sParticipantName
	 
  Call Main()
   
  Sub Main()
		Dim oParticipant
		'***************
		On Error Resume Next
		Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")		
		oParticipant.ChangeStatus Session("sAppServer"), CLng(Request.QueryString("id")), CLng(Request.QueryString("status"))
		Set oParticipant = Nothing
		If (Err.number <> 0) Then
			Set Session("oError") = Err
			Response.Redirect(Application("smallErrPage"))
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