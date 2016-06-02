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
		sParticipantName = Request.Form("txtName") & " " & Request.Form("txtLastName1")
		nUserId = CLng(Request.Form("txtUserId"))
		If CLng(nUserId) = 0 Then
			Set oParticipants = Server.CreateObject("MHParticipantsMgr.CParticipants")
			Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
			Set oExtendedAttrs = oParticipant.ExtendedAttributes(Session("sAppServer"), 0)
			oExtendedAttrs("LastName1") = Request.Form("txtLastName1")
			oExtendedAttrs("LastName2") = Request.Form("txtLastName2")
			oExtendedAttrs("FirstName") = Request.Form("txtName")
			oExtendedAttrs("IsFemale")  = CBool(Request.Form("chkIsFemale"))
			oExtendedAttrs("JobEMail")  = Request.Form("txtJobEMail")
			oExtendedAttrs("JobPhone")  = Request.Form("txtJobPhone")
			nUserId = oParticipants.Append(Session("sAppServer"), CStr(sParticipantName), 1, _
																		 CStr(Request.Form("txtParticipantKey")), "", 1, _
																		 (oExtendedAttrs), Request.Form("txtHistoricDate"))

			Set oParticipants = Nothing
		Else
			Set oParticipant = Server.CreateObject("MHParticipantsMgr.CParticipant")
			Set oExtendedAttrs = oParticipant.ExtendedAttributes(Session("sAppServer"), CLng(nUserId))
			oExtendedAttrs("LastName1") = Request.Form("txtLastName1")
			oExtendedAttrs("LastName2") = Request.Form("txtLastName2")
			oExtendedAttrs("FirstName") = Request.Form("txtName")
			oExtendedAttrs("IsFemale")  = CBool(Request.Form("chkIsFemale"))
			oExtendedAttrs("JobEmail")  = Request.Form("txtJobEMail")
			oExtendedAttrs("JobPhone")  = Request.Form("txtJobPhone")
			nUserId = oParticipant.Save(Session("sAppServer"), CLng(nUserId), CStr(sParticipantName), _
																  CStr(Request.Form("txtParticipantKey")), "", (oExtendedAttrs), Date)			
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
<BODY onload='window.close();'>
</BODY>
</HTML>