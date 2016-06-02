<% 
	Option Explicit	
	Response.Expires = -1
	
	If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If	
%>
<% RSDispatch %>

<!--#INCLUDE FILE="rs.asp"-->

<SCRIPT RUNAT=SERVER Language=javascript>

	function IServerScripts()
	{
		this.CanChangePassword = Function('sUserName', 'sPassword', 'return CanChangePassword(sUserName, sPassword)');		
		this.WhoHasUserKey = Function('sUserKey', 'nUserId', 'return WhoHasUserKey(sUserKey, nUserId)');
		this.IsDate = Function('sDate', 'return IsDateOK(sDate)');
	}
	public_description = new IServerScripts();  

</SCRIPT>

<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">

Function CanChangePassword(sUserName, sPassword)
  Dim oIdentity, nTemp
  '*********************************************
	On Error Resume Next	
	Set oIdentity = Server.CreateObject("AOIdentity.CServices")
	nTemp = oIdentity.UserId(Session("sAppServer"), CStr(sUserName), CStr(sPassword), Session.SessionID)
	CanChangePassword = (nTemp = Session("uid"))	
	Set oIdentity = Nothing
End Function

Function WhoHasUserKey(sUserKey, nUserId)
  Dim oParticipants, nTemp
  '******************************************
	On Error Resume Next	
	Set oParticipants = Server.CreateObject("MHParticipantsMgr.CParticipants")
	WhoHasUserKey = oParticipants.WhoHasUserKey(Session("sAppServer"), CStr(sUserKey), CLng(nUserId))
	Set oParticipants = Nothing
End Function


Function IsDateOK(sDate)
	IsDateOK = IsDate(sDate)
End Function

</SCRIPT>