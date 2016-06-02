<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim nScriptTimeout
	Dim gsErrNumber, gsErrSource, gsErrDescription
 
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600	
  Call Main()
  Server.ScriptTimeout = nScriptTimeout
   
  Sub Main()
		Dim aWorkItemsIds, oMHInboxes, nWorkItemId, gnDateStatus, i
		'************************************************************
		Set oMHInboxes = Server.CreateObject("MHInboxExplorer.CExplorer")				
		aWorkItemsIds = Split(Request.Form("txtPendingTasks"), ",")
    For i = LBound(aWorkItemsIds) To UBound(aWorkItemsIds)
			nWorkItemId = aWorkItemsIds(i)
			oMHInboxes.AssignWorkItem Session("sAppServer"), CLng(nWorkItemId), Session("uid")			
		Next
		Set oMHInboxes = Nothing		
		If (Err.number = 0) Then
			Response.Redirect("../inbox.asp")
		Else
			gsErrNumber = Err.number
			gsErrSource = Err.source
			gsErrDescription = Err.description
		End If
  End Sub
%>