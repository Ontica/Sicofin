<%
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	ElseIf Len(Request.QueryString("id")) = 0 Then 
		Response.Redirect Application("main_page")
	End If
			
	Call Main()
	
	Sub Main()
		Dim oWorkflowEngine, nTaskId, sAppPath
		'*************************************
		nTaskId = Request.QueryString("id")
		Set oWorkflowEngine = Server.CreateObject("EWMEngine.CEngine")
		sAppPath = oWorkflowEngine.TaskApplicationPath(Session("sAppServer"), CLng(nTaskId), Session("uid"))
		Set oWorkflowEngine = Nothing
		Response.Redirect sAppPath
	End Sub	
%>
