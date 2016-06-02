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
		Dim oExplorer, oRecordset, nItemId, sURL
		'********************************************************
		'On Error Resume Next
		Set oExplorer = Server.CreateObject("MHInboxExplorer.CExplorer")
		
		Set oRecordset = oExplorer.GetWorkItem(Session("sAppServer"), CLng(Request.QueryString("id")))
		nItemId = oRecordset("itemId")
		oRecordset.Close
		Set oRecordset = Nothing
		Set oExplorer = Nothing
		
		sURL = "/empiria/financial_accounting/transactions/pending_voucher_viewer.asp?id=" & nItemId
		Response.Redirect(sURL)
	End Sub
%>
