<%
  Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim gsReturnPage, gsCancelPage
 
	gsReturnPage = "../pages/not_labor_days.asp?id=" & Request.QueryString("Id")
	gsCancelPage = Session("main_page")
	
	Call SaveNotLaborWeekdays(CLng(Request.QueryString("Id")))
		    
  Sub SaveNotLaborWeekdays(nCalendarId)
		Dim aNotLaborWeekDays()
		Dim oCalendar, nItemsCount, i
		'*************************************************
		On Error Resume Next
		Set oCalendar = Server.CreateObject("AOCalendar.CManager")
		nItemsCount = CLng(Request.Form("chkNotLaborDay").Count)
		If nItemsCount <> 0 Then
			ReDim aNotLaborWeekDays(nItemsCount - 1)
			For i = 0 To (nItemsCount - 1)
				aNotLaborWeekDays(i) = Request.Form("chkNotLaborDay").Item(i + 1)
			Next
			oCalendar.SaveNotLaborWeekdays Session("sAppServer"), CLng(nCalendarId), aNotLaborWeekDays
		Else
			oCalendar.SaveNotLaborWeekdays Session("sAppServer"), CLng(nCalendarId)
		End If		
		Set oCalendar = Nothing
		If (Err.number = 0) Then
			If Len(gsReturnPage) <> 0 Then
				Response.Redirect(gsReturnPage)
			End If
		Else		
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("../programs/exception.asp")					
		End If
  End Sub
%>
