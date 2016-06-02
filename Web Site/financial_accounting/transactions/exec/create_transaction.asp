<%
	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	
  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	Dim nScriptTimeout
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
	
	Sub Main()	
		Dim oVoucherUS, vGralLedgers, nTransactionId, dApplicationDate, dElaborationDate, i
		'***********************************************************************************
		'On Error Resume Next
		If Len(Request.Form("txtOutOfPeriodDate")) <> 0 Then			
			dApplicationDate = Request.Form("txtOutOfPeriodDate")
		ElseIf Len(Request.Form("txtAlternativeDate")) <> 0 Then
			dApplicationDate = Request.Form("txtAlternativeDate")
		Else
			dApplicationDate = Request.Form("cboApplicationDates")			
		End If
		dElaborationDate = Now()		
		If (Len(Request.Form("cboGralLedgers")) <> 0) Then
			Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")
			If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
				vGralLedgers = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGLCategories")), ",")
				vGralLedgers = Split(vGralLedgers, ",")
			Else
				vGralLedgers = CLng(Request.Form("cboGralLedgers"))
			End If							                       								
			If IsArray(vGralLedgers) Then
				For i = LBound(vGralLedgers) To UBound(vGralLedgers)
					Execute CLng(vGralLedgers(i)), Request.Form("txtVoucherType"), dApplicationDate, dElaborationDate
				Next					
			Else
				nTransactionId = Execute(CLng(vGralLedgers), Request.Form("txtVoucherType"), dApplicationDate, dElaborationDate)
			End If
			Set oVoucherUS = Nothing
		Else
			nTransactionId = Execute(0, Request.Form("txtVoucherType"), dApplicationDate, dElaborationDate)
		End If
	  If (Err.number = 0) Then	
			If nTransactionId > 0 Then
				Response.Redirect("../voucher_editor.asp?id=" & nTransactionId)
			ElseIf nTransactionId = 0 Then
				Response.Redirect("../pending_vouchers.asp")
			End If
		Else
			Session("nErrNumber") = "&H" & Hex(Err.number)
			Session("sErrSource") = Err.source
			Session("sErrDescription") = Err.description
			Session("sErrPage") = Request.ServerVariables("URL")
		  Response.Redirect("../programs/exception.asp")
		End If
	End Sub	

	Function Execute(nGralLedgerId, nVoucherType, dApplicationDate, dElaborationDate)		
		Dim oObject
		'******************************************************************************		
		Select Case nVoucherType
			Case 28
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")				                      				
 				Execute = oObject.Execute(Session("sAppServer"), _
 																	CLng(nVoucherType), CLng(Session("uid")), _
 																  CLng(Request.Form("cboSources")), _
 																  CStr(Request.Form("txtDescription")), _
																	CDate(dApplicationDate), CDate(dElaborationDate))				
			Case 29				
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")
				Execute = oObject.Execute(Session("sAppServer"), _
																	CLng(nVoucherType), CLng(Session("uid")), _
																	CLng(Request.Form("cboSources")), _
																	CStr(Request.Form("txtDescription")), _
																	CDate(dApplicationDate), CDate(dElaborationDate), _
																	CLng(nGralLedgerId))
			Case 30
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")
				Execute = oObject.Execute(Session("sAppServer"), _
																	CLng(nVoucherType), CLng(Session("uid")), _
																	CLng(Request.Form("cboSources")), _
																	CStr(Request.Form("txtDescription")), _
																	CDate(dApplicationDate), CDate(dElaborationDate), _
																	CLng(nGralLedgerId))				
			Case 64
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")
 				Execute = oObject.CancelTransaction(Session("sAppServer"), CLng(Session("uid")), 0, _
 																						CLng(nGralLedgerId), CStr(Request.Form("txtVoucherNumber")), _
 																						CStr(Request.Form("txtDescription")), _
 																						CLng(Request.Form("cboSources")), CLng(nVoucherType), _
 																						CDate(dApplicationDate), CDate(dElaborationDate))                                   																						
			Case 120
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")
				Execute = oObject.TransferBalances(Session("sAppServer"), CLng(Session("uid")), 0, _
																					 CLng(nGralLedgerId), CStr(Request.Form("txtDescription")), _ 
																					 CLng(Request.Form("cboSources")), CLng(nVoucherType), _
																					 CDate(dApplicationDate), CDate(dElaborationDate), _
																					 CStr(Request.Form("txtFromAccount")), _
																					 CStr(Request.Form("txtFromSubsAccount")), _
																					 CLng(Request.Form("cboFromSector")), _
																					 CLng(Request.Form("cboFromCurrency")), _
																					 "", "", 0, _
																					 CLng(Request.Form("cboToCurrency")), Request.Form("txtExchangeRate"))
			Case 122
				Set oObject = Server.CreateObject("EFARulesEngine.CTransactions")
				Execute = oObject.Execute(Session("sAppServer"), _
 																  CLng(nVoucherType), CLng(Session("uid")), _
 																  CLng(Request.Form("cboSources")), _ 
 																  CStr(Request.Form("txtDescription")), _ 
 																  CDate(dApplicationDate), CDate(dElaborationDate), _
 																  CLng(nGralLedgerId), 0, _
 																	CStr(Request.Form("txtVoucherNumber")))				
			Case Else				
				Set oObject = Server.CreateObject("AOGLVoucher.CServer")
				Execute = oObject.CreateTransaction(Session("sAppServer"), _
																 CLng(Request.Form("txtTransactionType")), _
																 CLng(Request.Form("txtVoucherType")), _				
																 CLng(nGralLedgerId), _														
																 CDate(dApplicationDate), CDate(dElaborationDate), _
																 CStr(Request.Form("txtDescription")), _
																 CLng(Request.Form("cboSources")), Session("uid"))
		End Select
		Set oObject = Nothing
	End Function
%>