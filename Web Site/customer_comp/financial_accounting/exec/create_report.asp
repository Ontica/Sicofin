<%
  Option Explicit
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1

  If (CLng(Session("uid")) = 0) Then
		Response.Redirect Application("exit_page")
	End If
	
	Dim oReports, sFileName, nScriptTimeout
	
	nScriptTimeout  = Server.ScriptTimeout
	Server.ScriptTimeout = 3600
	Call Main()
	Server.ScriptTimeout = nScriptTimeout
					
	Sub Main()
		Dim oVoucherUS, vGralLedgers, bRounded, sTitle, sTemp, dExcRateDate, bPrintInCascade, bTotal
		
		'On Error Resume Next
		Set oVoucherUS = Server.CreateObject("AOGLVoucherUS.CVoucher")		
		Set oReports = Server.CreateObject("SCFFixedReports.CReports")
				
		If (Len(Request.Form("cboGralLedgers")) <> 0 ) Then		
			If (Len(Request.Form("txtFromGL")) = 0) Then
				If CLng(Request.Form("cboGralLedgers")) = 0 Then		'Es la consolidada
					sTemp = oVoucherUS.GetGLGroupArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), ",")
					vGralLedgers = Split(sTemp, ",")
				Else
					vGralLedgers = CLng(Request.Form("cboGralLedgers"))
				End If
			Else
				vGralLedgers = oVoucherUS.GetGLRangeArray(Session("sAppServer"), CLng(Request.Form("cboGLGroups")), CLng(Request.Form("txtFromGL")), CLng(Request.Form("txtToGL")))
			End If
		End If
		If Len(Request.Form("chkRounded")) <> 0 Then			
			bRounded = CLng(Request.Form("chkRounded")) = 1
		End If
  	If Len(Request.Form("chkTotal")) <> 0 Then
			bTotal = False
		Else
			bTotal = True
		End If
		If Len(Request.Form("txtExchangeRateDate")) <> 0 Then
			dExcRateDate = Request.Form("txtExchangeRateDate")
		Else
			dExcRateDate = Date()
		End If
		
  	If Len(Request.Form("chkPrintInCascade")) <> 0 Then
			bPrintInCascade = True
		Else
			bPrintInCascade = False
		End If

		Select Case CLng(Request.QueryString("id"))
			Case 1
				sFileName = Report_01(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														 dExcRateDate, Request.Form("txtSigner1Name"), Request.Form("txtSigner1Title"), _
														 Request.Form("txtSigner2Name"), Request.Form("txtSigner2Title"), _
														 Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 3				
				sFileName = Report_03(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														  Request.Form("cboVoucherDatesMode"), Request.Form("cboVoucherStatus"))
			Case 9
				sFileName = Report_9(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))
			
			Case 18
				vGralLedgers = Array(1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32)
				sFileName = Report_18(9, vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															dExcRateDate, Request.Form("cboExchangeRateTypes"), _
															Request.Form("cboExchangeRateCurrencies"))
			Case 51
				sFileName = Report_51(vGralLedgers, Request.Form("cboPatterns"), _
															Request.Form("txtFinalDate"), dExcRateDate, bRounded, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 53			    
				sFileName = Report_53(vGralLedgers, Request.Form("cboPatterns"), _
															Request.Form("txtFinalDate"), dExcRateDate, bRounded, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 55
				sFileName = Report_55(vGralLedgers, Request.Form("cboPatterns"), _
															Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															dExcRateDate, bRounded, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 58
				sFileName = Report_58(vGralLedgers, Request.Form("cboPatterns"), _
															Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), dExcRateDate, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 59			    
				sFileName = Report_59(vGralLedgers, Request.Form("txtFinalDate"), dExcRateDate, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 60			    
				sFileName = Report_60(Request.Form("txtFinalDate"))
			
			Case 61
			  sFileName = Report_61(vGralLedgers, Request.Form("txtFinalDate"), dExcRateDate, _
									            Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 62
				sFileName = Report_62(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															Request.Form("txtAccountList"))
			Case 64
				sFileName = Report_64(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), bPrintInCascade)

			Case 65			    
				sFileName = Report_65(vGralLedgers, Request.Form("cboPatterns"), _
															Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), dExcRateDate, _
															Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"), _
															bPrintInCascade, bTotal)
     	Case 68
				sFileName = Report_68(vGralLedgers, Request.Form("txtFinalDate"))
				
     	Case 69
				sFileName = Report_69(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))		
	
			Case 70				
				sFileName = Report_70(Request.Form("txtFinalDate"))
				
			Case 73
			  sFileName = Report_73(Request.Form("cboGLGroups"), vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))
				
			Case 74
			  sFileName = Report_74(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															Request.Form("txtInitialDate2"), Request.Form("txtFinalDate2"), _
															dExcRateDate, Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
     	Case 96			    
				sFileName = Report_96(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
														  dExcRateDate, _
														  Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
      Case 99
				sFileName = Report_99(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															dExcRateDate, Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
      Case 101
				vGralLedgers = Array(1,2,3,4,5,6,7,8,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32)
				sFileName = Report_101(9, vGralLedgers, Request.Form("cboPatterns"), Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															dExcRateDate, Request.Form("cboExchangeRateTypes"), _
															Request.Form("cboExchangeRateCurrencies"), _
															Request.Form("txtAccountList"), Request.Form("cboTittles"))
			Case 102
				sFileName = Report_102(vGralLedgers, Request.Form("cboPatterns"), Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															dExcRateDate, Request.Form("cboExchangeRateTypes"), _
															Request.Form("cboExchangeRateCurrencies"), _
															Request.Form("txtAccountList"))
			Case 106
				sFileName = Report_106(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))

			Case 107
				sFileName = Report_107(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))

			Case 108
				sFileName = Report_108(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))

			Case 109
				sFileName = Report_109(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))

			Case 110
				sFileName = Report_110(vGralLedgers, Request.Form("cboPatterns"), _
															 Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															 Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
															 Request.Form("txtFromSubsAccount"), Request.Form("txtToSubsAccount"), _
															 bPrintInCascade)
			Case 111
				sFileName = Report_111(vGralLedgers, Request.Form("cboPatterns"), _
															 Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), bPrintInCascade, _
															 Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
															 Request.Form("txtFromSubsAccount"), Request.Form("txtToSubsAccount"))
			Case 113
				sFileName = Report_113(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
				                       Request.Form("txtFromAccount"), Request.Form("txtToAccount"), _
				                       Request.Form("txtFromSubsAccount"), Request.Form("txtToSubsAccount"))

			Case 115
			  sFileName = Report_115(Request.Form("txtFinalDate"), Request.Form("cboParticipants"), _
			                         Request.Form("cboParticipantOrder"), Request.Form("cboParticipantStatus"))
			Case 116
			  sFileName = Report_116(Request.Form("cboStdAccountType"), Request.Form("txtFinalDate"), _
			                         Request.Form("txtFromAccount"), Request.Form("txtToAccount"))
     	Case 117
			  sFileName = Report_117(Request.Form("cboAccountOrder"))
			Case 120				
				sFileName = Report_120(Request.Form("txtFinalDate"), Request.Form("cboConFidOptions"))

     	Case 126
			  sFileName = Report_126(vGralLedgers, Request.Form("txtFinalDate"))

     	Case 127
			  sFileName = Report_127(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
			                         Request.Form("cboOptionToDisplay"))

     	Case 128
				sFileName = Report_128(vGralLedgers, Request.Form("txtFinalDate"), dExcRateDate, _
                               Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 144
				sFileName = Report_144(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))
                               
			Case 149
				'vGralLedgers = Array(269,270,267,268,264,265,266)
				sFileName = Report_149(vGralLedgers, Request.Form("txtFromAccount"), _
				                       Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															 Request.Form("txtInitialDate2"), Request.Form("txtFinalDate2"), bPrintInCascade)
			Case 150
				'vGralLedgers = Array(269,270,267,268,264,265,266)
				sFileName = Report_150(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															Request.Form("txtInitialDate2"), Request.Form("txtFinalDate2"))
			Case 151
				'vGralLedgers = Array(269,270,267,268,264,265,266)
				sFileName = Report_151(vGralLedgers, Request.Form("txtFromAccount"), _
				                       Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
				                       dExcRateDate, _
															 Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))
			Case 152
				sFileName = Report_152(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
				                       dExcRateDate, _
															 Request.Form("cboExchangeRateTypes"), Request.Form("cboExchangeRateCurrencies"))

			Case 153
				'vGralLedgers = Array(269,270,267,268,264,265,266)
				sFileName = Report_153(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"), _
															 Request.Form("txtInitialDate2"), Request.Form("txtFinalDate2"))
			Case 156
				sFileName = Report_156(vGralLedgers, Request.Form("txtInitialDate"), Request.Form("txtFinalDate"))

		End Select
		sFileName = oReports.URLFilesPath & sFileName
		'If Err.number <> 0 Then
		'   Response.Write Err.Number & " " & Err.description & " " & Err.source
		'End If		
		Set oReports = Nothing
	End Sub  
                             
	Function Report_01(nId, dInitialDate, dFinalDate, dExcRateDate, _
										 sSigner1Name, sSigner1Title, sSigner2Name, sSigner2Title, _
										 nExchangeRateType, nExcRateCurrency)
		Report_01 = oReports.Notebook(Session("sAppServer"), CLng(nId), _
																  CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
																  CStr(sSigner1Name), CStr(sSigner1Title), _
																  CStr(sSigner2Name), CStr(sSigner2Title), _
																  CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function
	
	Function Report_03(aGL, dFromDate, dToDate, bByAfectationDate, bUpdatedVouchers)	
		Report_03 = oReports.Vouchers(Session("sAppServer"), aGL, CDate(dFromDate), CDate(dToDate), _
																	CBool(bByAfectationDate), CBool(bUpdatedVouchers))
	End Function		

	Function Report_9(aGL, dInitialDate, dFinalDate)
		Report_9 = oReports.Report9(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function
		
	Function Report_18(nGralLedgerId, aGL, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_18 = oReports.Report18(Session("sAppServer"), nGralLedgerId, aGL, _
																	CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
															    CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function

	Function Report_51(aGL, sPattern, dFinalDate, dExcRateDate, bRounded, nExchangeRateType, nExcRateCurrency)
		Report_51 = oReports.Report51_52(Session("sAppServer"), CStr(sPattern), aGL , _
	                                   CDate(dFinalDate), CDate(dExcRateDate), _
	                                   CBool(bRounded), CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function
                            
	Function Report_53(aGL, sPattern, dFinalDate, dExcRateDate, bRounded, nExchangeRateType, nExcRateCurrency)	
		Report_53 = oReports.Report53_54(Session("sAppServer"), CStr(sPattern), aGL , _
	                                   CDate(dFinalDate), CDate(dExcRateDate), _
	                                   CBool(bRounded), CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function

	Function Report_55(aGL, sPattern, dInitialDate, dFinalDate, dExcRateDate, bRounded, nExchangeRateType, nExcRateCurrency)
		Report_55 = oReports.Report55_56(Session("sAppServer"), CStr(sPattern), aGL, _
															       CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
															       CBool(bRounded), CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function

	Function Report_58(aGL, sPattern, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_58 = oReports.Report58(Session("sAppServer"), aGL, CStr(sPattern), _
															    CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
															    CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function
                         
	Function Report_59(nGL, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_59 = oReports.Report59(Session("sAppServer"), CLng(nGL), CDate(dFinalDate), CDate(dExcRateDate), _
																	CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function
	
	Function Report_60(dFinalDate)
		Report_60 = oReports.Report60(Session("sAppServer"), CDate(dFinalDate), Array("6205-02-04","6206-01-04","6206-03-06"))
  End Function    
  
	Function Report_61(nGL, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
	  Report_61 = oReports.Report61(Session("sAppServer"), CLng(nGL), CDate(dFinalDate), CDate(dExcRateDate), _
																	CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function
        
	Function Report_62(aGL, dInitialDate, dFinalDate, sAccountList)
		Report_62 = oReports.Report62(Session("sAppServer"), CStr(sAccountList), aGL, _
															    CDate(dInitialDate), CDate(dFinalDate))
	End Function
                        
	Function Report_64(aGL, dInitialDate, dFinalDate, bPrintInCascade)
		Report_64 = oReports.Report64(Session("sAppServer"), aGL, _
		                              CDate(dInitialDate), CDate(dFinalDate), _
		                              CStr(""), CStr(""), CBool(bPrintInCascade))
	End Function

	Function Report_65(aGL, sPattern, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency, bPrintInCascade, bTotal)
		Report_65 = oReports.Report65 (Session("sAppServer"), CStr(sPattern), aGL, _
	                                CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
	                                CLng(nExchangeRateType), CLng(nExcRateCurrency), CBool(bPrintInCascade), CBool(bTotal))
  End Function

	Function Report_68(aGL, dFinalDate)
		Report_68 = oReports.Report68(Session("sAppServer"), aGL, CDate(dFinalDate))
  End Function

	Function Report_69(aGL, dInitialDate, dFinalDate)
		If dFinalDate > CDate("31/12/2001") Then
			Report_69 = oReports.Report69(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate), Array("'2116-01'","'6390-09'","'2116-01','2116-02'"))
		Else
			Report_69 = oReports.Report69(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate), Array("'2116-01'","'6390-08'","'2116-01','2116-02'"))
		End If
	End Function

  Function Report_70(dFinalDate)
		Report_70 = oReports.Report70(Session("sAppServer"), CDate(dFinalDate))
	End Function

	Function Report_73(nGralLedgerGroup, aGL, dInitialDate, dFinalDate)
		Report_73 = oReports.Report73(Session("sAppServer"), CLng(nGralLedgerGroup), aGL, CStr("&&&&-&&-&&-&&-&&-&&-&&"), _
		                              CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_74(aGL, dInitialDate, dFinalDate, dInitialDate2, dFinalDate2, dExcRateDate, nExchangeRateType, nExcRateCurrency)
    Report_74 = oReports.Report74_76(Session("sAppServer"), aGL, _
	                                   CDate(dInitialDate), CDate(dFinalDate), _
	                                   CDate(dInitialDate2), CDate(dFinalDate2), _
                                     CDate(dExcRateDate), _
                                     CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function

	Function Report_96(aGL, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_96 = oReports.Report96(Session("sAppServer"), aGL, _
	                                CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
	                                CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function
      
	Function Report_99(aGL, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_99 = oReports.Report99(Session("sAppServer"), aGL, _
																	CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
																  CLng(nExchangeRateType), CLng(nExcRateCurrency))
	End Function

	Function Report_101(nGralLedgerId, aGL, sPattern, dInitialDate, dFinalDate, _
	                    dExcRateDate, nExchangeRateType, nExcRateCurrency, sAccountList, sTittle)
		Report_101 = oReports.Report101(Session("sAppServer"),	CLng(nGralLedgerId), aGL, CStr(sPattern), _
																	  CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
															      CLng(nExchangeRateType), CLng(nExcRateCurrency), CStr(sAccountList), CStr(sTittle))
  End Function
  
	Function Report_102(aGL, sPattern, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency, sAccountList)
		Report_102 = oReports.Report102(Session("sAppServer"), CStr(sPattern), aGL, _
																	  CDate(dInitialDate), CDate(dFinalDate), CDate(dExcRateDate), _
															      CLng(nExchangeRateType), CLng(nExcRateCurrency), CStr(sAccountList))
	End Function

	Function Report_106(aGL, dInitialDate, dFinalDate)
		Report_106 = oReports.Report106(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_107(aGL, dInitialDate, dFinalDate)
		Report_107 = oReports.Report107(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_108(aGL, dInitialDate, dFinalDate)
		Report_108 = oReports.Report108(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_109(aGL, dInitialDate, dFinalDate)
		Report_109 = oReports.Report109(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_110(aGL, sPattern, dInitialDate, dFinalDate, sFromAccount, sToAccount, _
										  sFromSubsAccount, sToSubsAccount, bPrintInCascade)
		Report_110 = oReports.Report110(Session("sAppServer"), CStr(sPattern), aGL, _
		                                CDate(dInitialDate), CDate(dFinalDate), _
		                                CStr(sFromAccount), CStr(sToAccount), _
		                                CStr(sFromSubsAccount), CStr(sToSubsAccount), CBool(bPrintInCascade))
	End Function
                        
	Function Report_111(aGL, sPattern, dInitialDate, dFinalDate, bPrintInCascade, _
											sFromAccount, sToAccount, sFromSubsAccount, sToSubsAccount)
		Dim oBalanceReporter
		'****************************************************************************
		Set oBalanceReporter = Server.CreateObject("EFABalanceReporter.CReporter")
		
		Report_111 = oBalanceReporter.Report111_112(Session("sAppServer"), aGL, CStr(sPattern), _
		                                    CDate(dInitialDate), CDate(dFinalDate), CBool(bPrintInCascade), _
		                                    CStr(sFromAccount), CStr(sToAccount), 0, 0, 0, 0, 0, , , _
		                                    CStr(sFromSubsAccount), CStr(sToSubsAccount))
		Set oBalanceReporter = Nothing
	End Function

	Function Report_113(aGL, dInitialDate, dFinalDate, sFromAccount, sToAccount, sFromSubsAccount, sToSubsAccount)		
		Report_113 = oReports.Report113(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate), _
		                                CStr(sFromAccount), CStr(sToAccount), CStr(sFromSubsAccount), CStr(sToSubsAccount))
	End Function

	Function Report_115(dFinalDate, sParticipant, sParticipantOrder, sParticipantStatus)
	  Report_115 = oReports.Report115(Session("sAppServer"), CStr(sParticipant), CDate(dFinalDate), _
	                                  CStr(sParticipantOrder), CStr(sParticipantStatus))
  End Function
  
	Function Report_116(nStdAccountType, dFinalDate, sFromAccount, sToAccount)
	  Report_116 = oReports.Report116(Session("sAppServer"), CLng(nStdAccountType), CDate(dFinalDate), _
	                                  CStr(sFromAccount), CStr(sToAccount))
  End Function

	Function Report_117(sAccountOrder)
    Report_117 = oReports.Report117(Session("sAppServer"), CLng(2), CStr(sAccountOrder))
  End Function

	Function Report_120(dFinalDate, nSourceGroup)
		Report_120 = oReports.Param(Session("sAppServer"), CDate(dFinalDate),CLng(5),CLng(10),CLng(nSourceGroup))
	End Function

	Function Report_126(aGL, dFinalDate)
	  Report_126 = oReports.Report126(Session("sAppServer"), CDate(dFinalDate), aGL)
  End Function
   
	Function Report_127(aGL, dInitialDate, dFinalDate, sOptionToDisplay)
	  Report_127 = oReports.Report127(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate), CStr(sOptionToDisplay))
  End Function

	Function Report_128(aGL, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_128 = oReports.Report128(Session("sAppServer"), CDate(dFinalDate), CDate(dExcRateDate), _
																    CStr("5011"), aGL, CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function

	Function Report_144(aGL, dInitialDate, dFinalDate)
		Report_144 = oReports.Report144(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function

	Function Report_149(aGL, sFromAccount, dInitialDate, dFinalDate, dInitialDate2, dFinalDate2, bPrintInCascade)
		Report_149 = oReports.Report149(Session("sAppServer"), CLng(13), aGL, CStr("&&&&"), CDate(dInitialDate), CDate(dFinalDate), _
		                                CDate(dInitialDate2), CDate(dFinalDate2), CBool(bPrintInCascade), CStr(sFromAccount), CStr(sFromAccount))
  End Function

	Function Report_150(aGL, dInitialDate, dFinalDate, dInitialDate2, dFinalDate2)
		Report_150 = oReports.Report150(Session("sAppServer"), CLng(13), aGL, CStr("&&&&"), CDate(dInitialDate), CDate(dFinalDate), _
		                                CDate(dInitialDate2), CDate(dFinalDate2))
  End Function

	Function Report_151(aGL, sFromAccount, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_151 = oReports.Report151(Session("sAppServer"), CLng(13), aGL, CStr("&&&&"), CDate(dInitialDate), CDate(dFinalDate), _
		                                CDate(dExcRateDate), CStr("5010-99"), CStr(sFromAccount), _
		                                "", "", CLng(0), CLng(0), CLng(0), CLng(nExchangeRateType), CLng(nExcRateCurrency), _
		                                True, CLng(4), True)
  End Function
  
	Function Report_152(aGL, dInitialDate, dFinalDate, dExcRateDate, nExchangeRateType, nExcRateCurrency)
		Report_152 = oReports.Report152(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate), _
		                                Array("'5001%'","'5011%'","'5000%'","'5010%'"), _
		                                CDate(dExcRateDate), CLng(nExchangeRateType), CLng(nExcRateCurrency))
  End Function

	Function Report_153(aGL, dInitialDate, dFinalDate, dInitialDate2, dFinalDate2)
	  Report_153 = oReports.Report153(Session("sAppServer"), CLng(13), aGL, CDate(dInitialDate), CDate(dFinalDate), _
	                                CDate(dInitialDate2), CDate(dFinalDate2))
  End Function

	Function Report_156(aGL, dInitialDate, dFinalDate)
		Report_156 = oReports.Report156(Session("sAppServer"), aGL, CDate(dInitialDate), CDate(dFinalDate))
	End Function
  

%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<meta http-equiv="Pragma" content="no-cache">
<link REL="stylesheet" TYPE="text/css" HREF="/empiria/resources/applications.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function showRightButtonMsg() {
  var sMsg;
  
  sMsg = "Para obtener una copia del reporte en su equipo, se requiere hacer\n" +
         "clic con el botón derecho del ratón y seleccionar la opción\n" + 
         "'Guardar destino como...'\n\n" + 
         "Gracias."
	alert(sMsg);	
}

function showReportInBrowser() {	
	window.open('<%=sFileName%>', 'dummy', "menubar=yes,toolbar=yes,scrollbars=yes,status=yes,location=no");
	return true;
}

//-->
</SCRIPT>
</head>
<body>
<table bgColor="khaki" width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td><font size=2><b>La información solicitada está lista.</b></font></td>	
</tr>
<tr>
	<td><font size=2><b>¿Qué desea hacer?</b></font></td>	
</tr>
</table>
<br>
<table width="100%" border="0" cellspacing="0" cellpadding="4">
<tr>
	<td>
		<a href="<%=sFileName%>" onclick="showRightButtonMsg();return false;">
			<img src="/empiria/images/download.jpg" border=0>
		</a>
	</td>	
	<td valign=middle>
		<a href="<%=sFileName%>" onclick="showRightButtonMsg();return false;">
			Si se desea obtener una copia del reporte en su equipo, haga clic sobre esta liga 
			con el botón derecho del ratón y seleccione la opción 'Guardar destino como...'
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>
		<a href="" onclick="showReportInBrowser();return false;">
			<img src="/empiria/images/view.jpg" border=0>
		</a>
	</td>
	<td valign=middle>
		<a href="" onclick="showReportInBrowser();return false;">	
			Ver la información solicitada en una página nueva del navegador.
		</a>
		<br><br>
	</td>	
</tr>
<tr>
	<td>&nbsp;</td>
	<td>
		<a href="" onclick="window.history.back();">
			Cerrar esta ventana y perder la información obtenida.
		</a>
		<br>
	</td>	
</tr>
</table>
</body>
</html>