VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Balance Reporter (EFABalanceReporter)           Parte   : User Services                     *
'* Clase      : CReporter                                       Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Proporciona las balanzas de comprobación del sistema de contabilidad finanaciera mediante   *
'*              archivos Microsoft Excel® 9.0, archivos de texto, XML o HTML.                               *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CReporter"

Private Const cnBlancosCorte As Long = 76
Private Const cnPageLines As Long = 45
Private Const cnFromCurrency As Long = 1

Private Type TMoney
  SI     As Currency
  Cargos As Currency
  Abonos As Currency
  SA     As Currency
End Type

Private Type TMoneyComparative
  SA     As Currency
  SA_2   As Currency
End Type

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

Public Property Get URLFilesPath() As String
  On Error GoTo ErrHandler
    URLFilesPath = GetConstant(cnURLFilesPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "URLFilesPath", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function Balances(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                         sPattern As String, dFromDate As Date, dToDate As Date, _
                         Optional sFromAccount As String, Optional sToAccount As String, _
                         Optional nTransactionTypeId As Long, Optional nVoucherTypeId As Long, _
                         Optional nCurrencyTypeId As Long, _
                         Optional nExcRateTypeId As Long, Optional nExcRateCurrencyId As Long, _
                         Optional dExcRateDate As Date, Optional bPrintInCascade As Boolean, _
                         Optional bCascadeDates As Boolean = False, _
                         Optional nRowsToDisplay As Long = 4, Optional bShowSubsAccts As Boolean, _
                         Optional bShowAvgColumn As Boolean, _
                         Optional bConsolidateExchangeRateCurrency As Boolean) As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    If bShowAvgColumn Then
      Balances = BalancesWithAvg(sAppServer, nGralLedgerGroupId, vGL, sPattern, dFromDate, dToDate, _
                                 sFromAccount, sToAccount, nTransactionTypeId, nVoucherTypeId, _
                                 nCurrencyTypeId, nExcRateTypeId, nExcRateCurrencyId, _
                                 dExcRateDate, bPrintInCascade, bCascadeDates, nRowsToDisplay, bShowSubsAccts)
    ElseIf (Not bShowAvgColumn) Then
      Balances = BalancesWithoutAvg(sAppServer, nGralLedgerGroupId, vGL, sPattern, dFromDate, dToDate, _
                                    sFromAccount, sToAccount, nTransactionTypeId, nVoucherTypeId, _
                                    nCurrencyTypeId, nExcRateTypeId, nExcRateCurrencyId, _
                                    dExcRateDate, bPrintInCascade, bCascadeDates, nRowsToDisplay, bShowSubsAccts, _
                                    bConsolidateExchangeRateCurrency)
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "Balances", Err.Number
End Function

Public Function BalancesComparative(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                    sPattern As String, dIniDate As Date, dEndDate As Date, _
                                    dIniDate2 As Date, dEndDate2 As Date, _
                                    Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                                    Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, Optional bPrintInCascade As Boolean = False, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False, _
                                    Optional bWithDiference As Boolean = False, _
                                    Optional bWithPercent As Boolean = False) As String
  Dim oRs               As Recordset
  Dim oTextFile         As New CTextFile
  Dim sGl               As String
  Dim sAccType          As String
  Dim i                 As Long
  Dim oTempBalances     As CTempBalances
  '***************************************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.Open_ GetConstant(cnGeneratedFilesPath)
    If Not bPrintInCascade Then
      Set oRs = oTempBalances.Comparative(sAppServer, sPattern, vGL, dIniDate, dEndDate, dIniDate2, dEndDate2, sIniAccount, sEndAccount, _
                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                     nOptionToDisplay, bWithAux)

      If (IsArray(vGL)) Then
        sGl = "TODAS LAS CONTABILIDADES"
      Else
        sGl = oTempBalances.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      End If
      sAccType = Mid$(oTempBalances.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)

      ConstructComparativeBalances sAppServer, sGl, sAccType, sPattern, dIniDate, dEndDate, _
                                   dIniDate2, dEndDate2, oTextFile, oRs, nTransactionTypeId, nVoucherTypeId, _
                                   nExcRateCurrency, nOptionToDisplay, bWithDiference, bWithPercent

    Else
      For i = LBound(vGL) To UBound(vGL)
        Set oRs = oTempBalances.Comparative(sAppServer, sPattern, vGL(i), dIniDate, dEndDate, dIniDate2, dEndDate2, sIniAccount, sEndAccount, _
                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                     nOptionToDisplay, bWithAux)
                                     
        sGl = oTempBalances.GetObjName(sAppServer, CLng(vGL(i)), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sAccType = Mid$(oTempBalances.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        
        ConstructComparativeBalances sAppServer, sGl, sAccType, sPattern, dIniDate, dEndDate, _
                                     dIniDate2, dEndDate2, oTextFile, oRs, nTransactionTypeId, nVoucherTypeId, _
                                     nExcRateCurrency, nOptionToDisplay, bWithDiference, bWithPercent
        oRs.Close
      Next i
    End If
    BalancesComparative = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalancesComparative", Err.Number
End Function

Public Function Report111_112(sAppServer As String, vGL As Variant, _
                              sPattern As String, dIniDate As Date, dEndDate As Date, bPrintInCascade As Boolean, _
                              Optional sIniAccount As String, Optional sEndAccount As String, _
                              Optional nTransactionTypeId As Long, Optional nVoucherTypeId As Long, _
                              Optional nCurrencyType As Long, _
                              Optional nExcRateType As Long, Optional nExcRateCurrency As Long, _
                              Optional dExcDate As Date, Optional nOptionToDisplay As Long = 4, _
                              Optional sFromSubsAccount As String, Optional sToSubsAccount As String, _
                              Optional bConsolidateExchangeRateCurrency As Boolean) As String
  Dim oBalanceEngine As New EFABalanceEngine.CEngine
  Dim oBalances      As CTempBalances
  Dim oRs           As Recordset
  Dim oRSGenerator  As CTempBalances
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  Dim bPrintAccount As Boolean
  
  Dim sCuenta       As String
  Dim sSubcuenta    As String
  Dim sMon          As String
  
  Dim tAcumulador   As TMoney
  Dim tSubCuenta    As TMoney
  Dim tCuenta       As TMoney
  Dim tbalance      As TMoney
  Dim tMoneda       As TMoney
  
  Dim sGl           As String
  Dim sCurrency     As String
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  
  Dim vGL2          As Variant
  Dim vGLToSend     As Variant
  Dim i             As Long
  '***************************
  On Error GoTo ErrHandler
    oTextFile.Open_ GetConstant(cnGeneratedFilesPath)
    vGL2 = vGL
    If Not IsArray(vGL2) Then
      vGL2 = Array(vGL2)
    End If
    nPage = 1
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        nGl = CLng(vGLToSend)
        Set oBalances = New CTempBalances
        Set oRSGenerator = New CTempBalances
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      Else
        Set oBalances = New CTempBalances
        Set oRSGenerator = New CTempBalances
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          nGl = vGLToSend(0)
        Else
          nGl = vGLToSend
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(nGl), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        End If
      End If
        
      nAccType = oBalances.AccountType(sAppServer, nGl)
      sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

      Set oRs = oBalanceEngine.BalancesRS(sAppServer, sPattern, vGLToSend, True, True, True, _
                                          dIniDate, dEndDate, sIniAccount, sEndAccount, _
                                          nTransactionTypeId, nVoucherTypeId, nCurrencyType, 0, "", _
                                          nExcRateType, nExcRateCurrency, dExcDate, nOptionToDisplay, _
                                          sFromSubsAccount, sToSubsAccount, False, bConsolidateExchangeRateCurrency)
        sDTime = GetFmtDate
        oTextFile.PrintNewLine
        
        bPrintAccount = True
        nLine = 100
        
        With oRs
          If Not (.BOF And .EOF) Then
            sCuenta = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 4)
            sSubcuenta = Mid(!NUMERO_CUENTA_ESTANDAR, 5) & !clave_sector
            sMon = !ID_MONEDA
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            Do While Not (.EOF)
              PrintHeader111_112 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, "REPORTE DE MOVIMIENTOS Y SALDOS MENSUALES POR AUXILIARES", sDTime, sAccType
              If nLine = 16 Then
                bPrintAccount = True
              End If
              If ((sSubcuenta <> Mid(!NUMERO_CUENTA_ESTANDAR, 5) & !clave_sector) Or (sCuenta <> Mid(!NUMERO_CUENTA_ESTANDAR, 1, 4)) Or (sMon <> !ID_MONEDA)) Then
                PrintTot111 oTextFile, nLine, tAcumulador, tSubCuenta, sSubcuenta, Mid(!NUMERO_CUENTA_ESTANDAR, 5) & !clave_sector, "    TOTAL POR SUBCUENTA"
                bPrintAccount = True
              End If
              If ((sCuenta <> Mid(!NUMERO_CUENTA_ESTANDAR, 1, 4)) Or (sMon <> !ID_MONEDA)) Then
                PrintTot111 oTextFile, nLine, tSubCuenta, tCuenta, sCuenta, Mid(!NUMERO_CUENTA_ESTANDAR, 1, 4), "    TOTAL POR CUENTA"
              End If
              If (sMon <> !ID_MONEDA) Then
                PrintTot111 oTextFile, nLine, tCuenta, tMoneda, sMon, !ID_MONEDA, "    TOTAL POR MONEDA"
                sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
                nLine = 55
                oTextFile.PageFeed (nLine)
                PrintHeader111_112 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, "REPORTE DE MOVIMIENTOS Y SALDOS MENSUALES POR AUXILIARES", sDTime, sAccType
              End If
      
              PrintDetalleAux oTextFile, nLine, tAcumulador, _
                              !NUMERO_CUENTA_ESTANDAR, !clave_sector, !nombre_cuenta, _
                              !numero_cuenta_auxiliar, !nombre_cuenta_auxiliar, _
                              !SALDO_INICIAL, !Cargos, !Abonos, !Saldo_Actual, !fecha_ultimo_movimiento, bPrintAccount
              .MoveNext
            Loop
            PrintTot111 oTextFile, nLine, tAcumulador, tSubCuenta, sSubcuenta, "", "    TOTAL POR SUBCUENTA"
            PrintTot111 oTextFile, nLine, tSubCuenta, tCuenta, sCuenta, "", "    TOTAL POR CUENTA"
            PrintTot111 oTextFile, nLine, tCuenta, tMoneda, sMon, "", "    TOTAL POR MONEDA"
            .Close
          Else
            oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
          End If
        End With
        If bPrintInCascade Then
          oTextFile.PageFeed nLine
          ClearMoney tAcumulador
          ClearMoney tSubCuenta
          ClearMoney tCuenta
          ClearMoney tMoneda
          ClearMoney tbalance
        Else
          Exit For
        End If
    Next i
    Report111_112 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report111_112", Err.Number
End Function

Public Function Report19(sAppServer As String, sPatron As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New CTempBalances
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sNat          As String
  Dim sMon          As String
  Dim sAcc          As String
  
  Dim tAcumulador   As TMoney
  Dim tCuenta       As TMoney
  Dim tMoneda       As TMoney
  Dim tNaturaleza   As TMoney
  
  Dim sCurrency     As String
  Dim sNatureMsg    As String
  
  Dim nCargosD      As Currency
  Dim nAbonosD      As Currency
  Dim nSAD          As Currency
  
  Dim nCargosA      As Currency
  Dim nAbonosA      As Currency
  Dim nSAA          As Currency
  
  Dim nTotCargos    As Currency
  Dim nTotAbonos    As Currency
  Dim nTotSA        As Currency
  '************************************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.Open_ GetConstant(cnGeneratedFilesPath)
    Set oRs = oRSGenerator.GeneraRecordset19(sAppServer, sPatron, vGL, dIniDate, dEndDate, sIniAccount, _
                                             sEndAccount, nTransactionTypeId, nVoucherTypeId, _
                                             nCurrencyType, nRowsToDisplay, bWithAux)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sNat = !NATURALEZA
        sMon = !o_id_moneda
        sAcc = !NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        Do While Not (.EOF)
          PrintHeader19 oTextFile, nLine, nPage, sCurrency, dIniDate, dEndDate, sDTime
          
          If ((sAcc <> (!NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector)) Or (sMon <> !o_id_moneda) Or (sNat <> !NATURALEZA)) Then
            PrintTot19 oTextFile, nLine, tAcumulador, tCuenta, sAcc, (!NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector), "  SUMA DE DELEGACIONES  "
          End If
          If (sNat <> !NATURALEZA) Or (sMon <> !o_id_moneda) Then
            If (sNat = "D") Then
              sNatureMsg = "  TOTAL DEUDORAS  "
              nCargosD = tCuenta.Cargos
              nTotCargos = nTotCargos + nCargosD
              nAbonosD = tCuenta.Abonos
              nTotAbonos = nTotAbonos + nAbonosD
              nSAD = tCuenta.SA
              nTotSA = nTotSA + nSAD
            ElseIf (sNat = "A") Then
              sNatureMsg = "  TOTAL ACREEDORAS  "
              nCargosA = tCuenta.Cargos
              nTotCargos = nTotCargos + nCargosA
              nAbonosA = tCuenta.Abonos
              nTotAbonos = nTotAbonos + nAbonosA
              nSAA = tCuenta.SA
              nTotSA = nTotSA + nSAA
            End If
            PrintTot19 oTextFile, nLine, tCuenta, tNaturaleza, sNat, !NATURALEZA, sNatureMsg
            oTextFile.PrintNewLine
            PrintHeader19 oTextFile, nLine, nPage, sCurrency, dIniDate, dEndDate, sDTime
          End If
          If ((sMon <> !o_id_moneda)) Then ' Or (sNat <> !NATURALEZA)) Then
            PrintQtyLines oTextFile
            oTextFile.PrintLine SpcFill(" ", 30) & SpcFill("  TOTAL POR MONEDA  ", 30) & "  " & _
                                SpcFillAmount(nCargosD + nCargosA, 22) & "  " & _
                                SpcFillAmount(nAbonosD + nAbonosA, 22) & "  " & _
                                SpcFillAmount(nSAD + nSAA, 22)
            PrintQtyLines oTextFile
            sMon = !o_id_moneda
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            nCargosD = 0
            nAbonosD = 0
            nSAD = 0
            nCargosA = 0
            nAbonosA = 0
            nSAA = 0
            
            oTextFile.PageFeed nLine
            PrintHeader19 oTextFile, nLine, nPage, sCurrency, dIniDate, dEndDate, sDTime
          End If
  
          PrintDetalle19 oTextFile, nLine, tAcumulador, _
                         !NUMERO_CUENTA_ESTANDAR, !clave_sector, !NUMERO_MAYOR & "  " & !NOMBRE_MAYOR, _
                         !Cargos, !Abonos, !Saldo_Actual
          .MoveNext
        Loop
        .Close
        If (sNat = "D") Then
          sNatureMsg = "  TOTAL DEUDORAS  "
        ElseIf (sNat = "A") Then
          sNatureMsg = "  TOTAL ACREEDORAS  "
        End If
        PrintTot19 oTextFile, nLine, tAcumulador, tCuenta, sAcc, "", "  SUMA DE DELEGACIONES  "
        If (sNat = "D") Then
          sNatureMsg = "  TOTAL DEUDORAS  "
          nCargosD = tCuenta.Cargos
          nTotCargos = nTotCargos + nCargosD
          nAbonosD = tCuenta.Abonos
          nTotAbonos = nTotAbonos + nAbonosD
          nSAD = tCuenta.SA
          nTotSA = nTotSA + nSAD
          'Debug.Print nCargosD & " " & nAbonosD & " " & nSAD
        ElseIf (sNat = "A") Then
          sNatureMsg = "  TOTAL ACREEDORAS  "
          nCargosA = tCuenta.Cargos
          nTotCargos = nTotCargos + nCargosA
          nAbonosA = tCuenta.Abonos
          nTotAbonos = nTotAbonos + nAbonosA
          nSAA = tCuenta.SA
          nTotSA = nTotSA + nSAA
          'Debug.Print nCargosA & " " & nAbonosA & " " & nSAA
        End If
        PrintTot19 oTextFile, nLine, tCuenta, tNaturaleza, sNat, "", sNatureMsg
        PrintQtyLines oTextFile
        oTextFile.PrintLine SpcFill(" ", 30) & SpcFill("  TOTAL POR MONEDA  ", 30) & "  " & _
                            SpcFillAmount(nCargosD + nCargosA, 22) & "  " & _
                            SpcFillAmount(nAbonosD + nAbonosA, 22) & "  " & _
                            SpcFillAmount(nSAD + nSAA, 22)
        oTextFile.PrintNewLine
        oTextFile.PrintLine SpcFill(" ", 30) & SpcFill("  TOTAL DEL REPORTE ", 30) & "  " & _
                            SpcFillAmount(nTotCargos, 22) & "  " & _
                            SpcFillAmount(nTotAbonos, 22) & "  " & _
                            SpcFillAmount(nTotSA, 22)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report19 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report19", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function BalancesWithAvg(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                 sPattern As String, dIniDate As Date, dEndDate As Date, _
                                 Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                                 Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                 Optional nCurrencyType As Long = 0, _
                                 Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                                 Optional dExcDate As Date, Optional bPrintInCascade As Boolean = False, _
                                 Optional bCascadeDates As Boolean = False, _
                                 Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New CTempBalances
  Dim oBalance      As New CTempBalances
  Dim oTextFile     As New CTextFile
  Dim sGl           As String
  Dim sAccType      As String
  Dim i                As Long
  Dim sReportTitle     As String

  Dim j                As Date
  Dim dIniDateTmp      As Date
  Dim dEndDateTmp      As Date
  Dim dAccDate         As Date
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.Open_ GetConstant(cnGeneratedFilesPath)
    dAccDate = CDate("01/01/1990")
    sReportTitle = "BALANZA DE COMPROBACIÓN"
    
    If Not bPrintInCascade Then
      If (IsArray(vGL)) Then
        sGl = "TODAS LAS CONTABILIDADES"
      Else
       sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      End If
      sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
      
      If Not bCascadeDates Then
        Set oRs = oBalance.BalanceGeneral(sAppServer, sPattern, vGL, dIniDate, dEndDate, sIniAccount, sEndAccount, _
                                          nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                          nOptionToDisplay, bWithAux, False, True)

        ConstructBalancesWithAvg sAppServer, sGl, sAccType, sPattern, dIniDate, dEndDate, _
                                 oTextFile, oRs, sReportTitle, dAccDate, nTransactionTypeId, nVoucherTypeId, _
                                 nExcRateCurrency, nOptionToDisplay
      Else
        
        For j = dIniDate To dEndDate
          dIniDateTmp = j
          dEndDateTmp = j
          sReportTitle = "DIARIO GENERAL"
          dAccDate = dIniDateTmp

          Set oRs = oBalance.BalanceGeneral(sAppServer, sPattern, vGL, dIniDateTmp, dEndDateTmp, sIniAccount, sEndAccount, _
                                            nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                            nOptionToDisplay, bWithAux, False, True)
        
          ConstructBalancesWithAvg sAppServer, sGl, sAccType, sPattern, dIniDateTmp, dEndDateTmp, _
                                   oTextFile, oRs, sReportTitle, dAccDate, nTransactionTypeId, nVoucherTypeId, _
                                   nExcRateCurrency, nOptionToDisplay
          oRs.Close
        Next j

      End If
    Else
      For i = LBound(vGL) To UBound(vGL)
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL(i)), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        If Not bCascadeDates Then
          Set oRs = oBalance.BalanceGeneral(sAppServer, sPattern, vGL(i), dIniDate, dEndDate, sIniAccount, sEndAccount, _
                                            nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                            nOptionToDisplay, bWithAux, False, True)
                                     
        
          ConstructBalancesWithAvg sAppServer, sGl, sAccType, sPattern, dIniDate, dEndDate, _
                                 oTextFile, oRs, sReportTitle, dAccDate, nTransactionTypeId, nVoucherTypeId, _
                                 nExcRateCurrency, nOptionToDisplay
          oRs.Close
        Else
          For j = dIniDate To dEndDate
          dIniDateTmp = j
          dEndDateTmp = j
          sReportTitle = "DIARIO GENERAL"
          dAccDate = dIniDateTmp

          Set oRs = oBalance.BalanceGeneral(sAppServer, sPattern, vGL(i), dIniDateTmp, dEndDateTmp, sIniAccount, sEndAccount, _
                                            nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                            nOptionToDisplay, bWithAux, False, True)
          
          ConstructBalancesWithAvg sAppServer, sGl, sAccType, sPattern, dIniDateTmp, dEndDateTmp, _
                                   oTextFile, oRs, sReportTitle, dAccDate, nTransactionTypeId, nVoucherTypeId, _
                                   nExcRateCurrency, nOptionToDisplay
          oRs.Close
        Next j
        End If
      Next i
    End If
    BalancesWithAvg = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalancesWithAvg", Err.Number
End Function

Private Function BalancesWithoutAvg(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                    sPattern As String, dIniDate As Date, dEndDate As Date, _
                                    Optional sIniAccount As String, Optional sEndAccount As String, _
                                    Optional nTransactionTypeId As Long, Optional nVoucherTypeId As Long, _
                                    Optional nCurrencyType As Long, _
                                    Optional nExcRateType As Long, Optional nExcRateCurrency As Long, _
                                    Optional dExcDate As Date, Optional bPrintInCascade As Boolean, _
                                    Optional bCascadeDates As Boolean, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean, _
                                    Optional bConsolidateExchangeRateCurrency As Boolean) As String
  Dim oRs           As Object
  Dim oRSGenerator  As New CTempBalances
  Dim oBalance      As CTempBalances
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tAcumulador   As TMoney
  Dim tDeudoras     As TMoney
  Dim tAcreedoras   As TMoney
  Dim tGroup        As TMoney
  Dim tbalance      As TMoney
  Dim tMoneda       As TMoney
  Dim sGl           As String
  Dim sAccType      As String
  Dim sCurrency     As String
  
  Dim vGL2             As Variant
  Dim vGLToSend        As Variant
  Dim i                As Long
  Dim j                As Date
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  Dim dIniDateTmp      As Date
  Dim dEndDateTmp      As Date
  Dim dAccDate         As Date
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.Open_ GetConstant(cnGeneratedFilesPath)
    If nOptionToDisplay = 1 Then
      sAccIncludes = "(Sólo cuentas con saldo actual)"
    ElseIf nOptionToDisplay = 2 Then
      sAccIncludes = "(Sólo cuentas con cargos y abonos)"
    ElseIf nOptionToDisplay = 4 Then
      sAccIncludes = "(Todas las cuentas sin importar su saldo)"
    ElseIf (nOptionToDisplay = 5) Then ' SOBREGIRADOS
      sAccIncludes = " (Saldos Sobregirados) "
    Else
      sAccIncludes = " "
    End If
    
    sTransactionType = TransactionTypeDescription(sAppServer, nTransactionTypeId)
    sVoucherType = VoucherTypeDescription(sAppServer, nVoucherTypeId)
    
    vGL2 = vGL
    If Not IsArray(vGL2) Then
      vGL2 = Array(vGL2)
    End If
    dAccDate = CDate("01/01/1990")
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        Set oBalance = New CTempBalances
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        Set oBalance = New CTempBalances
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        End If
      End If
      sReportTitle = "BALANZA DE COMPROBACIÓN"
      For j = dIniDate To dEndDate
        If bCascadeDates Then
          dIniDateTmp = j
          dEndDateTmp = j
          Set oBalance = New CTempBalances
          sReportTitle = "DIARIO GENERAL"
          dAccDate = dIniDateTmp
        Else
          dIniDateTmp = dIniDate
          dEndDateTmp = dEndDate
        End If
        
        Set oRs = CreateObject("ADODB.Recordset")
        Set oRs = oBalance.BalanceGeneral(sAppServer, sPattern, vGLToSend, dIniDateTmp, dEndDateTmp, sIniAccount, sEndAccount, _
                                          nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                          nOptionToDisplay, bWithAux, False, False, False, bConsolidateExchangeRateCurrency)
        sDTime = GetFmtDate
        oTextFile.PrintNewLine
        nLine = 100
        nPage = 1
        With oRs
          If Not (.BOF And .EOF) Then
            sNat = !NATURALEZA
            sMon = !clave_moneda
            sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
            Do While Not (.EOF)
              PrintHeader oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                             sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
                             
              If ((sGpo <> Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)) Or (sNat <> !NATURALEZA) Or (sMon <> !clave_moneda)) Then
                PrintGroup oTextFile, nLine, tAcumulador, tGroup, sGpo, Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
              End If
              If ((sNat <> !NATURALEZA) Or (sMon <> !clave_moneda)) Then
                PrintNature oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, !NATURALEZA, sCurrency
              End If
              If (sMon <> !clave_moneda) Then
                PrintCurrency oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, !clave_moneda, sCurrency
                sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
                sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
              
                oTextFile.PageFeed nLine
                PrintHeader oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                               sReportTitle, "KEY", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
              End If
  
              If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPattern)) Then
                PrintDetalle oTextFile, nLine, tAcumulador, !NUMERO_CUENTA_ESTANDAR, !clave_sector, !nombre_cuenta, _
                             !SALDO_INICIAL, !Cargos, !Abonos, !Saldo_Actual, !rol_cuenta, 0, nOptionToDisplay
              ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPattern)) Then
                nLine = nLine - 1
              End If
              .MoveNext
            Loop
          Else
            PrintHeader oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                          sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
          End If
          PrintGroup oTextFile, nLine, tAcumulador, tGroup, sGpo, " "
          PrintNature oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, sNat, sCurrency
          PrintCurrency oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, sMon, sCurrency
          PrintBalance oTextFile, tMoneda
          .Close
          .ActiveConnection = Nothing
          Set oBalance = Nothing
        End With
        If bCascadeDates Then
          oTextFile.PageFeed nLine
          ClearMoney tAcumulador
          ClearMoney tDeudoras
          ClearMoney tAcreedoras
          ClearMoney tGroup
          ClearMoney tbalance
          ClearMoney tMoneda
        Else
          Exit For
        End If
      Next j
      If bPrintInCascade Then
        oTextFile.PageFeed nLine
        ClearMoney tAcumulador
        ClearMoney tDeudoras
        ClearMoney tAcreedoras
        ClearMoney tGroup
        ClearMoney tbalance
        ClearMoney tMoneda
      Else
        Exit For
      End If
    Next i
    BalancesWithoutAvg = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalancesWithoutAvg", Err.Number
End Function

Private Sub AddMoney(tAcumulador As TMoney, SI As Currency, _
                     Cargos As Currency, Abonos As Currency, _
                     SA As Currency)
  On Error GoTo ErrHandler
    tAcumulador.SI = tAcumulador.SI + SI
    tAcumulador.Cargos = tAcumulador.Cargos + Cargos
    tAcumulador.Abonos = tAcumulador.Abonos + Abonos
    tAcumulador.SA = tAcumulador.SA + SA
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddMoney", Err.Number
End Sub

Private Sub AddMoneyComparative(tAcumulador As TMoneyComparative, SA As Currency, _
                                SA_2 As Currency)
  On Error GoTo ErrHandler
    tAcumulador.SA = tAcumulador.SA + SA
    tAcumulador.SA_2 = tAcumulador.SA_2 + SA_2
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddMoneyComparative", Err.Number
End Sub

Private Sub ClearMoney(tAcumulador As TMoney)
  On Error GoTo ErrHandler
    tAcumulador.SI = 0
    tAcumulador.Cargos = 0
    tAcumulador.Abonos = 0
    tAcumulador.SA = 0
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ClearMoney", Err.Number
End Sub
 
 Private Sub ClearMoneyComparative(tAcumulador As TMoneyComparative)
  On Error GoTo ErrHandler
    tAcumulador.SA = 0
    tAcumulador.SA_2 = 0
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ClearMoneyComparative", Err.Number
End Sub

 Private Sub ConstructBalancesWithAvg(sAppServer As String, sGl As String, sAccType As String, _
                                 sPattern As String, dIniDate As Date, dEndDate As Date, _
                                 oTextFile As CTextFile, oRs As Recordset, sReportTitle As String, dAccDate As Date, _
                                 Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                 Optional nExcRateCurrency As Long = 0, _
                                 Optional nOptionToDisplay As Long = 4)
  Dim oRSGenerator  As New CTempBalances
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tAcumulador   As TMoney
  Dim tDeudoras     As TMoney
  Dim tAcreedoras   As TMoney
  Dim tGroup        As TMoney
  Dim tbalance      As TMoney
  Dim tMoneda       As TMoney
  Dim sCurrency     As String
  
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  
  Dim j                As Date
  '*********************************************************************************************
  On Error GoTo ErrHandler
    If nOptionToDisplay = 1 Then
      sAccIncludes = "(Sólo cuentas con saldo actual)"
    ElseIf nOptionToDisplay = 2 Then
      sAccIncludes = "(Sólo cuentas con cargos y abonos)"
    ElseIf nOptionToDisplay = 4 Then
      sAccIncludes = "(Todas las cuentas sin importar su saldo)"
    ElseIf (nOptionToDisplay = 5) Then ' SOBREGIRADOS
      sAccIncludes = " (Saldos Sobregirados) "
    Else
      sAccIncludes = " "
    End If
    
    sTransactionType = TransactionTypeDescription(sAppServer, nTransactionTypeId)
    sVoucherType = VoucherTypeDescription(sAppServer, nVoucherTypeId)
    
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        '.MoveFirst
        sNat = !NATURALEZA
        sMon = !o_id_moneda
        sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
        Do While Not (.EOF)
          PrintHeaderAvg oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                         sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
          If ((sGpo <> Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)) Or (sNat <> !NATURALEZA) Or (sMon <> !o_id_moneda)) Then
            PrintGroup oTextFile, nLine, tAcumulador, tGroup, sGpo, Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
          End If
          If ((sNat <> !NATURALEZA) Or (sMon <> !o_id_moneda)) Then
            PrintNature oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, !NATURALEZA, sCurrency
          End If
          If (sMon <> !o_id_moneda) Then
            PrintCurrency oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, !o_id_moneda, sCurrency
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
          
            oTextFile.PageFeed nLine
            PrintHeaderAvg oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                           sReportTitle, "KEY", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
          End If
  
          If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPattern)) Then
              PrintDetalleProm oTextFile, nLine, tAcumulador, !NUMERO_CUENTA_ESTANDAR, !clave_sector, !nombre_cuenta, _
                         !SALDO_INICIAL, !Cargos, !Abonos, !Saldo_Actual, !rol_cuenta, !from_movs, !saldo_promedio
            
          ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPattern)) Then
            nLine = nLine - 1
          End If
          .MoveNext
          'oRsProm.MoveNext
        Loop
      Else
        PrintHeaderAvg oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                       sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
      End If
      PrintGroup oTextFile, nLine, tAcumulador, tGroup, sGpo, " "
      PrintNature oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, sNat, sCurrency
      PrintCurrency oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, sMon, sCurrency
      PrintBalance oTextFile, tMoneda
    End With
    oTextFile.PageFeed nLine
    Exit Sub
ErrHandler:
  RaiseError ClassId, "ConstructBalancesWithAvg", Err.Number
End Sub

Private Sub ConstructComparativeBalances(sAppServer As String, sGl As String, sAccType As String, _
                                    sPattern As String, dIniDate As Date, dEndDate As Date, _
                                    dIniDate2 As Date, dEndDate2 As Date, oTextFile As CTextFile, oRs As Recordset, _
                                    Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                    Optional nExcRateCurrency As Long = 0, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithDiference As Boolean = False, _
                                    Optional bWithPercent As Boolean = False)
  
  Dim oRSGenerator      As New CTempBalances
  Dim nLine             As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon              As String
  Dim tAcumulador       As TMoneyComparative
  Dim tDeudoras         As TMoneyComparative
  Dim tAcreedoras       As TMoneyComparative
  Dim tGroup            As TMoneyComparative
  Dim tbalance          As TMoneyComparative
  Dim tMoneda           As TMoneyComparative
  Dim sCurrency         As String
  
  Dim sReportTitle      As String
  Dim sAccIncludes      As String
  Dim sTransactionType  As String
  Dim sVoucherType      As String
  '*********************************************************************************************
  On Error GoTo ErrHandler

    If nOptionToDisplay = 1 Then
      sAccIncludes = "(Sólo cuentas con saldo actual)"
    ElseIf nOptionToDisplay = 2 Then
      sAccIncludes = "(Sólo cuentas con cargos y abonos)"
    ElseIf nOptionToDisplay = 4 Then
      sAccIncludes = "(Todas las cuentas sin importar su saldo)"
    ElseIf (nOptionToDisplay = 5) Then ' SOBREGIRADOS
      sAccIncludes = " (Saldos Sobregirados) "
    Else
      sAccIncludes = " "
    End If
    
    sTransactionType = TransactionTypeDescription(sAppServer, nTransactionTypeId)
    sVoucherType = VoucherTypeDescription(sAppServer, nVoucherTypeId)
    
    sReportTitle = "BALANZA DE COMPARACIÓN"
      
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sNat = !NATURALEZA
        sMon = !o_id_moneda
        sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
        Do While Not (.EOF)
          PrintHeaderComparative oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                                    dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, _
                                    sTransactionType, sVoucherType, nExcRateCurrency, bWithDiference, bWithPercent
          If ((sGpo <> Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)) Or (sNat <> !NATURALEZA) Or (sMon <> !o_id_moneda)) Then
            PrintGroupComparative oTextFile, nLine, tAcumulador, tGroup, sGpo, Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2), bWithDiference, bWithPercent
          End If
          If ((sNat <> !NATURALEZA) Or (sMon <> !o_id_moneda)) Then
            PrintNatureComparative oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, !NATURALEZA, sCurrency, bWithDiference, bWithPercent
          End If
          If (sMon <> !o_id_moneda) Then
            PrintCurrencyComparative oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, !o_id_moneda, sCurrency, bWithDiference, bWithPercent
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
          
            oTextFile.PageFeed nLine
            PrintHeaderComparative oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                                 dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, sTransactionType, _
                                 sVoucherType, nExcRateCurrency, bWithDiference, bWithPercent
          End If

          If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPattern)) Then
            PrintDetalleComparative oTextFile, nLine, tAcumulador, !NUMERO_CUENTA_ESTANDAR, !clave_sector, !nombre_cuenta, _
                         !Saldo_Actual, !SALDO_ACTUAL_2, !rol_cuenta, !from_movs, bWithDiference, bWithPercent
          ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPattern)) Then
            nLine = nLine - 1
          End If
          .MoveNext
        Loop
      Else
        'PrintRepHeaderComparative oTextFile, nLine, nPage, sGL, sCurrency, dIniDate, dEndDate, _
                                 dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, sTransactionType, _
                                 sVoucherType, nExcRateCurrency, bWithDiference, bWithPercent
      End If
      PrintGroupComparative oTextFile, nLine, tAcumulador, tGroup, sGpo, " ", bWithDiference, bWithPercent
      PrintNatureComparative oTextFile, nLine, tGroup, tDeudoras, tAcreedoras, sNat, sNat, sCurrency, bWithDiference, bWithPercent
      PrintCurrencyComparative oTextFile, nLine, tMoneda, tbalance, tDeudoras, tAcreedoras, sMon, sMon, sCurrency, bWithDiference, bWithPercent
      PrintBalanceComparative oTextFile, tMoneda, bWithDiference, bWithPercent
    End With
    oTextFile.PageFeed nLine
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ConstructComparativeBalances", Err.Number
End Sub

Public Function FormatLine(nLineLength As Long, sLeftString As String, _
                           sCenterString As String, sRightString As String)
  Dim sTemp As String, n1 As Long, n2 As Long
  '************************************************************************
  If (Len(sLeftString) <> 0) And (Len(sCenterString) <> 0) And (Len(sRightString) <> 0) Then
    If Len(sLeftString & sCenterString & sRightString) < nLineLength Then
      n1 = (nLineLength - Len(sCenterString)) \ 2 - Len(sLeftString)
      n2 = nLineLength - (Len(sLeftString & sCenterString & sRightString)) - n1
    Else
      n1 = 2
      n2 = 2
      sRightString = Mid(sRightString, Len(sRightString) - 2)
      sLeftString = Mid(sLeftString, Len(sLeftString) - 2)
    End If
    sTemp = sLeftString & Space(IIf(n1 < 0, 1, n1)) & sCenterString & Space(n2) & sRightString
  End If
  
  If (Len(sLeftString) = 0) And (Len(sCenterString) <> 0) And (Len(sRightString) = 0) Then
    If Len(sCenterString) < nLineLength Then
      n1 = (nLineLength - Len(sCenterString)) \ 2
      n2 = n1 + (n1 Mod 2)
    Else
      n1 = 2
      n2 = 2
      sCenterString = Mid(sCenterString, nLineLength - 4)
    End If
    sTemp = Space(n1) & sCenterString & Space(n2)
  End If
  
  If (Len(sLeftString) <> 0) And (Len(sCenterString) = 0) And (Len(sRightString) <> 0) Then
    If Len(sLeftString & sRightString) < nLineLength Then
      n1 = (nLineLength - Len(sLeftString & sRightString))
      'n2 = n1 + (n1 Mod 2)
    Else
      n1 = 2
      'n2 = 2
      sRightString = Mid(sRightString, Len(sRightString) - 2)
    End If
    sTemp = sLeftString & Space(n1) & sRightString
  End If
  FormatLine = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "FormatLine", Err.Number
End Function

Private Function GetCuentaSubCuenta(ByVal sNumCuenta As String, _
                                   Optional sCuenta As String = "", Optional sSubcuenta As String = "", _
                                   Optional nSpaces As Long = 1, Optional bIsIni As Boolean = True) As String
  Dim nLen As Long
  Dim nPos As Long
  '***********************
  On Error GoTo ErrorHandler
  sCuenta = Mid(sNumCuenta, 1, 4)
  sSubcuenta = Mid(sNumCuenta, 6, 17)
  Do
    If (InStr(sSubcuenta, "-") = 0) Then
      Exit Do
    End If
    nPos = InStr(sSubcuenta, "-")
    sSubcuenta = Mid(sSubcuenta, 1, nPos - 1) & Mid(sSubcuenta, nPos + 1, 17)
  Loop
  nLen = Len(sSubcuenta)
  If bIsIni Then
    If (nLen < 12) Then
      sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
    End If
  Else
    If (nLen < 12) Then
      sSubcuenta = sSubcuenta & Mid("999999999999", 1, 12 - nLen)
    End If
  End If
  
  GetCuentaSubCuenta = sCuenta & Space(nSpaces) & sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetCuentaSubcuenta", Err.Number
End Function

Private Function GetFmtDate() As String
  On Error GoTo ErrHandler
    GetFmtDate = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetFmtDate", Err.Number
End Function

Private Sub MoreMoney(tRes As TMoney, tMinuendo As TMoney, tSustraendo As TMoney)
  On Error GoTo ErrHandler
    tRes.SI = tMinuendo.SI - tSustraendo.SI
    tRes.Cargos = tMinuendo.Cargos + tSustraendo.Cargos
    tRes.Abonos = tMinuendo.Abonos + tSustraendo.Abonos
    tRes.SA = tMinuendo.SA - tSustraendo.SA
  Exit Sub
ErrHandler:
  RaiseError ClassId, "MoreMoney", Err.Number
End Sub

Private Sub MoreMoneyComparative(tRes As TMoneyComparative, tMinuendo As TMoneyComparative, tSustraendo As TMoneyComparative)
  On Error GoTo ErrHandler
    tRes.SA = tMinuendo.SA - tSustraendo.SA
    tRes.SA_2 = tMinuendo.SA_2 - tSustraendo.SA_2
  Exit Sub
ErrHandler:
  RaiseError ClassId, "MoreMoneyComparative", Err.Number
End Sub

Private Sub PrintBalance(oTextFile As CTextFile, tAcumulador As TMoney, Optional nLen As Long = 0)
  On Error GoTo ErrHandler
    If (nLen = 0) Then
      PrintQtyLines oTextFile
    Else
      PrintQtyLines oTextFile, nLen
    End If
    oTextFile.PrintLine SpcFill("    TOTAL  CONSOLIDADO ", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SI, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintBalance", Err.Number
End Sub

Private Sub PrintBalanceComparative(oTextFile As CTextFile, tAcumulador As TMoneyComparative, bWithDif As Boolean, _
                                    bWithPercent As Boolean)
  Dim nPercent As Currency
  '*****************************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    If tAcumulador.SA_2 = 0 Then
      nPercent = 0
    Else
      nPercent = ((tAcumulador.SA_2 - tAcumulador.SA) / tAcumulador.SA_2) * 100
    End If
    If bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL  CONSOLIDADO ", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL  CONSOLIDADO ", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22)
    ElseIf Not bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL  CONSOLIDADO ", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf Not bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL  CONSOLIDADO ", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintBalanceComparative", Err.Number
End Sub

Private Sub PrintCurrency(oTextFile As CTextFile, nLine As Long, tMoneda As TMoney, _
                          tbalance As TMoney, tDeudoras As TMoney, tAcreedoras As TMoney, _
                          sMon As String, sMoneda As String, sCurrencyDesc, Optional nLen As Long = 0)
  On Error GoTo ErrHandler
    If (nLen = 0) Then
      PrintQtyLines oTextFile
    Else
      PrintQtyLines oTextFile, nLen
    End If
    MoreMoney tbalance, tDeudoras, tAcreedoras
    oTextFile.PrintLine SpcFill("    TOTAL MONEDA (" & sCurrencyDesc & ")", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tbalance.SI, 22) & _
                        "  " & SpcFillAmount(tbalance.Cargos, 22) & _
                        "  " & SpcFillAmount(tbalance.Abonos, 22) & _
                        "  " & SpcFillAmount(tbalance.SA, 22)
    oTextFile.PrintNewLine
    AddMoney tMoneda, tbalance.SI, tbalance.Cargos, tbalance.Abonos, tbalance.SA
    ClearMoney tDeudoras
    ClearMoney tAcreedoras
    nLine = nLine + 1
    sMon = sMoneda
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintCurrency", Err.Number
End Sub

Private Sub PrintCurrencyComparative(oTextFile As CTextFile, nLine As Long, tMoneda As TMoneyComparative, _
                          tbalance As TMoneyComparative, tDeudoras As TMoneyComparative, tAcreedoras As TMoneyComparative, _
                          sMon As String, sMoneda As String, sCurrencyDesc, bWithDif As Boolean, bWithPercent As Boolean)
  Dim nPercent As Currency
  '**********************************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    MoreMoneyComparative tbalance, tDeudoras, tAcreedoras
    If tbalance.SA_2 = 0 Then
      nPercent = 0
    Else
      nPercent = ((tbalance.SA_2 - tbalance.SA) / tbalance.SA_2) * 100
    End If
    If bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL MONEDA (" & sCurrencyDesc & ")", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                        "  " & SpcFillAmount(tbalance.SA, 22) & _
                        "  " & SpcFillAmount((tbalance.SA_2 - tbalance.SA), 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL MONEDA (" & sCurrencyDesc & ")", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                        "  " & SpcFillAmount(tbalance.SA, 22) & _
                        "  " & SpcFillAmount((tbalance.SA_2 - tbalance.SA), 22)
    ElseIf Not bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL MONEDA (" & sCurrencyDesc & ")", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                        "  " & SpcFillAmount(tbalance.SA, 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf Not bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL MONEDA (" & sCurrencyDesc & ")", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                        "  " & SpcFillAmount(tbalance.SA, 22)
    End If
    oTextFile.PrintNewLine
    AddMoneyComparative tMoneda, tbalance.SA, tbalance.SA_2
    ClearMoneyComparative tDeudoras
    ClearMoneyComparative tAcreedoras
    nLine = nLine + 1
    sMon = sMoneda
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintCurrencyComparative", Err.Number
End Sub

Public Function PrintDate(dIniDate As Date, dEndDate As Date) As String
  
  If (dIniDate = dEndDate) Then
    PrintDate = "Movimientos del " & Format$(dIniDate, "d"" de ""MMMM"" de ""yyyy")
  End If
  If (dIniDate <> dEndDate) Then
    If (Day(dIniDate) = 1) And ((Month(dIniDate) + 1) = Month(dEndDate + 1)) And (Month(dIniDate) = Month(dEndDate)) Then
      PrintDate = "Movimientos del mes de " & Format(dEndDate, "MMMM"" del año ""YYYY")
    ElseIf ((Day(dIniDate) = 1) And (Month(dIniDate) = 1)) And ((Day(dEndDate) = 31) And (Month(dEndDate) = 12)) And (Year(dIniDate) = Year(dEndDate)) Then
      PrintDate = "Movimientos del año " & Format(dEndDate, "YYYY")
    Else
      PrintDate = "Movimientos del día " & Format(dIniDate, "dd/mmm/yyyy") & " al día " & Format(dEndDate, "dd/mmm/yyyy")
    End If
  End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "PrintDate", Err.Number
End Function

Private Sub PrintDetalle(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, _
                         sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                         nSI As Currency, nCargos As Currency, nAbonos As Currency, nSA As Currency, _
                         sROL As String, nFROM_MOVS As Long, Optional nOptionToDisplay As Long = 0)
  Dim sChar  As String
  '***************************************************************************************************
  On Error GoTo ErrHandler
    If ((Len(sCuenta) = 4) And (sCLAVE_SECTOR = "00")) Or (nOptionToDisplay = 5) Then
      AddMoney tAcumulador, nSI, nCargos, nAbonos, nSA
      sChar = ""
    End If
    If (sROL = "S") Then
      sChar = sChar & " "
    ElseIf (sROL <> "S") Then
      sChar = sChar & "*"
    End If
    oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSI, 22) & "  " & SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & SpcFillAmount(nSA, 22) & sChar
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalle", Err.Number
End Sub

Private Sub PrintDetalleProm(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, _
                             sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                         nSI As Currency, nCargos As Currency, nAbonos As Currency, nSA As Currency, _
                         sROL As String, nFROM_MOVS As Long, nSaldo_Promedio As Currency)
  Dim sChar  As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    If ((Len(sCuenta) = 4) And (sCLAVE_SECTOR = "00")) Then
      AddMoney tAcumulador, nSI, nCargos, nAbonos, nSA
      sChar = ""
    End If
    If (sROL = "S") Then
      sChar = sChar & " "
    ElseIf (sROL <> "S") Then
      sChar = sChar & "*"
    End If
    oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSI, 22) & "  " & SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & _
                        "  " & SpcFillAmount(nSA, 22) & "  " & SpcFillAmount(nSaldo_Promedio, 22) & sChar
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalleProm", Err.Number
End Sub

Private Sub PrintDetalle19(oTextFile As CTextFile, nLine As Long, _
                           tAcumulador As TMoney, _
                           sAccount As String, sCveSector As String, sGl As String, _
                           nCargos As Currency, nAbonos As Currency, nSA As Currency)
  On Error GoTo ErrHandler
    oTextFile.PrintLine SpcFill(sAccount & "  " & sCveSector, 28) & "  " & SpcFill(sGl, 30) & "  " & _
                        SpcFillAmount(nCargos, 22) & "  " & _
                        SpcFillAmount(nAbonos, 22) & "  " & _
                        SpcFillAmount(nSA, 22)
    AddMoney tAcumulador, 0, nCargos, nAbonos, nSA
    nLine = nLine + 1
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot19", Err.Number
End Sub

Private Sub PrintGroup(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, tGroup As TMoney, _
                       sGpo As String, sGroup As String, Optional nLen As Long = 0)
  Dim txtLabel As String
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    If nLen = 0 Then
      PrintQtyLines oTextFile
    Else
      PrintQtyLines oTextFile, nLen
    End If
    AddMoney tGroup, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    oTextFile.PrintLine SpcFill("    TOTAL GRUPO " & sGpo & "00", cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SI, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sGpo = sGroup
    ClearMoney tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintGroup", Err.Number
End Sub

Private Sub PrintHeader(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                        sGl As String, sCC As String, sCurrency As String, _
                        sGLN As String, sCCname As String, sCurrName As String, _
                        dIniDate As Date, dEndDate As Date, _
                        sTittleReport As String, sReportKey As String, sDateTime As String, _
                        Optional sAccountInclude As String, Optional sAccType As String, _
                        Optional sTransactionType As String, Optional sVoucherType As String, _
                        Optional nExcRateCurrency As Long, Optional dAccDate As Date)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      If nLine >= 101 Then
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
      Else
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
      End If
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", PrintDate(dIniDate, dEndDate), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "", sAccountInclude, "")
      If (dAccDate <> "01/01/1990") Then
        oTextFile.PrintLine "FECHA MTO.         :" & dAccDate
      End If
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      oTextFile.PrintLine FormatLine(174, "T. DE CONTABILIDAD : " & sAccType, "", sVoucherType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, "MONEDA             : " & sCurrency & " " & sCurrName, "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, "VALORIZADA A       : " & sCurrency & " " & sCurrName, "", sDateTime)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO ANTERIOR                   CARGOS                   ABONOS             SALDO ACTUAL "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader", Err.Number
End Sub

Private Sub PrintHeaderAvg(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCC As String, sCurrency As String, _
                           sGLN As String, sCCname As String, sCurrName As String, _
                           dIniDate As Date, dEndDate As Date, _
                           sTittleReport As String, sReportKey As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional dAccDate As Date)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", PrintDate(dIniDate, dEndDate), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "", sAccountInclude, "")
      If (dAccDate <> "01/01/1990") Then
        oTextFile.PrintLine "FECHA MTO.         :" & dAccDate
      End If
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      oTextFile.PrintLine FormatLine(174, "T. DE CONTABILIDAD : " & sAccType, "", sVoucherType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, "MONEDA             : " & sCurrency & " " & sCurrName, "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, "VALORIZADA A       : " & sCurrency & " " & sCurrName, "", sDateTime)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO ANTERIOR                   CARGOS                   ABONOS             SALDO ACTUAL           SALDO_PROMEDIO "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderAvg", Err.Number
End Sub

Private Sub PrintHeader111_112(oTextFile As CTextFile, nLine As Long, nPage As Long, sGl As String, _
                               sCurrency As String, dIniDate As Date, dEndDate As Date, _
                               sTittleReport As String, sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                      BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                                          FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                            INSTITUCION BANCA DE DESARROLLO                                                                    PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                             " & sTittleReport
      oTextFile.PrintLine "BNO-670315-CDO                                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "MONEDA               : " & sCurrency
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA   SUBCUENTA      DESCRIPCION           F. ULT. MOV.     AUXILIAR          DESCRIPCION                    SALDO ANTERIOR                   CARGOS                   ABONOS             SALDO ACTUAL "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader111_112", Err.Number
End Sub

Private Sub PrintHeader19(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                             sCurrency As String, _
                             dIniDate As Date, dEndDate As Date, _
                             sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                             BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                    INSTITUCION BANCA DE DESARROLLO                                              PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                         BALANZA DE CONSOLIDACION DE SUCURSALES FORANEAS"
      oTextFile.PrintLine "BNO-670315-CDO                                                           " & " AL: " & Day(dEndDate) & " DE " & MonthName(Month(dEndDate)) & " DE " & Year(dEndDate)
      oTextFile.PrintNewLine
      oTextFile.PrintLine "MONEDA          : " & sCurrency
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CUENTA                        DELEGACIONES                                    CARGOS                   ABONOS             SALDO ACTUAL "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 14
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader19", Err.Number
End Sub

Private Sub PrintHeaderComparative(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           dIniDate As Date, dEndDate As Date, dIniDate2 As Date, dEndDate2 As Date, _
                           sTittleReport As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional bWithDif As Boolean = False, _
                           Optional bWithPercent As Boolean = False)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", PrintDate(dIniDate2, dEndDate2) & " Y " & PrintDate(dIniDate, dEndDate), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "", sAccountInclude, "")
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      oTextFile.PrintLine FormatLine(174, "T. DE CONTABILIDAD : " & sAccType, "", sVoucherType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, "MONEDA             : " & sCurrency, "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, "VALORIZADA A       : " & sCurrency, "", sDateTime)
      End If
      oTextFile.PrintNewLine
      If bWithDif And bWithPercent Then
        oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO SEGUNDO PERIODO       SALDO PRIMER PERIODO          DIFERENCIA          VARIACION "
        oTextFile.PrintLine "                                                                                                                                                                     (%)    "
      ElseIf bWithDif And Not bWithPercent Then
        oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO SEGUNDO PERIODO       SALDO PRIMER PERIODO          DIFERENCIA                    "
      ElseIf Not bWithDif And bWithPercent Then
        oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO SEGUNDO PERIODO       SALDO PRIMER PERIODO          VARIACION                     "
        oTextFile.PrintLine "                                                                                                                                                 (%)                        "
      ElseIf Not bWithDif And Not bWithPercent Then
        oTextFile.PrintLine "CTA   SUBCUENTA             NOMBRE                                                  SALDO SEGUNDO PERIODO       SALDO PRIMER PERIODO                                        "
      End If
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderComparative", Err.Number
End Sub

Private Sub PrintGroupComparative(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoneyComparative, _
                                  tGroup As TMoneyComparative, sGpo As String, sGroup As String, _
                                  bWithDif As Boolean, bWithPercent As Boolean)
  Dim txtLabel As String
  Dim nPercent As Currency
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    AddMoneyComparative tGroup, tAcumulador.SA, tAcumulador.SA_2
    If tAcumulador.SA_2 = 0 Then
      nPercent = 0
    Else
      nPercent = ((tAcumulador.SA_2 - tAcumulador.SA) / tAcumulador.SA_2) * 100
    End If
    If bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL GRUPO " & sGpo & "00", cnBlancosCorte) & _
                          "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                         "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                         "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22) & _
                         "  " & SpcFillAmount(nPercent, 22)
    ElseIf bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL GRUPO " & sGpo & "00", cnBlancosCorte) & _
                          "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                          "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                          "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22)
    ElseIf Not bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL GRUPO " & sGpo & "00", cnBlancosCorte) & _
                          "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                         "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                         "  " & SpcFillAmount(nPercent, 22)
    ElseIf Not bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill("    TOTAL GRUPO " & sGpo & "00", cnBlancosCorte) & _
                          "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                         "  " & SpcFillAmount(tAcumulador.SA, 22)
    End If
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sGpo = sGroup
    ClearMoneyComparative tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintGroupComparative", Err.Number
End Sub

Private Sub PrintNatureComparative(oTextFile As CTextFile, nLine As Long, _
                        tAcumulador As TMoneyComparative, tDeudoras As TMoneyComparative, _
                        tAcreedoras As TMoneyComparative, _
                        sNat As String, sNature As String, sCurrency As String, _
                        bWithDif As Boolean, bWithPercent As Boolean)
  Dim txtLabel As String
  Dim nPercent As Currency
  '****************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    If (sNat = "D") Then
      txtLabel = "    TOTAL DEUDORAS [" & sCurrency & "]"
      AddMoneyComparative tDeudoras, tAcumulador.SA, tAcumulador.SA_2
    ElseIf (sNat = "A") Then
      txtLabel = "    TOTAL ACREEDORAS [" & sCurrency & "]"
      AddMoneyComparative tAcreedoras, tAcumulador.SA, tAcumulador.SA_2
    End If
    If tAcumulador.SA_2 = 0 Then
      nPercent = 0
    Else
      nPercent = ((tAcumulador.SA_2 - tAcumulador.SA) / tAcumulador.SA_2) * 100
    End If
    If bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill(txtLabel, cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill(txtLabel, cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount((tAcumulador.SA_2 - tAcumulador.SA), 22)
    ElseIf Not bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill(txtLabel, cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22) & _
                        "  " & SpcFillAmount(nPercent, 22)
    ElseIf Not bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill(txtLabel, cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SA_2, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
   End If
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sNat = sNature
    ClearMoneyComparative tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintNatureComparative", Err.Number
End Sub

Private Sub PrintQtyLines(oTextFile As CTextFile, Optional nLen As Long = 0, Optional nIniSpace As Long)
  Dim sLine           As String
  '*************************************************
  On Error GoTo ErrorHandler
  '182 columnas
    If (nLen = 0) Then
      sLine = "==================================================" & _
              "==================================================" & _
              "==================================================" & _
              "===================================================="
      oTextFile.PrintLine sLine
      oTextFile.PrintNewLine
    Else
      sLine = Space(nIniSpace) & "==================================================" & _
              "==================================================" & _
              "==================================================" & _
              "===================================================="
      sLine = Left(sLine, nLen + nIniSpace)
      oTextFile.PrintLine sLine
    End If
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintQtyLines", Err.Number
End Sub

Private Sub PrintTot111(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, tTot As TMoney, _
                        sOld As String, sNew As String, sMsg As String)
  '****************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    AddMoney tTot, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    oTextFile.PrintLine SpcFill(sMsg, cnBlancosCorte + 22) & _
                        "      " & SpcFillAmount(tAcumulador.SI, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
    oTextFile.PrintNewLine
    nLine = nLine + 5
    sOld = sNew
    ClearMoney tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot111", Err.Number
End Sub

Private Sub PrintTot19(oTextFile As CTextFile, nLine As Long, _
                       tAcumulador As TMoney, tTot As TMoney, _
                       sVal As String, sNewVal As String, sMsg As String)
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    oTextFile.PrintLine SpcFill(" ", 30) & SpcFill(sMsg, 30) & "  " & _
                        SpcFillAmount(tAcumulador.Cargos, 22) & "  " & _
                        SpcFillAmount(tAcumulador.Abonos, 22) & "  " & _
                        SpcFillAmount(tAcumulador.SA, 22)
    oTextFile.PrintNewLine
    AddMoney tTot, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    ClearMoney tAcumulador
    sVal = sNewVal
    nLine = nLine + 3
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot19", Err.Number
End Sub

Private Sub PrintDetalleAux(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, _
                            sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                            sCLAVE_AUX As Variant, sNOMBRE_AUX As Variant, nSI As Currency, _
                            nCargos As Currency, nAbonos As Currency, nSA As Currency, _
                            dUltMov As Variant, bPrintAccount As Boolean)
  Dim sLabel  As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    AddMoney tAcumulador, nSI, nCargos, nAbonos, nSA
    
    If (bPrintAccount) Then
      sLabel = GetCuentaSubCuenta(sCuenta) & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 20)
    ElseIf Not (bPrintAccount) Then
      sLabel = SpcFill(" ", 41)
    End If
    bPrintAccount = False
    
    oTextFile.PrintLine sLabel & "  " & SpcFill(dUltMov, 10) & "  " & SpcFill(sCLAVE_AUX, 20) & "  " & SpcFill(sNOMBRE_AUX, 25) & "  " & _
              SpcFillAmount(nSI, 22) & "  " & SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & SpcFillAmount(nSA, 22)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalleAux", Err.Number
End Sub

Private Sub PrintDetalleComparative(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoneyComparative, _
                         sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                         nSA As Currency, nSA_2 As Currency, _
                         sROL As String, nFROM_MOVS As Long, bWithDif As Boolean, bWithPercent As Boolean)
  Dim sChar    As String
  Dim nPercent As Currency
  '******************************************************************************************************
  On Error GoTo ErrHandler
    If ((Len(sCuenta) = 4) And (sCLAVE_SECTOR = "00")) Then
      AddMoneyComparative tAcumulador, nSA, nSA_2
      sChar = ""
    End If
    If (sROL = "S") Then
      sChar = sChar & " "
    ElseIf (sROL <> "S") Then
      sChar = sChar & "*"
    End If
    If nSA_2 = 0 Then
      nPercent = 0
    Else
      nPercent = ((nSA_2 - nSA) / nSA_2) * 100
    End If
    If bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSA_2, 22) & "  " & SpcFillAmount(nSA, 22) & "  " & SpcFillAmount(nSA_2 - nSA, 22) & _
                        "  " & SpcFillAmount(nPercent, 22) & sChar
    ElseIf bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSA_2, 22) & "  " & SpcFillAmount(nSA, 22) & "  " & SpcFillAmount(nSA_2 - nSA, 22) & sChar
    ElseIf Not bWithDif And bWithPercent Then
      oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSA_2, 22) & "  " & SpcFillAmount(nSA, 22) & "  " & SpcFillAmount(nPercent, 22) & sChar
    ElseIf Not bWithDif And Not bWithPercent Then
      oTextFile.PrintLine SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 48) & "  " & _
                        SpcFillAmount(nSA_2, 22) & "  " & SpcFillAmount(nSA, 22) & sChar
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalleComparative", Err.Number
End Sub

Private Sub PrintNature(oTextFile As CTextFile, nLine As Long, _
                        tAcumulador As TMoney, tDeudoras As TMoney, tAcreedoras As TMoney, _
                        sNat As String, sNature As String, sCurrency As String, Optional nLen As Long = 0)
  Dim txtLabel As String
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    If nLen = 0 Then
      PrintQtyLines oTextFile
    Else
      PrintQtyLines oTextFile, nLen
    End If
    If (sNat = "D") Then
      txtLabel = "    TOTAL DEUDORAS [" & sCurrency & "]"
      AddMoney tDeudoras, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    ElseIf (sNat = "A") Then
      txtLabel = "    TOTAL ACREEDORAS [" & sCurrency & "]"
      AddMoney tAcreedoras, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    End If
    oTextFile.PrintLine SpcFill(txtLabel, cnBlancosCorte) & _
                        "  " & SpcFillAmount(tAcumulador.SI, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sNat = sNature
    ClearMoney tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintNature", Err.Number
End Sub

Private Function SpcFill(ByVal sStr As Variant, nLen As Long) As String
  On Error GoTo ErrorHandler
    If IsNull(sStr) Then
      sStr = ""
    End If
    If (Len(sStr) < nLen) Then
      SpcFill = sStr & Space(nLen - Len(sStr))
    Else
      SpcFill = Mid(sStr, 1, nLen)
    End If
  Exit Function
ErrorHandler:
  RaiseError ClassId, "SpcFill", Err.Number
End Function

Private Function SpcFillAmount(ByVal cAmount As String, nLen As Long, Optional nHowManyDec As Long = 2) As String
  Dim sFmtAmount As String
  Dim sSign      As String
  '************************************************************************
  On Error GoTo ErrorHandler
    sSign = " "
    If (cAmount < 0) Then
      cAmount = cAmount * -1
      sSign = "-"
    End If
    sFmtAmount = FormatNumber(cAmount, nHowManyDec)
    If (nLen - Len(sFmtAmount)) > 0 Then
      SpcFillAmount = Space(nLen - Len(sFmtAmount)) & sFmtAmount & sSign
    Else
      SpcFillAmount = sFmtAmount & sSign
    End If
  Exit Function
ErrorHandler:
  RaiseError ClassId, "SpcFillAmount", Err.Number
End Function

Private Function TransactionTypeDescription(sAppServer As String, nTransactionTypeId As Long) As String
  Dim oRSGenerator  As New CTempBalances
  '************************************
  If (nTransactionTypeId < 0) Then
    TransactionTypeDescription = "Todas las transacciones excepto: " & oRSGenerator.GetObjName(sAppServer, nTransactionTypeId, "AO_OBJECTS", "OBJECT_NAME", "OBJECT_ID")
  ElseIf (nTransactionTypeId > 0) Then
    TransactionTypeDescription = "Tipos de transacción: " & oRSGenerator.GetObjName(sAppServer, nTransactionTypeId, "AO_OBJECTS", "OBJECT_NAME", "OBJECT_ID")
  Else
    TransactionTypeDescription = " "
  End If
Exit Function
ErrHandler:
  RaiseError ClassId, "TransactionTypeDescription", Err.Number
End Function

Private Function VoucherTypeDescription(sAppServer As String, nVoucherTypeId As Long)
  Dim oRSGenerator  As New CTempBalances
  '************************************
  If (nVoucherTypeId < 0) Then
    VoucherTypeDescription = "Todas las pólizas excepto: " & oRSGenerator.GetObjName(sAppServer, nVoucherTypeId, "AO_OBJECTS", "OBJECT_NAME", "OBJECT_ID")
  ElseIf (nVoucherTypeId > 0) Then
    VoucherTypeDescription = "Tipos de póliza: " & oRSGenerator.GetObjName(sAppServer, nVoucherTypeId, "AO_OBJECTS", "OBJECT_NAME", "OBJECT_ID")
  Else
  VoucherTypeDescription = " "
  End If
Exit Function
ErrHandler:
  RaiseError ClassId, "VoucherTypeDescription", Err.Number
End Function

