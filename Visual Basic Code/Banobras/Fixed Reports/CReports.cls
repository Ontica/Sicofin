VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components\Banobras                    Sistema : Financial Accounting              *
'* Componente : Fixed Reports (SCFFixedReports)                 Parte   : Business services                 *
'* Clase      : CReports                                        Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Generador de los reportes contables fijos                                                   *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CReports"

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

Public Function ReportsTable(sAppServer As String, nUserId As Long) As String
  Const cRow = "<TR><TD><A href='' onclick='return callEditor(<@NUMBER@>);'>" & _
               "<@NAME@></TD><TD><@DESCRIPTION@></TD></TR>" & vbCrLf
  Dim oRecordset As New Recordset, sSQL As String, sHTML As String, sTemp As String
  '********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM AOObjects WHERE entityId = 13 ORDER BY presentationOrder"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = Replace(cRow, "<@NUMBER@>", !objectKey)
        sTemp = Replace(sTemp, "<@NAME@>", !objectName)
        sHTML = sHTML & Replace(sTemp, "<@DESCRIPTION@>", IIf(IsNull(!objectDescription), "&nbsp;", !objectDescription))
        .MoveNext
      Loop
      .Close
    End With
    ReportsTable = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "ReportsTable", Err.Number
End Function

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

Public Function PrintDateGeneral(dIniDate As Date, dEndDate As Date) As String
  If (dIniDate = dEndDate) Then
    PrintDateGeneral = " Del " & Format$(dIniDate, "d"" de ""MMMM"" de ""yyyy")
  End If
  If (dIniDate <> dEndDate) Then
    If (Day(dIniDate) = 1) And (Month(dIniDate) + 1 = Month(dEndDate + 1)) And (Month(dIniDate) = Month(dEndDate)) Then
      PrintDateGeneral = " Del mes de " & Format(dEndDate, "MMMM"" del año ""YYYY")
    ElseIf ((Day(dIniDate) = 1) And (Month(dIniDate) = 1)) And ((Day(dEndDate) = 31) And (Month(dEndDate) = 12)) And (Year(dIniDate) = Year(dEndDate)) Then
      PrintDateGeneral = " Del año " & Format(dEndDate, "YYYY")
    Else
      PrintDateGeneral = " Del día " & Format(dIniDate, "dd/mmm/yyyy") & " al día " & Format(dEndDate, "dd/mmm/yyyy")
    End If
  End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "PrintDateGeneral", Err.Number
End Function

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
  
Public Function CheckAvailableExchangeRates(sAppServer As String, dFromDate As Date, dToDate As Date, _
                                            nExchangeRateTypeId As Long, dExchangeRateDate As Date) As String
  Dim oUsedCurrenciesRS As New Recordset, sSQL As String, sGralLedgers As String, sTemp As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_moneda, currency_name, o_id_moneda FROM ao_currencies, " & _
           "( SELECT DISTINCT id_moneda FROM " & _
           "(SELECT id_moneda FROM cof_movimiento, cof_transaccion " & _
           "WHERE (cof_movimiento.id_transaccion = cof_transaccion.id_transaccion) AND " & _
           "('" & dFromDate & "' <= cof_transaccion.fecha_afectacion) AND " & _
           "(cof_transaccion.fecha_afectacion <= '" & dToDate & "') " & _
           ") UNION (" & _
           " SELECT id_moneda FROM cof_saldo_inicial " & _
           " WHERE (fecha_saldo_inicial <= '" & dToDate & "') ) ) " & _
           " WHERE (ao_currencies.currency_id = id_moneda) " & _
           " ORDER BY o_id_moneda, currency_name "
    With oUsedCurrenciesRS
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      Do While Not .EOF
        If Not ExistsExchangeRate(sAppServer, !ID_MONEDA, nExchangeRateTypeId, dExchangeRateDate) Then
          If Len(sTemp) = 0 Then
            sTemp = "(" & !o_id_moneda & ") " & !CURRENCY_NAME
          Else
            sTemp = sTemp & vbCrLf & "(" & !o_id_moneda & ") " & !CURRENCY_NAME
          End If
        End If
        .MoveNext
      Loop
      .Close
    End With
    CheckAvailableExchangeRates = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CheckAvailableExchangeRates", Err.Number
End Function

Private Function ExistsExchangeRate(sAppServer As String, nCurrencyId As Long, _
                                    nExchangeRateTypeId As Long, dExchangeRateDate As Date) As Boolean
  Dim oExchangeRatesRS As New Recordset, sSQL As String
  '***************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT exchange_rate FROM ao_exchange_rates " & _
           "WHERE (from_currency_id = 1) AND (to_currency_id = " & nCurrencyId & ") AND " & _
           "(exchange_rate_type_id = " & nExchangeRateTypeId & ") AND " & _
           "(from_date = '" & dExchangeRateDate & "')"
    ExistsExchangeRate = False
    With oExchangeRatesRS
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If (!EXCHANGE_RATE <> 0) Then
          ExistsExchangeRate = True
        End If
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExistsExchangeRate", Err.Number
End Function

Private Function TransactionTypeDescription(sAppServer As String, nTransactionTypeId As Long) As String
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  '*************************************************
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
  Dim oRSGenerator As New SCFFixedReportsDS.CServer
  '***********************************************
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

Public Function Vouchers(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date, _
                         bByAfectationDate As Boolean, bUpdatedVoucher As Boolean) As String
  Dim oRs           As Recordset, oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tCargos       As Double
  Dim tAbonos       As Double
  Dim tCargosGroup  As Double
  Dim tAbonosGroup  As Double
  Dim nNumVoucher      As Long
  Dim nNumVoucherGroup As Long
  Dim sReportTitle  As String
  Dim sGl           As String
  Dim sGLTemp       As String
  
  Dim nAccType      As Long
  Dim nGl           As Long
  Dim sAccType      As String
  
  'Dim sVoucher     As String
  Dim vGL2          As Variant
  Dim i             As Long
  '*******************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    If bUpdatedVoucher Then
      sReportTitle = " POLIZAS ACTUALIZADAS EN EL PERIODO "
    Else
      sReportTitle = " POLIZAS PENDIENTES DE ACTUALIZAR EN EL PERIODO "
    End If
    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
    Else
      nGl = CLng(vGL)
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    nLine = 100
    nPage = 1
    sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    Set oRs = oRSGenerator.VouchersRS(sAppServer, vGL, dIniDate, dEndDate, bByAfectationDate, bUpdatedVoucher)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
      With oRs
        If Not (.BOF And .EOF) Then
          Do While Not (.EOF)
            If (!id_mayor <> nGl) Then
              nGl = !id_mayor
              sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
            End If
            PrintHeaderVouchers oTextFile, nLine, nPage, sGl, sAccType, dIniDate, dEndDate, sReportTitle, sDTime
            If (sGl <> sGLTemp) Or (nLine = 16) Then
              oTextFile.PrintLine "Mayor: " & sGl
              sGLTemp = sGl
            End If
            oTextFile.PrintLine SpcFill(!numero_transaccion, 16) & " " & _
                                SpcFill(!FECHA_AFECTACION, 10) & "   " & _
                                SpcFill(!FECHA_REGISTRO, 10) & "   " & _
                                SpcFill(!ParticipantName, 23) & "  " & _
                                SpcFill(!CONCEPTO_TRANSACCION, 53) & "  " & _
                                SpcFillAmount(!Cargos, 22) & "  " & _
                                SpcFillAmount(!Abonos, 22) & "  " & _
                                IIf((!Cargos - !Abonos) <> 0, "¡¡ N O  B A L A N C E A D A !!", "")
            nLine = nLine + 1
            nNumVoucher = nNumVoucher + 1
            nNumVoucherGroup = nNumVoucherGroup + 1
            tCargosGroup = tCargosGroup + !Cargos
            tAbonosGroup = tAbonosGroup + !Abonos
            tCargos = tCargos + !Cargos
            tAbonos = tAbonos + !Abonos
            .MoveNext
            If Not .EOF Then
              If (!id_mayor <> nGl) Then
                PrintQtyLines oTextFile, 182
                oTextFile.PrintLine " POLIZAS POR MAYOR: " & SpcFill(nNumVoucherGroup, 6) & Space(97) & SpcFillAmount(tCargosGroup, 22) & "  " & SpcFillAmount(tAbonosGroup, 22)
                nNumVoucherGroup = 0
                tCargosGroup = 0
                tAbonosGroup = 0
                oTextFile.PrintNewLine
                nLine = nLine + 3
              End If
            End If
          Loop
          PrintQtyLines oTextFile, 182
          oTextFile.PrintLine " POLIZAS POR MAYOR: " & SpcFill(nNumVoucherGroup, 6) & Space(97) & SpcFillAmount(tCargosGroup, 22) & "  " & SpcFillAmount(tAbonosGroup, 22)
          oTextFile.PrintNewLine
          PrintQtyLines oTextFile, 182
          oTextFile.PrintLine "  TOTAL DE POLIZAS: " & SpcFill(nNumVoucher, 6) & Space(97) & SpcFillAmount(tCargos, 22) & "  " & SpcFillAmount(tAbonos, 22)
        End If
        .Close
      End With
    Vouchers = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Vouchers", Err.Number
End Function

Public Function Param(sAppServer As String, dDate As Date, _
                      nRegla As Long, nLevel As Long, _
                      Optional nSourceGroup As Long = 0) As String
  Dim nType         As Long
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim nEntityGroup  As Long
  '************************************************************
  
  If (nSourceGroup < 0) Then
    nEntityGroup = 16
  ElseIf (nSourceGroup > 0) Then
    nEntityGroup = 9
  Else
    nEntityGroup = 0
  End If
  nType = CLng(oRSGenerator.GetObjName(sAppServer, nRegla, "COF_REGLA_CONTABLE", "TIPO_REGLA", "ID_REGLA_CONTABLE"))
  'nType = 2
  
  If (nType = 1) Or (nType = 3) Then
    Param = ParamWithDestiny(sAppServer, dDate, nRegla, nLevel, nEntityGroup, Abs(nSourceGroup))
  ElseIf (nType = 2) Then
    Param = ParamWithoutDestiny(sAppServer, dDate, nRegla, nLevel)
  End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "Param", Err.Number
End Function

Private Function ParamWithDestiny(sAppServer As String, dDate As Date, nRegla As Long, nLevel As Long, _
                                  Optional nEntityGroup As Long = 0, Optional nSourceGroup As Long) As String
  Dim oRs           As Recordset, oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim nNumAcc         As Long
  Dim nNumAccByGroup  As Long
  Dim nNumGroup       As Long
  
  Dim sReportTitle    As String
  Dim sGl             As String
  Dim sVoucher        As String
  Dim i             As Long
  '*******************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sReportTitle = " REPORTE DE CUENTAS PARA LA CONCENTRACION FIDUCIARIA "
      Set oRs = oRSGenerator.ParamWithDestiny(sAppServer, dDate, nRegla, nLevel, nEntityGroup, nSourceGroup)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
    
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sReportTitle = sReportTitle & "  VERSION " & SpcFillAmount(!VERSION_REGLA, 4)
          Do While Not (.EOF)
            PrintHeaderParamWithDestiny oTextFile, nLine, nPage, "", "", "", CDate("01/01/1990"), CDate("01/01/1990"), sReportTitle, sDTime, ""
            If (nNumGroup <> !ID_AGRUPADOR_ORIGEN) Or (nLine = 16) Then
               oTextFile.PrintLine "DESCRIPCION DEL GRUPO: " & SpcFill(!NOMBRE_GRUPO, 60)
               nNumGroup = !ID_AGRUPADOR_ORIGEN
               oTextFile.PrintNewLine
               nLine = nLine + 2
            End If
            oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_ORIGEN, 22) & "   " & _
                                SpcFill(!CLAVE_SECTOR_ORIGEN, 10) & "     " & _
                                IIf(!ID_MONEDA_ORIGEN < 0, SpcFill("(" & !ABBREV_MONEDA_ORIGEN & ")", 12), SpcFill(!ABBREV_MONEDA_ORIGEN, 12)) & "          " & _
                                SpcFill(IIf(!NUMERO_CUENTA_ORIGEN = !NUMERO_CUENTA_DESTINO, " ", !NUMERO_CUENTA_DESTINO), 22) & "   " & _
                                SpcFill(!CLAVE_SECTOR_DESTINO, 10) & "      " & _
                                SpcFill(!NUMERO_CUENTA_AUXILIAR_DESTINO, 18) & "             " & _
                                SpcFill(IIf(!NUMERO_CUENTA_ORIGEN = !NUMERO_CUENTA_SOB, " ", !NUMERO_CUENTA_SOB), 22) & " " & _
                                SpcFill(!CLAVE_SECTOR_SOB, 10) & "     " & _
                                SpcFill(!NUMERO_CUENTA_AUXILIAR_SOB, 18)
                                
            nLine = nLine + 1
            nNumAcc = nNumAcc + 1
            nNumAccByGroup = nNumAccByGroup + 1
            .MoveNext
            If Not .EOF Then
              If (nNumGroup <> !ID_AGRUPADOR_ORIGEN) Then
                oTextFile.PrintLine "======================================================================================================================================================================================"
                oTextFile.PrintLine "NUMERO DE CUENTAS POR GRUPO: " & SpcFill(nNumAccByGroup, 10)
                oTextFile.PrintNewLine
                nNumAccByGroup = 0
              End If
            End If
          Loop
          oTextFile.PrintLine "======================================================================================================================================================================================"
          oTextFile.PrintLine "NUMERO DE CUENTAS POR GRUPO: " & SpcFill(nNumAccByGroup, 10)
          oTextFile.PrintNewLine
          oTextFile.PrintLine "======================================================================================================================================================================================"
          oTextFile.PrintLine " TOTAL DE CUENTA: " & SpcFill(nNumAcc, 10)
          .Close
        Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN "
        End If
      End With
    ParamWithDestiny = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParamWithDestiny", Err.Number
End Function

Private Function ParamWithoutDestiny(sAppServer As String, dDate As Date, _
                                    nRegla As Long, nLevel As Long) As String
  Dim oRs           As Recordset, oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim nNumAcc         As Long
  Dim nNumAccByGroup  As Long
  Dim nNumGroup       As Long
  
  Dim sReportTitle    As String
  Dim sGl             As String
  Dim sVoucher        As String
  Dim i             As Long
  '*******************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    sReportTitle = " REPORTE DE CUENTAS PARA LA SECTORIZACION "
      sGl = oRSGenerator.GetObjName(sAppServer, 9, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      Set oRs = oRSGenerator.ParamWithoutDestiny(sAppServer, dDate, nRegla, nLevel)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
    
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sReportTitle = sReportTitle & "  VERSION " & SpcFillAmount(!VERSION_REGLA, 4)
          Do While Not (.EOF)
            PrintHeaderParamWithoutDestiny oTextFile, nLine, nPage, sGl, "", "", CDate("01/01/1990"), CDate("01/01/1990"), sReportTitle, sDTime, ""
            If (nNumGroup <> !ID_AGRUPADOR_ORIGEN) Or (nLine = 16) Then
               oTextFile.PrintLine "DESCRIPCION DEL GRUPO: " & SpcFill(!NOMBRE_GRUPO, 60)
               nNumGroup = !ID_AGRUPADOR_ORIGEN
               oTextFile.PrintNewLine
               nLine = nLine + 2
            End If
            
            oTextFile.PrintLine SpcFill(!numero_grupo_cuenta, 22) & "   " & _
                                SpcFill(!nombre_grupo_cuenta, 35) & "     " & _
                                SpcFill(!NUMERO_CUENTA_ORIGEN, 22) & "          " & _
                                SpcFill(!CLAVE_SECTOR_ORIGEN, 10) & "     " & _
                                IIf(!ID_MONEDA_ORIGEN < 0, SpcFill("(" & !ABBREV_MONEDA_ORIGEN & ")", 12), SpcFill(!ABBREV_MONEDA_ORIGEN, 12))
                                
            nLine = nLine + 1
            nNumAcc = nNumAcc + 1
            nNumAccByGroup = nNumAccByGroup + 1
            .MoveNext
            If Not .EOF Then
              If (nNumGroup <> !ID_AGRUPADOR_ORIGEN) Then
                oTextFile.PrintLine "======================================================================================================================================================================================"
                oTextFile.PrintLine "NUMERO DE CUENTAS POR GRUPO: " & SpcFill(nNumAccByGroup, 10)
                oTextFile.PrintNewLine
                nNumAccByGroup = 0
              End If
            End If
          Loop
          oTextFile.PrintLine "======================================================================================================================================================================================"
          oTextFile.PrintLine "NUMERO DE CUENTAS POR GRUPO: " & SpcFill(nNumAccByGroup, 10)
          oTextFile.PrintNewLine
          oTextFile.PrintLine "======================================================================================================================================================================================"
          oTextFile.PrintLine " TOTAL DE CUENTA: " & SpcFill(nNumAcc, 10)
          .Close
        Else
          oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
        End If
      End With
    ParamWithoutDestiny = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParamWithoutDestiny", Err.Number
End Function

Public Function Report9(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  
  Dim sDateTime     As String
  Dim sGl           As String
  Dim sSubAcc       As String
  Dim sAcc          As String
  Dim sMoneda       As String
  Dim bPrintAcc     As Boolean
  
  Dim tSubAccount   As TMoney
  Dim tAccount      As TMoney
  Dim tCurrency       As TMoney
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  
  Dim nCargo        As Double
  Dim nAbono        As Double
  Dim sLabel        As String
  Dim sConcepto     As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset9(sAppServer, vGL, dIniDate, dEndDate)
    sDateTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sSubAcc = GetSubCuenta(!cuenta) & !clave_sector
        sAcc = Mid(!cuenta, 1, 4)
        sMoneda = !MONEDA
        bPrintAcc = True
        Do While Not (.EOF)
          PrintHeader9 oTextFile, nLine, nPage, dIniDate, dEndDate, sGl, SpcFillAmount(!MONEDA, 4, 0) & "  " & !CURRENCY_NAME, sDateTime, bPrintAcc, sAccType
          
          If (sSubAcc <> (GetSubCuenta(!cuenta) & !clave_sector)) Then
            PrintQtyLines oTextFile, 54, 128
            oTextFile.PrintLine Space(28) & " TOTAL POR SUBCUENTA  " & Space(78) & _
                                SpcFillAmount(tSubAccount.Cargos, 22) & "  " & SpcFillAmount(tSubAccount.Abonos, 22)
            sSubAcc = GetSubCuenta(!cuenta) & !clave_sector
            ClearMoney tSubAccount
            bPrintAcc = True
            nLine = nLine + 2
          End If
          If (sAcc <> Mid(!cuenta, 1, 4)) Then
            PrintQtyLines oTextFile, 54, 128
            oTextFile.PrintLine Space(28) & " TOTAL POR CUENTA     " & Space(78) & _
                                SpcFillAmount(tAccount.Cargos, 22) & "  " & SpcFillAmount(tAccount.Abonos, 22)
            sAcc = Mid(!cuenta, 1, 4)
            ClearMoney tSubAccount
            ClearMoney tAccount
            bPrintAcc = True
            nLine = nLine + 2
          End If
          
          If (!MONEDA <> sMoneda) Then
            PrintQtyLines oTextFile, 54, 128
            oTextFile.PrintLine Space(28) & " TOTAL POR MONEDA     " & Space(78) & _
                                SpcFillAmount(tCurrency.Cargos, 22) & "  " & SpcFillAmount(tCurrency.Abonos, 22)
            sMoneda = !MONEDA
            ClearMoney tCurrency
            bPrintAcc = True
            nLine = nLine + 2
            oTextFile.PageFeed nLine
            PrintHeader9 oTextFile, nLine, nPage, dIniDate, dEndDate, sGl, SpcFillAmount(!MONEDA, 4, 0) & "  " & !CURRENCY_NAME, sDateTime, bPrintAcc, sAccType
          End If
          
          If (!tipo_movimiento = "D") Then
            AddMoney tSubAccount, 0, !monto, 0, 0
            AddMoney tAccount, 0, !monto, 0, 0
            AddMoney tCurrency, 0, !monto, 0, 0

            nCargo = !monto
            nAbono = 0
          ElseIf (!tipo_movimiento = "H") Then
            AddMoney tSubAccount, 0, 0, !monto, 0
            AddMoney tAccount, 0, 0, !monto, 0
            AddMoney tCurrency, 0, 0, !monto, 0
            
            nCargo = 0
            nAbono = !monto
          End If
          
          If (bPrintAcc) Then
            sLabel = SpcFill(!cuenta, 22) & "  " & !clave_sector & "  "
            bPrintAcc = False
          ElseIf Not (bPrintAcc) Then
            sLabel = SpcFill(" ", 28)
          End If
          sConcepto = Replace(CStr(!CONCEPTO), String(1, adCR) & String(1, adLF), " ")
          
          oTextFile.PrintLine sLabel & _
                              SpcFill(!POLIZA, 22) & "  " & SpcFill(!Aux, 22) & "  " & _
                              SpcFill(sConcepto, 50) & "  " & _
                              SpcFillAmount(nCargo, 22) & "  " & SpcFillAmount(nAbono, 22)
          nLine = nLine + 1
          .MoveNext
        Loop
        PrintQtyLines oTextFile, 54, 128
        oTextFile.PrintLine Space(28) & " TOTAL POR SUBCUENTA  " & Space(78) & _
                            SpcFillAmount(tSubAccount.Cargos, 22) & "  " & SpcFillAmount(tSubAccount.Abonos, 22)
                            
        PrintQtyLines oTextFile, 54, 128
          
        oTextFile.PrintLine Space(28) & " TOTAL POR CUENTA     " & Space(78) & _
                            SpcFillAmount(tAccount.Cargos, 22) & "  " & SpcFillAmount(tAccount.Abonos, 22)
                    
        PrintQtyLines oTextFile, 54, 128

        oTextFile.PrintLine Space(28) & " TOTAL POR MONEDA     " & Space(78) & _
                            SpcFillAmount(tCurrency.Cargos, 22) & "  " & SpcFillAmount(tCurrency.Abonos, 22)
                            
        .Close
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report9 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report9", Err.Number
End Function

Public Function Report18(sAppServer As String, nGl As Variant, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         dExcDate As Date, Optional nExcRateType As Long = 46, _
                         Optional nExcRateCurrency As Long = 1) As String
                         
  Dim oRsBalance      As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  
  Dim nTotalDeud          As Currency
  Dim nTotalBDeud         As Currency
  Dim nTotalAcre          As Currency
  Dim nTotalBAcre         As Currency
  
  Dim nLine           As Long
  Dim nPage           As Long
  
  Dim sDateTime       As String
  Dim skindacc        As String
  Dim sNat            As String
  '**************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = GetFmtDate
    nLine = 100
    nPage = 1
     
    Set oRsBalance = oRSGenerator.GeneraRecordset18(sAppServer, nGl, vGL, dIniDate, dEndDate, _
                                                     dExcDate, nExcRateType, nExcRateCurrency)
    With oRsBalance
      If Not (.BOF And .EOF) Then
        Do While Not .EOF()
          sNat = !nat
          If sNat = "D" Then
            skindacc = "DEUDORA"
          Else
            skindacc = "ACREEDORA"
          End If
          PrintHeader18 oTextFile, nLine, nPage, dEndDate, "", skindacc, sDateTime
          oTextFile.PrintLine SpcFill(!cuenta, 6) & " " & SpcFill(!nombre, 40) & "  " & _
                              SpcFillAmount(!SA, 25) & "  " & SpcFillAmount(!sa_otros, 25) & " " & _
                              SpcFillAmount((!SA + !sa_otros), 25)
          nLine = nLine + 1
          If sNat = "D" Then
            nTotalDeud = nTotalDeud + !SA
            nTotalBDeud = nTotalBDeud + !sa_otros
          Else
            nTotalAcre = nTotalAcre + !SA
            nTotalBAcre = nTotalBAcre + !sa_otros
          End If
          
          .MoveNext
          If Not .EOF Then
            If (!nat <> sNat) Then
              oTextFile.PrintLine "                                              ===================================================================================="
              oTextFile.PrintLine "                TOTAL CUENTA " & skindacc & "        " & _
                                  SpcFillAmount(nTotalDeud, 30) & "  " & SpcFillAmount(nTotalBDeud, 25) & " " & SpcFillAmount((nTotalDeud + nTotalBDeud), 25)
              nLine = 100
            End If
          End If
        Loop
        oTextFile.PrintLine "                                              ===================================================================================="
        oTextFile.PrintLine "                TOTAL CUENTA " & skindacc & "      " & _
                            SpcFillAmount(nTotalAcre, 30) & "  " & SpcFillAmount(nTotalBAcre, 25) & " " & SpcFillAmount((nTotalAcre + nTotalBAcre), 25)
        oTextFile.PrintNewLine
        oTextFile.PrintLine "                                              ===================================================================================="
        oTextFile.PrintLine "                TOTAL FINAL  " & "               " & _
                            SpcFillAmount((nTotalDeud - nTotalAcre), 30) & "  " & SpcFillAmount((nTotalBDeud - nTotalBAcre), 25) & " " & SpcFillAmount(((nTotalDeud + nTotalBDeud) - (nTotalAcre + nTotalBAcre)), 25)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
      .Close
    End With
    Set oRsBalance = Nothing
    Report18 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report18", Err.Number
End Function

Public Function Report19(sAppServer As String, sPatron As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
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
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset19(sAppServer, sPatron, vGL, dIniDate, dEndDate, sIniAccount, _
                                             sEndAccount, nTransactionTypeId, nVoucherTypeId, _
                                             nCurrencyType, nRowsToDisplay, bWithAux)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sNat = !naturaleza
        sMon = !o_id_moneda
        sAcc = !NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        Do While Not (.EOF)
          PrintHeader19 oTextFile, nLine, nPage, sCurrency, dIniDate, dEndDate, sDTime
          
          If ((sAcc <> (!NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector)) Or (sMon <> !o_id_moneda) Or (sNat <> !naturaleza)) Then
            PrintTot19 oTextFile, nLine, tAcumulador, tCuenta, sAcc, (!NUMERO_CUENTA_ESTANDAR & "  " & !clave_sector), "  SUMA DE DELEGACIONES  "
          End If
          If (sNat <> !naturaleza) Or (sMon <> !o_id_moneda) Then
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
            PrintTot19 oTextFile, nLine, tCuenta, tNaturaleza, sNat, !naturaleza, sNatureMsg
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
        ElseIf (sNat = "A") Then
          sNatureMsg = "  TOTAL ACREEDORAS  "
          nCargosA = tCuenta.Cargos
          nTotCargos = nTotCargos + nCargosA
          nAbonosA = tCuenta.Abonos
          nTotAbonos = nTotAbonos + nAbonosA
          nSAA = tCuenta.SA
          nTotSA = nTotSA + nSAA
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

Public Function Report51_52(sAppServer As String, sPatron As String, vGL As Variant, _
                            dEndDate As Date, dExcDate As Date, Optional bWithRound As Boolean = False, _
                            Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  Dim sNat          As String
  
  Dim nS_MN         As Currency
  Dim nS_OTRA       As Currency
  Dim nS_Tot        As Currency
  Dim tTipoCuenta   As TMoney
  Dim tTotNat       As TMoney
  Dim tTotDeud      As TMoney
  Dim tTotAcre      As TMoney
  Dim sCharD_A      As String
  
  Dim nNumDec       As Long
  Dim nTCuenta      As Long
  Dim sFileName     As String
  
  Dim cGroupName    As Collection
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  '***********************************************************************************
  On Error GoTo ErrHandler
    Set cGroupName = GetGroupsNames
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    Set oRs = oRSGenerator.GeneraRecordset51_a_54(sAppServer, vGL, dEndDate, dExcDate, nExcRateType, nExcRateCurrency)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      nGl = vGL(0)
      sGl = " TODAS LAS CONTABILIDADES "
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        
    nLine = 100
    nPage = 1
    nTCuenta = 6
    
    If (bWithRound) Then
      nNumDec = 0
    ElseIf Not (bWithRound) Then
      nNumDec = 2
    End If
    With oRs
      If Not (.BOF And .EOF) Then
        sNat = !naturaleza
        If (sNat = "D") Then
          sCharD_A = "   1  CTAS DEUDORAS"
        ElseIf (sNat = "A") Then
          sCharD_A = "   2  CTAS ACREEDORAS"
        End If
        Do While Not (.EOF)
          PrintHeader51_52 oTextFile, nLine, nPage, sGl, dEndDate, sCharD_A, sDTime, sAccType
            
          If (Mid(!cuenta, 1, 1) = nTCuenta) Or (sNat <> !naturaleza) Then
             ' Corte por Tipo de Cuenta
            If (sNat = "D") Then
              PrintTot51_52 oTextFile, nLine, tTipoCuenta, tTotDeud, nNumDec, "   SUBTOTAL"
            ElseIf (sNat = "A") Then
              PrintTot51_52 oTextFile, nLine, tTipoCuenta, tTotAcre, nNumDec, "   SUBTOTAL"
            End If
             nTCuenta = 10
          End If
          If (sNat <> !naturaleza) Then
            ' Corte por Naturaleza
            sNat = !naturaleza
            If (!naturaleza = "D") Then
              sCharD_A = "   1  CTAS DEUDORAS"
            ElseIf (!naturaleza = "A") Then
              sCharD_A = "   2  CTAS ACREEDORAS"
            End If
            PrintTot51_52 oTextFile, nLine, tTotDeud, tTotDeud, nNumDec, "     TOTAL DEUDORAS ", sCharD_A, True, True
            nTCuenta = 6
          End If
          
          If (bWithRound) Then
            nS_MN = Round(!SALDO_MN)
            nS_OTRA = Round(!SALDO_OTRA)
            nS_Tot = Round(!SALDO_TOTAL)
          ElseIf Not (bWithRound) Then
            nS_MN = !SALDO_MN
            nS_OTRA = !SALDO_OTRA
            nS_Tot = !SALDO_TOTAL
          End If
          
          AddMoney tTipoCuenta, nS_MN, nS_OTRA, nS_Tot, 0
          oTextFile.PrintLine "  " & SpcFill(!cuenta & "00", 6) & "  " & SpcFill(cGroupName(!cuenta), 30) & "  " & _
                              SpcFillAmount(nS_MN, 22, nNumDec) & "  " & SpcFillAmount(nS_OTRA, 22, nNumDec) & "  " & _
                              SpcFillAmount(nS_Tot, 22, nNumDec)
  
          .MoveNext
        Loop
        .Close
        PrintTot51_52 oTextFile, nLine, tTipoCuenta, tTotAcre, nNumDec, "   SUBTOTAL"
        PrintTot51_52 oTextFile, nLine, tTotAcre, tTotAcre, nNumDec, "     TOTAL ACREEDORAS ", , , True
        PrintQtyLines oTextFile
        oTextFile.PrintLine SpcFill("TOTAL DEL REPORTE ", 42) & _
                          SpcFillAmount((tTotDeud.SI - tTotAcre.SI), 22, nNumDec) & "  " & _
                          SpcFillAmount((tTotDeud.Cargos - tTotAcre.Cargos), 22, nNumDec) & "  " & _
                          SpcFillAmount((tTotDeud.Abonos - tTotAcre.Abonos), 22, nNumDec)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report51_52 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
    RaiseError ClassId, "Report51_52", Err.Number
End Function

Public Function Report53_54(sAppServer As String, sPatron As String, vGL As Variant, _
                            dEndDate As Date, dExcDate As Date, Optional bWithRound As Boolean = False, _
                            Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  Dim sNat          As String
  
  Dim nS_MN         As Currency
  Dim nS_OTRA       As Currency
  Dim nS_Tot        As Currency
  Dim tTotNat       As TMoney
  
  Dim tTotDeud      As TMoney
  Dim tTotAcre      As TMoney
  Dim sCharD_A      As String
  
  Dim nNumDec       As Long
  Dim sFileName     As String
  Dim cGroupName    As Collection
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  '***********************************************************************************
  On Error GoTo ErrHandler
    Set cGroupName = GetGroupsNames
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset51_a_54(sAppServer, vGL, dEndDate, dExcDate, nExcRateType, nExcRateCurrency)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = " TODAS LAS CONTABILIDADES "
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
    nLine = 100
    nPage = 1
    
    If (bWithRound) Then
      nNumDec = 0
    ElseIf Not (bWithRound) Then
      nNumDec = 2
    End If
    
    With oRs
      If Not (.BOF And .EOF) Then
        sNat = !naturaleza
        If (sNat = "D") Then
          sCharD_A = " TOTAL DEUDORAS"
        ElseIf (sNat = "A") Then
          sCharD_A = " TOTAL ACREDORAS"
        End If
        Do While Not (.EOF)
          PrintHeader53_54 oTextFile, nLine, nPage, sGl, dEndDate, sDTime, sAccType
            
          If (sNat <> !naturaleza) Then
            ' Corte por Naturaleza
            PrintTot53_54 oTextFile, nLine, tTotDeud, nNumDec, sCharD_A
            
            sNat = !naturaleza
            If (!naturaleza = "D") Then
              sCharD_A = " TOTAL DEUDORAS"
            ElseIf (sNat = "A") Then
              sCharD_A = " TOTAL ACREDORAS"
            End If
          End If
          
          If (bWithRound) Then
            nS_MN = Round(!SALDO_MN)
            nS_OTRA = Round(!SALDO_OTRA)
            nS_Tot = Round(!SALDO_TOTAL)
          ElseIf Not (bWithRound) Then
            nS_MN = !SALDO_MN
            nS_OTRA = !SALDO_OTRA
            nS_Tot = !SALDO_TOTAL
          End If
          
          If (!naturaleza = "D") Then
            AddMoney tTotDeud, nS_MN, nS_OTRA, nS_Tot, 0
          ElseIf (sNat = "A") Then
            AddMoney tTotAcre, nS_MN, nS_OTRA, nS_Tot, 0
          End If
          
          oTextFile.PrintLine "  " & SpcFill(!cuenta & "00", 13) & "  " & _
                              SpcFillAmount(nS_MN, 22, nNumDec) & "  " & SpcFillAmount(nS_OTRA, 22, nNumDec) & "  " & _
                              SpcFillAmount(nS_Tot, 22, nNumDec) & "  " & SpcFill(cGroupName(!cuenta), 50)
  
          .MoveNext
        Loop
        .Close
        ' Total por Naturaleza
        PrintTot53_54 oTextFile, nLine, tTotAcre, nNumDec, sCharD_A
        PrintQtyLines oTextFile
        oTextFile.PrintLine SpcFill("TOTAL DEL REPORTE ", 17) & _
                          SpcFillAmount(tTotDeud.SI - tTotAcre.SI, 22, nNumDec) & "  " & _
                          SpcFillAmount(tTotDeud.Cargos - tTotAcre.Cargos, 22, nNumDec) & "  " & _
                          SpcFillAmount(tTotDeud.Abonos - tTotAcre.Abonos, 22, nNumDec)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report53_54 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
    RaiseError ClassId, "Report53_54", Err.Number
End Function

Public Function Report55_56(sAppServer As String, sPatron As String, vGL As Variant, _
                            dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                            Optional bWithRound As Boolean = False, _
                            Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  Dim tAcumulador   As TMoney
  Dim tDeudoras     As TMoney
  Dim tAcreedoras   As TMoney
  
  Dim sGl           As String
  
  Dim nS_MN         As Currency
  Dim nS_OTRA       As Currency
  Dim sCharD_A      As String
  
  Dim nNumDec       As Long
  Dim sFileName     As String
  
  Dim sCurrName     As String
  Dim nExcRate      As Currency
  
  Dim nAccType      As Long
  Dim sAccType      As String
  Dim nGl           As Long
  Dim sAccount      As String
  '************************************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    Set oRs = oRSGenerator.GeneraRecordset55_56(sPatron, sAppServer, vGL, dIniDate, dEndDate, _
                                                dExcDate, nExcRateType, nExcRateCurrency)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = " TODAS LAS CONTABILIDADES "
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
 
    If (nExcRateCurrency = 1) Then
      sCurrName = "PESOS"
    ElseIf (nExcRateCurrency <> 1) Then
      sCurrName = oRSGenerator.GetObjName(sAppServer, nExcRateCurrency, "AO_CURRENCIES", "CURRENCY_NAME", "CURRENCY_ID")
    End If
    
    nLine = 100
    nPage = 1
    
    If (bWithRound) Then
      nNumDec = 0
    ElseIf Not (bWithRound) Then
      nNumDec = 2
    End If

    With oRs
      If Not (.BOF And .EOF) Then
        sAccount = !NUMERO_CUENTA_ESTANDAR
        Do While Not (.EOF)
          PrintHeader55_56 oTextFile, nLine, nPage, sGl, sCurrName, dIniDate, dEndDate, sDTime, sAccType
  
          If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPatron)) Then
            If (bWithRound) Then
              nS_MN = Round(!Saldo_actual_MN)
              nS_OTRA = Round(!Saldo_actual_ME)
            ElseIf Not (bWithRound) Then
              nS_MN = (!Saldo_actual_MN)
              nS_OTRA = (!Saldo_actual_ME)
            End If
            
            If (!naturaleza = "D") Then
              sCharD_A = "1"
            ElseIf (!naturaleza = "A") Then
              If (sCharD_A <> "2") Then
                oTextFile.PrintNewLine
                oTextFile.PrintLine Space(30) & "=================================================="
                oTextFile.PrintLine "    TOTAL DEUDORAS      " & Space(6) & _
                                    SpcFillAmount(tDeudoras.SI, 22) & "  " & SpcFillAmount(tDeudoras.Cargos, 22)
                oTextFile.PrintNewLine
                nLine = nLine + 4
              End If
              sCharD_A = "2"
            End If
            
            If (Left(sAccount, 2) <> Left(!NUMERO_CUENTA_ESTANDAR, 2)) Then
              oTextFile.PrintLine Space(30) & "=================================================="
              oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(6) & _
                                  SpcFillAmount(tAcumulador.SI, 22) & "  " & SpcFillAmount(tAcumulador.Cargos, 22)
              oTextFile.PrintNewLine
              ClearMoney tAcumulador
              nLine = nLine + 3
            End If
            
            If (Right(!NUMERO_CUENTA_ESTANDAR, 2) <> "00") Then
              oTextFile.PrintLine sCharD_A & " " & SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & !clave_sector & "  " & _
                                  SpcFillAmount(nS_MN, 22, nNumDec) & "  " & SpcFillAmount(nS_OTRA, 22, nNumDec) & "  " & _
                                 SpcFill(!nombre_cuenta, 50)
              nLine = nLine + 1
            End If
            
            If ((Len(!NUMERO_CUENTA_ESTANDAR) = 4) And (!clave_sector = "00")) Then
              AddMoney tAcumulador, nS_MN, nS_OTRA, 0, 0
              If !naturaleza = "D" Then
                AddMoney tDeudoras, nS_MN, nS_OTRA, 0, 0
              Else
                AddMoney tAcreedoras, nS_MN, nS_OTRA, 0, 0
              End If
            End If
            
            sAccount = !NUMERO_CUENTA_ESTANDAR
          ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPatron)) Then
            nLine = nLine - 1
          End If
          .MoveNext
        Loop
          oTextFile.PrintLine Space(30) & "=================================================="
          oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(6) & _
                                  SpcFillAmount(tAcumulador.SI, 22) & "  " & SpcFillAmount(tAcumulador.Cargos, 22)
                                    
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(30) & "=================================================="
          oTextFile.PrintLine "    TOTAL ACREEDORAS    " & Space(6) & _
                                  SpcFillAmount(tAcreedoras.SI, 22) & "  " & SpcFillAmount(tAcreedoras.Cargos, 22)
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(30) & "=================================================="
          oTextFile.PrintLine "    TOTAL DEL REPORTE   " & Space(6) & _
                                  SpcFillAmount((tDeudoras.SI - tAcreedoras.SI), 22) & "  " & _
                                  SpcFillAmount((tDeudoras.Cargos - tAcreedoras.Cargos), 22)
         ClearMoney tAcumulador
        .Close
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report55_56 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
    RaiseError ClassId, "Report55_56", Err.Number
End Function

Public Function Report58(sAppServer As String, vGL As Variant, ByVal sPatron As String, _
                         dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, _
                         Optional nExcRateCurrency As Long = 2) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim dFecha        As Date
  Dim sAccount      As String
  Dim sSector       As String
  Dim nSaldoMN      As Double
  Dim nSaldoME      As Double
  Dim nSaldoMNA     As Double
  Dim nSaldoMEA     As Double
  Dim sAccName      As String
  Dim tAcumulador   As TMoney
  Dim tDeudoras     As TMoney
  Dim tAcreedoras   As TMoney

  Dim nLine         As Long
  Dim nPage         As Long
  Dim sNature       As String
  Dim sNat          As String
  Dim sDateTime     As String

  Dim DebugAccount  As String
  Dim sAccType      As String
  
  Dim nAccType      As Long
  Dim nGl           As Long
  Dim sGl           As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    If (IsArray(vGL)) Then
      sGl = " TODAS LAS CONTABILIDADES "
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    Set oRs = oRSGenerator.GeneraRecordset58(sPatron, sAppServer, vGL, dIniDate, dEndDate, _
                                             dExcDate, nExcRateType, nExcRateCurrency)

    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

    DebugAccount = "0000"

    With oRs
      If Not (.BOF And .EOF) Then
        nPage = 1
        nLine = 100
        dFecha = !fecha_rev
        sAccount = !NUMERO_CUENTA_ESTANDAR
        sAccName = !nombre_cuenta
        sSector = !clave_sector
        nSaldoMN = 0
        nSaldoME = 0
        nSaldoMNA = 0
        nSaldoMEA = 0
        Do While Not (.EOF)
          If (!naturaleza = "D") Then
            sNature = "1"
            sNat = "1"
          Else
            sNature = "2"
          End If
  
          If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
            If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPatron)) Then
              If (Right(sAccount, 2) <> "00") Then
                PrintHeader58 oTextFile, nLine, nPage, sGl, "", dIniDate, dEndDate, sDateTime, sAccType
                nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
                nSaldoMEA = (nSaldoMEA + (nSaldoME * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
              
                oTextFile.PrintLine sNature & " " & SpcFill(sAccount, 22) & " " & sSector & " " & _
                                  SpcFillAmount(nSaldoMNA, 22) & "  " & _
                                  SpcFillAmount(nSaldoMEA, 22) & "  " & SpcFill(sAccName, 35)
              
                nLine = nLine + 1
              End If
              If (Len(sAccount) = 4) Then
                AddMoney tAcumulador, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
                If (!naturaleza = "D") Then
                  AddMoney tDeudoras, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
                Else
                  AddMoney tAcreedoras, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
                End If
              End If
            ElseIf (Len(!cuenta) > Len(sPatron)) Then
              'Siguiente
            End If
            If (Left(sAccount, 2) <> Left(!NUMERO_CUENTA_ESTANDAR, 2)) Then
              oTextFile.PrintLine Space(26) & "=================================================="
              oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(4) & _
                                  SpcFillAmount(tAcumulador.SI, 22) & "  " & SpcFillAmount(tAcumulador.Cargos, 22)
              oTextFile.PrintNewLine
              ClearMoney tAcumulador
              nLine = nLine + 3
            End If
            If (sNature <> sNat) Then
              oTextFile.PrintLine Space(26) & "=================================================="
              oTextFile.PrintLine "    TOTAL DEUDORAS      " & Space(4) & _
                                  SpcFillAmount(tDeudoras.SI, 22) & "  " & SpcFillAmount(tDeudoras.Cargos, 22)
              oTextFile.PrintNewLine
              sNat = sNature
              nLine = nLine + 3
            End If

            nSaldoMN = 0
            nSaldoME = 0
            nSaldoMNA = 0
            nSaldoMEA = 0
            sAccount = !NUMERO_CUENTA_ESTANDAR
            sAccName = !nombre_cuenta
            sSector = !clave_sector
            dFecha = !fecha_rev
          End If
  
          If (dFecha = #1/1/1990#) Then
             nSaldoMN = !saldo_inicial_mn
             nSaldoME = !Saldo_Inicial_me
             nSaldoMNA = 0
             nSaldoMEA = 0
             dFecha = dIniDate
          ElseIf (dFecha <> #1/1/1990#) Then
            nSaldoMNA = nSaldoMNA + (nSaldoMN * (!fecha_rev - dFecha))
            nSaldoMEA = nSaldoMEA + (nSaldoME * (!fecha_rev - dFecha))
            If (!naturaleza = "D") Then
              nSaldoMN = nSaldoMN + !cargos_mn - !abonos_mn
              nSaldoME = nSaldoME + !cargos_me - !abonos_me
            ElseIf (!naturaleza = "A") Then
              nSaldoMN = nSaldoMN - !cargos_mn + !abonos_mn
              nSaldoME = nSaldoME - !cargos_me + !abonos_me
            End If
            dFecha = !fecha_rev
          End If
  '        If (Mid(!cuenta, 1, 4) = DebugAccount) Then
  '          Debug.Print !cuenta & " " & !fecha_ult & " " & SpcFillAmount(!si_mn, 20) & " " & _
  '                      SpcFillAmount(!cargos_mn, 20) & " " & SpcFillAmount(!abonos_mn, 20) & " " & _
  '                      SpcFillAmount(nSaldoMN, 20) & " " & SpcFillAmount(nSaldoMNA, 20) & " "
  '        End If
          .MoveNext
        Loop
        'Para pruebas 9 dic 2000
        'PrintHeader58 oTextFile, nLine, nPage, "", "", dIniDate, dEndDate, sDateTime, sAccType
  '      Debug.Print "Suma de Saldos " & (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) & " Entre " & (dEndDate - dIniDate + 1) & " Dias "
        If (Right(sAccount, 2) <> "00") Then
          nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
          nSaldoMEA = (nSaldoMEA + (nSaldoME * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
          oTextFile.PrintLine sNature & " " & SpcFill(sAccount, 22) & " " & sSector & " " & _
                              SpcFillAmount(nSaldoMNA, 22) & "  " & _
                              SpcFillAmount(nSaldoMEA, 22) & "  " & SpcFill(sAccName, 35)
          nLine = nLine + 1
        End If
        If (Len(sAccount) = 4) Then
          AddMoney tAcumulador, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
          If (sNature = "1") Then
            AddMoney tDeudoras, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
          Else
            AddMoney tAcreedoras, CCur(nSaldoMNA), CCur(nSaldoMEA), 0, 0
          End If
        End If
        oTextFile.PrintLine Space(26) & "=================================================="
        oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(4) & _
                            SpcFillAmount(tAcumulador.SI, 22) & "  " & SpcFillAmount(tAcumulador.Cargos, 22)
                            
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(26) & "=================================================="
        oTextFile.PrintLine "    TOTAL ACREEDORAS    " & Space(4) & _
                            SpcFillAmount(tAcreedoras.SI, 22) & "  " & SpcFillAmount(tAcreedoras.Cargos, 22)
        
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(26) & "=================================================="
        oTextFile.PrintLine "    TOTAL DEL REPORTE   " & Space(4) & _
                                SpcFillAmount((tDeudoras.SI - tAcreedoras.SI), 22) & "  " & _
                                SpcFillAmount((tDeudoras.Cargos - tAcreedoras.Cargos), 22)
                            
                            
  '      Debug.Print "Saldo Promedio " & nSaldoMNA
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report58 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report58", Err.Number
End Function

Public Function Report59(sAppServer As String, nGl As Long, _
                         dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String

  Dim oRsBalance      As ADODB.Recordset
  Dim oRsPosting      As ADODB.Recordset
  Dim oRsAccNames     As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  Dim sGLDesc         As String
  Dim sLongDate       As String
  
  Dim nTotSubAcc      As Currency
  Dim nTotAcc         As Currency
  Dim nLine           As Long
  Dim nPage           As Long
  Dim nCurr           As Long
  Dim nSect           As Long
  Dim nExcRate        As Currency
  
  Dim sDateTime       As String
  Dim sFileName       As String
  Dim skindacc        As String
  Dim sNat            As String
  Dim sAcc            As String
  Dim sAccOr          As String
  Dim sSubAcc         As String
  Dim sSubAccOr       As String
  Dim nDisFor         As Long
  Dim nDisBac         As Long
  Dim sName           As String
  Dim sNameSub        As String
  Dim sHeadAcc        As String
  Dim sAccount        As String
  Dim sOnlyAccount    As String
  Dim nAccount        As Double
  '*********************************************************************************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sLongDate = Format(dEndDate, "d") & " de " & Format(dEndDate, "MMMM") & " de " & Year(dEndDate)
    
  nLine = 100
  nPage = 1
  
    Set oRsBalance = oRSGenerator.GeneraRecordset59(sAppServer, nGl, dEndDate, dExcDate, _
                                                     nExcRateType, nExcRateCurrency)
    Set oRsAccNames = oRSGenerator.GeneraRSNames(sAppServer, nGl)
    sGLDesc = Space(2) & oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", _
                                                  "ID_MAYOR", "NOMBRE_MAYOR")
    With oRsBalance
      If Not (.BOF And .EOF) Then
        sOnlyAccount = Mid(!cuenta, 1, 4)
        Do While Not .EOF()
          If (!st <> 0) Then
            nCurr = !ID_MONEDA
            nSect = !sector
            sAcc = !cuenta
            sSubAccOr = GetSubCuenta(!cuenta)
            
            If (sOnlyAccount <> Mid(!cuenta, 1, 4)) Then
              PrintQtyLines oTextFile, 139
              oTextFile.PrintLine SpcFill(" TOTAL DE LA CUENTA:  " & sOnlyAccount & "  " & sName, 55) & Space(59) & _
                          SpcFillAmount(nAccount, 25)

              sOnlyAccount = Mid(!cuenta, 1, 4)
              nAccount = 0
              oTextFile.PageFeed nLine
              nLine = 100
            End If
            If sAccount <> Mid(!cuenta, 1, 4) Then
              nLine = 100
              sAccount = Mid(!cuenta, 1, 4)
              sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 4), nGl)
            End If
            
            sHeadAcc = sAccount & "  " & SpcFill(sName, 35) & Space(10) & " Sector: " & ZeroFill(!sector, 2)
            PrintHeader59 oTextFile, nLine, nPage, sLongDate, sGLDesc, skindacc, sHeadAcc, sDateTime
            
            If ((sSubAcc = GetSubCuenta(!cuenta) & ZeroFill(!sector, 2)) And nLine = 20) Then
              sNameSub = GETAccName(oRsAccNames, !cuenta, nGl)
              oTextFile.PrintLine SpcFill("Subcuenta " & sSubAcc & " " & sNameSub, 60)
              nLine = nLine + 1
            End If
            
            If (sSubAcc <> GetSubCuenta(!cuenta) & ZeroFill(!sector, 2)) Then
              sNameSub = GETAccName(oRsAccNames, !cuenta, nGl)
              sSubAcc = GetSubCuenta(!cuenta) & ZeroFill(!sector, 2)
              oTextFile.PrintLine SpcFill("Subcuenta " & sSubAcc & " " & sNameSub, 60)
              nLine = nLine + 1
            End If
            
            oTextFile.PrintLine SpcFill(!nom_aux, 30) & Space(5) & SpcFill(!num_aux, 17) & "  " & _
                                SpcFill(!ID_MONEDA & " " & !nom_moneda, 20) & Space(5) & SpcFillAmount(!st, 20) & "  " & _
                                SpcFillAmount(!exc_rate, 10, 5) & " " & SpcFillAmount((!s_val), 25)
            nAccount = nAccount + !s_val
            nLine = nLine + 1
            nTotSubAcc = nTotSubAcc + (!s_val)
            nTotAcc = nTotAcc + (!s_val)
            .MoveNext
            If Not (.EOF) Then
              If (sAcc = !cuenta) And (nCurr <> !ID_MONEDA) And (nCurr = 1) Then
                oTextFile.PrintLine Space(100) & "======================================="
                oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda NACIONAL ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
                oTextFile.PrintNewLine
                nTotSubAcc = 0
                nLine = nLine + 3
              End If
              If (sAcc = !cuenta) And (nCurr <> !ID_MONEDA) And (!ID_MONEDA = 1) Then
                oTextFile.PrintLine Space(100) & "======================================="
                oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda EXTRANJERA ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
                oTextFile.PrintNewLine
                nTotSubAcc = 0
                nLine = nLine + 3
              End If
              If (sAcc <> !cuenta) And (nCurr = 1) Then
                oTextFile.PrintLine Space(100) & "======================================="
                oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda NACIONAL ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
                oTextFile.PrintNewLine
                nTotSubAcc = 0
                nLine = nLine + 3
              End If
              If (sAcc <> !cuenta) And (nCurr <> 1) Then
                oTextFile.PrintLine Space(100) & "======================================="
                oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda EXTRANJERA ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
                oTextFile.PrintNewLine
                nTotSubAcc = 0
                nLine = nLine + 3
              End If
              If (nSect <> !sector) Or (sOnlyAccount <> Mid(!cuenta, 1, 4)) Then
                oTextFile.PrintLine Space(100) & "======================================="
                oTextFile.PrintLine "Total Sector " & Space(86) & SpcFillAmount(nTotAcc, 40)
                oTextFile.PrintNewLine
                nTotSubAcc = 0
                nTotAcc = 0
                If (sOnlyAccount <> Mid(!cuenta, 1, 4)) Then
                Else
                   oTextFile.PageFeed nLine
                   nLine = 100
                End If
              End If
            End If
          Else
            .MoveNext
          End If
        Loop
          oTextFile.PrintLine Space(100) & "======================================="
          If nCurr = 1 Then
            oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda NACIONAL ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
          Else
            oTextFile.PrintLine "Total " & SpcFill("Subcuenta " & sSubAcc & " Moneda EXTRANJERA ", 50) & Space(53) & SpcFillAmount(nTotSubAcc, 30)
          End If
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(100) & "======================================="
          oTextFile.PrintLine "Total Sector " & Space(86) & SpcFillAmount(nTotAcc, 40)
          oTextFile.PrintNewLine
          PrintQtyLines oTextFile, 139
          oTextFile.PrintLine SpcFill(" TOTAL DE LA CUENTA:  " & sOnlyAccount & "  " & sName, 55) & Space(59) & _
                              SpcFillAmount(nAccount, 25)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
      .Close
    End With
    Set oRsBalance = Nothing
  Report59 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report59", Err.Number
End Function

Public Function Report60(sAppServer As String, dCuteDate As Date, _
                         vAccounts As Variant)
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  
  Dim sAccounts     As String
  Dim sDateTime     As String
  Dim SA            As Currency
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nCurrencyId   As Long
  Dim sCurrency     As String
  Dim nSumCurrency  As Currency
  '********************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    nLine = 100
    nPage = 1
    If (IsArray(vAccounts)) Then
      sAccounts = Join(vAccounts, "','")
    Else
      sAccounts = vAccounts
    End If
    Set oRs = oRSGenerator.GeneraRecordset60(sAppServer, dCuteDate, sAccounts)
    With oRs
      If Not (.BOF And .EOF) Then
        nCurrencyId = !ID_MONEDA
        Do While Not .EOF()
          If !naturaleza = "D" Then
            SA = !SI + !Cargos - !Abonos
          ElseIf !naturaleza = "A" Then
            SA = !SI - !Cargos + !Abonos
          End If
          If nCurrencyId <> !ID_MONEDA Then
            'Total Moneda
            PrintHeader60 oTextFile, nLine, nPage, dCuteDate, sDateTime, sAccounts
            sCurrency = oRSGenerator.GetObjName(sAppServer, nCurrencyId, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            oTextFile.PrintLine "                    -----------------------------"
            oTextFile.PrintLine " Total " & SpcFill(Mid(sCurrency, 6), 17) & " " & SpcFillAmount(nSumCurrency, 22)
            oTextFile.PrintNewLine
            nLine = nLine + 3
            PrintHeader60 oTextFile, nLine, nPage, dCuteDate, sDateTime, sAccounts
            If (nLine > 13) Then
              Printlines60 oTextFile
            ElseIf (nLine = 1) Then
              'No pinta nada
            End If
            nLine = nLine + 2
            nSumCurrency = 0
          End If
          If (SA <> 0) Then
            PrintHeader60 oTextFile, nLine, nPage, dCuteDate, sDateTime, sAccounts
            oTextFile.PrintLine "  " & _
                              SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & " " & _
                              SpcFillAmount(SA, 22) & "  " & _
                              SpcFill(!nombre_cuenta_auxiliar, 35)
            nLine = nLine + 1
          ElseIf (SA > 0) Then
            'siguiente
          End If
          nCurrencyId = !ID_MONEDA
          nSumCurrency = nSumCurrency + SA
          .MoveNext
        Loop
        sCurrency = oRSGenerator.GetObjName(sAppServer, nCurrencyId, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        oTextFile.PrintLine "                    -----------------------------"
        oTextFile.PrintLine " Total " & SpcFill(Mid(sCurrency, 6), 17) & " " & SpcFillAmount(nSumCurrency, 22)
      End If
    End With
    Report60 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report60", Err.Number
End Function

Public Function Report61(sAppServer As String, nGl As Long, _
                         dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, _
                         Optional nExcRateCurrency As Long = 1) As String
  Dim oRsBalance    As Recordset
  Dim oRsAccNames   As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim sLongDate     As String
  Dim sGLDesc       As String
  
  Dim nTotSubAcc    As Currency
  Dim nTotAcc       As Currency
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nCurr         As Long
  Dim nSect         As Long
  Dim nExcRate      As Currency
  
  Dim sDateTime     As String
  Dim sFileName     As String
  Dim skindacc      As String
  Dim sNat          As String
  Dim sAcc          As String
  Dim sAccOr        As String
  Dim sSubAcc       As String
  Dim nDisFor       As Long
  Dim nDisBac       As Long
  Dim sName         As String
  Dim bChangeAcc    As Boolean
  Dim nAccType      As Long
  Dim sAccType      As String
  '*********************************************************************************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sLongDate = Format(dEndDate, "d") & " de " & Format(dEndDate, "MMMM") & " de " & Year(dEndDate)
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
   
    nLine = 100
    nPage = 1
     
    Set oRsBalance = oRSGenerator.GeneraRecordset61(sAppServer, nGl, dEndDate, dExcDate, nExcRateType, nExcRateCurrency)
    Set oRsAccNames = oRSGenerator.GeneraRSNames(sAppServer, nGl)
    sGLDesc = Space(2) & oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", _
                                                  "ID_MAYOR", "NOMBRE_MAYOR")
    With oRsBalance
      If Not (.BOF And .EOF) Then
      Do While Not .EOF()
        sNat = !nat
        If sNat = "D" Then
          skindacc = "DEUDORA"
        Else
          skindacc = "ACREEDORA"
        End If
        If (!st <> 0) Then
          nCurr = !ID_MONEDA
          PrintHeader61 oTextFile, nLine, nPage, sLongDate, sGLDesc, skindacc, sDateTime, sAccType
    
          sAcc = Mid(!cuenta, 1, 4)
          sName = GETAccName(oRsAccNames, sAcc, nGl)
          oTextFile.PrintLine SpcFill(sAcc, 4) & Space(14) & SpcFill(sName, 30) & "  " & _
                               SpcFill(!ID_MONEDA & " " & !nom_moneda, 35)
          nLine = nLine + 1
          sSubAcc = Mid(!cuenta, 6, 2)
          nDisFor = 2
          nDisBac = 14
          If (sSubAcc <> "") Then
            sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 7), nGl)
            oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & SpcFill(sName, 30)
            nLine = nLine + 1
              sSubAcc = Mid(!cuenta, 9, 2)
              If (sSubAcc <> "") Then
                nDisFor = nDisFor + 2
                nDisBac = nDisBac - 2
                sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 10), nGl)
                oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & SpcFill(sName, 30)
                nLine = nLine + 1
                  sSubAcc = Mid(!cuenta, 12, 2)
                  If (sSubAcc <> "") Then
                    nDisFor = nDisFor + 2
                    nDisBac = nDisBac - 2
                    sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 13), nGl)
                    oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & " " & SpcFill(sName, 30)
                    nLine = nLine + 1
                      sSubAcc = Mid(!cuenta, 15, 2)
                      If (sSubAcc <> "") Then
                        nDisFor = nDisFor + 2
                        nDisBac = nDisBac - 2
                        sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 16), nGl)
                        oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & SpcFill(sName, 30)
                        nLine = nLine + 1
                          sSubAcc = Mid(!cuenta, 18, 2)
                          If (sSubAcc <> "") Then
                            nDisFor = nDisFor + 2
                            nDisBac = nDisBac - 2
                            sName = GETAccName(oRsAccNames, Mid(!cuenta, 1, 19), nGl)
                            oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & SpcFill(sName, 30)
                            nLine = nLine + 1
                              sSubAcc = Mid(!cuenta, 21, 2)
                              If (sSubAcc <> "") Then
                                nDisFor = nDisFor + 2
                                nDisBac = nDisBac - 2
                                sName = GETAccName(oRsAccNames, (!cuenta), nGl)
                                oTextFile.PrintLine Space(nDisFor) & SpcFill(sSubAcc, 2) & Space(nDisBac) & SpcFill(sName, 30)
                                nLine = nLine + 1
                              End If
                          End If
                      End If
                  End If
              End If
          End If
          nDisFor = nDisFor + 2
          nDisBac = nDisBac - 2
          oTextFile.PrintLine Space(nDisFor) & SpcFill(ZeroFill(!cve_sector, 2), 2) & Space(nDisBac) & SpcFill(!nom_sector, 30)
          oTextFile.PrintLine Space(90) & SpcFillAmount(!st, 35)
          nLine = nLine + 2
          nTotSubAcc = !st
          nTotAcc = nTotAcc + !st
          sAccOr = !cuenta
          nCurr = !ID_MONEDA
          nSect = !sector
        .MoveNext
          If Not (.EOF) Then
            Do
              oTextFile.PrintLine Space(12) & SpcFill(!num_aux, 17) & " " & _
                                  SpcFill(ZeroFill(!cve_sector, 2), 2) & " " & _
                                  SpcFill(!nom_aux, 30) & " " & _
                                  SpcFillAmount(!st, 35)
                      nLine = nLine + 1
                      nExcRate = 0
                      nExcRate = !exc_rate
              .MoveNext
              If Not (.EOF) Then
                If (sAccOr <> !cuenta) Or (nCurr <> !ID_MONEDA) Or (nSect <> !sector) Then
                  oTextFile.PrintLine Space(64) & "===================================="
                  oTextFile.PrintLine Space(14) & "TOTAL POR SUBCUENTA " & Space(30) & SpcFillAmount(nTotSubAcc, 35)
                  oTextFile.PrintLine Space(64) & "===================================="
                  nLine = nLine + 3
                  bChangeAcc = True
                  Exit Do
                End If
              Else
                Exit Do
              End If
            Loop
          End If
          
          If (.EOF) Then
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "TOTAL POR CUENTA " & Space(39) & SpcFillAmount(nTotAcc, 35)
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "CONVERSION MONEDA NACIONAL " & Space(10) & SpcFillAmount(nExcRate, 10, 5) & Space(8) & SpcFillAmount((nTotAcc * nExcRate), 35)
            oTextFile.PrintLine Space(100) & "========================="
            nTotAcc = 0
          ElseIf (nCurr <> !ID_MONEDA) And Not (.EOF) Then
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "TOTAL POR CUENTA " & Space(39) & SpcFillAmount(nTotAcc, 35)
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "CONVERSION MONEDA NACIONAL " & Space(10) & SpcFillAmount(nExcRate, 10, 5) & Space(8) & SpcFillAmount((nTotAcc * nExcRate), 35)
            oTextFile.PrintLine Space(100) & "========================="
            nLine = nLine + 5
            nTotAcc = 0
          ElseIf (sAcc <> Mid(!cuenta, 1, 4)) And Not (.EOF) Then
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "TOTAL POR CUENTA " & Space(39) & SpcFillAmount(nTotAcc, 35)
            oTextFile.PrintLine Space(100) & "========================="
            oTextFile.PrintLine Space(34) & "CONVERSION MONEDA NACIONAL " & Space(10) & SpcFillAmount(nExcRate, 10, 5) & Space(8) & SpcFillAmount((nTotAcc * nExcRate), 35)
            oTextFile.PrintLine Space(100) & "========================="
            nLine = nLine + 5
            nTotAcc = 0
          End If
          If bChangeAcc Then
            oTextFile.PageFeed nLine
            nLine = 100
          End If
        Else
          .MoveNext
        End If
      Loop
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
      .Close
    End With
    Set oRsBalance = Nothing
  Report61 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report61", Err.Number
End Function

Public Function Report62(sAppServer As String, sAccounts As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  
  Dim sDateTime     As String
  Dim sGl           As String
  Dim sAcc          As String
  Dim sPoliza       As String
  Dim dFecAfe       As Date
  Dim sMoneda       As String
  Dim bPrintAcc     As Boolean
  
  Dim tPoliza       As TMoney
  Dim tFecha        As TMoney
  Dim tMoneda       As TMoney
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  
  Dim tAccount      As TMoney
  '*********************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset62(sAccounts, sAppServer, vGL, dIniDate, dEndDate)
    sDateTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = " TODAS LAS CONTABILIDADES "
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sAcc = !cuenta & "  " & !clave_sector
        sPoliza = !POLIZA
        dFecAfe = !FEC_AFE
        sMoneda = !MONEDA
        bPrintAcc = True
        Do While Not (.EOF)
          PrintHeader62 oTextFile, nLine, nPage, dIniDate, dEndDate, sGl, dFecAfe, sDateTime, bPrintAcc, sAccType
          
          If (sPoliza <> !POLIZA) Or (sAcc <> (!cuenta & "  " & !clave_sector)) Or (!FEC_AFE <> dFecAfe) Then
            PrintTotPoliza oTextFile, nLine, bPrintAcc, tPoliza, tFecha, tMoneda, sPoliza, !POLIZA, _
                           sAcc, (!cuenta & "  " & !clave_sector), "    TOTAL POR POLIZA "
          End If
          If (!FEC_AFE <> dFecAfe) Or (sAcc <> (!cuenta & "  " & !clave_sector)) Then
            PrintTotFecha oTextFile, nLine, bPrintAcc, tFecha, dFecAfe, !FEC_AFE, "    TOTAL POR FECHA "
          End If
          
          If (sAcc <> (!cuenta & "  " & !clave_sector)) Then
            nLine = nLine + 5
            bPrintAcc = True
            oTextFile.PrintLine Space(100) & "=================================================================================="
            oTextFile.PrintLine SpcFill("   TOTAL POR CUENTA ", 132) & SpcFillAmount(tAccount.Cargos, 22) & "  " & SpcFillAmount(tAccount.Abonos, 22)
            sAcc = !cuenta & "  " & !clave_sector
            oTextFile.PrintNewLine
            ClearMoney tAccount
          End If
          
          If (!MONEDA <> sMoneda) Then
            PrintTotMoneda oTextFile, nLine, bPrintAcc, tMoneda, sMoneda, !MONEDA, "    TOTAL POR MONEDA "
          End If
          
          PrintDetalle62 oTextFile, bPrintAcc, tPoliza, _
                        !cuenta, !clave_sector, !POLIZA, !MONEDA, !CONCEPTO, !Aux, !tipo_movimiento, !monto
          
          If (!tipo_movimiento = "D") Then
            AddMoney tAccount, 0, !monto, 0, 0
          ElseIf (!tipo_movimiento = "H") Then
            AddMoney tAccount, 0, 0, !monto, 0
          Else
            AddMoney tAccount, 0, 0, 0, 0
          End If
          
          .MoveNext
        Loop
        .Close
        PrintTotPoliza oTextFile, nLine, bPrintAcc, tPoliza, tFecha, tMoneda, sPoliza, " ", sAcc, " ", "    TOTAL POR POLIZA "
        PrintTotFecha oTextFile, nLine, bPrintAcc, tFecha, "01/01/1990 ", "01/01/1990 ", "    TOTAL POR FECHA "
        
        oTextFile.PrintLine Space(100) & "=================================================================================="
        oTextFile.PrintLine SpcFill("   TOTAL POR CUENTA ", 132) & SpcFillAmount(tAccount.Cargos, 22) & "  " & SpcFillAmount(tAccount.Abonos, 22)
       
        PrintTotMoneda oTextFile, nLine, bPrintAcc, tMoneda, "", "", "    TOTAL POR MONEDA "
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report62 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report62", Err.Number
End Function


Public Function Report64(sAppServer As String, vGL As Variant, _
                          dIniDate As Date, dEndDate As Date, _
                          Optional sIniAccount As String, Optional sEndAccount As String, _
                          Optional bPrintInCascade As Boolean = False) As String
  Dim oRs           As Object
  Dim oRSGenerator  As SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sAccount      As String
  Dim sCurrencyId   As String
  Dim sCurrency     As String
  
  Dim tAcumulador   As TMoney
  Dim tCuenta       As TMoney
  Dim tMoneda       As TMoney
  
  Dim nBalance      As Currency
  Dim sGl           As String
  
  
  Dim bPrintSI      As Boolean
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  
  Dim vGLToSend     As Variant
  Dim vGL2          As Variant
  Dim i             As Long
  Dim sAccName      As String
  Dim sNature       As String
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sEndAccount = sEndAccount & "-99-99-99-99-99-99"
    vGL2 = vGL
    If Not IsArray(vGL2) Then
      vGL2 = Array(vGL2)
    End If
    
    nPage = 1
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        nGl = CLng(vGLToSend)
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      Else
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          nGl = vGLToSend(0)
          sGl = "TODAS LAS CONTABILIDADES"
        Else
          nGl = CLng(vGLToSend)
          sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        End If
      End If

      nAccType = oRSGenerator.AccountType(sAppServer, nGl)
      sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

        Set oRs = CreateObject("ADODB.Recordset")
        Set oRs = oRSGenerator.GeneraRecordset64(sAppServer, vGLToSend, dIniDate, dEndDate, sIniAccount, sEndAccount)
        sDTime = GetFmtDate
        oTextFile.PrintNewLine
        
        nLine = 100
        With oRs
          If Not (.BOF And .EOF) Then
            sCurrencyId = !ID_MONEDA
            sAccount = !cuenta
            sAccName = !nombre_cuenta
            sNature = !naturaleza
            nBalance = !SI
            bPrintSI = True
            sCurrency = !o_id_moneda & "  " & !nombre_moneda
            PrintHeader64 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                             sNature, sAccType
            oTextFile.PrintLine " Cuenta : " & sAccount & " " & sAccName
            Do While Not (.EOF)
              PrintHeader64 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                             sNature, sAccType
                                
              If ((sAccount <> Mid(!cuenta, 1, 4)) Or (sCurrencyId <> !ID_MONEDA) Or (sNature <> !naturaleza)) Then
                AddMoney tAcumulador, 0, 0, 0, nBalance
                nBalance = !SI
                bPrintSI = True
                PrintTot oTextFile, nLine, tAcumulador, tCuenta, sAccount, !cuenta, "   TOTAL POR CUENTA "
                If ((sNature <> !naturaleza) And (sCurrencyId = !ID_MONEDA)) Then
                  oTextFile.PageFeed (nLine)
                  nLine = 90
                  sNature = !naturaleza
                  PrintHeader64 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                                sNature, sAccType
                Else
                  oTextFile.PrintNewLine
                End If
                If (sCurrencyId <> !ID_MONEDA) Then
                  PrintTot oTextFile, nLine, tCuenta, tMoneda, sCurrencyId, !ID_MONEDA, "   TOTAL POR MONEDA ", False
                  'AddMoney tTotal, tMoneda.SI, tMoneda.Cargos, tMoneda.Abonos, tMoneda.SA
                  ClearMoney tMoneda
                  sCurrency = !o_id_moneda & "  " & !nombre_moneda
                  sNature = !naturaleza
                  oTextFile.PageFeed (nLine)
                  nLine = 90
                  PrintHeader64 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                                sNature, sAccType
                End If
                sAccName = !nombre_cuenta
                oTextFile.PrintLine " Cuenta : " & sAccount & " " & sAccName
              End If
              
              If (bPrintSI) Then
                oTextFile.PrintLine SpcFill(" ", 125) & " SALDO INICIAL: " & SpcFillAmount(nBalance, 22)
                bPrintSI = False
              ElseIf Not (bPrintSI) Then
               '
              End If
              
              PrintDetalle64 oTextFile, !naturaleza, tAcumulador, !fecha_ult, IIf(IsNull(!concepto_movimiento), "", !concepto_movimiento), nBalance, !Cargos, !Abonos
              
              .MoveNext
            Loop
            AddMoney tAcumulador, 0, 0, 0, nBalance
            PrintTot oTextFile, nLine, tAcumulador, tCuenta, sAccount, "", "   TOTAL POR CUENTA "
            
            PrintTot oTextFile, nLine, tCuenta, tMoneda, sCurrencyId, "", "    TOTAL POR MONEDA", False
    
            .Close
          Else
            'oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
          End If
        End With
        If bPrintInCascade Then
          oTextFile.PageFeed nLine
          nBalance = 0
          ClearMoney tAcumulador
          ClearMoney tCuenta
          ClearMoney tMoneda
        Else
          Exit For
        End If
    Next i
    Report64 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report64", Err.Number
End Function

Public Function Report65(sAppServer As String, ByVal sPatron As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1, _
                         Optional bPrintInCascade As Boolean = False, Optional bPrintTotal As Boolean = True) As String
  Dim oRs            As Object
  Dim oRSGenerator   As SCFFixedReportsDS.CServer
  Dim oTextFile      As New CTextFile
  Dim dFecha         As Date
  Dim sAccount       As String
  Dim sSector        As String
  Dim nSaldoMN       As Double
  Dim nSaldoME       As Double
  Dim nSaldoMNA      As Double
  Dim nSaldoMEA      As Double
  Dim sAccName       As String
  Dim tAcumulador    As TMoney
  Dim tDeudoras      As TMoney
  Dim tAcreedoras    As TMoney
  
  Dim nLine          As Long
  Dim nPage          As Long
  Dim sNature        As String
  Dim sNat           As String
  Dim sDateTime      As String
  
  Dim sGl            As String
  Dim nGl            As Long
  Dim nAccType       As Long
  Dim sAccType       As String
  
  Dim vGL2           As Variant
  Dim vGLToSend      As Variant
  Dim i              As Long
  '**************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    vGL2 = vGL
    If Not IsArray(vGL2) Then
      vGL2 = Array(vGL2)
    End If
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        nGl = CLng(vGLToSend)
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      Else
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          nGl = vGLToSend(0)
        Else
          nGl = vGLToSend
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(nGl), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        End If
      End If
        nAccType = oRSGenerator.AccountType(sAppServer, nGl)
        sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
        sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
        sPatron = "&&&&"
        
        Set oRs = CreateObject("ADODB.Recordset")
        Set oRs = oRSGenerator.GeneraRecordset58(sPatron, sAppServer, vGLToSend, dIniDate, dEndDate, dExcDate, _
                                                 nExcRateType, nExcRateCurrency, nReport:=65)
        With oRs
          nPage = 1
          nLine = 100
          dFecha = !fecha_rev
          sAccount = !NUMERO_CUENTA_ESTANDAR
          sAccName = !nombre_cuenta
          sSector = !clave_sector
          nSaldoMN = 0
          nSaldoME = 0
          nSaldoMNA = 0
          nSaldoMEA = 0
          Do While Not .EOF
            If (!naturaleza = "D") Then
              sNature = "1"
              sNat = "1"
            Else
              sNature = "2"
            End If
            
            If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
              If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPatron)) Then
                PrintHeader65 oTextFile, nLine, nPage, sGl, "", "", dIniDate, dEndDate, sDateTime, sAccType
                nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1)))
                nSaldoMEA = (nSaldoMEA + (nSaldoME * ((dEndDate - dFecha) + 1)))
                oTextFile.PrintLine SpcFill(sAccount, 6) & " " & SpcFill(sAccName, 40) & " " & _
                                    SpcFillAmount(nSaldoMNA, 27) & "  " & _
                                    SpcFillAmount(nSaldoMEA, 27) & "  " & _
                                    SpcFillAmount(nSaldoMNA + nSaldoMEA, 27) & " " & _
                                    SpcFillAmount((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1), 27)
                nLine = nLine + 1
                
                If (Len(sAccount) = 4) Then
                  AddMoney tAcumulador, CCur(nSaldoMNA), CCur(nSaldoMEA), CCur(nSaldoMNA + nSaldoMEA), CCur((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1))
                  If (!naturaleza = "D") Then
                    AddMoney tDeudoras, CCur(nSaldoMNA), CCur(nSaldoMEA), CCur(nSaldoMNA + nSaldoMEA), CCur((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1))
                  ElseIf (!naturaleza = "A") Then
                    AddMoney tAcreedoras, CCur(nSaldoMNA), CCur(nSaldoMEA), CCur(nSaldoMNA + nSaldoMEA), CCur((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1))
                  End If
                End If
                
              ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPatron)) Then
                'Siguiente
              End If
              If (Left(sAccount, 2) <> Left(!NUMERO_CUENTA_ESTANDAR, 2)) Then
                If bPrintTotal Then
                  oTextFile.PrintLine Space(46) & "========================================================================================================================="
                  oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(24) & _
                                      SpcFillAmount(tAcumulador.SI, 27) & "  " & _
                                     SpcFillAmount(tAcumulador.Cargos, 27) & "  " & _
                                     SpcFillAmount(tAcumulador.Abonos, 27) & " " & _
                                     SpcFillAmount(tAcumulador.SA, 27)
                  oTextFile.PrintNewLine
                  nLine = nLine + 3
                End If
                  ClearMoney tAcumulador
              End If
              
              If (sNature <> sNat) Then
                If bPrintTotal Then
                  oTextFile.PrintLine Space(46) & "========================================================================================================================="
                  oTextFile.PrintLine "    TOTAL DEUDORAS      " & Space(24) & _
                                     SpcFillAmount(tDeudoras.SI, 27) & "  " & SpcFillAmount(tDeudoras.Cargos, 27) & "  " & _
                                     SpcFillAmount(tDeudoras.Abonos, 27) & " " & SpcFillAmount(tDeudoras.SA, 27)
                  oTextFile.PrintNewLine
                  nLine = nLine + 4
                End If
                  sNat = sNature
              End If
              
              nSaldoMN = 0
              nSaldoME = 0
              nSaldoMNA = 0
              nSaldoMEA = 0
              sAccount = !NUMERO_CUENTA_ESTANDAR
              sAccName = !nombre_cuenta
              sSector = !clave_sector
              dFecha = !fecha_rev
            End If
            
            If (dFecha = #1/1/1990#) Then
               nSaldoMN = !saldo_inicial_mn
               nSaldoME = !Saldo_Inicial_me
               nSaldoMNA = 0
               nSaldoMEA = 0
               dFecha = dIniDate
            ElseIf (dFecha <> #1/1/1990#) Then
              nSaldoMNA = nSaldoMNA + (nSaldoMN * (!fecha_rev - dFecha))
              nSaldoMEA = nSaldoMEA + (nSaldoME * (!fecha_rev - dFecha))
              If (!naturaleza = "D") Then
                nSaldoMN = nSaldoMN + !cargos_mn - !abonos_mn
                nSaldoME = nSaldoME + !cargos_me - !abonos_me
              ElseIf (!naturaleza = "A") Then
                nSaldoMN = nSaldoMN - !cargos_mn + !abonos_mn
                nSaldoME = nSaldoME - !cargos_me + !abonos_me
              End If
              dFecha = !fecha_rev
            End If
            .MoveNext
          Loop
          'De aquí hasta el otro letrero JMCOTA el 27/Ene/2002
          PrintHeader65 oTextFile, nLine, nPage, sGl, "", "", dIniDate, dEndDate, sDateTime, sAccType
          nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1)))
          nSaldoMEA = (nSaldoMEA + (nSaldoME * ((dEndDate - dFecha) + 1)))
          oTextFile.PrintLine SpcFill(sAccount, 6) & " " & SpcFill(sAccName, 40) & " " & _
                              SpcFillAmount(nSaldoMNA, 27) & "  " & _
                              SpcFillAmount(nSaldoMEA, 27) & "  " & _
                              SpcFillAmount(nSaldoMNA + nSaldoMEA, 27) & " " & _
                              SpcFillAmount((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1), 27)
          nLine = nLine + 1
          
          If (Len(sAccount) = 4) Then
            AddMoney tAcumulador, CCur(nSaldoMNA), CCur(nSaldoMEA), CCur(nSaldoMNA + nSaldoMEA), CCur((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1))
            AddMoney tAcreedoras, CCur(nSaldoMNA), CCur(nSaldoMEA), CCur(nSaldoMNA + nSaldoMEA), CCur((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1))
          End If
          'Fin del cambio del 27/Ene/2002 JMCOTA
          
          If bPrintTotal Then
            oTextFile.PrintLine Space(46) & "========================================================================================================================="
            oTextFile.PrintLine "    TOTAL DEL GRUPO " & Left(sAccount, 2) & "00" & Space(24) & _
                                SpcFillAmount(tAcumulador.SI, 27) & "  " & _
                                SpcFillAmount(tAcumulador.Cargos, 27) & "  " & _
                                SpcFillAmount(tAcumulador.Abonos, 27) & " " & _
                                SpcFillAmount(tAcumulador.SA, 27)
            
            oTextFile.PrintNewLine
            oTextFile.PrintLine Space(46) & "========================================================================================================================="
            oTextFile.PrintLine "    TOTAL ACREEDORAS    " & Space(24) & _
                                SpcFillAmount(tAcreedoras.SI, 27) & "  " & SpcFillAmount(tAcreedoras.Cargos, 27) & "  " & _
                                SpcFillAmount(tAcreedoras.Abonos, 27) & "  " & SpcFillAmount(tAcreedoras.SA, 27)
            
            oTextFile.PrintNewLine
            oTextFile.PrintLine Space(46) & "========================================================================================================================="
            oTextFile.PrintLine "    TOTAL DEL REPORTE   " & Space(24) & _
                                    SpcFillAmount((tDeudoras.SI - tAcreedoras.SI), 27) & "  " & _
                                    SpcFillAmount((tDeudoras.Cargos - tAcreedoras.Cargos), 27) & "  " & _
                                    SpcFillAmount((tDeudoras.Abonos - tAcreedoras.Abonos), 27) & "  " & _
                                    SpcFillAmount((tDeudoras.SA - tAcreedoras.SA), 27)
          End If
        End With
      If bPrintInCascade Then
        oTextFile.PageFeed nLine
        nSaldoMN = 0
        nSaldoME = 0
        nSaldoMNA = 0
        nSaldoMEA = 0
        ClearMoney tAcumulador
        ClearMoney tDeudoras
        ClearMoney tAcreedoras
      Else
        Exit For
      End If
    Next i
    Report65 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report65", Err.Number
End Function

Public Function Report68(sAppServer As String, vGL As Variant, dEndDate As Date) As String
  Dim oRsBalance    As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim sGLname       As String
  Dim sGLnum        As String
  
  Dim nTotal        As Currency
  
  Dim nLine         As Long
  Dim nPage         As Long
  Dim nCurr         As Long
  
  Dim sDateTime     As String
  Dim sFileName     As String
  Dim skindacc      As String
  Dim sNat          As String
  Dim nTypeStdAcc   As Long
  '*************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  
    nLine = 100
    nPage = 1
    
    If (IsArray(vGL)) Then
      nTypeStdAcc = oRSGenerator.TypeStdAccount(sAppServer, CLng(vGL(0)))
    Else
      nTypeStdAcc = oRSGenerator.TypeStdAccount(sAppServer, CLng(vGL))
    End If
    
    Set oRsBalance = oRSGenerator.GeneraRecordset68(sAppServer, vGL, dEndDate)
     
    With oRsBalance
      If Not (.BOF And .EOF) Then
        nCurr = !ID_MONEDA
        Do While Not .EOF()
          If !st <> 0 Then
              PrintHeader68 oTextFile, nLine, nPage, dEndDate, "", (!ID_MONEDA & "    " & !nom_moneda), sDateTime, nTypeStdAcc
      
              oTextFile.PrintLine Space(5) & SpcFill(ZeroFill(!num_mayor, 4), 4) & Space(4) & SpcFill(GetCuentaSubCuenta(!cuenta) & ZeroFill(!sector, 2), 20) & "  " & _
                                  SpcFill(!nombre, 35) & "  " & SpcFillAmount(!st, 30)
              nLine = nLine + 1
              nTotal = nTotal + !st
          End If
          .MoveNext
          If Not .EOF Then
            If (!ID_MONEDA <> nCurr) Then
              oTextFile.PrintLine "                                                                       ================================"
              oTextFile.PrintLine "                                                     TOTAL POR MONEDA: " & SpcFillAmount(nTotal, 31)
              nCurr = !ID_MONEDA
              nLine = nLine + 2
              oTextFile.PageFeed nLine
              'nLine = nLine + 2
              nTotal = 0
            End If
          End If
        Loop
        oTextFile.PrintLine "                                                                       ================================"
        oTextFile.PrintLine "                                                     TOTAL POR MONEDA: " & SpcFillAmount(nTotal, 31)
        .Close
      End If
    End With
    Set oRsBalance = Nothing
    Report68 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report68", Err.Number
End Function

Public Function Report69(sAppServer As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         vAccountGroups As Variant, Optional vAccountGroupName As Variant) As String
                                   
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tbalance      As TMoney
  Dim tTotal        As TMoney
  Dim sGl           As String
  Dim sGLNumber     As String
  Dim sAccType      As String
  Dim sGroup        As String
  
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
    Else
      sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
      sReportTitle = "RELACIÓN MENSUAL DE HONORARIOS FIDUCIARIOS"
      Set oRs = oRSGenerator.GeneraRecordset69(sAppServer, vGL, dIniDate, dEndDate, vAccountGroups)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sGroup = !order_id_grupo_mayor
          Do While Not (.EOF)
            PrintHeader69 oTextFile, nLine, nPage, dIniDate, dEndDate, _
                                       vAccountGroupName, sReportTitle, sDTime
                                        
            If (sGroup <> !order_id_grupo_mayor) Then
              oTextFile.PrintLine Space(67) & _
                                  "======================" & "      " & "======================" & "      " & _
                                  "======================" & "      " & "======================"
              If (sGroup = 1) Then
                oTextFile.PrintLine Space(11) & "  " & SpcFill(" TOTAL FIDEICOMISO ", 52) & "  " & _
                                  SpcFillAmount(tbalance.SI, 22) & "     " & SpcFillAmount(tbalance.Cargos, 22) & "     " & _
                                  SpcFillAmount(tbalance.Abonos, 22) & "     " & _
                                  SpcFillAmount(tbalance.SA, 22)
                                  
              Else
                oTextFile.PrintLine Space(11) & "  " & SpcFill(" TOTAL MANDATOS    ", 52) & "  " & _
                                  SpcFillAmount(tbalance.SI, 22) & "     " & SpcFillAmount(tbalance.Cargos, 22) & "     " & _
                                  SpcFillAmount(tbalance.Abonos, 22) & "     " & _
                                  SpcFillAmount(tbalance.SA, 22)
                                  
              End If
              oTextFile.PrintNewLine
              ClearMoney tbalance
              sGroup = !order_id_grupo_mayor
            End If
            oTextFile.PrintLine SpcFill(!NUMERO_MAYOR, 6) & "       " & SpcFill(!NOMBRE_MAYOR, 52) & "  " & _
                                SpcFillAmount(!Abonos, 22) & "     " & SpcFillAmount(!CARGOS_2, 22) & "     " & _
                                SpcFillAmount(!SALDO_ACTUAL_2, 22) & "     " & _
                                SpcFillAmount(!SALDO_ACTUAL_3, 22)
            nLine = nLine + 1
            AddMoney tbalance, !Abonos, !CARGOS_2, !SALDO_ACTUAL_2, !SALDO_ACTUAL_3
            AddMoney tTotal, !Abonos, !CARGOS_2, !SALDO_ACTUAL_2, !SALDO_ACTUAL_3
            .MoveNext
          Loop
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(67) & _
                            "======================" & "      " & "======================" & "      " & _
                            "======================" & "      " & "======================"
                                
          If (sGroup = 1) Then
            oTextFile.PrintLine Space(11) & "  " & SpcFill(" TOTAL FIDEICOMISO ", 52) & "  " & _
                              SpcFillAmount(tbalance.SI, 22) & "     " & SpcFillAmount(tbalance.Cargos, 22) & "     " & _
                              SpcFillAmount(tbalance.Abonos, 22) & "     " & _
                              SpcFillAmount(tbalance.SA, 22)
                              
          Else
            oTextFile.PrintLine Space(11) & "  " & SpcFill(" TOTAL MANDATO     ", 52) & "  " & _
                              SpcFillAmount(tbalance.SI, 22) & "     " & SpcFillAmount(tbalance.Cargos, 22) & "     " & _
                              SpcFillAmount(tbalance.Abonos, 22) & "     " & _
                              SpcFillAmount(tbalance.SA, 22)
                              
          End If
          
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(67) & _
                            "======================" & "      " & "======================" & "      " & _
                            "======================" & "      " & "======================"
                                
          oTextFile.PrintLine Space(11) & "  " & SpcFill(" TOTAL GENERAL ", 52) & "  " & _
                            SpcFillAmount(tTotal.SI, 22) & "     " & SpcFillAmount(tTotal.Cargos, 22) & "     " & _
                            SpcFillAmount(tTotal.Abonos, 22) & "     " & _
                            SpcFillAmount(tTotal.SA, 22)

        Else
          oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
        End If
        .Close
        Set oRSGenerator = Nothing
      End With
    
    Report69 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report69", Err.Number
End Function

Public Function Report70(sAppServer As String, dCuteDate As Date) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nMayor        As String
  Dim sMayorName    As String
  Dim sNumeroMayor  As String
  Dim sAcc1101      As Currency
  Dim sAcc1103      As Currency
  Dim sAcc1202      As Currency
  Dim sAcc1204      As Currency
  Dim sAcc1206      As Currency
  Dim sAcc1101AM    As Currency
  Dim sAcc1103AM    As Currency
  Dim sAcc1202AM    As Currency
  Dim sAcc1204AM    As Currency
  Dim sAcc1206AM    As Currency
  Dim sAcc1101AC    As Currency
  Dim sAcc1103AC    As Currency
  Dim sAcc1202AC    As Currency
  Dim sAcc1204AC    As Currency
  Dim sAcc1206AC    As Currency
  Dim nTipoMayor    As Long
  Dim sTipo         As String
  Dim nID_moneda    As Long
  
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As Variant
  Dim sDateTime     As String
  Dim sCurrency     As String
  '*************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset70(sAppServer, dCuteDate)
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    nLine = 100
    nPage = 1
    sAcc1101 = 0
    sAcc1103 = 0
    sAcc1202 = 0
    sAcc1204 = 0
    sAcc1206 = 0
    sAcc1101AM = 0
    sAcc1103AM = 0
    sAcc1202AM = 0
    sAcc1204AM = 0
    sAcc1206AM = 0
    sAcc1101AC = 0
    sAcc1103AC = 0
    sAcc1202AC = 0
    sAcc1204AC = 0
    sAcc1206AC = 0
    With oRs
      If Not (.BOF And .EOF) Then
        nMayor = !id_mayor
        sMayorName = !NOMBRE_MAYOR
        nTipoMayor = !object
        nID_moneda = !ID_MONEDA
        sNumeroMayor = !NUMERO_MAYOR
        If !object = 1 Then
          sTipo = "TOTAL  FIDEICOMISO"
        ElseIf !object = 2 Then
          sTipo = "TOTAL  MANDATO"
        End If
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        Do While Not .EOF
          ' Corte para Fideicomiso o Mandato
          If nMayor <> !id_mayor Then
             nLines70 oTextFile, nLine, nPage, 1, sCurrency, dCuteDate, sDateTime
             oTextFile.PrintLine SpcFill(sNumeroMayor, 4) & " " & SpcFill(sMayorName, 40) & "  " & _
                                 SpcFillAmount(sAcc1101, 23) & " " & _
                                 SpcFillAmount(sAcc1202, 23) & " " & _
                                 SpcFillAmount(sAcc1204, 23) & " " & _
                                 SpcFillAmount(sAcc1103, 23) & " " & _
                                 SpcFillAmount(sAcc1206, 23)
             nLine = nLine + 1
             nMayor = !id_mayor
             sMayorName = !NOMBRE_MAYOR
             sNumeroMayor = !NUMERO_MAYOR
             sAcc1101AM = sAcc1101AM + sAcc1101
             sAcc1103AM = sAcc1103AM + sAcc1103
             sAcc1202AM = sAcc1202AM + sAcc1202
             sAcc1204AM = sAcc1204AM + sAcc1204
             sAcc1206AM = sAcc1206AM + sAcc1206
             sAcc1101 = 0
             sAcc1103 = 0
             sAcc1202 = 0
             sAcc1204 = 0
             sAcc1206 = 0
          End If
          Select Case !cuenta
              Case "1101"
                 sAcc1101 = !SA
              Case "1103"
                 sAcc1103 = !SA
              Case "1202"
                 sAcc1202 = !SA
              Case "1204"
                 sAcc1204 = !SA
              Case "1206"
                 sAcc1206 = !SA
          End Select
          If nTipoMayor <> !object Then
            oTextFile.PrintLine "                                             ------------------------------------------------------------------------------------------------------------------------------"
            oTextFile.PrintLine SpcFill(" ", 4) & " " & SpcFill(sTipo, 40) & "  " & _
                                SpcFillAmount(sAcc1101AM, 23) & " " & _
                                SpcFillAmount(sAcc1202AM, 23) & " " & _
                                SpcFillAmount(sAcc1204AM, 23) & " " & _
                                SpcFillAmount(sAcc1103AM, 23) & " " & _
                                SpcFillAmount(sAcc1206AM, 23)
            oTextFile.PrintLine "                                             =============================================================================================================================="
            oTextFile.PrintNewLine
            nLine = nLine + 3
            sAcc1101AC = sAcc1101AC + sAcc1101AM
            sAcc1103AC = sAcc1103AC + sAcc1103AM
            sAcc1202AC = sAcc1202AC + sAcc1202AM
            sAcc1204AC = sAcc1204AC + sAcc1204AM
            sAcc1206AC = sAcc1206AC + sAcc1206AM
            sAcc1101AM = 0
            sAcc1103AM = 0
            sAcc1202AM = 0
            sAcc1204AM = 0
            sAcc1206AM = 0
            nTipoMayor = !object
            If !object = 1 Then
              sTipo = "TOTAL  FIDEICOMISO"
            ElseIf !object = 2 Then
              sTipo = "TOTAL  MANDATO"
            End If
          End If
  '        ' Corte para Moneda
          If nID_moneda <> !ID_MONEDA Then
            oTextFile.PrintLine "                                             ------------------------------------------------------------------------------------------------------------------------------"
            oTextFile.PrintLine SpcFill(" ", 4) & " " & SpcFill("TOTAL  MONEDA", 40) & "  " & _
                                SpcFillAmount(sAcc1101AC, 23) & " " & _
                                SpcFillAmount(sAcc1202AC, 23) & " " & _
                                SpcFillAmount(sAcc1204AC, 23) & " " & _
                                SpcFillAmount(sAcc1103AC, 23) & " " & _
                                SpcFillAmount(sAcc1206AC, 23)
            oTextFile.PrintLine "                                             =============================================================================================================================="
            oTextFile.PrintNewLine
            nLine = nLine + 2
            sAcc1101AC = 0
            sAcc1103AC = 0
            sAcc1202AC = 0
            sAcc1204AC = 0
            sAcc1206AC = 0
            nID_moneda = !ID_MONEDA
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            PrintLines70 oTextFile, sCurrency
          End If
          .MoveNext
        Loop
        .MoveLast
        Select Case !cuenta
            Case "1101"
               sAcc1101 = !SA
            Case "1103"
               sAcc1103 = !SA
            Case "1202"
               sAcc1202 = !SA
            Case "1204"
               sAcc1204 = !SA
            Case "1206"
               sAcc1206 = !SA
        End Select
        nLines70 oTextFile, nLine, nPage, 1, sCurrency, dCuteDate, sDateTime
        oTextFile.PrintLine SpcFill(sNumeroMayor, 4) & " " & SpcFill(sMayorName, 40) & "  " & _
                            SpcFillAmount(sAcc1101, 23) & " " & _
                            SpcFillAmount(sAcc1202, 23) & " " & _
                            SpcFillAmount(sAcc1204, 23) & " " & _
                            SpcFillAmount(sAcc1103, 23) & " " & _
                            SpcFillAmount(sAcc1206, 23)
        nLine = nLine + 1
        nMayor = !id_mayor
        sMayorName = !NOMBRE_MAYOR
        sNumeroMayor = !NUMERO_MAYOR
        sAcc1101AM = sAcc1101AM + sAcc1101
        sAcc1103AM = sAcc1103AM + sAcc1103
        sAcc1202AM = sAcc1202AM + sAcc1202
        sAcc1204AM = sAcc1204AM + sAcc1204
        sAcc1206AM = sAcc1206AM + sAcc1206
        'Corte
        oTextFile.PrintLine "                                             =============================================================================================================================="
        oTextFile.PrintLine SpcFill(" ", 4) & " " & SpcFill(sTipo, 40) & "  " & _
                            SpcFillAmount(sAcc1101AM, 23) & " " & _
                            SpcFillAmount(sAcc1202AM, 23) & " " & _
                            SpcFillAmount(sAcc1204AM, 23) & " " & _
                            SpcFillAmount(sAcc1103AM, 23) & " " & _
                            SpcFillAmount(sAcc1206AM, 23)
        oTextFile.PrintNewLine
        sAcc1101AC = sAcc1101AC + sAcc1101AM
        sAcc1103AC = sAcc1103AC + sAcc1103AM
        sAcc1202AC = sAcc1202AC + sAcc1202AM
        sAcc1204AC = sAcc1204AC + sAcc1204AM
        sAcc1206AC = sAcc1206AC + sAcc1206AM
        oTextFile.PrintLine "                                             =============================================================================================================================="
        oTextFile.PrintLine SpcFill(" ", 4) & " " & SpcFill("TOTAL  MONEDA", 40) & "  " & _
                            SpcFillAmount(sAcc1101AC, 23) & " " & _
                            SpcFillAmount(sAcc1202AC, 23) & " " & _
                            SpcFillAmount(sAcc1204AC, 23) & " " & _
                            SpcFillAmount(sAcc1103AC, 23) & " " & _
                            SpcFillAmount(sAcc1206AC, 23)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report70 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report70", Err.Number
End Function

Public Function Report71(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date, _
                         dExcDate As Date, vAccountGroups As Variant, vAccountGroupName As Variant, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As Recordset, oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tbalance      As TMoney
  Dim tTotal        As TMoney
  Dim sGl           As String
  Dim sGLNumber     As String
  Dim sAccType      As String
  Dim sGroup        As String
  
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
    Else
      sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
      sReportTitle = "REPORTE ANUAL DE PATRIMONIOS DE FIDEICOMISOS Y MANDATOS"
      Set oRs = oRSGenerator.GeneraRecordset71(sAppServer, vGL, dIniDate, dEndDate, vAccountGroups, nExcRateType:=nExcRateType, _
                                                nExcRateCurrency:=nExcRateCurrency, dExcDate:=dExcDate)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sGroup = !id_grupo_mayor
          Do While Not (.EOF)
            PrintHeader71 oTextFile, nLine, nPage, dIniDate, dEndDate, _
                                       vAccountGroupName, sReportTitle, sDTime
            If (sGroup <> !id_grupo_mayor) Then
              oTextFile.PrintLine Space(50) & _
                                "======================" & "   " & "======================" & "   " & _
                                "======================" & "   " & "======================"
                                
              oTextFile.PrintLine Space(6) & "  " & SpcFill(" TOTAL FIDEICOMISO ", 40) & "  " & _
                                SpcFillAmount(tbalance.SI, 22) & "  " & SpcFillAmount(tbalance.Cargos, 22) & "  " & _
                                SpcFillAmount(tbalance.Abonos, 22) & "  " & _
                                SpcFillAmount((tbalance.SI + tbalance.Cargos + tbalance.Abonos), 22)
              oTextFile.PrintNewLine
              ClearMoney tbalance
              sGroup = !id_grupo_mayor
            End If
            oTextFile.PrintLine SpcFill(!NUMERO_MAYOR, 6) & "  " & SpcFill(!NOMBRE_MAYOR, 40) & "  " & _
                                SpcFillAmount(!Saldo_Actual, 22) & "  " & SpcFillAmount(!SALDO_ACTUAL_2, 22) & "  " & _
                                SpcFillAmount(!SALDO_ACTUAL_3, 22) & "  " & _
                                SpcFillAmount((!Saldo_Actual + !SALDO_ACTUAL_2 + !SALDO_ACTUAL_3), 22)
            nLine = nLine + 1
            AddMoney tbalance, !Saldo_Actual, !SALDO_ACTUAL_2, !SALDO_ACTUAL_3, 0
            AddMoney tTotal, !Saldo_Actual, !SALDO_ACTUAL_2, !SALDO_ACTUAL_3, 0
            .MoveNext
          Loop
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(50) & _
                            "======================" & "   " & "======================" & "   " & _
                            "======================" & "   " & "======================"
                                
          oTextFile.PrintLine Space(6) & "  " & SpcFill(" TOTAL MANDATO ", 40) & "  " & _
                            SpcFillAmount(tbalance.SI, 22) & "  " & SpcFillAmount(tbalance.Cargos, 22) & "  " & _
                            SpcFillAmount(tbalance.Abonos, 22) & "  " & _
                            SpcFillAmount((tbalance.SI + tbalance.Cargos + tbalance.Abonos), 22)
                                  
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(50) & _
                            "======================" & "   " & "======================" & "   " & _
                            "======================" & "   " & "======================"
                                
          oTextFile.PrintLine Space(6) & "  " & SpcFill(" TOTAL GENERAL ", 40) & "  " & _
                            SpcFillAmount(tTotal.SI, 22) & "  " & SpcFillAmount(tTotal.Cargos, 22) & "  " & _
                            SpcFillAmount(tTotal.Abonos, 22) & "  " & _
                            SpcFillAmount((tTotal.SI + tTotal.Cargos + tTotal.Abonos), 22)

        Else
          PrintHeader71 oTextFile, nLine, nPage, dIniDate, dEndDate, _
                        vAccountGroupName, sReportTitle, sDTime
        End If
                                
        .Close
        Set oRSGenerator = Nothing
      End With
    
    Report71 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report71", Err.Number
End Function

Public Function Report73(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                         sPattern As String, dIniDate As Date, dEndDate As Date, _
                         Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                         Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                         Optional nCurrencyType As Long = 0, _
                         Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                         Optional dExcDate As Date, Optional bPrintInCascade As Boolean = False, _
                         Optional bCascadeDates As Boolean = False, _
                         Optional nOptionToDisplay As Long = 5, Optional bWithAux As Boolean = False) As String
                         
  Dim oRs           As Object
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
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
    oTextFile.GenerateRandomFile
    oTextFile.Open_
   
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
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = Mid$(oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR"), 10)
        End If
      End If
      sReportTitle = "REPORTE POR MONEDAS, EN FIDEICOMISOS Y MANDATOS"
      For j = dIniDate To dEndDate
        If bCascadeDates Then
          dIniDateTmp = j
          dEndDateTmp = j
          sReportTitle = "DIARIO GENERAL"
          dAccDate = dIniDateTmp
        Else
          dIniDateTmp = dIniDate
          dEndDateTmp = dEndDate
        End If
        
        Set oRs = CreateObject("ADODB.Recordset")
        Set oRs = oRSGenerator.GeneraRecordset73(sAppServer, sPattern, vGLToSend, dIniDateTmp, dEndDateTmp, sIniAccount, sEndAccount, _
                                          nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                          nOptionToDisplay, bWithAux, False, False)
        sDTime = GetFmtDate
        oTextFile.PrintNewLine
        nLine = 100
        nPage = 1
        With oRs
          If Not (.BOF And .EOF) Then
            sNat = !naturaleza
            sMon = !o_id_moneda
            sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
            Do While Not (.EOF)
              PrintHeader73 oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                             sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
                             

                PrintDetalle73 oTextFile, nLine, tAcumulador, !NOMBRE_MAYOR, !NUMERO_CUENTA_ESTANDAR, !clave_sector, !nombre_cuenta, _
                             !saldo_inicial, !Cargos, !Abonos, !Saldo_Actual, !rol_cuenta, !from_movs, nOptionToDisplay
              .MoveNext
            Loop
          Else
            PrintHeader73 oTextFile, nLine, nPage, sGl, "", sCurrency, "", "", "", dIniDate, dEndDate, _
                          sReportTitle, "", sDTime, sAccIncludes, sAccType, sTransactionType, sVoucherType, nExcRateCurrency, dAccDate
          End If
          PrintBalance oTextFile, tAcumulador, 180
          .Close
          .ActiveConnection = Nothing
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
    Report73 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report73", Err.Number
End Function

Public Function Report74_76(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date, _
                            dIniDate2 As Date, dEndDate2 As Date, dExcDate As Date, _
                            Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  
  Dim oRs             As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  
  Dim tActivo         As TMoney
  Dim t6320           As TMoney
  Dim tTotal          As TMoney
  
  Dim tTActivo         As TMoney
  Dim tT6320           As TMoney
  Dim tTTotal          As TMoney
  
  Dim nLine           As Long
  Dim nPage           As Long
  
  Dim sDateTime       As String
  Dim sFileName       As String
  
  Dim nGLKind         As Long
  '**************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = GetFmtDate
    nLine = 100
    nPage = 1
     
    Set oRs = oRSGenerator.GeneraRecordset74(sAppServer, "&&&&", vGL, dIniDate, dEndDate, dIniDate2, dEndDate2, "", "", nExcRateType, nExcRateCurrency, dExcDate)
    With oRs
      If Not (.BOF And .EOF) Then
        nGLKind = !ORDEN
        Do While Not .EOF()
          PrintHeader74_76 oTextFile, nLine, nPage, dIniDate, dEndDate, dIniDate2, dEndDate2, sDateTime
          If (nGLKind <> !ORDEN) Then
              If (nGLKind = 1) Then
                oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
                oTextFile.PrintLine SpcFill(" TOTAL FIDEICOMISOS : ", 36) & _
                                    SpcFillAmount(tActivo.SI, 22) & " " & SpcFillAmount(tActivo.Cargos, 22) & " " & _
                                    SpcFillAmount(tActivo.Abonos, 22) & " " & _
                                    SpcFillAmount(t6320.SI, 22) & " " & SpcFillAmount(t6320.Cargos, 22) & " " & _
                                    SpcFillAmount(t6320.Abonos, 22) & " " & _
                                    SpcFillAmount(tTotal.SI, 22) & " " & SpcFillAmount(tTotal.Cargos, 22) & " " & _
                                    SpcFillAmount(tTotal.Abonos, 22)
              ElseIf (nGLKind = 2) Then
                oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
                oTextFile.PrintLine SpcFill(" TOTAL MANDATOS     : ", 36) & _
                                    SpcFillAmount(tActivo.SI, 22) & " " & SpcFillAmount(tActivo.Cargos, 22) & " " & _
                                    SpcFillAmount(tActivo.Abonos, 22) & " " & _
                                    SpcFillAmount(t6320.SI, 22) & " " & SpcFillAmount(t6320.Cargos, 22) & " " & _
                                    SpcFillAmount(t6320.Abonos, 22) & " " & _
                                    SpcFillAmount(tTotal.SI, 22) & " " & SpcFillAmount(tTotal.Cargos, 22) & " " & _
                                    SpcFillAmount(tTotal.Abonos, 22)
              End If

              oTextFile.PrintNewLine
              nLine = nLine + 3
              AddMoney tTActivo, tActivo.SI, tActivo.Cargos, tActivo.Abonos, 0
              AddMoney tT6320, t6320.SI, t6320.Cargos, t6320.Abonos, 0
              AddMoney tTTotal, tTotal.SI, tTotal.Cargos, tTotal.Abonos, 0
  
              ClearMoney tActivo
              ClearMoney t6320
              ClearMoney tTotal
              nGLKind = !ORDEN
          End If
          oTextFile.PrintLine SpcFill(!NUMERO_MAYOR, 4) & " " & SpcFill(!NOMBRE_MAYOR, 30) & " " & _
                              SpcFillAmount(!SALDO_ACTUAL_2, 22) & " " & SpcFillAmount(!Saldo_Actual, 22) & " " & _
                              SpcFillAmount(!SALDO_ACTUAL_2 - !Saldo_Actual, 22) & " " & _
                              SpcFillAmount(!SALDO_ACTUAL_4, 22) & " " & SpcFillAmount(!SALDO_ACTUAL_3, 22) & " " & _
                              SpcFillAmount(!SALDO_ACTUAL_4 - !SALDO_ACTUAL_3, 22) & " " & _
                              SpcFillAmount(!SALDO_ACTUAL_2 + !SALDO_ACTUAL_4, 22) & " " & SpcFillAmount(!Saldo_Actual + !SALDO_ACTUAL_3, 22) & " " & _
                              SpcFillAmount((!SALDO_ACTUAL_2 + !SALDO_ACTUAL_4) - (!Saldo_Actual + !SALDO_ACTUAL_3), 22)
                              
          AddMoney tActivo, !SALDO_ACTUAL_2, !Saldo_Actual, (!SALDO_ACTUAL_2 - !Saldo_Actual), 0
          AddMoney t6320, !SALDO_ACTUAL_4, !SALDO_ACTUAL_3, (!SALDO_ACTUAL_4 - !SALDO_ACTUAL_3), 0
          AddMoney tTotal, (!SALDO_ACTUAL_2 + !SALDO_ACTUAL_4), (!Saldo_Actual + !SALDO_ACTUAL_3), _
                           ((!SALDO_ACTUAL_2 + !SALDO_ACTUAL_4) - (!Saldo_Actual + !SALDO_ACTUAL_3)), 0
          nLine = nLine + 1
          .MoveNext
        Loop
        .Close
        If (nGLKind = 1) Then
          oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
          oTextFile.PrintLine SpcFill(" TOTAL FIDEICOMISOS : ", 36) & _
                              SpcFillAmount(tActivo.SI, 22) & " " & SpcFillAmount(tActivo.Cargos, 22) & " " & _
                              SpcFillAmount(tActivo.Abonos, 22) & " " & _
                              SpcFillAmount(t6320.SI, 22) & " " & SpcFillAmount(t6320.Cargos, 22) & " " & _
                              SpcFillAmount(t6320.Abonos, 22) & " " & _
                              SpcFillAmount(tTotal.SI, 22) & " " & SpcFillAmount(tTotal.Cargos, 22) & " " & _
                              SpcFillAmount(tTotal.Abonos, 22)
        ElseIf (nGLKind = 2) Then
          oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
          oTextFile.PrintLine SpcFill(" TOTAL MANDATOS     : ", 36) & _
                              SpcFillAmount(tActivo.SI, 22) & " " & SpcFillAmount(tActivo.Cargos, 22) & " " & _
                              SpcFillAmount(tActivo.Abonos, 22) & " " & _
                              SpcFillAmount(t6320.SI, 22) & " " & SpcFillAmount(t6320.Cargos, 22) & " " & _
                              SpcFillAmount(t6320.Abonos, 22) & " " & _
                              SpcFillAmount(tTotal.SI, 22) & " " & SpcFillAmount(tTotal.Cargos, 22) & " " & _
                              SpcFillAmount(tTotal.Abonos, 22)
        End If

        oTextFile.PrintNewLine
        nLine = nLine + 3
        AddMoney tTActivo, tActivo.SI, tActivo.Cargos, tActivo.Abonos, 0
        AddMoney tT6320, t6320.SI, t6320.Cargos, t6320.Abonos, 0
        AddMoney tTTotal, tTotal.SI, tTotal.Cargos, tTotal.Abonos, 0
        
        oTextFile.PrintNewLine
        oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
        oTextFile.PrintLine SpcFill(" TOTAL GENERAL  : ", 36) & _
                            SpcFillAmount(tTActivo.SI, 22) & " " & SpcFillAmount(tTActivo.Cargos, 22) & " " & _
                            SpcFillAmount(tTActivo.Abonos, 22) & " " & _
                            SpcFillAmount(tT6320.SI, 22) & " " & SpcFillAmount(tT6320.Cargos, 22) & " " & _
                            SpcFillAmount(tT6320.Abonos, 22) & " " & _
                            SpcFillAmount(tTTotal.SI, 22) & " " & SpcFillAmount(tTTotal.Cargos, 22) & " " & _
                            SpcFillAmount(tTTotal.Abonos, 22)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Set oRs = Nothing
    Report74_76 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report74_76", Err.Number
End Function

Public Function Report96(sAppServer As String, vGL As Variant, dIniDate As Date, _
                         dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String

  Dim oRsBalance      As ADODB.Recordset
  Dim oRsPosting      As ADODB.Recordset
  Dim oRsAccNames     As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  'Dim sGLDesc         As String
  Dim sLongDate       As String
  
  Dim nTotDeud        As Currency
  Dim nTotAcre        As Currency
  Dim nTotDeudMN      As Currency
  Dim nTotAcreMN      As Currency
  Dim nLine           As Long
  Dim nPage           As Long
  Dim nCurr           As Long
  Dim nExcRate        As Double
  
  Dim sDateTime       As String
  Dim sFileName       As String
  
  Dim sCurrName       As String
  
  '*********************************************************************************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    'oTextFile.Open_ "C:\dev2001\", "Report96.txt"
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sLongDate = Format(dEndDate, "d") & " de " & Format(dEndDate, "MMMM") & " de " & Year(dEndDate)
    
    nLine = 100
    nPage = 1
  
    Set oRsBalance = oRSGenerator.GeneraRecordset96(sAppServer, vGL, dIniDate, dEndDate, dExcDate, nExcRateType, nExcRateCurrency)
    With oRsBalance
      Do While Not .EOF()
        nCurr = !ID_MONEDA
        nExcRate = !exc_rate
        sCurrName = !nom_moneda
        
        If (!st <> 0 And Not IsNull(!st)) Then
          If (!st > 0) Then
            nTotDeud = nTotDeud + !st
            nTotDeudMN = nTotDeudMN + (!st * !exc_rate)
          Else
            nTotAcre = nTotAcre + (!st * -1)
            nTotAcreMN = nTotAcreMN + ((!st * -1) * !exc_rate)
          End If
          
          PrintHeader96 oTextFile, nLine, nPage, sLongDate, " ", " ", " ", sDateTime
          oTextFile.PrintLine SpcFill(GetCuentaSubCuenta(!cuenta) & !cve_sector, 19) & "  " & SpcFill(!nom_cuenta, 21) & "  " & _
                              SpcFill(IIf(IsNull(!num_aux), " ", !num_aux), 18) & "  " & SpcFill(IIf(IsNull(!nom_aux), " ", !nom_aux), 30) & "  " & SpcFill(!ID_MONEDA, 3) & "  " & _
                              IIf(!st > 0, SpcFillAmount(!st, 20), SpcFill(" ", 20)) & "  " & IIf(!st < 0, SpcFillAmount((!st * -1), 20), SpcFill(" ", 20))
          nLine = nLine + 1
          .MoveNext
          If Not .EOF Then
            If (!ID_MONEDA <> nCurr) Then
              oTextFile.PrintLine Space(100) & "============================================"
              oTextFile.PrintLine Space(54) & SpcFill(" Total Moneda: " & sCurrName, 35) & Space(12) & SpcFillAmount(nTotDeud, 20) & " " & SpcFillAmount(nTotAcre, 20)
              oTextFile.PrintLine Space(100) & "============================================"
              oTextFile.PrintLine Space(40) & " Total en Moneda Nacional  Tipo Cambio       " & SpcFillAmount(nExcRate, 8, 6) & Space(7) & SpcFillAmount(nTotDeud * nExcRate, 20) & " " & SpcFillAmount(nTotAcre * nExcRate, 20)
              nTotDeud = 0
              nTotAcre = 0
              oTextFile.PrintNewLine
              nLine = nLine + 5
            End If
          End If
        Else
          .MoveNext
        End If
      Loop
        oTextFile.PrintLine Space(100) & "============================================"
        oTextFile.PrintLine Space(54) & SpcFill(" Total Moneda: " & sCurrName, 35) & Space(12) & SpcFillAmount(nTotDeud, 20) & " " & SpcFillAmount(nTotAcre, 20)
        oTextFile.PrintLine Space(100) & "============================================"
        oTextFile.PrintLine Space(40) & " Total en Moneda Nacional  Tipo Cambio       " & SpcFillAmount(nExcRate, 8, 6) & Space(7) & SpcFillAmount(nTotDeud * nExcRate, 20) & " " & SpcFillAmount(nTotAcre * nExcRate, 20)
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(100) & "============================================"
        oTextFile.PrintLine Space(55) & " TOTAL CUENTA EN MONEDA NACIONAL       " & Space(7) & SpcFillAmount(nTotDeudMN, 20) & " " & SpcFillAmount(nTotAcreMN, 20)
        oTextFile.PrintNewLine
        
      .Close
    End With
    Set oRsBalance = Nothing
  Report96 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report96", Err.Number
End Function

Public Function Report99(sAppServer As String, vGL As Variant, dIniDate As Date, _
                         dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String

  Dim oRsBalance      As ADODB.Recordset
  Dim oRsPosting      As ADODB.Recordset
  Dim oRsAccNames     As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  Dim sLongDate       As String
  
  Dim nTot            As Currency
  Dim nTotAcre        As Currency
  Dim nTotDeudMN      As Currency
  Dim nTotAcreMN      As Currency
  Dim nLine           As Long
  Dim nPage           As Long
  Dim nCurr           As Long
  Dim nExcRate        As Currency
  
  Dim sDateTime       As String
  Dim sFileName       As String
  Dim skindacc        As String
  Dim sNat            As String
  Dim sCurrName       As String
  Dim sAcc            As String
  Dim sHeadAcc        As String
  Dim sGl             As String
  
  Dim nGl             As Long
  '*********************************************************************************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = Format(Date, "dd/mm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sLongDate = Format(dEndDate, "MMMM") & " de " & Year(dEndDate)
    
    If (IsArray(vGL)) Then
      sGl = "0000 AGENCIA CONSOLIDADORA BANOBRAS"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    nLine = 100
    nPage = 1
    
    Set oRsBalance = oRSGenerator.GeneraRecordset99(sAppServer, "&&&&", vGL, dIniDate, dEndDate, nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency, dExcDate:=dExcDate, bWithAux:=False)
    With oRsBalance
      If Not (.BOF And .EOF) Then
        Do While Not .EOF()
          sNat = !naturaleza
          If sNat = "D" Then
            skindacc = "1 CTAS DEUDORAS"
          Else
            skindacc = "2 CTAS ACREEDORAS"
          End If
          nCurr = !ID_MONEDA
          'nExcRate = !exc_rate
          'sCurrName = !nom_moneda
          
          If (!Saldo_Actual <> 0 And Not IsNull(!Saldo_Actual)) Then
            PrintHeader99 oTextFile, nLine, nPage, sLongDate, sGl, skindacc, " ", sDateTime
            If (sAcc <> !NUMERO_CUENTA_ESTANDAR) Or (nLine = 20) Then
              sHeadAcc = SpcFill(!NUMERO_CUENTA_ESTANDAR & " " & !nombre_cuenta, 45)
              sAcc = !NUMERO_CUENTA_ESTANDAR
            Else
              sHeadAcc = SpcFill(" ", 45)
            End If
            oTextFile.PrintLine sHeadAcc & "  " & SpcFill(!CURRENCY_NAME, 25) & "  " & SpcFill(ZeroFill(!ID_MONEDA, 2), 4) & "  " & _
                                SpcFillAmount(!Saldo_Actual, 20) & Space(5) & "  " & SpcFillAmount(!EXCHANGE_RATE, 8, 6) & "  " & SpcFillAmount(!Saldo_Actual * !EXCHANGE_RATE, 25)
            nLine = nLine + 1
            nTot = nTot + (!Saldo_Actual * !EXCHANGE_RATE)
            .MoveNext
            If (Not .EOF) Then
              If (sAcc <> !NUMERO_CUENTA_ESTANDAR) Then
                oTextFile.PrintLine Space(120) & "========================"
                oTextFile.PrintLine Space(5) & " TOTAL POR CUENTA " & Space(96) & SpcFillAmount(nTot, 25)
                nTot = 0
                nLine = nLine + 3
                oTextFile.PrintNewLine
                oTextFile.PrintNewLine
              End If
              If (sNat <> !naturaleza) Then
                nLine = 100
              End If
            End If
          Else
            .MoveNext
          End If
        Loop
        oTextFile.PrintLine Space(120) & "========================"
        oTextFile.PrintLine Space(5) & " TOTAL POR CUENTA " & Space(96) & SpcFillAmount(nTot, 25)
        .Close
      End If
    End With
    Set oRsBalance = Nothing
  Report99 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report99", Err.Number
End Function

Public Function Report101(sAppServer As String, nGl As Long, vGL As Variant, sPatron As String, _
                         dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1, _
                         Optional sAccounts As String, Optional sReportType As String) As String
  Dim oRs           As ADODB.Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim dFecha        As Date
  Dim dFechaFin     As Date
  Dim sAccount      As String
  Dim sSector       As String
  Dim nSaldoMN      As Double
  Dim nSaldoME      As Double
  Dim nSaldoMNA     As Double
  Dim nSaldoMEA     As Double
  Dim nSaldoCurrM   As Double
  Dim nSaldoCurrD   As Double
  Dim nExcRate      As Double
  Dim sAccName      As String

  Dim nLine         As Long
  Dim nPage         As Long
  Dim sNature       As String
  Dim sDateTime     As String

  Dim DebugAccount  As String
  Dim sAccType      As String
  Dim sCurrency     As String
  Dim nCurrency     As Long
  
  Dim nAccType      As Long
  Dim nGlStdAcc     As Long
  Dim sGl           As String
  
  Dim sReportTittle As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    If (IsArray(vGL)) Then
      sGl = " TODAS LAS CONTABILIDADES "
      nGlStdAcc = vGL(0)
    Else
      nGlStdAcc = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGlStdAcc, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    If (sReportType = "CI") Then
      sReportTittle = "ANALISIS DE SALDOS ACUMULADOS PARA EL CALCULO DE INDICE INFLACIONARIO DE LA INSTITUCION"
    ElseIf (sReportType = "PM") Then
      sReportTittle = "ANALISIS DE SALDOS ACUMULADOS PARA EL CALCULO DE LA POSICION MONETARIA"
    End If
    
    Set oRs = oRSGenerator.GeneraRecordset101(sPatron, sAppServer, nGl, vGL, dIniDate, dEndDate, _
                                             dExcDate, nExcRateType, nExcRateCurrency, sAccounts)

    nAccType = oRSGenerator.AccountType(sAppServer, nGlStdAcc)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

    DebugAccount = "0000"

    With oRs
      If Not (.BOF And .EOF) Then
        nPage = 1
        nLine = 100
        dFecha = !fecha_rev
        sAccount = !NUMERO_CUENTA_ESTANDAR
        sAccName = !nombre_cuenta
        sSector = !clave_sector
        nExcRate = !EXCHANGE_RATE
        nSaldoMN = 0
        nSaldoME = 0
        nSaldoMNA = 0
        nSaldoMEA = 0
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
        nCurrency = !ID_MONEDA
        PrintHeader101 oTextFile, nLine, nPage, sGl, sCurrency, !EXCHANGE_RATE, dIniDate, dEndDate, sDateTime, sAccType, sReportTittle
        oTextFile.PrintLine "Cuenta : " & SpcFill(sAccount, 22) & "  " & SpcFill(sAccName, 35)
         
        Do While Not (.EOF)
          If (!naturaleza = "D") Then
            sNature = "1"
          Else
            sNature = "2"
          End If
          PrintHeader101 oTextFile, nLine, nPage, sGl, sCurrency, !EXCHANGE_RATE, dIniDate, dEndDate, sDateTime, sAccType, sReportTittle
                    
          If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
            If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPatron)) Then
              dFechaFin = dFecha
              Do While dFechaFin <= dEndDate
                oTextFile.PrintLine SpcFill(sAccount, 22) & " " & sSector & "  " & SpcFill(dFechaFin, 35) & " " & SpcFillAmount(nSaldoMN, 22) & "  " & _
                                SpcFillAmount(nSaldoME, 22) & "  " & SpcFillAmount((nSaldoMN + nSaldoME), 22)
                nSaldoMNA = nSaldoMNA + nSaldoMN
                nSaldoMEA = nSaldoMEA + nSaldoME
                dFechaFin = dFechaFin + 1
                nLine = nLine + 1
              Loop
              oTextFile.PrintNewLine
              oTextFile.PrintLine Space(2) & "Total de Saldos :     " & " " & Space(37) & " " & SpcFillAmount(nSaldoMNA, 22) & "  " & _
                                SpcFillAmount(nSaldoMEA, 22) & "  " & SpcFillAmount((nSaldoMNA + nSaldoMEA), 22)
              oTextFile.PrintNewLine
              nSaldoCurrM = nSaldoCurrM + nSaldoMNA
              nSaldoCurrD = nSaldoCurrD + nSaldoMEA

              nSaldoMNA = (nSaldoMNA) / (dEndDate - dIniDate + 1)
              nSaldoMEA = (nSaldoMEA) / (dEndDate - dIniDate + 1)
                                  
              oTextFile.PrintLine Space(103) & " ================================== "
              oTextFile.PrintLine Space(103) & " PROMEDIO " & SpcFillAmount((nSaldoMNA + nSaldoMEA), 22)
              oTextFile.PrintNewLine
              nLine = nLine + 5
            ElseIf (Len(!cuenta) > Len(sPatron)) Then
              'Siguiente
            End If
            nSaldoMN = 0
            nSaldoME = 0
            nSaldoMNA = 0
            nSaldoMEA = 0
            sAccount = !NUMERO_CUENTA_ESTANDAR
            sAccName = !nombre_cuenta
            sSector = !clave_sector
            dFecha = !fecha_rev
            If (nCurrency <> !ID_MONEDA) Then
              oTextFile.PrintNewLine
              oTextFile.PrintLine Space(2) & "TOTAL MONEDA    :     " & " " & SpcFill(sCurrency, 37) & " " & SpcFillAmount(nSaldoCurrM, 22) & "  " & _
                                SpcFillAmount(nSaldoCurrD, 22) & "  " & SpcFillAmount((nSaldoCurrM + nSaldoCurrD), 22)
                                
              oTextFile.PrintNewLine
              oTextFile.PrintLine Space(2) & "TOTAL MONEDA NACIONAL " & " " & Space(37) & " " & SpcFillAmount((nSaldoCurrM * nExcRate), 22) & "  " & _
                                SpcFillAmount((nSaldoCurrD * nExcRate), 22) & "  " & SpcFillAmount((nSaldoCurrM + nSaldoCurrD) * nExcRate, 22)
              oTextFile.PrintNewLine
                                
              sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
              sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
              nCurrency = !ID_MONEDA
              nExcRate = !EXCHANGE_RATE
              nSaldoCurrM = 0
              nSaldoCurrD = 0
              nLine = 100
            End If
            PrintHeader101 oTextFile, nLine, nPage, sGl, sCurrency, !EXCHANGE_RATE, dIniDate, dEndDate, sDateTime, sAccType, sReportTittle
            oTextFile.PrintLine "Cuenta : " & SpcFill(sAccount, 22) & "  " & SpcFill(sAccName, 35)
            nLine = nLine + 1
          End If
  
          If (dFecha = #1/1/1990#) Then
             nSaldoMN = !saldo_inicial
             nSaldoME = !Saldo_Inicial_all
             nSaldoMNA = 0
             nSaldoMEA = 0
             dFecha = dIniDate
          ElseIf (dFecha <> #1/1/1990#) Then
            If (dFecha = !fecha_rev) Then
              If (!naturaleza = "D") Then
                nSaldoMN = nSaldoMN + !Cargos - !Abonos
                nSaldoME = nSaldoME + !cargos_all - !abonos_all
              ElseIf (!naturaleza = "A") Then
                nSaldoMN = nSaldoMN - !Cargos + !Abonos
                nSaldoME = nSaldoME - !cargos_all + !abonos_all
              End If
              oTextFile.PrintLine SpcFill(sAccount, 22) & " " & sSector & "  " & SpcFill(dFecha, 35) & " " & SpcFillAmount(nSaldoMN, 22) & "  " & _
                                SpcFillAmount(nSaldoME, 22) & "  " & SpcFillAmount((nSaldoMN + nSaldoME), 22)
              nSaldoMNA = nSaldoMNA + nSaldoMN
              nSaldoMEA = nSaldoMEA + nSaldoME
              nLine = nLine + 1
            Else
              Do While dFecha <= (!fecha_rev - 1)
                oTextFile.PrintLine SpcFill(sAccount, 22) & " " & sSector & "  " & SpcFill(dFecha, 35) & " " & SpcFillAmount(nSaldoMN, 22) & "  " & _
                                SpcFillAmount(nSaldoME, 22) & "  " & SpcFillAmount((nSaldoMN + nSaldoME), 22)
                nSaldoMNA = nSaldoMNA + nSaldoMN
                nSaldoMEA = nSaldoMEA + nSaldoME
                
                dFecha = dFecha + 1
                nLine = nLine + 1
              Loop
                            
              If (!naturaleza = "D") Then
                nSaldoMN = nSaldoMN + !Cargos - !Abonos
                nSaldoME = nSaldoME + !cargos_all - !abonos_all
              ElseIf (!naturaleza = "A") Then
                nSaldoMN = nSaldoMN - !Cargos + !Abonos
                nSaldoME = nSaldoME - !cargos_all + !abonos_all
              End If
              oTextFile.PrintLine SpcFill(sAccount, 22) & " " & sSector & "  " & SpcFill(!fecha_rev, 35) & " " & SpcFillAmount(nSaldoMN, 22) & "  " & _
                                SpcFillAmount(nSaldoME, 22) & "  " & SpcFillAmount((nSaldoMN + nSaldoME), 22)
              nSaldoMNA = nSaldoMNA + nSaldoMN
              nSaldoMEA = nSaldoMEA + nSaldoME
              nLine = nLine + 1
            End If
            dFecha = !fecha_rev + 1
          End If
          .MoveNext
        Loop
        
        dFechaFin = dFecha
        Do While dFechaFin <= dEndDate
          oTextFile.PrintLine SpcFill(sAccount, 22) & " " & sSector & "  " & SpcFill(dFechaFin, 35) & " " & SpcFillAmount(nSaldoMN, 22) & "  " & _
                          SpcFillAmount(nSaldoME, 22) & "  " & SpcFillAmount((nSaldoMN + nSaldoME), 22)
          nSaldoMNA = nSaldoMNA + nSaldoMN
          nSaldoMEA = nSaldoMEA + nSaldoME
          dFechaFin = dFechaFin + 1
        Loop
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(2) & "Total de Saldos :     " & " " & Space(37) & " " & SpcFillAmount(nSaldoMNA, 22) & "  " & _
                          SpcFillAmount(nSaldoMEA, 22) & "  " & SpcFillAmount((nSaldoMNA + nSaldoMEA), 22)
        oTextFile.PrintNewLine
        
        nSaldoCurrM = nSaldoCurrM + nSaldoMNA
        nSaldoCurrD = nSaldoCurrD + nSaldoMEA
              
        nSaldoMNA = (nSaldoMNA) / (dEndDate - dIniDate + 1)
        nSaldoMEA = (nSaldoMEA) / (dEndDate - dIniDate + 1)
                            
        oTextFile.PrintLine Space(103) & " ================================== "
        oTextFile.PrintLine Space(103) & " PROMEDIO " & SpcFillAmount((nSaldoMNA + nSaldoMEA), 22)
        oTextFile.PrintNewLine
                     
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(2) & "TOTAL MONEDA    :     " & " " & SpcFill(sCurrency, 37) & " " & SpcFillAmount(nSaldoCurrM, 22) & "  " & _
                          SpcFillAmount(nSaldoCurrD, 22) & "  " & SpcFillAmount((nSaldoCurrM + nSaldoCurrD), 22)
                          
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(2) & "TOTAL MONEDA NACIONAL " & " " & Space(37) & " " & SpcFillAmount((nSaldoCurrM * nExcRate), 22) & "  " & _
                          SpcFillAmount((nSaldoCurrD * nExcRate), 22) & "  " & SpcFillAmount((nSaldoCurrM + nSaldoCurrD) * nExcRate, 22)
        oTextFile.PrintNewLine
        
        
        nLine = nLine + 1
      Else
        oTextFile.PrintLine " NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report101 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report101", Err.Number
End Function

Public Function Report102(sAppServer As String, sPatron As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                         Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1, _
                         Optional sAccounts As String) As String
  Dim oRs            As Recordset
  Dim oRSGenerator   As New SCFFixedReportsDS.CServer
  Dim oTextFile      As New CTextFile
  Dim dFecha         As Date
  Dim sAccount       As String
  Dim sSector        As String
  Dim nSaldoMN       As Double
  Dim nSaldoME       As Double
  Dim nSaldoMNA      As Double
  Dim nSaldoMEA      As Double
  Dim sAccName       As String
  Dim nTotDeudoras   As Double
  Dim nTotAcreedoras As Double
  
  Dim nLine          As Long
  Dim nPage          As Long
  Dim sNature        As String
  Dim sDateTime      As String
  
  Dim nGl            As Long
  Dim nAccType       As Long
  Dim sAccType       As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    If (IsArray(vGL)) Then
      nGl = vGL(0)
    Else
      nGl = vGL
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sPatron = "&&&&"
    
    Set oRs = oRSGenerator.GeneraRecordset58(sPatron, sAppServer, vGL, dIniDate, dEndDate, dExcDate, nExcRateType, _
                                             nExcRateCurrency, sAccounts, 102)
    With oRs
      If Not (.BOF And .EOF) Then
        nPage = 1
        nLine = 100
        dFecha = !fecha_rev
        sAccount = !NUMERO_CUENTA_ESTANDAR
        sAccName = !nombre_cuenta
        sSector = !clave_sector
        nSaldoMN = 0
        nSaldoME = 0
        nSaldoMNA = 0
        nSaldoMEA = 0
        nTotAcreedoras = 0
        nTotDeudoras = 0
        Do While Not (.EOF)
          If (!naturaleza = "D") Then
            sNature = "1"
          Else
            sNature = "2"
          End If
          
          If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
            If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPatron)) Then
              PrintHeader102 oTextFile, nLine, nPage, "", "", "", dIniDate, dEndDate, sDateTime, sAccType
              nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1)))
              nSaldoMEA = (nSaldoMEA + (nSaldoME * ((dEndDate - dFecha) + 1)))
              oTextFile.PrintLine SpcFill(sAccount, 6) & " " & SpcFill(sAccName, 40) & " " & _
                                  SpcFillAmount(nSaldoMNA, 27) & "  " & _
                                  SpcFillAmount(nSaldoMEA, 27) & "  " & _
                                  SpcFillAmount(nSaldoMNA + nSaldoMEA, 27) & " " & _
                                  SpcFillAmount((nSaldoMNA + nSaldoMEA) / (dEndDate - dIniDate + 1), 27)
              nLine = nLine + 1
            ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPatron)) Then
              'Siguiente
            End If
            nSaldoMN = 0
            nSaldoME = 0
            nSaldoMNA = 0
            nSaldoMEA = 0
            sAccount = !NUMERO_CUENTA_ESTANDAR
            sAccName = !nombre_cuenta
            sSector = !clave_sector
            dFecha = !fecha_rev
          End If
          
          If (dFecha = #1/1/1990#) Then
             nSaldoMN = !saldo_inicial_mn
             nSaldoME = !Saldo_Inicial_me
             nSaldoMNA = 0
             nSaldoMEA = 0
             dFecha = dIniDate
          ElseIf (dFecha <> #1/1/1990#) Then
            nSaldoMNA = nSaldoMNA + (nSaldoMN * (!fecha_rev - dFecha))
            nSaldoMEA = nSaldoMEA + (nSaldoME * (!fecha_rev - dFecha))
            If (!naturaleza = "D") Then
              nSaldoMN = nSaldoMN + !cargos_mn - !abonos_mn
              nSaldoME = nSaldoME + !cargos_me - !abonos_me
            ElseIf (!naturaleza = "A") Then
              nSaldoMN = nSaldoMN - !cargos_mn + !abonos_mn
              nSaldoME = nSaldoME - !cargos_me + !abonos_me
            End If
            dFecha = !fecha_rev
          End If
          .MoveNext
        Loop
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Report102 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report102", Err.Number
End Function

Public Function Report106(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset106(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader106 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 60) & "     " & _
                              SpcFill(IIf(IsNull(!Programa), 0, !Programa), 10) & "            " & _
                              SpcFill(IIf(IsNull(!SubPrograma), 0, !SubPrograma), 10)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report106 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report106", Err.Number
End Function

Public Function Report107(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset107(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader107 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 50) & "     " & _
                              SpcFill(IIf(IsNull(!RFC), "", !RFC), 20) & "  " & _
                              SpcFill(IIf(IsNull(!CURP), "", !CURP), 20) & "  " & _
                              SpcFill(IIf(IsNull(!CLAVEIVA), "", !CLAVEIVA), 4) & "  " & _
                              SpcFillAmount(IIf(IsNull(!TASAIVA), 0, !TASAIVA), 8) & "  " & _
                              SpcFill(IIf(IsNull(!TIPORETENCION), "", !TIPORETENCION), 4)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report107 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report107", Err.Number
End Function

Public Function Report108(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset108(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader108 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 50) & "     " & _
                              SpcFill(IIf(IsNull(!FECHAADQUISICION), "", !FECHAADQUISICION), 14) & "  " & _
                              SpcFillAmount(IIf(IsNull(!VALORHISTORICO), 0, !VALORHISTORICO), 18)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report108 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report108", Err.Number
End Function

Public Function Report109(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset109(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader109 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 60)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report109 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report109", Err.Number
End Function

Public Function Report110(sAppServer As String, sPatron As String, vGL As Variant, _
                          dIniDate As Date, dEndDate As Date, _
                          Optional sFromAccount As String, Optional sToAccount As String, _
                          Optional sFromSubsAccount As String, Optional sToSubsAccount As String, _
                          Optional bPrintInCascade As Boolean = False) As String
  Dim oRs           As Object
  Dim oRSGenerator  As SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sCuenta       As String
  Dim sSubcuenta    As String
  Dim sAux          As String
  Dim sAuxHead      As String
  Dim dFechaMov     As String
  Dim sMon          As String
  
  Dim tAcumulador   As TMoney
  Dim tSubCuenta    As TMoney
  Dim tCuenta       As TMoney
  Dim tAux          As TMoney
  Dim tFechaMov     As TMoney
  Dim tMoneda       As TMoney
  Dim tTotal        As TMoney
  Dim tVoid         As TMoney
  
  Dim nBalance      As Currency
  
  Dim sGl           As String
  Dim sCurrency     As String
  
  Dim bPrintSI      As Boolean
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  
  Dim bSubAccCh     As Boolean
  Dim bPaintSubAcc  As Boolean
  
  Dim vGLToSend     As Variant
  Dim vGL2          As Variant
  Dim i             As Long
  Dim sNameAux      As String
  Dim sNameAcc      As String
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    vGL2 = vGL
    If Not IsArray(vGL2) Then
      vGL2 = Array(vGL2)
    End If
    
    nPage = 1
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        nGl = CLng(vGLToSend)
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
      Else
        Set oRSGenerator = New SCFFixedReportsDS.CServer
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          nGl = vGLToSend(0)
          sGl = "TODAS LAS CONTABILIDADES"
        Else
          nGl = CLng(vGLToSend)
          sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        End If
      End If

      nAccType = oRSGenerator.AccountType(sAppServer, nGl)
      sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")

      Set oRs = CreateObject("ADODB.Recordset")
      Set oRs = oRSGenerator.GeneraRecordset110(sPatron, sAppServer, vGLToSend, dIniDate, dEndDate, _
                                                sFromAccount, sToAccount, sFromSubsAccount, sToSubsAccount)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
        
      nLine = 100
      With oRs
        If Not (.BOF And .EOF) Then
          sMon = !ID_MONEDA
          sCuenta = Mid(!cuenta, 1, 4)
          sSubcuenta = Mid(!cuenta, 5) & !clave_sector
          sAux = !clave_aux
          sAuxHead = !clave_aux
          sNameAux = !nombre_aux
          sNameAcc = !nombre_cuenta
          dFechaMov = !fecha_ult
          
          nBalance = !SI
          bPrintSI = True
          bSubAccCh = False
          bPaintSubAcc = False
          sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
          
          Do While Not (.EOF)
            PrintHeader110 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                              !cuenta, !clave_sector, !clave_aux, !nombre_cuenta, !N_SECTOR, !nombre_aux, sAccType

            If ((dFechaMov <> !fecha_ult) And (!fecha_ult <> #1/1/1990#)) Or _
               ((sAux <> !clave_aux) Or (sSubcuenta <> (Mid(!cuenta, 5) & !clave_sector)) Or _
               (sCuenta <> Mid(!cuenta, 1, 4)) Or (sMon <> !ID_MONEDA)) Then
              PrintTot oTextFile, nLine, tFechaMov, tVoid, dFechaMov, !fecha_ult, "   TOTAL POR FECHA " & dFechaMov
            End If
            AddMoney tFechaMov, 0, !Cargos, !Abonos, nBalance
            
            If ((sAux <> !clave_aux) Or (sSubcuenta <> (Mid(!cuenta, 5) & !clave_sector)) Or _
                (sCuenta <> Mid(!cuenta, 1, 4)) Or (sMon <> !ID_MONEDA)) Then
              AddMoney tAcumulador, 0, 0, 0, nBalance
              nBalance = !SI
              bPrintSI = True
              PrintTot oTextFile, nLine, tAcumulador, tAux, sAux, !clave_aux, "   TOTAL POR AUXILIAR " & sAux
              sNameAux = !nombre_aux
              If (sSubcuenta <> (Mid(!cuenta, 5) & !clave_sector)) And (nLine <> 19) Then
                bSubAccCh = True
                sNameAcc = !nombre_cuenta
              End If
            End If
            If ((sSubcuenta <> (Mid(!cuenta, 5) & !clave_sector)) Or (sCuenta <> Mid(!cuenta, 1, 4)) Or _
                (sMon <> !ID_MONEDA)) Then
              PrintTot oTextFile, nLine, tAux, tSubCuenta, sSubcuenta, (Mid(!cuenta, 5) & !clave_sector), "    TOTAL POR SUBCUENTA " & sCuenta & Left(sSubcuenta, Len(sSubcuenta) - 2)
            End If
            If ((sCuenta <> Mid(!cuenta, 1, 4)) Or (sMon <> !ID_MONEDA)) Then
              PrintTot oTextFile, nLine, tSubCuenta, tCuenta, sCuenta, Mid(!cuenta, 1, 4), "   TOTAL POR CUENTA " & sCuenta
              oTextFile.PageFeed (nLine)
              nLine = 90
              PrintHeader110 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                                !cuenta, !clave_sector, !clave_aux, !nombre_cuenta, !N_SECTOR, !nombre_aux, sAccType
            End If
            If (sMon <> !ID_MONEDA) Then
              PrintTot oTextFile, nLine, tCuenta, tMoneda, sMon, !ID_MONEDA, "   TOTAL POR MONEDA ", False
              AddMoney tTotal, tMoneda.SI, tMoneda.Cargos, tMoneda.Abonos, tMoneda.SA
              ClearMoney tMoneda
              sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
              oTextFile.PageFeed (nLine)
              nLine = 90
              PrintHeader110 sAppServer, oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, sDTime, _
                                !cuenta, !clave_sector, !clave_aux, !nombre_cuenta, !N_SECTOR, !nombre_aux, sAccType
            End If
            
            If (bSubAccCh) And (nLine <> 19) Then
              oTextFile.PrintLine " SUBCUENTA:  " & !cuenta & "  " & _
                                  oRSGenerator.GetObjNameStr(sAppServer, !cuenta, "COF_CUENTA_ESTANDAR_HIST", "NOMBRE_CUENTA_ESTANDAR", "NUMERO_CUENTA_ESTANDAR")
              bSubAccCh = False
              bPaintSubAcc = True
            ElseIf (nLine = 19) Then
              bSubAccCh = False
            End If
            
            If ((sAux <> sAuxHead) And (nLine <> 19)) Or bPaintSubAcc Then
              oTextFile.PrintLine "  AUXILIAR:  " & sAux & "  " & sNameAux
              sAuxHead = sAux
              bPaintSubAcc = False
            ElseIf (nLine = 19) Then
              sAuxHead = sAux
            End If
            
            If (bPrintSI) Then
              oTextFile.PrintLine SpcFill(" ", 125) & " SALDO INICIAL: " & SpcFillAmount(nBalance, 22)
              bPrintSI = False
            ElseIf Not (bPrintSI) Then
             '
            End If
            PrintDetalle110 oTextFile, !naturaleza, tAcumulador, !fecha_ult, !POLIZA, !CONCEPTO, nBalance, !Cargos, !Abonos
            .MoveNext
          Loop
          AddMoney tAcumulador, 0, 0, 0, nBalance
          PrintTot oTextFile, nLine, tAcumulador, tAux, sAux, "", "    TOTAL POR AUXILIAR " & sAux
          PrintTot oTextFile, nLine, tAux, tSubCuenta, sSubcuenta, "", "    TOTAL POR SUBCUENTA " & sCuenta & Left(sSubcuenta, Len(sSubcuenta) - 2)
  
          PrintTot oTextFile, nLine, tSubCuenta, tCuenta, sCuenta, "", "    TOTAL POR CUENTA " & sCuenta
          
          PrintTot oTextFile, nLine, tCuenta, tMoneda, sMon, "", "    TOTAL POR MONEDA", False
  
          AddMoney tTotal, tMoneda.SI, tMoneda.Cargos, tMoneda.Abonos, tMoneda.SA
          oTextFile.PrintLine SpcFill("    TOTAL DEL REPORTE", cnBlancosCorte + 13) & _
                            "  " & SpcFillAmount(tTotal.Cargos, 22) & _
                            "  " & SpcFillAmount(tTotal.Abonos, 22)
                            '& _
                            "  " & SpcFillAmount(tTotal.SA, 22)
          .Close
        Else
          'oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
        End If
      End With
      If bPrintInCascade Then
        oTextFile.PageFeed nLine
        nBalance = 0
        ClearMoney tAcumulador
        ClearMoney tSubCuenta
        ClearMoney tAux
        ClearMoney tMoneda
        ClearMoney tTotal
      Else
        Exit For
      End If
    Next i
    Report110 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report110", Err.Number
End Function

Public Function Report113(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date, _
                         Optional sFromAccount As String = "", Optional sToAccount As String = "", _
                         Optional sFromSubsAccount As String = "", Optional sToSubsAccount As String = "", _
                         Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                         Optional nCurrencyType As Long = 0, _
                         Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                         Optional dExcDate As Date, Optional nOptionToDisplay As Long = 3) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  Dim bPrintAccount As Boolean
  
  Dim sCuenta       As String
  Dim sSubcuenta    As String
  
  Dim sMon          As String
  Dim sAux          As String
  
  Dim tMoneda       As TMoney
  Dim tAux          As TMoney
  
  Dim sGl           As String
  Dim sCurrency     As String
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  '***************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset113(sAppServer, vGL, dIniDate, dEndDate, sFromAccount, sToAccount, _
                                 sFromSubsAccount, sToSubsAccount, nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                 nOptionToDisplay, True, False)      '
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
    bPrintAccount = True
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        sAux = !NUMERO_CUENTA_AUXILIAR
        sMon = !ID_MONEDA
        sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
        Do While Not (.EOF)
          PrintHeader113 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, "REPORTE DE AUXILIARES POR CUENTA", sDTime, sAccType, bPrintAccount
          
          If (sAux <> !NUMERO_CUENTA_AUXILIAR) Then
            PrintQtyLines oTextFile, 102, 80
            oTextFile.PrintLine Space(58) & "TOTAL POR AUXILIAR  " & "  " & _
                                SpcFillAmount(tAux.SI, 22) & "  " & _
                                SpcFillAmount(tAux.Cargos, 22) & "  " & _
                                SpcFillAmount(tAux.Abonos, 22) & "  " & _
                                SpcFillAmount(tAux.SA, 22)
            ClearMoney tAux
            bPrintAccount = True
            sAux = !NUMERO_CUENTA_AUXILIAR
            nLine = nLine + 1
          End If
          
          If (sMon <> !ID_MONEDA) Then
            PrintQtyLines oTextFile, 102, 80
            oTextFile.PrintLine Space(58) & "TOTAL POR MONEDA    " & "  " & _
                                SpcFillAmount(tMoneda.SI, 22) & "  " & _
                                SpcFillAmount(tMoneda.Cargos, 22) & "  " & _
                                SpcFillAmount(tMoneda.Abonos, 22) & "  " & _
                                SpcFillAmount(tMoneda.SA, 22)
            sMon = !ID_MONEDA
            sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
            ClearMoney tMoneda
            oTextFile.PageFeed nLine
            PrintHeader113 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, "REPORTE DE AUXILIARES POR CUENTA", sDTime, sAccType, bPrintAccount
          End If
          
          If (bPrintAccount) Then
            oTextFile.PrintLine "AUXILIAR: " & SpcFill(!NUMERO_CUENTA_AUXILIAR, 22)
            bPrintAccount = False
          End If
          
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & SpcFill(!clave_sector, 2) & "  " & _
                              SpcFill(!nombre_cuenta_estandar, 50) & "  " & _
                              SpcFillAmount(!saldo_inicial, 22) & "  " & _
                              SpcFillAmount(!Cargos, 22) & "  " & _
                              SpcFillAmount(!Abonos, 22) & "  " & _
                              SpcFillAmount(!Saldo_Actual, 22)
          AddMoney tAux, !saldo_inicial, !Cargos, !Abonos, !Saldo_Actual
          AddMoney tMoneda, !saldo_inicial, !Cargos, !Abonos, !Saldo_Actual
          nLine = nLine + 1
  
          .MoveNext
        Loop
        PrintQtyLines oTextFile, 102, 80
        oTextFile.PrintLine Space(58) & "TOTAL POR AUXILIAR  " & "  " & _
                            SpcFillAmount(tAux.SI, 22) & "  " & _
                            SpcFillAmount(tAux.Cargos, 22) & "  " & _
                            SpcFillAmount(tAux.Abonos, 22) & "  " & _
                            SpcFillAmount(tAux.SA, 22)
        
        PrintQtyLines oTextFile, 102, 80
        oTextFile.PrintLine Space(58) & "TOTAL POR MONEDA    " & "  " & _
                            SpcFillAmount(tMoneda.SI, 22) & "  " & _
                            SpcFillAmount(tMoneda.Cargos, 22) & "  " & _
                            SpcFillAmount(tMoneda.Abonos, 22) & "  " & _
                            SpcFillAmount(tMoneda.SA, 22)
        
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    oRs.Close
    Set oRs = Nothing
    Report113 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report113", Err.Number
End Function

Public Function Report115(sAppServer As String, sParticipantType As String, Optional ByVal dDate As Variant, _
                          Optional ByVal sSQLOrder As String, Optional ByVal sStatus As String) As String
  Dim oRs               As Recordset
  Dim oRSGenerator      As New SCFFixedReportsDS.CServer
  Dim oTextFile         As New CTextFile
  Dim sDateTime         As String
  Dim nLine             As Long
  Dim nPage             As Long
  Dim dTemp                 As Date
  Dim sParticipantTypeDesc  As String
  Dim nParticipantCount     As Long
  '**********************************************************************************************************
  On Error GoTo ErrorHandler
  oTextFile.GenerateRandomFile
  oTextFile.Open_
  
  If IsMissing(dDate) Then
    dTemp = Date
  Else
    dTemp = CDate(dDate)
  End If
  
  If sParticipantType = "U" Then
    sParticipantTypeDesc = "USUARIOS"
  ElseIf (sParticipantType = "O") Then
    sParticipantTypeDesc = "ORGANIZACIONES"
  ElseIf (sParticipantType = "S") Then
    sParticipantTypeDesc = "SISTEMAS"
  Else
    sParticipantTypeDesc = "ROLES"
  End If
  
  sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  nLine = 100
  nPage = 1
  Set oRs = oRSGenerator.GeneraRecordset115(sAppServer, sParticipantType, dTemp, sSQLOrder, sStatus)
          ' , bopen
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not .EOF()
          PrintHeader115 oTextFile, nLine, nPage, sDateTime, sParticipantTypeDesc
          If (!Status = "A") Then
            sStatus = "Activo"
          ElseIf (!Status = "S") Then
            sStatus = "Suspendido"
          Else
            sStatus = "Borrado"
          End If
          oTextFile.PrintLine SpcFill(!ParticipantKey, 24) & "   " & SpcFill(!ParticipantName, 64) & "   " & _
                              SpcFill(sStatus, 10)
          nParticipantCount = nParticipantCount + 1
          nLine = nLine + 1
          .MoveNext
        Loop
        oTextFile.PrintNewLine
        oTextFile.PrintLine "=========================================================================================================="
        oTextFile.PrintLine "TOTAL DE " & sParticipantTypeDesc & ": " & nParticipantCount
        .Close
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Set oRs = Nothing
  
  Report115 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report115", Err.Number
End Function

Public Function Report116(sAppServer As String, nStdAccountTypeId As Long, Optional ByVal dDate As Variant, _
                          Optional ByVal sFromAccount As String, Optional ByVal sToAccount As String) As String
                          
  Dim oRs            As Recordset
  Dim oRSGenerator   As New SCFFixedReportsDS.CServer
  Dim oTextFile      As New CTextFile
  Dim sDateTime      As String
  Dim nLine          As Long
  Dim nPage          As Long
  Dim dTemp          As Date
  Dim sRolDesc       As String
  '**************************************
  On Error GoTo ErrorHandler
  oTextFile.GenerateRandomFile
  oTextFile.Open_
  
  If IsMissing(dDate) Then
    dTemp = Date
  Else
    dTemp = dDate
  End If
  sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  nLine = 100
  nPage = 1
     
    Set oRs = oRSGenerator.GeneraRecordset116(sAppServer, nStdAccountTypeId, dTemp, sFromAccount, sToAccount)
      ' , bopen
 
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not .EOF()
           PrintHeader116 oTextFile, nLine, nPage, nStdAccountTypeId, sDateTime
           If (!rol_cuenta = "S") Then
             sRolDesc = "Sumaria"
           ElseIf (!rol_cuenta = "P") Then
             sRolDesc = "Detalle"
           ElseIf (!rol_cuenta = "C") Then
             sRolDesc = "Control"
           End If
           
           oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & SpcFill(!nombre_cuenta_estandar, 50) & "  " & _
                               SpcFill(!clave_sector, 10) & "   " & SpcFill(!naturaleza, 17) & " " & SpcFill(sRolDesc, 15) & "  " & _
                               SpcFill(!tipo_cuenta, 18)
           nLine = nLine + 1
           .MoveNext
           Loop
         .Close
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Set oRs = Nothing
    Report116 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
   RaiseError ClassId, "Report116", Err.Number
End Function

Public Function Report117(sAppServer As String, nStdAccountTypeId As Long, sSQLOrder As String) As String
  Dim oRs              As Recordset
  Dim oRSGenerator     As New SCFFixedReportsDS.CServer
  Dim oTextFile        As New CTextFile
  Dim sDateTime        As String
  Dim nLine            As Long
  Dim nPage            As Long
  Dim sGralLedgerGroup As String
  Dim nGralLedGerCount As Long
  '*******************************************************************************************************
  On Error GoTo ErrorHandler
  oTextFile.GenerateRandomFile
  oTextFile.Open_
  sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  nLine = 100
  nPage = 1
    Set oRs = oRSGenerator.GeneraRecordset117(sAppServer, nStdAccountTypeId, sSQLOrder)
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not .EOF()
          PrintHeader117 oTextFile, nLine, nPage, sDateTime, sSQLOrder
          If (sSQLOrder = "") Then
            If (sGralLedgerGroup <> !NOMBRE_GRUPO_MAYOR) Or (nLine = 12) Then
              oTextFile.PrintLine "      GRUPO: " & SpcFill(!NOMBRE_GRUPO_MAYOR, 96)
              oTextFile.PrintNewLine
              sGralLedgerGroup = !NOMBRE_GRUPO_MAYOR
              nLine = nLine + 2
            End If
            oTextFile.PrintLine SpcFill(!NUMERO_MAYOR, 10) & " " & SpcFill(!NOMBRE_MAYOR, 116)
            nLine = nLine + 1
          Else
            oTextFile.PrintLine SpcFill(!NUMERO_MAYOR, 10) & " " & SpcFill(!NOMBRE_MAYOR, 60) & "  " & _
                                SpcFill(Trim(Mid$(!NOMBRE_GRUPO_MAYOR, InStr(!NOMBRE_GRUPO_MAYOR, " "))), 58)
            nLine = nLine + 1
          End If
          
          nGralLedGerCount = nGralLedGerCount + 1
          
          .MoveNext
          If Not (.EOF) Then
            If (sSQLOrder = "") Then
              If (sGralLedgerGroup <> !NOMBRE_GRUPO_MAYOR) Then
                oTextFile.PrintNewLine
                oTextFile.PrintLine "    TOTAL DE MAYORES: " & nGralLedGerCount
                oTextFile.PrintNewLine
                oTextFile.PrintNewLine
                nGralLedGerCount = 0
                nLine = nLine + 4
              End If
            End If
          End If
        Loop
        oTextFile.PrintNewLine
        oTextFile.PrintLine "    TOTAL DE MAYORES: " & nGralLedGerCount
        .Close
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    Set oRs = Nothing
    Report117 = oTextFile.FileName
  oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report117", Err.Number
End Function

Public Function Report126(sAppServer As String, dCuteDate As Date, vGL As Variant) As String
  Dim oRs             As ADODB.Recordset
  Dim oRSGenerator    As New SCFFixedReportsDS.CServer
  Dim oTextFile       As New CTextFile
  
  Dim nLine           As Long
  Dim nPage           As Long
  
  Dim nTTot_Base      As Currency
  Dim nTTot_Base_IVA  As Currency
  Dim nTTot_RET       As Currency
  Dim nTTot_IVA       As Currency
  
  Dim sGl             As String
  Dim nGl             As Long
  Dim sDTime          As String
  '********************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRS_5031(sAppServer, dCuteDate, vGL)
    sDTime = GetFmtDate
    
    If (IsArray(vGL)) Then
      sGl = "AGENCIA CONSOLIDADORA BANOBRAS"
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NOMBRE_MAYOR", "ID_MAYOR")
    End If
    
    nLine = 100
    nPage = 1
    
    If ((oRs.BOF) And (oRs.EOF)) Then
      PrintHeader126 oTextFile, nLine, nPage, dCuteDate, sGl, sDTime, "5031"
      oTextFile.PrintLine " NO EXISTE INFORMACION "
      Exit Function
    End If
    
    With oRs
      Do While Not .EOF()
        PrintHeader126 oTextFile, nLine, nPage, dCuteDate, sGl, sDTime, "5031"
        oTextFile.PrintLine SpcFill(!RFC, 15) & "  " & _
                            SpcFill(!CURP, 19) & "  " & _
                            SpcFill(!nombre_cuenta_auxiliar, 35) & "  " & _
                            SpcFill(!CLAVE_RETENCION, 9) & "  " & _
                            SpcFillAmount(!Tot_Base, 16) & "  " & _
                            SpcFillAmount(!Tot_Base_IVA, 16) & "  " & _
                            SpcFillAmount(!Tot_RET, 16) & "  " & _
                            SpcFillAmount(!Tot_IVA, 16)
        nLine = nLine + 1
        
        nTTot_Base = nTTot_Base + !Tot_Base
        nTTot_Base_IVA = nTTot_Base_IVA + !Tot_Base_IVA
        nTTot_RET = nTTot_RET + !Tot_RET
        nTTot_IVA = nTTot_IVA + !Tot_IVA

        .MoveNext
      Loop
      .Close
      
      PrintHeader126 oTextFile, nLine, nPage, dCuteDate, sGl, sDTime, "5031"
      
      oTextFile.PrintLine "========================================================================================================================================================================================="
      oTextFile.PrintLine SpcFill("   TOTAL ", 84) & "  " & _
                          SpcFillAmount(nTTot_Base, 16) & "  " & _
                          SpcFillAmount(nTTot_Base_IVA, 16) & "  " & _
                          SpcFillAmount(nTTot_RET, 16) & "  " & _
                          SpcFillAmount(nTTot_IVA, 16)
   End With
   Report126 = oTextFile.FileName
   oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report126", Err.Number
End Function

Public Function Report127(sAppServer As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, sOptionToDisplay As String) As String
  Dim oBalance      As New EFABalanceEngine.CEngine
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oRs           As Recordset
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAccType      As Long
  Dim sAccType      As String
  Dim nTotal        As Double
  '***************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oBalance.BalancesRS(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGL, True, True, False, dIniDate, dEndDate, _
                                  "5201", "5201-99-99-99-99-99-99", nAccountType:=6)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
    nAccType = oRSGenerator.AccountType(sAppServer, nGl)
    sAccType = oRSGenerator.GetObjName(sAppServer, nAccType, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
    
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader127 oTextFile, nLine, nPage, sGl, dEndDate, "DETERMINACION DE LA BASE DEL IVA", sDTime, sAccType, sOptionToDisplay
          
          If (sOptionToDisplay = "M") Then
            oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & SpcFill(!clave_sector, 2) & "  " & _
                                SpcFill(!nombre_cuenta_estandar, 55) & "  " & _
                                SpcFillAmount1(!Saldo_Actual - !saldo_inicial, 22)
            nTotal = nTotal + (!Saldo_Actual - !saldo_inicial)
          Else
            oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & SpcFill(!clave_sector, 2) & "  " & _
                                SpcFill(!nombre_cuenta, 55) & "  " & _
                                SpcFillAmount1(!Saldo_Actual, 22)
            nTotal = nTotal + !Saldo_Actual
          End If
          nLine = nLine + 1
          .MoveNext
        Loop
        PrintQtyLines oTextFile, 112
        oTextFile.PrintLine "TOTAL  " & Space(78) & _
                            SpcFillAmount(nTotal, 22)
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
      End If
    End With
    oRs.Close
    Set oRs = Nothing
    Report127 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report127", Err.Number
End Function

Public Function Report128(sAppServer As String, dCuteDate As Date, dExcDate As Date, _
                          sAccounts As String, vGL As Variant, _
                          Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As String
  Dim oGenRS_A          As SCFFixedReportsDS.CServer
  Dim oTextFile         As New CTextFile
  Dim oBalance_A        As SCFFixedReportsDS.CBalances
  Dim oRS_A             As Object
  Dim sDateTime         As String
  Dim sDate             As String
  
  Dim nExchanges        As Double
  Dim nExcByCurr        As Double
  
  Dim sAcc              As String
  Dim nLine             As Long
  Dim nPage             As Long
  
  Dim nSum              As Double
  Dim nSumMN            As Double
  Dim nMn               As Double
  Dim nTot              As Double
  Dim nSaldo_A          As Double
  Dim sCurrencyName     As String
  
  Dim dInitialDate      As Date
  '**************************************
  On Error GoTo ErrorHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    sDateTime = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
    sDate = Format(dCuteDate, "d"" de ""MMMM"" de ""yyyy")
    dInitialDate = dCuteDate - CLng(Day(dCuteDate) - 1)
    
    nLine = 100
    nPage = 1
    Set oGenRS_A = New SCFFixedReportsDS.CServer
    Set oBalance_A = New SCFFixedReportsDS.CBalances
    nExchanges = oGenRS_A.GetExchangeRate(sAppServer, 2, nExcRateType, 1, dExcDate)
   
    If (sAccounts = 5011) Then
      sAcc = 5010
    Else
      sAcc = 5050
    End If
   
    Set oRS_A = CreateObject("ADODB.Recordset")
    Set oRS_A = oBalance_A.BalanceGeneral(sAppServer, "&&&&", vGL, dInitialDate, dCuteDate, sAccounts, sAccounts, bOnlyLastLevels:=True)
    
    With oRS_A
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
            PrintHeader128 oTextFile, nLine, nPage, sAccounts, sDate, sDateTime
            If (Len(!NUMERO_CUENTA_ESTANDAR) = 4) And (!clave_sector = "00") And (!NUMERO_CUENTA_ESTANDAR = sAccounts) Then
              sCurrencyName = oGenRS_A.GetObjName(sAppServer, !ID_MONEDA, "ao_currencies", "currency_name", "currency_id")
              nExcByCurr = oGenRS_A.GetExchangeRate(sAppServer, !ID_MONEDA, nExcRateType, 1, dExcDate)
              oTextFile.PrintLine SpcFill(!o_id_moneda, 8) & "  " & SpcFill(sCurrencyName, 30) & "  " & _
                                  SpcFillAmount1(!Saldo_Actual, 22) & "     " & Space(11) & _
                                  SpcFillAmount((nExcByCurr / nExchanges), 5, 6) & " " & _
                                  SpcFillAmount1((!Saldo_Actual * (nExcByCurr / nExchanges)), 33)
              nSum = nSum + ((!Saldo_Actual * nExcByCurr) / nExchanges)
              nSumMN = nSumMN + (!Saldo_Actual * nExcByCurr)
              nLine = nLine + 1
            End If
          .MoveNext
        Loop
        .Close
        .ActiveConnection = Nothing
        
        nSaldo_A = BalanceAccount(sAppServer, dCuteDate, dExcDate, sAcc, sAcc, vGL, "&&&&", nExcRateType, 1, True)
        nMn = nSum * nExchanges
        nTot = nSaldo_A - nSumMN
        oTextFile.PrintLine Space(85) & "  ______________________________________ "
        oTextFile.PrintLine "Posición larga (corta)                    " & SpcFillAmount1(nSum, 82)
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(52) & "T.C.(**)  " & Space(54) & SpcFillAmount(nExchanges, 3, 6)
        oTextFile.PrintLine Space(85) & "  ______________________________________ "
        oTextFile.PrintNewLine
        oTextFile.PrintLine "Valorización en Moneda Nacional                               " & SpcFillAmount1(nSumMN, 62)
        oTextFile.PrintNewLine
        oTextFile.PrintLine "Saldo cta. " & sAcc & " al " & sDate & "              " & SpcFillAmount1(nSaldo_A, 71)
        oTextFile.PrintLine Space(85) & "  ______________________________________ "
        oTextFile.PrintNewLine
        oTextFile.PrintNewLine
        oTextFile.PrintLine "Utilidad o (Pérdida) cambiaria por operaciones propias                   " & SpcFillAmount1(nTot, 51)
        oTextFile.PrintLine Space(85) & "  ====================================== "
      End If
    End With
    Set oRS_A = Nothing
    Set oBalance_A = Nothing
    Set oGenRS_A = Nothing
    
    Report128Medium oTextFile, sAppServer, dCuteDate, dExcDate, vGL, nExcRateType, nExcRateCurrency
    Report128End oTextFile, sAppServer, dInitialDate, dCuteDate, dExcDate, vGL, nExcRateType, nExcRateCurrency, nExchanges
   
    Report128 = oTextFile.FileName
    'oTextFile.FilePath
    oTextFile.Close_
  Exit Function
ErrorHandler:
  RaiseError ClassId, "Report128", Err.Number
End Function

Private Sub Report128Medium(oTextFile_B As CTextFile, sAppServer_B As String, dCuteDate As Date, dExcDate As Date, _
                           vGL_B As Variant, Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1)
  
  Dim nSaldo_2          As Double
  Dim nSaldo_3          As Double
  Dim nSaldo_4          As Double
  Dim nSaldo_5          As Double
  Dim nSaldo_6          As Double
  '**************************************
  On Error GoTo ErrorHandler
    
    oTextFile_B.PrintLine "INTEGRACIÓN DE LA UTILIDAD O (PÉRDIDA) CAMBIARIA"
    oTextFile_B.PrintNewLine
    nSaldo_2 = BalanceAccount(sAppServer_B, dCuteDate, dExcDate, "5205-01", "5205-01", vGL_B, "&&&&-&&", nExcRateType, 1, True, True)
    
    nSaldo_3 = BalanceAccount(sAppServer_B, dCuteDate, dExcDate, "5205-04", "5205-04-99-99-99-99-99", vGL_B, "&&&&-&&", nExcRateType, 1, False, True)

    nSaldo_4 = nSaldo_2 + nSaldo_3
    nSaldo_5 = BalanceAccount(sAppServer_B, dCuteDate, dExcDate, "5205-99", "5205-99-99-99-99-99-99", vGL_B, "&&&&-&&", nExcRateType, 1, False, True)
    nSaldo_6 = nSaldo_4 + nSaldo_5
    oTextFile_B.PrintLine "Utilidad (Pérdida) por compraventa física de divisas      " & Space(5) & SpcFillAmount1(nSaldo_2, 22)
    oTextFile_B.PrintLine "Utilidad (Pérdida) por valorización de divisas            " & Space(5) & SpcFillAmount1(nSaldo_3, 22)
    oTextFile_B.PrintLine "                                                          " & Space(5) & "--------------------------"
    oTextFile_B.PrintLine "  Subtotal                                                " & Space(5) & SpcFillAmount1(nSaldo_4, 22)
    oTextFile_B.PrintLine "Reexpresión de saldos por criterio de la Circular No. 1448" & Space(5) & SpcFillAmount1(nSaldo_5, 22) & "  " & SpcFillAmount1(nSaldo_6, 22)
    oTextFile_B.PrintLine "                                                          " & Space(5) & "--------------------------" & "  " & "=========================="
  
    oTextFile_B.PrintNewLine
    oTextFile_B.PrintLine "(*) La equivalencia de divisas a dólares americanos, es la que dio a conocer el Banco de Mexico el " & Format$(dExcDate, "d"" de ""MMMM"" de ""yyyy")
    
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "Report128Medium", Err.Number
End Sub


Private Sub Report128End(oTextFile_C As CTextFile, sAppServer_C As String, dIniDate As Date, dCuteDate As Date, _
                        dExcDate As Date, vGL_C As Variant, _
                        Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1, _
                        Optional nExchanges As Double)
  Dim oGeneraRS_C       As SCFFixedReportsDS.CServer
  Dim oBal_C            As SCFFixedReportsDS.CBalances
  Dim oRs_C             As ADODB.Recordset
  Dim sDate             As String
  
  Dim nExcByCurr        As Double
  
  Dim sAcc_C            As String
  Dim sAccounts_C       As String
  
  Dim nSum              As Double
  Dim nSumMN            As Double
  Dim nMn               As Double
  Dim nTot              As Double
  Dim nSaldo_C          As Double
  Dim sCurrencyName     As String
  
  '**************************************
  On Error GoTo ErrorHandler
    Set oGeneraRS_C = New SCFFixedReportsDS.CServer
    Set oBal_C = New SCFFixedReportsDS.CBalances
    sDate = Format(dCuteDate, "d"" de ""MMMM"" de ""yyyy")
   
    sAccounts_C = 5051
    
    oTextFile_C.PrintNewLine
    oTextFile_C.PrintNewLine
    oTextFile_C.PrintNewLine
    oTextFile_C.PrintLine "     Posición del mecanismo FOAEM en monedas extranjeras con saldos al " & sDate
    oTextFile_C.PrintNewLine
    oTextFile_C.PrintLine "                                                                          Equivalencia                   Valuación en     "
    oTextFile_C.PrintLine "         MONEDA                                       " & sAccounts_C & "              dólar americano(*)             dólares americanos   "
    oTextFile_C.PrintLine "        _______                                       ____             __________________               _____________________ "
    
    
    If (sAccounts_C = 5011) Then
      sAcc_C = 5010
    Else
      sAcc_C = 5050
    End If
    
    Set oRs_C = oBal_C.BalanceGeneral(sAppServer_C, "&&&&", vGL_C, CDate("02/01/1990"), dCuteDate, _
                                      sAccounts_C, sAccounts_C & "-99-99-99-99-99-99", bOnlyLastLevels:=False)
    With oRs_C
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
          If (Len(!NUMERO_CUENTA_ESTANDAR) = 4) And (!clave_sector = "00") And (!NUMERO_CUENTA_ESTANDAR = sAccounts_C) Then
            sCurrencyName = oGeneraRS_C.GetObjName(sAppServer_C, !ID_MONEDA, "ao_currencies", "currency_name", "currency_id")
            nExcByCurr = oGeneraRS_C.GetExchangeRate(sAppServer_C, !ID_MONEDA, nExcRateType, 1, dExcDate)
            oTextFile_C.PrintLine SpcFill(!o_id_moneda, 8) & "  " & SpcFill(sCurrencyName, 30) & "  " & _
                                  SpcFillAmount1(!Saldo_Actual, 22) & "     " & Space(11) & _
                                  SpcFillAmount((nExcByCurr / nExchanges), 5, 6) & " " & _
                                  SpcFillAmount1((!Saldo_Actual * (nExcByCurr / nExchanges)), 33)
            nSum = nSum + ((!Saldo_Actual * nExcByCurr) / nExchanges)
            nSumMN = nSumMN + (!Saldo_Actual * nExcByCurr)
          End If
          .MoveNext
        Loop
        .Close
        .ActiveConnection = Nothing
        
        nSaldo_C = BalanceAccount(sAppServer_C, dCuteDate, dExcDate, sAcc_C, sAcc_C & "-99-99-99-99-99-99", vGL_C, "&&&&", nExcRateType, 1, 0, False)
        nMn = nSum * nExchanges
        nTot = nSaldo_C - nSumMN
        oTextFile_C.PrintLine Space(85) & "  ______________________________________ "
        oTextFile_C.PrintLine "Posición larga (corta)                    " & SpcFillAmount1(nSum, 82)
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintLine Space(52) & "T.C.(**)  " & Space(54) & SpcFillAmount(nExchanges, 3, 6)
        oTextFile_C.PrintLine Space(85) & "  ______________________________________ "
        oTextFile_C.PrintNewLine
        
        oTextFile_C.PrintLine "Valorización en Moneda Nacional                               " & SpcFillAmount1(nSumMN, 62)
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintLine "Total cta. " & sAcc_C & " Operaciones Pasivas de Mecanismos FOAEM            " & SpcFillAmount1(nSaldo_C, 57)
        oTextFile_C.PrintLine Space(85) & "  ______________________________________ "
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintLine "Utilidad o (Pérdida) cambiaria por operaciones del mecanismo FOAEM              " & SpcFillAmount1(nTot, 44)
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintNewLine
        oTextFile_C.PrintLine " * Con base a cotización de divisas extranjeras a la compra del día " & sDate
        oTextFile_C.PrintLine " **Tipo de cambio interbancario a la compra del día " & sDate
      End If
    End With
    Set oBal_C = Nothing
    Set oGeneraRS_C = Nothing
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "Report128End", Err.Number
End Sub

Public Function Report144(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset144(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader144 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 50) & "     " & _
                              SpcFill(IIf(IsNull(!TipoTasa), "", !TipoTasa), 10) & "  " & _
                              SpcFillAmount(IIf(IsNull(!Tasa), 0, !Tasa), 5) & "  " & _
                              SpcFillAmount(IIf(IsNull(!MontoCreditoReestruct), 0, !MontoCreditoReestruct), 20) & "      " & _
                              SpcFill(IIf(IsNull(!FechaInicioAmort), "", !FechaInicioAmort), 10) & "   " & _
                              SpcFill(IIf(IsNull(!FechaFinAmort), "", !FechaFinAmort), 10)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report144 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report144", Err.Number
End Function

Public Function Report149(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                    sPattern As String, dIniDate As Date, dEndDate As Date, _
                                    dIniDate2 As Date, dEndDate2 As Date, bPrintInCascade As Boolean, _
                                    Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                                    Optional ByVal nTransactionTypeId As Long = 0, _
                                    Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = True) As String
  Dim oRs               As Object
  Dim oRSGenerator      As New SCFFixedReportsDS.CServer
  Dim oBalance          As SCFFixedReportsDS.CBalances
  Dim oTextFile         As New CTextFile
  Dim nLine             As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon              As String
  Dim tAcumulador       As TMoneyComparative
  Dim tDeudoras         As TMoneyComparative
  Dim tAcreedoras       As TMoneyComparative
  Dim tGroup            As TMoneyComparative
  Dim tbalance          As TMoneyComparative
  Dim tMoneda           As TMoneyComparative
  Dim sGl               As String
  Dim sGLNumber         As String
  Dim sAccType          As String
  Dim sCurrency         As String
  Dim nPercent          As Currency
  Dim nSaldo_A          As Double
  Dim nSaldo_B          As Double
  
  Dim vGL2              As Variant
  Dim vGLToSend         As Variant
  Dim i                 As Long
  Dim sReportTitle      As String
  Dim sAccIncludes      As String
  Dim sTransactionType  As String
  Dim sVoucherType      As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    If sIniAccount = "" Then
      sIniAccount = "5232-01-01"
      sEndAccount = "5232-01-01"
    End If
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
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL2(i)
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sGLNumber = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        End If
      End If
      sReportTitle = "COMPARATIVO DE CUENTAS POR AUXILIARES"
      Set oRs = CreateObject("ADODB.Recordset")
      Set oRs = oBalance.Comparative(sAppServer, sPattern, vGLToSend, dIniDate, dEndDate, dIniDate2, dEndDate2, sIniAccount, sEndAccount, _
                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                     nOptionToDisplay, bWithAux)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sNat = !naturaleza
          sMon = !o_id_moneda
          sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
          sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
          sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
          Do While Not (.EOF)
            PrintHeader149 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                           dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, _
                           sTransactionType, sVoucherType, nExcRateCurrency, True, True, sIniAccount
            If (sMon <> !o_id_moneda) Then
              sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
              sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
            
              oTextFile.PageFeed nLine
              PrintHeader149 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                             dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, sTransactionType, _
                             sVoucherType, nExcRateCurrency, True, True, sIniAccount
            End If
            nSaldo_A = (!Saldo_Actual - !saldo_inicial)
            nSaldo_B = (!SALDO_ACTUAL_2 - !saldo_inicial_2)
            If (nSaldo_A <> 0) Then
              nPercent = ((nSaldo_B / nSaldo_A) - 1) * 100
            Else
              nPercent = 0
            End If
            If (nSaldo_A <> 0 Or nSaldo_B <> 0) Then
              oTextFile.PrintLine SpcFill(sGLNumber, 6) & "  " & SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "  " & _
                                 SpcFill(!nombre_cuenta_auxiliar, 40) & _
                                 SpcFillAmount(nSaldo_A, 22) & "  " & SpcFillAmount(nSaldo_B, 22) & "  " & _
                                 SpcFillAmount(nSaldo_B - nSaldo_A, 22) & _
                                 "  " & SpcFillAmount(nPercent, 22)
              nLine = nLine + 1
              AddMoneyComparative tbalance, CCur(nSaldo_A), CCur(nSaldo_B)
            End If
            .MoveNext
          Loop
        Else
          PrintHeader149 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                         dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, sTransactionType, _
                         sVoucherType, nExcRateCurrency, True, True, sIniAccount
        End If
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(30) & "  " & SpcFill("Totales : ", 40) & _
                            SpcFillAmount(tbalance.SA, 22) & "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                            "  " & SpcFillAmount(tbalance.SA_2 - tbalance.SA, 22)
                            
        If bPrintInCascade Then
          oTextFile.PageFeed nLine
          ClearMoneyComparative tbalance
        Else
          Exit For
        End If
      End With
      oRs.Close
      Set oRs = Nothing
      Set oBalance = Nothing
      nSaldo_A = 0
      nSaldo_B = 0
      nPercent = 0
    Next i
    Report149 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report149", Err.Number
End Function

Public Function Report150(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                    sPattern As String, dIniDate As Date, dEndDate As Date, _
                                    dIniDate2 As Date, dEndDate2 As Date, _
                                    Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                                    Optional ByVal nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, Optional bPrintInCascade As Boolean = True, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = True) As String
  Dim oRs           As Object
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer, oBalance As SCFFixedReportsDS.CBalances
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tAcumulador   As TMoneyComparative
  Dim tDeudoras     As TMoneyComparative
  Dim tAcreedoras   As TMoneyComparative
  Dim tGroup        As TMoneyComparative
  Dim tbalance      As TMoneyComparative
  Dim tMoneda       As TMoneyComparative
  Dim sGl           As String
  Dim sGLNumber     As String
  Dim sAccType      As String
  Dim sCurrency     As String
  Dim nPercent      As Currency
  Dim nSaldo_A      As Double
  Dim nSaldo_B      As Double
  
  Dim vGL2             As Variant
  Dim vGLToSend        As Variant
  Dim i                As Long
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
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
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL2(i)
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sGLNumber = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        End If
      End If
      sReportTitle = "COMPARATIVO DE CUENTAS POR AUXILIARES"
      Set oRs = CreateObject("ADODB.Recordset")
      Set oRs = oBalance.Comparative(sAppServer, sPattern, vGLToSend, dIniDate, dEndDate, dIniDate2, dEndDate2, sIniAccount, sEndAccount, _
                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                     nOptionToDisplay, bWithAux)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sNat = !naturaleza
          sMon = !o_id_moneda
          sGpo = Mid(!NUMERO_CUENTA_ESTANDAR, 1, 2)
          sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
          sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
          Do While Not (.EOF)
            PrintHeader150 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                           dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, _
                           sTransactionType, sVoucherType, nExcRateCurrency, True, True, sIniAccount
            nSaldo_A = (!Saldo_Actual - !saldo_inicial)
            nSaldo_B = (!SALDO_ACTUAL_2 - !saldo_inicial_2)
            If (nSaldo_A <> 0) Then
              nPercent = ((nSaldo_B / nSaldo_A) - 1) * 100
            Else
              nPercent = 0
            End If
            If (nSaldo_A <> 0 Or nSaldo_B <> 0) Then
              oTextFile.PrintLine SpcFill(sGLNumber, 6) & "  " & SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "  " & _
                                 SpcFill(!o_id_moneda, 4) & "  " & SpcFill(GetCuentaSubCuenta(!NUMERO_CUENTA_ESTANDAR), 32) & "  " & _
                                 SpcFillAmount(nSaldo_A, 22) & "  " & SpcFillAmount(nSaldo_B, 22) & "  " & _
                                 SpcFillAmount(nSaldo_B - nSaldo_A, 22) & _
                                 "  " & SpcFillAmount(nPercent, 22)
              nLine = nLine + 1
              AddMoneyComparative tbalance, CCur(nSaldo_A), CCur(nSaldo_B)
            End If
            .MoveNext
          Loop
        Else
        
          PrintHeader150 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                         dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, sTransactionType, _
                         sVoucherType, nExcRateCurrency, True, True, sIniAccount
        End If
        oTextFile.PrintNewLine
        oTextFile.PrintLine Space(30) & "  " & SpcFill("Totales : ", 40) & _
                            SpcFillAmount(tbalance.SA, 22) & "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                            "  " & SpcFillAmount(tbalance.SA_2 - tbalance.SA, 22)
        If bPrintInCascade Then
          oTextFile.PageFeed nLine
          ClearMoneyComparative tbalance
        Else
          Exit For
        End If
      End With
      oRs.Close
      Set oRs = Nothing
      Set oBalance = Nothing
    Next i
    Report150 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report150", Err.Number
End Function

Public Function Report151(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                          sPattern As String, dIniDate As Date, dEndDate As Date, dExcDate As Date, _
                          sAccount As String, sAccount2 As String, _
                          Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                          Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                          Optional nCurrencyType As Long = 0, _
                          Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1, _
                          Optional bPrintInCascade As Boolean = False, _
                          Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False) As String
  Dim oRs           As Object
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer, oBalance As SCFFixedReportsDS.CBalances
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  Dim tAcumulador   As TMoneyComparative
  Dim tDeudoras     As TMoneyComparative
  Dim tAcreedoras   As TMoneyComparative
  Dim tGroup        As TMoneyComparative
  Dim tbalance      As TMoneyComparative
  Dim tMoneda       As TMoneyComparative
  Dim sGl           As String
  Dim sGLNumber     As String
  Dim sAccType      As String
  Dim sCurrency     As String
  Dim nPercent      As Currency
  Dim nSaldo        As Double
  Dim nSaldo_2      As Double
  Dim nExcType      As Double
  
  Dim vGL2             As Variant
  Dim vGLToSend        As Variant
  Dim i                As Long
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    If (sAccount2 = "") Then
      sAccount2 = "5011-99"
    End If
    
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
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL2(i)
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sGLNumber = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        Set oBalance = New SCFFixedReportsDS.CBalances
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        End If
      End If
      sReportTitle = "COMPARATIVO DE CUENTAS POR AUXILIARES"
      Set oRs = CreateObject("ADODB.Recordset")
      Set oRs = oBalance.BalanceCompByAccountRS(sAppServer, sPattern, vGLToSend, dIniDate, dEndDate, sAccount, sAccount2, sIniAccount, sEndAccount, _
                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                     nOptionToDisplay, bWithAux)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          sCurrency = oRSGenerator.GetObjName(sAppServer, !ID_MONEDA, "AO_CURRENCIES", "O_ID_MONEDA", "CURRENCY_ID", "CURRENCY_NAME")
          sCurrency = Mid(sCurrency, 3, Len(sCurrency) - 2)
          Do While Not (.EOF)
            PrintHeader151 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                                      sAccount, sAccount2, sReportTitle, sDTime, sAccIncludes, sAccType, _
                                       sTransactionType, sVoucherType, nExcRateCurrency, True, True, !exchange_rate_2
                                       
            nSaldo = (!Saldo_Actual - !saldo_inicial)
            nSaldo_2 = (!SALDO_ACTUAL_2 - !saldo_inicial_2)
            If (nSaldo <> 0) Then
              nPercent = ((nSaldo_2 / nSaldo) - 1) * 100
            Else
              nPercent = 0
            End If
            If (nSaldo_2 <> 0) Then
              nExcType = (nSaldo / nSaldo_2) * !exchange_rate_2
            Else
              nExcType = 0
            End If
            If (nSaldo <> 0 Or nSaldo_2 <> 0) Then
              oTextFile.PrintLine SpcFill(sGLNumber, 6) & "  " & SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "  " & SpcFill(!nombre_cuenta_auxiliar, 40) & _
                                 SpcFillAmount(nSaldo, 22) & "  " & SpcFillAmount(nSaldo_2, 22) & "  " & SpcFillAmount((nSaldo_2 - nSaldo), 22) & _
                                 "  " & SpcFillAmount(nPercent, 14) & "  " & SpcFillAmount(nExcType, 11, 7)
              nLine = nLine + 1
              AddMoneyComparative tbalance, CCur(nSaldo), CCur(nSaldo_2)
            End If
            .MoveNext
          Loop
          oTextFile.PrintNewLine
          oTextFile.PrintLine Space(30) & "  " & SpcFill("Totales : ", 40) & _
                              SpcFillAmount(tbalance.SA, 22) & "  " & SpcFillAmount(tbalance.SA_2, 22) & _
                             "  " & SpcFillAmount(tbalance.SA_2 - tbalance.SA, 22)
          If bPrintInCascade Then
            oTextFile.PageFeed nLine
            ClearMoneyComparative tbalance
          Else
            Exit For
          End If
        End If
      End With
      oRs.Close
      Set oRs = Nothing
      Set oBalance = Nothing
    Next i
    Report151 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report151", Err.Number
End Function

Public Function Report152(sAppServer As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         vAccountGroups As Variant, dExcDate As Date, _
                         Optional nExcRateType As Long = 49, _
                         Optional nExcRateCurrency As Long = 1) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sGl           As String
  
  Dim sReportTitle     As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
    Else
      sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
    
      sReportTitle = "ANÁLISIS DE SALDOS PARA VERIFICAR LA VALORIZACIÓN MENSUAL DE LOS FID. Y MOS."
      Set oRs = oRSGenerator.GeneraRecordset152(sAppServer, vGL, dIniDate, dEndDate, vAccountGroups, dExcDate, nCurrencyType:=-1, _
                                                nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency)
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      With oRs
        If Not (.BOF And .EOF) Then
          Do While Not (.EOF)
            PrintHeader152 oTextFile, nLine, nPage, dIniDate, dEndDate, _
                            sReportTitle, sDTime
                                        
            oTextFile.PrintLine SpcFill(!NUMERO_MAYOR_COMPLETO, 12) & "  " & _
                                SpcFill(!o_id_moneda, 4) & "     " & _
                                SpcFillAmount1(!Saldo_Actual, 22) & "     " & _
                                SpcFill(!NUMERO_MAYOR_COMPLETO, 12) & "  " & _
                                SpcFillAmount1(!SALDO_ACTUAL_2, 22) & "  " & _
                                SpcFillAmount(!EXCHANGE_RATE, 12, 6) & "          " & _
                                SpcFillAmount1((!Saldo_Actual * !EXCHANGE_RATE), 22) & "  " & _
                                SpcFillAmount(!SALDO_ACTUAL_2 - (!Saldo_Actual * !EXCHANGE_RATE), 22)
                                
            nLine = nLine + 1
            .MoveNext
          Loop
        Else
          oTextFile.PrintLine "NO SE ENCONTRO INFORMACIÓN"
        End If
        .Close
        Set oRSGenerator = Nothing
      End With
    
    Report152 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report152", Err.Number
End Function

Public Function Report153(sAppServer As String, nGralLedgerGroupId As Long, vGL As Variant, _
                                    dIniDate As Date, dEndDate As Date, _
                                    dIniDate2 As Date, dEndDate2 As Date, _
                                    Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                                    Optional ByVal nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, Optional bPrintInCascade As Boolean = False, _
                                    Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False) As String
  Dim oBalance As New SCFFixedReportsDS.CBalances
  Dim oRs           As Object
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oRsParam      As Recordset
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long, nPage As Long, sDTime As String, sGpo As String, sNat As String
  Dim sMon          As String
  
  Dim tLevel1       As TMoneyComparative
  Dim tLevel2       As TMoneyComparative
  Dim sGl           As String
  Dim sGLNumber     As String
  Dim sAccType      As String
  Dim sCurrency     As String
  Dim nPercent      As Currency
  Dim nSaldo_A      As Double
  Dim nSaldo_B      As Double
  
  Dim vGL2             As Variant
  Dim vGLToSend        As Variant
  Dim i                As Long
  Dim sReportTitle     As String
  Dim sAccIncludes     As String
  Dim sTransactionType As String
  Dim sVoucherType     As String
  
  Dim sAccountName     As String
  Dim nLevel1          As Long
  Dim nLevel2          As Long
  Dim nMoneda          As Long
  Dim sLevel1          As String
  Dim sLevel2          As String

  '*********************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
   
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
    
    Set oRsParam = oRSGenerator.GeneraParametros(sAppServer, 153)
    sReportTitle = "COMPARATIVO DE GASTOS DE OPERACION"
    
    For i = LBound(vGL2) To UBound(vGL2)
      If bPrintInCascade Then
        vGLToSend = vGL2(i)
        sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
        sGLNumber = oRSGenerator.GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR")
        sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
      Else
        vGLToSend = vGL
        If (IsArray(vGL)) Then
          sGl = "TODAS LAS CONTABILIDADES"
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        Else
          sGl = oRSGenerator.GetObjName(sAppServer, CLng(vGL), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
          sAccType = oRSGenerator.GetObjName(sAppServer, nGralLedgerGroupId, "COF_GRUPO_MAYOR", "NOMBRE_GRUPO_MAYOR", "ID_GRUPO_MAYOR")
        End If
      End If
      
      sDTime = GetFmtDate
      oTextFile.PrintNewLine
      nLine = 100
      nPage = 1
      
      With oRsParam
        If Not (.BOF And .EOF) Then
          nLevel1 = !level1
          nLevel2 = !level2
          sLevel1 = IIf(IsNull(!desc_level1), "", !desc_level1)
          sLevel2 = !Desc_Level2
        
          Do While Not .EOF
            PrintHeader153 oTextFile, nLine, nPage, sGl, sCurrency, dIniDate, dEndDate, _
                          dIniDate2, dEndDate2, sReportTitle, sDTime, sAccIncludes, sAccType, _
                          sTransactionType, sVoucherType, nExcRateCurrency, True, True, sIniAccount
            
            If (nLevel1 <> !level1) Then
              oTextFile.PrintNewLine
              If (tLevel1.SA <> 0) Then
                nPercent = ((tLevel1.SA_2 / tLevel1.SA) - 1) * 100
              Else
                nPercent = 0
              End If

              oTextFile.PrintLine Space(28) & "  " & SpcFill(sLevel1, 40) & _
                                SpcFillAmount(tLevel1.SA, 22) & "  " & SpcFillAmount(tLevel1.SA_2, 22) & _
                                "  " & SpcFillAmount(tLevel1.SA_2 - tLevel1.SA, 22) & _
                                SpcFillAmount(nPercent, 22)
               nLevel1 = !level1
               ClearMoneyComparative tLevel1
               oTextFile.PrintNewLine
            End If
            
            If (nLevel2 <> !level2) Then
              oTextFile.PrintNewLine
              If (tLevel2.SA <> 0) Then
                nPercent = ((tLevel2.SA_2 / tLevel2.SA) - 1) * 100
              Else
                nPercent = 0
              End If

              oTextFile.PrintLine Space(28) & "  " & SpcFill(sLevel1, 40) & _
                                SpcFillAmount(tLevel2.SA, 22) & "  " & SpcFillAmount(tLevel2.SA_2, 22) & _
                                "  " & SpcFillAmount(tLevel2.SA_2 - tLevel2.SA, 22) & _
                                SpcFillAmount(nPercent, 22)
               nLevel2 = !level2
               ClearMoneyComparative tLevel2
               oTextFile.PrintNewLine
            End If
            Set oRs = CreateObject("ADODB.Recordset")
            Set oRs = oBalance.BalanceCompLastLevels(sAppServer, vGLToSend, dIniDate, dEndDate, dIniDate2, dEndDate2, _
                                                     !NUMERO_CUENTA_ESTANDAR, !NUMERO_CUENTA_ESTANDAR, _
                                                     nTransactionTypeId, nVoucherTypeId, nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, _
                                                     nOptionToDisplay, bWithAux)
            With oRs
              If Not (.BOF And .EOF) Then
                  nMoneda = !o_id_moneda
                  nSaldo_A = (!Saldo_Actual - !saldo_inicial)
                  nSaldo_B = (!SALDO_ACTUAL_2 - !saldo_inicial_2)
                  If (nSaldo_A <> 0) Then
                    nPercent = ((nSaldo_B / nSaldo_A) - 1) * 100
                  Else
                    nPercent = 0
                  End If
                  If (nSaldo_A <> 0 Or nSaldo_B <> 0) Then
                    oTextFile.PrintLine SpcFill(!o_id_moneda, 4) & "  " & _
                                        SpcFill(!NUMERO_CUENTA_ESTANDAR, 22) & "  " & _
                                        SpcFill(!nombre_cuenta_estandar, 32) & "  " & _
                                        SpcFillAmount(nSaldo_A, 22) & "  " & SpcFillAmount(nSaldo_B, 22) & "  " & _
                                        SpcFillAmount(nSaldo_B - nSaldo_A, 22) & _
                                        SpcFillAmount(nPercent, 22)
                    nLine = nLine + 1
                    AddMoneyComparative tLevel1, CCur(nSaldo_A), CCur(nSaldo_B)
                    AddMoneyComparative tLevel2, CCur(nSaldo_A), CCur(nSaldo_B)
                  End If
              Else
                oTextFile.PrintLine SpcFill(1, 4) & "  " & _
                                    SpcFill(oRsParam!NUMERO_CUENTA_ESTANDAR, 22) & "  " & _
                                    SpcFill(" ", 32) & "  " & _
                                    SpcFillAmount(0, 22) & "  " & SpcFillAmount(0, 22) & "  " & _
                                    SpcFillAmount(0, 22) & _
                                    SpcFillAmount(0, 22)
                nLine = nLine + 1

              End If
            End With
            oRs.Close
            Set oRs = Nothing
            
            .MoveNext
          Loop
          oTextFile.PrintNewLine
          If (tLevel1.SA <> 0) Then
            nPercent = ((tLevel1.SA_2 / tLevel1.SA) - 1) * 100
          Else
            nPercent = 0
          End If

          oTextFile.PrintLine Space(28) & "  " & SpcFill(sLevel1, 40) & _
                            SpcFillAmount(tLevel1.SA, 22) & "  " & SpcFillAmount(tLevel1.SA_2, 22) & _
                            "  " & SpcFillAmount(tLevel1.SA_2 - tLevel1.SA, 22) & _
                            SpcFillAmount(nPercent, 22)
          oTextFile.PrintNewLine
            
          If (tLevel2.SA <> 0) Then
            nPercent = ((tLevel2.SA_2 / tLevel2.SA) - 1) * 100
          Else
            nPercent = 0
          End If

          oTextFile.PrintLine Space(28) & "  " & SpcFill(sLevel1, 40) & _
                            SpcFillAmount(tLevel2.SA, 22) & "  " & SpcFillAmount(tLevel2.SA_2, 22) & _
                            "  " & SpcFillAmount(tLevel2.SA_2 - tLevel2.SA, 22) & _
                            SpcFillAmount(nPercent, 22)
        Else
          oTextFile.PrintLine "NO SE ENCONTRARON LOS PARAMETROS PARA ESTE REPORTE"
        End If
        If bPrintInCascade Then
          oTextFile.PageFeed nLine
          ClearMoneyComparative tLevel1
          ClearMoneyComparative tLevel2
        Else
          Exit For
        End If

      End With
      oRsParam.Close
      Set oRsParam = Nothing
      Set oBalance = Nothing
    Next i
    Report153 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report153", Err.Number
End Function

Public Function Report156(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date) As String
  Dim oRs           As Recordset
  Dim oRSGenerator  As New SCFFixedReportsDS.CServer
  Dim oTextFile     As New CTextFile
  Dim nLine         As Long
  Dim nPage         As Long
  Dim sDTime        As String
  
  Dim sGl           As String
  
  Dim nGl           As Long
  Dim nAuxCount     As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    oTextFile.GenerateRandomFile
    oTextFile.Open_
    Set oRs = oRSGenerator.GeneraRecordset156(sAppServer, vGL, dIniDate, dEndDate)
    sDTime = GetFmtDate
    oTextFile.PrintNewLine
    
    If (IsArray(vGL)) Then
      sGl = "TODAS LAS CONTABILIDADES"
      nGl = vGL(0)
    Else
      nGl = vGL
      sGl = oRSGenerator.GetObjName(sAppServer, nGl, "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR", "NOMBRE_MAYOR")
    End If
        
    nLine = 100
    nPage = 1
    With oRs
      If Not (.BOF And .EOF) Then
        Do While Not (.EOF)
          PrintHeader156 oTextFile, nLine, nPage, sGl, dIniDate, dEndDate, sDTime
          oTextFile.PrintLine SpcFill(!NUMERO_CUENTA_AUXILIAR, 22) & "   " & _
                              SpcFill(IIf(IsNull(!nombre_cuenta_auxiliar), "", !nombre_cuenta_auxiliar), 50) & "     " & _
                              SpcFill(IIf(IsNull(!RFC), "", !RFC), 14) & "  " & _
                              SpcFill(IIf(IsNull(!DOMICILIO), "", !DOMICILIO), 50)
          
          nAuxCount = nAuxCount + 1
          
          .MoveNext
        Loop
        .Close
        PrintQtyLines oTextFile, 150
        oTextFile.PrintLine Space(8) & " TOTAL DE AUXILIARES: " & nAuxCount
      Else
        oTextFile.PrintLine "NO SE ENCONTRO INFORMACION"
      End If
    End With
    Report156 = oTextFile.FileName
    oTextFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Report156", Err.Number
End Function

'********************************************************************************
'      Encabezados de los Reportes
'********************************************************************************

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

Private Sub PrintHeader9(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          dIniDate As Date, dEndDate As Date, _
                          sGl As String, sCurrency As String, _
                          sDateTime As String, bPrintAcc As Boolean, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                                       FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                                                     PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                                                    BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C"
      oTextFile.PrintLine "                                                                             ACTUALIZACIÓN DE MOVIMIENTOS "
      oTextFile.PrintLine "                                                                            DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "MONEDA               : " & sCurrency
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA CONTABLE                GRUPO CONTROL           AUXILIAR                DESCRIPCION DE LA POLIZA                                             CARGO                    ABONO  "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 17
      bPrintAcc = True
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader9", Err.Number
End Sub

Private Sub PrintHeader18(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          dCuteDate As Date, sGLname As String, skindacc As String, sDateTime As String)
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                           FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                         PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                                 BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C"
      oTextFile.PrintLine "                                                BALANZA CONSOLIDADA POR CUENTA Y SU VALORIZACION"
      oTextFile.PrintLine "                                                                 AL: " & dCuteDate
      oTextFile.PrintNewLine
      oTextFile.PrintLine "TIPO DE CUENTA: " & skindacc
      oTextFile.PrintLine "CTA    NOMBRE                                              METROPOLITANA                  DELEGACIONES          TOTAL CONSOLIDADO"
      oTextFile.PrintLine "=================================================================================================================================="
      nPage = nPage + 1
      nLine = 15
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader18", Err.Number
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

Private Sub PrintHeader51_52(oTextFile As CTextFile, nLine As Long, nPage As Long, sGl As String, _
                             dEndDate As Date, sNat As String, sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                   BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                 C.N.B.-BANXICO SISTEMA DE INFORMACION FINANCIERA                                      PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                     RESUMEN ESTADO ANALITICO DE CUENTAS"
      oTextFile.PrintLine "BNO-670315-CDO" & "               " & (MonthName(Month(dEndDate))) & " de " & Year(dEndDate)
      oTextFile.PrintLine "COMISION NACIONAL BANCARIA"
      oTextFile.PrintLine "CLAVE BCO : 221"
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "NATURALEZA           : " & sNat
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CUENTA    D E S C R I P C I O N                  SALDO MON. NAL.          SALDO MON. EXT.            T O T A L E S "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 18
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader51_52", Err.Number
End Sub

Private Sub PrintHeader53_54(oTextFile As CTextFile, nLine As Long, nPage As Long, sGl As String, _
                             dEndDate As Date, sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                    BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                 C.N.B.-BANXICO SISTEMA DE INFORMACION FINANCIERA                                       PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                            COMPARACION DE GRUPOS"
      oTextFile.PrintLine "BNO-670315-CDO" & "               " & (MonthName(Month(dEndDate))) & " de " & Year(dEndDate)
      oTextFile.PrintLine "SALDOS A FIN DE MES (PESOS)"
      oTextFile.PrintLine "CLAVE BCO : 221"
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CUENTA           SALDO MON. NAL.          SALDO MON. EXT.            T O T A L E S   D E S C R I P C I O N"
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader53_54", Err.Number
End Sub

Private Sub PrintHeader55_56(oTextFile As CTextFile, nLine As Long, nPage As Long, sGl As String, _
                                sCurrency As String, dIniDate As Date, dEndDate As Date, _
                                sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                   BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                      FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                 C.N.B.-BANXICO SISTEMA DE INFORMACION FINANCIERA                                       PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                             ANALITICO DE CUENTAS"
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine "SALDOS A FIN DE MES EN (" & sCurrency & ")"
      oTextFile.PrintLine "PERIODO: " & Year(dEndDate) & ZeroFill(Month(dEndDate), 2)
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CUENTA                              SALDO MON. NAL.         SALDO MON. EXT.      D E S C R I P C I O N "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 17
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader55_56", Err.Number
End Sub

Private Sub PrintHeader58(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          sGl As String, sCurrency As String, _
                          dIniDate As Date, dEndDate As Date, _
                          sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintLine sDateTime
     oTextFile.PrintLine "                                                            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                  FECHA  : " & Date
     oTextFile.PrintLine "TECOYOTITLA NUM. 100                                         C.N.B.-BANXICO SISTEMA  DE INFORMACION FINANCIERA                                   PAGINA : " & nPage
     oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                      ANALITICO DE CUENTAS  "
     oTextFile.PrintLine "BNO-670315-CDO"
     oTextFile.PrintLine "SALDOS PROMEDIOS EN PESOS"
     oTextFile.PrintLine "PERIODO: " & Year(dEndDate) & ZeroFill(Month(dEndDate), 2)
     oTextFile.PrintLine "CLAVE BANCO : 221"
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CONTABILIDAD         : " & sGl
     oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CUENTA                              SALDO MON. NAL.        SALDO EN DOLARES   D E S C R I P C I O N "
     PrintQtyLines oTextFile
     nPage = nPage + 1
     nLine = 19
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader58", Err.Number
End Sub

Private Sub PrintHeader59(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                    dCuteDate As String, _
                                    sGLDesc As String, _
                                    skindacc As String, _
                                    sAccount As String, _
                                    sDateTime As String)
  
  On Error GoTo ErrorHandler
    If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                              PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                           BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C "
      oTextFile.PrintLine "                                                       CARTERA DE CREDITOS"
      oTextFile.PrintLine "                                                       CIFRAS DE SIFRRBCO"
      oTextFile.PrintLine "                                                    AL " & dCuteDate
      oTextFile.PrintLine ""
      oTextFile.PrintLine "                                                                            DELEGACION: " & sGLDesc
      oTextFile.PrintLine "Cuenta Contable: " & sAccount
      oTextFile.PrintLine "=============================================================================================================================================="
      oTextFile.PrintLine "Descripcion Auxiliar                   Auxiliar      MND                                      Saldo    T. Valorizacion        Saldo Valorizado"
      oTextFile.PrintLine "=============================================================================================================================================="
      nPage = nPage + 1
      nLine = 20
    End If
    Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader59", Err.Number
End Sub

Private Sub PrintHeader60(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                     dCuteDate As Date, sDateTime As String, sAccounts As String)
  On Error GoTo ErrorHandler
    If (nPage = 0) Then
           nPage = 1
    End If
    sAccounts = Replace(sAccounts, "','", " ")
    If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C          FECHA  : " & Format(Date, "dd/mmm/yyyy")
      oTextFile.PrintLine "TECOYOTITLA NUM. 100              RESUMEN MENSUAL POR AUXILIARES DE LAS CUENTAS              PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                    " & sAccounts
      oTextFile.PrintLine "                                                 AL : " & Format(dCuteDate, "dd/mmm/yyyy")
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      Printlines60 oTextFile
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader60", Err.Number
End Sub

Private Sub Printlines60(oTextFile As CTextFile)
  On Error GoTo ErrHandler
    oTextFile.PrintLine "  AUXILIAR                       SALDO ACTUAL     DESCRIPCION"
    oTextFile.PrintLine "==================================================================================================================="
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Printlines60", Err.Number
End Sub

Private Sub PrintHeader61(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                    dCuteDate As String, _
                                    sGLDesc As String, _
                                    skindacc As String, _
                                    sDateTime As String, _
                                    sAccType As String)
  
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "BMG5072"
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                              PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                           BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C "
      oTextFile.PrintLine "                                                 REPORTE DE CUENTAS DEUDORAS Y ACCREDORAS"
      oTextFile.PrintLine "                                                        AL: " & dCuteDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGLDesc
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "TIPO: " & skindacc
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader61", Err.Number
End Sub

Private Sub PrintHeader62(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          dIniDate As Date, dEndDate As Date, _
                          sGl As String, dAfecDate As Date, _
                          sDateTime As String, bPrintAcc As Boolean, sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                                       FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                                                     PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                                                    BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C"
      oTextFile.PrintLine "                                                                         MOVIMIENTO DE CUENTAS SELECCIONADAS"
      oTextFile.PrintLine "                                                                           DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "FECHA                : " & dAfecDate
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA SUBCUENTA                 VOLANTE              MON               DESCRIPCION DEL VOLANTE                      AUXILIAR                         CARGO                        ABONO  "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 17
      bPrintAcc = True
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader62", Err.Number
End Sub

Private Sub PrintHeader64(sAppServer As String, oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, dIniDate As Date, dEndDate As Date, _
                           sDateTime As String, sNature As String, _
                           sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                    INSTITUCION BANCA DE DESARROLLO                                                PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                            MAYOR GENERAL"
      oTextFile.PrintLine "BNO-670315-CDO                                                         " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "MONEDA               : " & sCurrency
      oTextFile.PrintLine "NATURALEZA           : " & IIf(sNature = "D", " 1 CTAS DEUDORAS", " 2 CTAS ACREEDORAS")
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  FECHA     NOMBRE                                                                                        CARGOS                   ABONOS                 SALDO"
      PrintQtyLines oTextFile, 172
      nPage = nPage + 1
      nLine = 11
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader64", Err.Number
End Sub

Private Sub PrintHeader65(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          sGl As String, sCurrency As String, _
                          sGLN As String, dIniDate As Date, dEndDate As Date, _
                          sDateTime As String, sAccType As String)
  Dim sDatePrint As String
  '******************************************************************************************
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintLine sDateTime
     oTextFile.PrintLine "                                                            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                  FECHA  : " & Format(Date, "dd/mmm/yyyy")
     oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                INSTITUCION BANCA DE DESARROLLO                                              PAGINA : " & nPage
     oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
     oTextFile.PrintLine "BNO-670315-CDO                                             SUMA Y PROMEDIOS DE SALDOS DIARIOS EN CUENTAS DE MAYOR"
     oTextFile.PrintLine "                                                                     AL " & Format$(dEndDate, "d "" de "" MMMM "" del "" YYYY")
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CONTABILIDAD         : " & sGl
     oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CUENTA DESCRIPCION                                        MONEDA NACIONAL             MONEDA EXTRANJERA           TOT. SUMA DE SALDOS       SALDOS DIARIO PROMEDIO  "
     oTextFile.PrintLine "                                                                                                                                                     MN             "
     PrintQtyLines oTextFile, 167
     nPage = nPage + 1
     nLine = 15
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader65", Err.Number
End Sub

Private Sub PrintHeader68(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          dCuteDate As Date, sGLname As String, _
                          skindcurr As String, sDateTime As String, _
                          nStdAccType As Long)
  Dim sAccHeader  As String
  '******************************************************************************
  On Error GoTo ErrorHandler
   If (nStdAccType = 1) Then
     sAccHeader = "DELEG."
   Else
     sAccHeader = "FID.  "
   End If
   If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                              PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                      BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C BANCA DE DESARROLLO"
      oTextFile.PrintLine "                                             REPORTE MENSUAL CONSOLIDADO DE CUENTA DE CHEQUES"
      oTextFile.PrintLine "                                                        DE FIDEICOMISOS Y MANDATOS"
      oTextFile.PrintLine "                                                        POR EL MES " & Format(dCuteDate, "MMMM ""DE"" YYYY")
      oTextFile.PrintNewLine
      oTextFile.PrintLine "MONEDA: " & skindcurr
      oTextFile.PrintNewLine
      oTextFile.PrintLine "NUM. " & sAccHeader & "  CTA    SUBCUENTA       NOMBRE                                                   MONTO"
      oTextFile.PrintLine "======================================================================================================="
      nPage = nPage + 1
      nLine = 15
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader68", Err.Number
End Sub

Private Sub PrintHeader69(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           dIniDate As Date, dEndDate As Date, vAccountGroupName As Variant, _
                           sTittleReport As String, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", " ", "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", sTittleReport & " AL MES DE " & Format(dEndDate, "MMMM"" DEL ""yyyy"), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, " ", " ", sDateTime & " -- Página " & nPage)
      oTextFile.PrintNewLine
      oTextFile.PrintLine "NUMERO       NOMBRE                                                            PROVISIONADOS                  PAGADO                  PAGADO                        PENDENTES "
      oTextFile.PrintLine "FID./MDTO.                                                                       EN EL MES                   EN EL MES            EN EL EJERCICIO                    DE PAGO  "
      PrintQtyLines oTextFile, 178
      nPage = nPage + 1
      nLine = 10
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader69", Err.Number
End Sub


Private Sub PrintHeader71(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           dIniDate As Date, dEndDate As Date, vAccountGroupName As Variant, _
                           sTittleReport As String, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", PrintDateGeneral(dIniDate, dEndDate), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, " ", " ", sDateTime & " -- Página " & nPage)
      oTextFile.PrintNewLine
      oTextFile.PrintLine "FID     NOMBRE                                                " & SpcFill(vAccountGroupName(0), 10) & "            " & SpcFill(vAccountGroupName(1), 10) & "                  " & SpcFill(vAccountGroupName(2), 10) & "     TOTAL COLS. 1-2-3"
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 10
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader71", Err.Number
End Sub

Private Sub PrintHeader73(oTextFile As CTextFile, nLine As Long, nPage As Long, _
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
      oTextFile.PrintNewLine
      oTextFile.PrintLine "FID           CTA   SUBCUENTA             NOMBRE                                    SALDO ANTERIOR                   CARGOS                   ABONOS             SALDO ACTUAL "
      PrintQtyLines oTextFile, 180
      nPage = nPage + 1
      nLine = 14
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader73", Err.Number
End Sub

Private Sub PrintHeader74_76(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                             dIniDate As Date, dEndDate As Date, dIniDate2 As Date, _
                             dEndDate2 As Date, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                                                                FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                        COMPARATIVO MENSUAL DE SALDOS DE FIDEICOMISOS Y MANDATOS                                                                              PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                 FECHA ACTUAL( " & Format$(dIniDate2, "MMMM"" de ""yyyy") & " ) - FECHA ANTERIOR( " & Format$(dIniDate, "MMMM"" de ""yyyy") & " )"
      oTextFile.PrintLine "BNO-670315-CDO                                       "
      oTextFile.PrintNewLine
      oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
      oTextFile.PrintLine "                                                                       A C T I V O                                                                6320-95                                                              T O T A L E S"
      oTextFile.PrintLine " FID N O M B R E                                    ACTUAL                ANTERIOR               VARIACION                  ACTUAL                ANTERIOR               VARIACION                  ACTUAL                ANTERIOR               VARIACION"
      oTextFile.PrintLine "==========================================================================================================================================================================================================================================================="
      nPage = nPage + 1
      nLine = 14
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader74_76", Err.Number
End Sub

Private Sub PrintHeader96(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                    dCuteDate As String, _
                                    sGLDesc As String, _
                                    skindacc As String, _
                                    sAccount As String, _
                                    sDateTime As String)
  
  On Error GoTo ErrorHandler
    If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                                FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                              PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                           BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C "
      oTextFile.PrintLine "                                     AUXILIARES DE LA CUENTA 1103 PARA LA RECLASIFICACION DE SALDOS"
      oTextFile.PrintLine "                                                   Fecha: " & dCuteDate
      oTextFile.PrintNewLine
      oTextFile.PrintLine "================================================================================================================================================"
      oTextFile.PrintLine "Cta  Subcuenta       Descripcion            Auxiliar            Descripcion                    MNA           Saldo Deudor        Saldo Acreedor "
      oTextFile.PrintLine "================================================================================================================================================"
      nPage = nPage + 1
      nLine = 20
    End If
    Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader96", Err.Number
End Sub

Private Sub PrintHeader101(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          sGl As String, sCurrency As String, nExcRate As Currency, _
                          dIniDate As Date, dEndDate As Date, _
                          sDateTime As String, sAccType As String, sReportTittle As String)
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintLine sDateTime
                                                                                                                                                                                          
     oTextFile.PrintLine "                                                            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                  FECHA  : " & Date
     oTextFile.PrintLine FormatLine(164, "TECOYOTITLA NUM. 100", sReportTittle, "PAGINA : " & nPage)
     oTextFile.PrintLine "C.P. 01030 MEXICO D.F.     "
     oTextFile.PrintLine "BNO-670315-CDO"
     oTextFile.PrintLine "                                                                    AL " & Format$(dEndDate, "d "" de "" MMMM "" del "" YYYY")
     oTextFile.PrintLine SpcFill("CODIGO MONEDA : " & sCurrency, 40) & " TIPO CAMBIO " & SpcFillAmount(nExcRate, 12, 6)
     oTextFile.PrintLine "CUENTA                           DESCRIPCION                              F O M E N T O             DELEGACIONES               TOTAL"
     PrintQtyLines oTextFile, 164
     nPage = nPage + 1
     nLine = 11
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader101", Err.Number
End Sub

Private Sub PrintHeader102(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                          sGl As String, sCurrency As String, _
                          sGLN As String, dIniDate As Date, dEndDate As Date, _
                          sDateTime As String, sAccType As String)
  Dim sDatePrint As String
  '*******************************************************************************
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintLine sDateTime
     oTextFile.PrintLine "                                                            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                  FECHA  : " & Format(Date, "dd/mmm/yyyy")
     oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                INSTITUCION BANCA DE DESARROLLO                                              PAGINA : " & nPage
     oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
     oTextFile.PrintLine "BNO-670315-CDO                             PROMEDIO DE SALDOS PARA CALCULAR EL COMPONENTE INFLACIONARIO PARA EFECTOS DE I.S.R."
     oTextFile.PrintLine "                                                                     AL " & Format$(dEndDate, "d "" de "" MMMM "" del "" YYYY")
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CONTABILIDAD         : TODAS LAS CONTABILIDADES "
     oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
     oTextFile.PrintNewLine
     oTextFile.PrintLine "CUENTA DESCRIPCION                                        MONEDA NACIONAL             MONEDA EXTRANJERA           TOT. SUMA DE SALDOS       SALDOS DIARIO PROMEDIO  "
     PrintQtyLines oTextFile
     nPage = nPage + 1
     nLine = 15
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader102", Err.Number
End Sub

Private Sub PrintHeader106(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                 sGl As String, dIniDate As Date, dEndDate As Date, _
                                 sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                   DIRECTORIO ALFABETICO POR AUXILIARES DE CARTERA                             PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                                   PROGRAMA           SUBPROGRAMA  "
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader106", Err.Number
End Sub

Private Sub PrintHeader107(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                 sGl As String, dIniDate As Date, dEndDate As Date, _
                                 sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                 DIRECTORIO ALFABETICO POR AUXILIARES ADMINISTRATIVOS                          PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                            RFC              CURP                      CLAVE IVA   TASA      TIPO RETENCION"
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader107", Err.Number
End Sub

Private Sub PrintHeader108(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                 sGl As String, dIniDate As Date, dEndDate As Date, _
                                 sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                   DIRECTORIO ALFABETICO POR AUXILIARES DE ACTIVOS                             PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                            ADQUISICION        VALOR HISTORICO"
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader108", Err.Number
End Sub

Private Sub PrintHeader109(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                 sGl As String, dIniDate As Date, dEndDate As Date, _
                                 sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                        DIRECTORIO ALFABETICO POR AUXILIARES                                   PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                                                                   "
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader109", Err.Number
End Sub

Private Sub PrintHeader110(sAppServer As String, oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, dIniDate As Date, dEndDate As Date, _
                           sDateTime As String, sCuenta As String, sCLAVE_SECTOR As String, sCLAVE_AUX As String, _
                           sNOMBRE_CUENTA As String, sN_SECTOR As String, sNOMBRE_AUX As String, _
                           sAccType As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                    INSTITUCION BANCA DE DESARROLLO                                                PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                AUXILIAR DE MOVIMIENTOS POR MONEDA"
      oTextFile.PrintLine "BNO-670315-CDO                                                         " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      nLine = 11
      PrintSubHead sAppServer, oTextFile, nLine, sCurrency, sCuenta, sCLAVE_SECTOR, sCLAVE_AUX, sNOMBRE_CUENTA, sN_SECTOR, sNOMBRE_AUX
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  FECHA     NO. POLIZA                 DESCRIPCION                                                        CARGOS                   ABONOS                 SALDO"
      PrintQtyLines oTextFile, 172
      nPage = nPage + 1
      nLine = nLine + 4
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader110", Err.Number
End Sub

Private Sub PrintHeader113(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                              sGl As String, sCurrency As String, dIniDate As Date, dEndDate As Date, _
                              sTittleReport As String, sDateTime As String, sAccType As String, _
                              bPrintAccount As Boolean)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                          BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                              FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                INSTITUCION BANCA DE DESARROLLO                                                        PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                             " & sTittleReport
      oTextFile.PrintLine "BNO-670315-CDO                                                     " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintLine "MONEDA               : " & sCurrency
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CTA   SUBCUENTA             DESCRIPCION                                                 SALDO ANTERIOR                   CARGOS                   ABONOS             SALDO ACTUAL "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
      bPrintAccount = True
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader113", Err.Number
End Sub

Private Sub PrintHeader115(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sDateTime As String, sParticipantTypeDesc As String)
  
  On Error GoTo ErrorHandler
     
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                  FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C"
      oTextFile.PrintLine "                                            REPORTE DE " & sParticipantTypeDesc
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CLAVE                                NOMBRE                                                   ESTADO     "
      oTextFile.PrintLine "=========================================================================================================="
       
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader115", Err.Number
End Sub

Private Sub PrintHeader116(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           nStdAccountTypeId As Long, sDateTime As String)
  Dim sStdAccTypeDesc As String
  '*******************************************************************************
  On Error GoTo ErrorHandler
    If (nStdAccountTypeId = 1) Then
      sStdAccTypeDesc = " BANCARIA "
    Else
      sStdAccTypeDesc = " FIDUCIARIA "
    End If
              
    If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                              FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                            PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                                   BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C"
      oTextFile.PrintLine "                                                                CATALOGO DE CUENTAS  "
      oTextFile.PrintLine "CONTABILIDAD: " & sStdAccTypeDesc
      oTextFile.PrintLine "NUMERO DE CUENTA        NOMBRE                                            SECTOR    NATURALEZA         ROL DE CUENTA   TIPO DE CUENTA "
      oTextFile.PrintLine "====================================================================================================================================="

      nPage = nPage + 1
      nLine = 11
    End If
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader116", Err.Number
End Sub

Private Sub PrintHeader117(oTextFile As CTextFile, nLine As Long, nPage As Long, sDateTime As String, _
                           sSQLOrder As String)

  On Error GoTo ErrorHandler
    If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine FormatLine(128, "TECOYOTITLA NUM. 100", " ", "FECHA  : " & Date)
      oTextFile.PrintLine FormatLine(128, "C.P. 01030 MEXICO D.F.", " ", "PAGINA : " & nPage)
      oTextFile.PrintLine FormatLine(128, "BNO-670315-CDO", "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", " ")
      oTextFile.PrintLine FormatLine(128, " ", "CATALOGO DE CONTABILIDADES", " ")
      oTextFile.PrintNewLine
      If sSQLOrder = "" Then
        oTextFile.PrintLine "NUMERO      NOMBRE "
      Else
        oTextFile.PrintLine "NUMERO      NOMBRE                                                     TIPO  "
      End If
      PrintQtyLines oTextFile, 128
      nPage = nPage + 1
      nLine = 12
    End If
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader116", Err.Number
End Sub

Private Sub PrintHeader127(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, dEndDate As Date, _
                           sTittleReport As String, sDateTime As String, sAccType As String, _
                           sOptionToDisplay As String)
                           
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine FormatLine(110, " ", "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", "FECHA  : " & Date)
      oTextFile.PrintLine FormatLine(110, "TECOYOTITLA NUM. 100", "INSTITUCION BANCA DE DESARROLLO", "PAGINA : " & nPage)
      oTextFile.PrintLine FormatLine(110, "C.P. 01030 MEXICO D.F.", sTittleReport & " DEL MES DE " & Format(dEndDate, "MMMM"" DE ""yyyy"), " ")
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : " & sAccType
      oTextFile.PrintNewLine
      oTextFile.PrintLine "CUENTA                      NOMBRE                                                              INTERESES    "
      If (sOptionToDisplay = "M") Then
        oTextFile.PrintLine "                                                                                                 EXENTOS    "
      ElseIf (sOptionToDisplay = "S") Then
        oTextFile.PrintLine "                                                                                                 EXENTOS    "
        oTextFile.PrintLine "                                                                                                ACUMULADO   "
      End If
      PrintQtyLines oTextFile, 112
      nPage = nPage + 1
      nLine = 14
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader127", Err.Number
End Sub

Private Sub PrintHeader128(oTextFile As CTextFile, nLine As Long, nPage As Long, sAccounts As String, sDate As String, sDateTime As String)
 On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintNewLine
     oTextFile.PrintLine "                                                     SUBDIRECCION DE CONTABILIDAD GENERAL                        "
     oTextFile.PrintLine "                                                   GERENCIA DE REGISTRO E INFORMACION CONTABLE                   "
     oTextFile.PrintNewLine
     oTextFile.PrintLine "_____________________________________________________________________________________________________________________________"
     oTextFile.PrintNewLine
     oTextFile.PrintLine "            Posición propia en monedas extranjeras con saldos al " & sDate
     oTextFile.PrintLine "                                                                          Equivalencia                   Valuación en     "
     oTextFile.PrintLine "         MONEDA                                       " & sAccounts & "              dólar americano(*)             dólares americanos   "
     oTextFile.PrintLine "        _______                                       ____             __________________               _____________________ "
     nPage = nPage + 1
     nLine = 1
   End If
 Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader128", Err.Number
End Sub

Private Sub PrintHeader144(oTextFile As CTextFile, nLine As Long, nPage As Long, sGl As String, _
                           dIniDate As Date, dEndDate As Date, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                 DIRECTORIO ALFABETICO POR AUXILIARES DE FIDEICOMISOS                          PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                        TIPO TASA        TASA          MONTO CREDITO     FECHA INICIO    FECHA FIN    "
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader144", Err.Number
End Sub

Private Sub PrintHeader149(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           dIniDate As Date, dEndDate As Date, dIniDate2 As Date, dEndDate2 As Date, _
                           sTittleReport As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional bWithDif As Boolean = False, _
                           Optional bWithPercent As Boolean = False, Optional sAccount As String)
  On Error GoTo ErrorHandler
   'nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", " ", "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      oTextFile.PrintLine FormatLine(174, " ", "Cuenta: " & sAccount, sVoucherType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, "MONEDA             : " & sCurrency, "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, "VALORIZADA A       : " & sCurrency, "", sDateTime)
      End If
      
      oTextFile.PrintNewLine
      oTextFile.PrintLine "FID     AUXILIAR                      D E S C R I P C I O N                       " & Format(dIniDate, "MMMM YYYY") & "            " & Format(dIniDate2, "MMMM YYYY") & "                VARIACION                     (%)"
      oTextFile.PrintLine "No.                                                                                                                                                                                 "
      PrintQtyLines oTextFile, 175
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader149", Err.Number
End Sub

Private Sub PrintHeader150(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           dIniDate As Date, dEndDate As Date, dIniDate2 As Date, dEndDate2 As Date, _
                           sTittleReport As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional bWithDif As Boolean = False, _
                           Optional bWithPercent As Boolean = False, Optional sAccount As String)
  On Error GoTo ErrorHandler
   'nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", " ", "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, " ", "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, " ", "", sDateTime)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintLine "FID     AUXILIAR                MO    CUENTA  SUBCUENTAS                          " & Format(dIniDate, "MMMM YYYY") & "            " & Format(dIniDate2, "MMMM YYYY") & "                VARIACION                     (%)"
      oTextFile.PrintLine "No.                                                                                                                                                                                 "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader150", Err.Number
End Sub

Private Sub PrintHeader151(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           dIniDate As Date, dEndDate As Date, sAccount As String, sAccount2 As String, _
                           sTittleReport As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional bWithDif As Boolean = False, _
                           Optional bWithPercent As Boolean = False, Optional nExcRate As Double)
  On Error GoTo ErrorHandler
   'nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", " ", "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      oTextFile.PrintLine FormatLine(174, " ", " Tipo de Cambio UDIS: " & nExcRate, sVoucherType)
      oTextFile.PrintLine FormatLine(174, " ", " Movimientos del mes de " & Format(dIniDate, "MMMM YYYY"), " ")
      oTextFile.PrintNewLine
      oTextFile.PrintLine "FID     AUXILIAR                      D E S C R I P C I O N                       " & "Cta. " & sAccount & "            " & "Cta. " & sAccount2 & "                VARIACION               (%)     T. de C. Prom."
      oTextFile.PrintLine "No.                                                                                    M.N                Valorizado en M.N.                                                     "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader151", Err.Number
End Sub

Private Sub PrintHeader152(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           dIniDate As Date, dEndDate As Date, _
                           sTittleReport As String, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", " ", "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", sTittleReport, "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, " ", " AL MES DE " & Format(dEndDate, "MMMM"" DEL ""yyyy"), sDateTime & " -- Página " & nPage)
      oTextFile.PrintNewLine
      oTextFile.PrintLine "NUMERO         MON                SALDOS DE               NUMERO           SALDOS DE          T. DE C.        SALDO MON. NAL. PROVISIONADOS               DIF"
      oTextFile.PrintLine " FID.                           5001 Y/O 5011              FID           5000 Y/O 5010                        CONVERTIDA EN MON. EXTRANJERA                  "
      PrintQtyLines oTextFile, 178
      nPage = nPage + 1
      nLine = 10
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader152", Err.Number
End Sub

Private Sub PrintHeader153(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           dIniDate As Date, dEndDate As Date, dIniDate2 As Date, dEndDate2 As Date, _
                           sTittleReport As String, sDateTime As String, _
                           Optional sAccountInclude As String, Optional sAccType As String, _
                           Optional sTransactionType As String, Optional sVoucherType As String, _
                           Optional nExcRateCurrency As Long, Optional bWithDif As Boolean = False, _
                           Optional bWithPercent As Boolean = False, Optional sAccount As String)
  On Error GoTo ErrorHandler
   'nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", " ", "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "CONTABILIDAD       : " & sGl, "", sTransactionType)
      If nExcRateCurrency = 0 Then
        oTextFile.PrintLine FormatLine(174, " ", "", sDateTime & " -- Página " & nPage)
      Else
        oTextFile.PrintLine FormatLine(174, " ", "", sDateTime)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintLine "MO     CUENTA                NOMBRE                                           " & Format(dIniDate, "MMMM YYYY") & "            " & Format(dIniDate2, "MMMM YYYY") & "                VARIACION                     (%)"
      oTextFile.PrintLine "No.                                                                                                                                                                                 "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader153", Err.Number
End Sub

Private Sub PrintHeader156(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                 sGl As String, dIniDate As Date, dEndDate As Date, _
                                 sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                     BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                 DIRECTORIO ALFABETICO POR AUXILIARES DE PROVEEDORES                           PAGINA : " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F."
      oTextFile.PrintLine "BNO-670315-CDO                                               " & " DEL: " & dIniDate & " AL: " & dEndDate
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "  AUXILIAR               DESCRIPCION                                               RFC              DOMICILIO    "
      PrintQtyLines oTextFile, 150
      nPage = nPage + 1
      nLine = 12
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader156", Err.Number
End Sub

Private Sub PrintSubHead(sAppServer As String, oTextFile As CTextFile, nLine As Long, sMoneda As String, sCuenta As String, _
                         sCLAVE_SECTOR As String, sCLAVE_AUX As String, _
                         sNombreCuenta As String, sNombreSector As String, sNombreAux As String)
 Dim oRSGenerator As New SCFFixedReportsDS.CServer
 '*****************************************************************************************************
  On Error GoTo ErrHandler
    oTextFile.PrintNewLine
    oTextFile.PrintLine " MONEDA         : " & sMoneda
    oTextFile.PrintLine " CUENTA         : " & Mid(sCuenta, 1, 4) & "  " & _
                        SpcFill(oRSGenerator.GetObjNameStr(sAppServer, Mid(sCuenta, 1, 4), "COF_CUENTA_ESTANDAR_HIST", "NOMBRE_CUENTA_ESTANDAR", "NUMERO_CUENTA_ESTANDAR"), 50)
    oTextFile.PrintLine " SUBCUENTA      : " & SpcFill(sCuenta, 22) & "  " & SpcFill(sNombreCuenta, 50) & "  " & _
                        "SECTOR: " & sCLAVE_SECTOR & "  " & sNombreSector
    If (sCLAVE_AUX = "0") Then
      sCLAVE_AUX = "NO MANEJA AUXILIARES"
    End If
    oTextFile.PrintLine " AUXILIAR       : " & sCLAVE_AUX & "  " & sNombreAux
    nLine = nLine + 4
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintSubHead", Err.Number
End Sub

Private Sub PrintDetalle64(oTextFile As CTextFile, sNature As String, tAcumulador As TMoney, _
                            dFecha As Date, sConcepto As String, _
                            nBalance As Currency, nCargos As Currency, nAbonos As Currency)
  On Error GoTo ErrHandler
    If (sNature = "D") Then
      nBalance = nBalance + nCargos - nAbonos
    ElseIf (sNature = "A") Then
      nBalance = nBalance - nCargos + nAbonos
    Else
    '
    End If
    AddMoney tAcumulador, 0, nCargos, nAbonos, 0
    If (dFecha <> CDate("01/01/1990")) Then
      oTextFile.PrintLine dFecha & "  " & SpcFill(sConcepto, 77) & "  " & _
                          SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & _
                         SpcFillAmount(nBalance, 22)
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalle64", Err.Number
End Sub

Private Sub PrintDetalle110(oTextFile As CTextFile, sNature As String, tAcumulador As TMoney, _
                            dFecha As Date, sPoliza As String, sConcepto As String, _
                            nBalance As Currency, nCargos As Currency, nAbonos As Currency)
  On Error GoTo ErrHandler
    If (sNature = "D") Then
      nBalance = nBalance + nCargos - nAbonos
    ElseIf (sNature = "A") Then
      nBalance = nBalance - nCargos + nAbonos
    Else
    '
    End If
    AddMoney tAcumulador, 0, nCargos, nAbonos, 0
    If (dFecha <> CDate("01/01/1990")) Then
      oTextFile.PrintLine dFecha & "  " & SpcFill(sPoliza, 25) & "  " & SpcFill(sConcepto, 50) & "  " & _
                          SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & _
                         SpcFillAmount(nBalance, 22)
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalle110", Err.Number
End Sub

Private Sub PrintDetalleAux(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, _
                            sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                            sCLAVE_AUX As String, sNOMBRE_AUX As String, nSI As Currency, _
                            nCargos As Currency, nAbonos As Currency, nSA As Currency, _
                            dUltMov As Date, bPrintAccount As Boolean)
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
    
    oTextFile.PrintLine sLabel & "  " & dUltMov & "  " & sCLAVE_AUX & "  " & SpcFill(sNOMBRE_AUX, 25) & "  " & _
              SpcFillAmount(nSI, 22) & "  " & SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & SpcFillAmount(nSA, 22)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalleAux", Err.Number
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

Private Sub PrintDetalle73(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, _
                         sNombre_Mayor As String, sCuenta As String, sCLAVE_SECTOR As String, sNOMBRE_CUENTA As String, _
                         nSI As Currency, nCargos As Currency, nAbonos As Currency, nSA As Currency, _
                         sROL As String, nFROM_MOVS As Long, Optional nOptionToDisplay As Long = 0)
  Dim sChar  As String
  '******************************************************************************************************
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
    oTextFile.PrintLine SpcFill(sNombre_Mayor, 12) & "  " & SpcFill(sCuenta, 22) & "  " & _
                        sCLAVE_SECTOR & "  " & SpcFill(sNOMBRE_CUENTA, 34) & "  " & _
                        SpcFillAmount(nSI, 22) & "  " & SpcFillAmount(nCargos, 22) & "  " & SpcFillAmount(nAbonos, 22) & "  " & SpcFillAmount(nSA, 22) & sChar
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalle73", Err.Number
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

Private Sub PrintDetalle62(oTextFile As CTextFile, bPrintAcc As Boolean, tPoliza As TMoney, _
                           sCuenta As String, sCLAVE_SECTOR As String, sPoliza As String, _
                           sMoneda As String, sConcepto As String, sAux As String, sTIPO_MOV As String, _
                           nMonto As Currency)
  Dim nCargo            As Currency
  Dim nAbono            As Currency
  Dim sLabel            As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    If (sTIPO_MOV = "D") Then
      AddMoney tPoliza, 0, nMonto, 0, 0
      nCargo = nMonto
      nAbono = 0
    ElseIf (sTIPO_MOV = "H") Then
      AddMoney tPoliza, 0, 0, nMonto, 0
      nCargo = 0
      nAbono = nMonto
    End If
    
    If (bPrintAcc) Then
      sLabel = SpcFill(sCuenta, 22) & "  " & sCLAVE_SECTOR & "  "
      bPrintAcc = False
    ElseIf Not (bPrintAcc) Then
      sLabel = SpcFill(" ", 28)
    End If
    oTextFile.PrintLine sLabel & _
                        SpcFill(sPoliza, 22) & "  " & SpcFill(sMoneda, 2) & "  " & _
                        SpcFill(sConcepto, 50) & "  " & SpcFill(sAux, 22) & "  " & _
                        SpcFillAmount(nCargo, 22) & "  " & SpcFillAmount(nAbono, 22)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintDetalle62", Err.Number
End Sub

Private Sub PrintTotPoliza(oTextFile As CTextFile, nLine As Long, bPrintAcc As Boolean, _
                           tPoliza As TMoney, tFecha As TMoney, _
                           tMoneda As TMoney, sPoliza As String, sNewPoliza As String, _
                           sAcc As String, sNewAcc As String, sMsg As String)
  On Error GoTo ErrHandler
    nLine = nLine + 5
    bPrintAcc = True
    AddMoney tFecha, 0, tPoliza.Cargos, tPoliza.Abonos, 0
    AddMoney tMoneda, 0, tPoliza.Cargos, tPoliza.Abonos, 0
    oTextFile.PrintLine Space(100) & "=================================================================================="
    oTextFile.PrintLine SpcFill(sMsg, 132) & SpcFillAmount(tPoliza.Cargos, 22) & "  " & SpcFillAmount(tPoliza.Abonos, 22)
    PrintQtyLines oTextFile
    ClearMoney tPoliza
    sPoliza = sNewPoliza
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintToPoliza", Err.Number
End Sub

Private Sub PrintTotFecha(oTextFile As CTextFile, nLine As Long, bPrintAcc As Boolean, _
                           tFecha As TMoney, _
                           dFecAfe As Date, dNewFec As Date, sMsg As String)
  On Error GoTo ErrHandler
    nLine = nLine + 5
    bPrintAcc = True
    oTextFile.PrintLine Space(100) & "=================================================================================="
    oTextFile.PrintLine SpcFill(sMsg, 132) & SpcFillAmount(tFecha.Cargos, 22) & "  " & SpcFillAmount(tFecha.Abonos, 22)
    PrintQtyLines oTextFile
    ClearMoney tFecha
    If (dFecAfe <> dNewFec) Then
      dFecAfe = dNewFec
      oTextFile.PrintLine "FECHA                : " & dFecAfe
      oTextFile.PrintLine "CTA SUBCUENTA                 VOLANTE              MON               DESCRIPCION DEL MOVIMIENTO                   AUXILIAR                         CARGO                        ABONO  "
    Else
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintToFecha", Err.Number
End Sub

Private Sub PrintTotMoneda(oTextFile As CTextFile, nLine As Long, bPrintAcc As Boolean, _
                           tMoneda As TMoney, _
                           sMoneda As String, sNewMoneda As String, sMsg As String)
  On Error GoTo ErrHandler
    nLine = nLine + 5
    bPrintAcc = True
    oTextFile.PrintLine Space(100) & "=================================================================================="
    oTextFile.PrintLine SpcFill(sMsg, 132) & SpcFillAmount(tMoneda.Cargos, 22) & "  " & SpcFillAmount(tMoneda.Abonos, 22)
    PrintQtyLines oTextFile
    ClearMoney tMoneda
    sMoneda = sNewMoneda
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintToFecha", Err.Number
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

Private Sub PrintTot(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, tTot As TMoney, _
                     sOld As String, sNew As String, sMsg As String, Optional bPrintSA As Boolean = True)
  '*************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile, 172
    AddMoney tTot, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    If bPrintSA Then
      oTextFile.PrintLine SpcFill(sMsg, cnBlancosCorte + 13) & _
                         "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                         "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                         "  " & SpcFillAmount(tAcumulador.SA, 22)
    Else
      oTextFile.PrintLine SpcFill(sMsg, cnBlancosCorte + 13) & _
                   "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                   "  " & SpcFillAmount(tAcumulador.Abonos, 22)
    End If
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sOld = sNew
    ClearMoney tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot", Err.Number
End Sub

Private Sub PrintTot111(oTextFile As CTextFile, nLine As Long, tAcumulador As TMoney, tTot As TMoney, _
                        sOld As String, sNew As String, sMsg As String)
  '****************************************************************************************************
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    AddMoney tTot, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, tAcumulador.SA
    oTextFile.PrintLine SpcFill(sMsg, cnBlancosCorte + 21) & _
                        "      " & SpcFillAmount(tAcumulador.SI, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Cargos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.Abonos, 22) & _
                        "  " & SpcFillAmount(tAcumulador.SA, 22)
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    nLine = nLine + 4
    sOld = sNew
    ClearMoney tAcumulador
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot111", Err.Number
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

Private Sub PrintTot51_52(oTextFile As CTextFile, nLine As Long, _
                          tAcumulador As TMoney, tTot As TMoney, nNumDec As Long, _
                          sMsg As String, _
                          Optional sNat As String = "", Optional bNatCute As Boolean = False, _
                          Optional bIsTot As Boolean)
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    oTextFile.PrintLine SpcFill(sMsg, 42) & _
                        SpcFillAmount(tAcumulador.SI, 22, nNumDec) & "  " & _
                        SpcFillAmount(tAcumulador.Cargos, 22, nNumDec) & "  " & _
                        SpcFillAmount(tAcumulador.Abonos, 22, nNumDec)
    oTextFile.PrintNewLine
    
    If Not bIsTot Then
      AddMoney tTot, tAcumulador.SI, tAcumulador.Cargos, tAcumulador.Abonos, 0
      ClearMoney tAcumulador
    End If
    nLine = nLine + 3
    
    If (bNatCute) Then
      oTextFile.PrintNewLine
      oTextFile.PrintLine "NATURALEZA       : " & sNat
      oTextFile.PrintNewLine
      nLine = nLine + 3
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot51_52", Err.Number
End Sub

Private Sub PrintTot53_54(oTextFile As CTextFile, nLine As Long, _
                          tAcumulador As TMoney, nNumDec As Long, _
                          sNat As String)
  On Error GoTo ErrHandler
    PrintQtyLines oTextFile
    oTextFile.PrintLine SpcFill(sNat, 17) & _
                        SpcFillAmount(tAcumulador.SI, 22, nNumDec) & "  " & _
                        SpcFillAmount(tAcumulador.Cargos, 22, nNumDec) & "  " & _
                        SpcFillAmount(tAcumulador.Abonos, 22, nNumDec)
    oTextFile.PrintNewLine
    oTextFile.PrintNewLine
    nLine = nLine + 4
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintTot53_54", Err.Number
End Sub


Private Function GetSubCuenta(ByVal sNumCuenta As String, _
                              Optional sCuenta As String = "", Optional sSubcuenta As String = "") As String
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
  If (nLen < 12) Then
    sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
  End If
  
  GetSubCuenta = sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetSubCuenta", Err.Number
End Function

Private Function GetSubCuenta1(ByVal sNumCuenta As String, _
                                   Optional sCuenta As String = "", Optional sSubcuenta As String = "") As String
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
  If (nLen < 12) Then
    sSubcuenta = sSubcuenta & Mid("000000000000", 1, 12 - nLen)
  End If
  
  GetSubCuenta1 = sSubcuenta
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetSubcuenta", Err.Number
End Function

Private Function GETAccName(oRs As Recordset, sNumAcc As String, nId As Long) As String
  Dim sName         As String
  '***********************************************************************************
 On Error GoTo ErrHandler
  With oRs
    .Filter = " NUM_CUENTA = '" & sNumAcc & "'"
    If (.BOF And .EOF) Then
      sName = "  "
    Else
      sName = !nom_cuenta
    End If
    .Filter = adFilterNone
  End With
   GETAccName = sName
  Exit Function
ErrHandler:
  RaiseError ClassId, "GETAccName", Err.Number
End Function


Private Sub nLinesNew(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                      sDateTime As String, dIniDate As Date, dEndDate As Date, dDate As Date, _
                      sGLname As String, nBan_o_Fid As Boolean)
  On Error GoTo ErrorHandler
   If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                               BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                           FECHA  : " & Format(Date, "dd/mmm/yyyy")
      oTextFile.PrintLine "                                                          INSTITUCION BANCA DE DESARROLLO                                    PAGINA : " & nPage
      oTextFile.PrintLine "                                                    POLIZAS DEL : " & Format(dIniDate, "dd/mmm/yyyy") & " AL : " & Format(dEndDate, "dd/mmm/yyyy")
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      PrintLinesHead oTextFile, dDate, sGLname, nBan_o_Fid
      nPage = nPage + 1
      nLine = 13
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "nLinesNew", Err.Number
End Sub

Private Sub PrintLinesHead(oTextFile As CTextFile, dDate As Date, sGLname As String, nBan_o_Fid As Boolean)
  On Error GoTo ErrorHandler
    If Not (nBan_o_Fid) Then
      oTextFile.PrintLine "DELEGACION  : " & sGLname
    ElseIf (nBan_o_Fid) Then
      oTextFile.PrintLine "FIDEICOMISO : " & sGLname
    End If
    oTextFile.PrintLine "FECHA : " & Format(dDate, "dd/mmm/yyyy")
    oTextFile.PrintLine "POLIZA           CONCEPTO                 FECHA_AFECT FECHA_REG    NOMBRE USUARIO                     CARGOS                   ABONOS              DIFERENCIA"
    oTextFile.PrintLine "=============================================================================================================================================================="
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintLinesHead", Err.Number
End Sub

Private Sub nLines70(oTextFile As CTextFile, nLine As Long, nPage As Long, nMoneda As Long, sMonedaname As String, dCuteDate As Date, sDateTime As String)
  On Error GoTo ErrorHandler
    If (nPage = 0) Then
      nPage = 1
    End If
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "REPORT70"
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                       BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                        FECHA  : " & Format(Date, "dd/mmm/yyyy")
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                             INSTITUCION DE BANCA DE DESARROLLO                                                 PAGINA : " & nPage
      oTextFile.PrintLine "                                                         DISPONIBILIDADES FINANCIERAS DE FIDEICOMISOS Y MANDATOS"
      oTextFile.PrintLine "BNO-670315-CDO                                                   POR TIPO DE MONEDA AL: " & Format(dCuteDate, "dd/mmm/yyyy")
      oTextFile.PrintNewLine
      PrintLines70 oTextFile, sMonedaname
      nPage = nPage + 1
      nLine = 15
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "nLines70", Err.Number
End Sub

Private Sub PrintLines70(oTextFile As CTextFile, sCurrencyName As String)
  On Error GoTo ErrHandler
    oTextFile.PrintLine "MONEDA          : " & sCurrencyName
    oTextFile.PrintLine "FID.  NOMBRE                                                      1101                     1202                     1204                     1103                     1206"
    oTextFile.PrintLine "                                                                  CAJA     VALORES EMITIDOS POR            OTROS VALORES             OTROS BANCOS  VALORES GUBERNAMENTALES"
    oTextFile.PrintLine "                                                                          INST. NAC. DE CREDITO                                                            U D I S    "
    oTextFile.PrintLine "===========================================================================================================================================================================    "
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintLines70", Err.Number
End Sub


Public Function BalanceAccount(sAppServer As String, dEndDate As Date, dPartialExcDate As Date, _
                           sAccount As String, sAccount_2 As String, vPartialGL As Variant, _
                           Optional sPartialPattern As String = "", Optional nPartialExcRateType As Long = 46, _
                           Optional nPartialExcRateCurrency As Long = 1, _
                           Optional bPartialOnlyLastLevels As Boolean, _
                           Optional bMovements As Boolean = False) As Currency
  Dim oPartialRS        As ADODB.Recordset
  Dim oPartialBalance   As SCFFixedReportsDS.CBalances
  Dim dInitialDate      As Date
  '********************************************************************************
  On Error GoTo ErrHandler
  dInitialDate = dEndDate - (Day(dEndDate) + 1)
  
  Set oPartialBalance = New SCFFixedReportsDS.CBalances
  Set oPartialRS = oPartialBalance.BalanceGeneral(sAppServer, sPartialPattern, vPartialGL, dInitialDate, dEndDate, _
                                sAccount, sAccount_2, nExcRateType:=nPartialExcRateType, nExcRateCurrency:=nPartialExcRateCurrency, dExcDate:=dPartialExcDate, _
                                bOnlyLastLevels:=bPartialOnlyLastLevels)
                                 
  With oPartialRS
    If (.BOF And .EOF) Then
      BalanceAccount = 0
    Else
      Do While Not .EOF
        If (!NUMERO_CUENTA_ESTANDAR = sAccount) And (!clave_sector = "00") Then
          If bMovements Then
            BalanceAccount = !Saldo_Actual - !saldo_inicial
          Else
            BalanceAccount = !Saldo_Actual
          End If
          Exit Do
        Else
          .MoveNext
        End If
      Loop
    End If
    .Close
    .ActiveConnection = Nothing
  End With
  Set oPartialRS = Nothing
  Set oPartialBalance = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceAccount", Err.Number
End Function


Private Sub PrintHeader99(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                    dCuteDate As String, _
                                    sGLDesc As String, _
                                    skindacc As String, _
                                    sAccount As String, _
                                    sDateTime As String)
  
  On Error GoTo ErrorHandler
    If (nLine >= cnPageLines) Then
      'Imprimir caracter de avance de pagina
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                                                                          FECHA  : " & Date
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                                                                                        PAGINA : " & nPage
      oTextFile.PrintLine "BNO-670315-CDO                           BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C "
      oTextFile.PrintLine "                                                   VALORIZACION DE MONEDAS EXTRANJERAS"
      oTextFile.PrintLine "                                                           " & dCuteDate
      oTextFile.PrintNewLine
      oTextFile.PrintLine "COMISION NACIONAL BANCARIA "
      oTextFile.PrintLine "CLAVE DEL BANCO: 221 "
      oTextFile.PrintLine "CONTABILIDAD : " & sGLDesc
      oTextFile.PrintLine "NATURALEZA   : " & skindacc
      oTextFile.PrintLine "================================================================================================================================================"
      oTextFile.PrintLine "                                                    CLASE DE       CVE. MON.          IMPORTE MON.           TIPO              EQUIVALENCIA EN "
      oTextFile.PrintLine "NUMERO Y NOMBRE DE LAS CUENTAS                       MONEDA                           EXTRANJERA             CAMBIO            MONEDA NACIONAL"
      oTextFile.PrintLine "================================================================================================================================================"
      nPage = nPage + 1
      nLine = 20
    End If
    Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeader99", Err.Number
End Sub


Private Sub PrintHeader126(oTextFile As CTextFile, _
                            nLine As Long, nPage As Long, _
                            dCuteDate As Date, _
                            sGl As String, sDTime As String, _
                            sRepNum As String)
  Dim sDate       As String
  '**********************************************************
  On Error GoTo ErrHandler
    If (nLine >= cnPageLines) Then
      If nLine <> 100 Then
        oTextFile.PageFeed (nLine - 2)
      End If
      sDate = Month(dCuteDate) & " DE " & Year(dCuteDate)
      oTextFile.PrintLine sDTime
      oTextFile.PrintLine sRepNum
      oTextFile.PrintLine "TECOYOTITLA NUM 100                            BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, B.N.C.                     FECHA: " & Date
      oTextFile.PrintLine "COLONIA LA FLORIDA                           DECLARACION ANUAL RETENCION DEL IMPUESTO SOBRE LA RENTA                   PAGINA:   " & nPage
      oTextFile.PrintLine "C.P. 01030 MEXICO D.F.                                       SALDOS AL MES " & sDate
      oTextFile.PrintLine "BNO-670315-CDO"
      oTextFile.PrintLine "AGENCIA : " & sGl
      oTextFile.PrintNewLine
      oTextFile.PrintLine "================================================================================================================================================================="
      oTextFile.PrintLine "  CLAVE REG.FEDERAL   CURP             NOMBRE                                  CLAVE        TOTAL BASE         TOTAL BASE      TOTAL RETENCION    TOTAL RETENCION"
      oTextFile.PrintLine "  DE CONTRIBUYENTES                                                          RETENCION        I.S.R.             I.V.A.             I.S.R.             I.V.A.    "
      oTextFile.PrintLine "================================================================================================================================================================="
      nLine = 12
      nPage = nPage + 1
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintHeader126", Err.Number
End Sub

Private Sub PrintHeaderVouchers(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sAccType As String, _
                           dIniDate As Date, dEndDate As Date, _
                           sTittleReport As String, sDateTime As String)
  On Error GoTo ErrorHandler
   nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      
      oTextFile.PrintLine FormatLine(174, "BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C", sTittleReport, "Tecoyotitla Num. 100")
      oTextFile.PrintLine FormatLine(174, "INSTITUCION DE BANCA DE DESARROLLO", PrintDateGeneral(dIniDate, dEndDate), "C.P. 01030 México D.F.")
      oTextFile.PrintLine FormatLine(174, "T. DE CONTABILIDAD : " & sAccType, " ", " ")
      oTextFile.PrintLine FormatLine(174, " ", " ", sDateTime & " -- Página " & nPage)
      oTextFile.PrintNewLine
      oTextFile.PrintLine "NUMERO POLIZA    AFECTACION  ELABORACION   ELABORADO POR            CONCEPTO                                                          CARGOS                      ABONOS "
      PrintQtyLines oTextFile, 182
      nPage = nPage + 1
      nLine = 16
   End If
  Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderVouchers", Err.Number
End Sub

Private Sub PrintHeaderParamWithDestiny(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                           sGl As String, sCurrency As String, _
                           sGLN As String, _
                           dIniDate As Date, dEndDate As Date, _
                           sTittleReport As String, sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
  nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                    BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                         INSTITUCION BANCA DE DESARROLLO                                                PAGINA : " & nPage
      oTextFile.PrintLine FormatLine(180, "C.P. 01030 MEXICO D.F.   ", sTittleReport, " ")
      oTextFile.PrintLine "BNO-670315-CDO      "
      oTextFile.PrintNewLine

      oTextFile.PrintLine "NUMERO CUENTA           CLAVE SECTOR    ABREV MONEDA          NUMERO CUENTA           CLAVE SECTOR    NUMERO AUXILIAR CUENTA          NUMERO CUENTA         CLAVE SECTOR    NUMERO AUXILIAR CUENTA "
      oTextFile.PrintLine "  ORIGEN                   ORIGEN          ORIGEN                DESTINO                 DESTINO             DESTINO                    SOBREGIRO             SOBREGIRO            SOBREGIRO       "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderParamWithDestiny", Err.Number
End Sub

Private Sub PrintHeaderParamWithoutDestiny(oTextFile As CTextFile, nLine As Long, nPage As Long, _
                                           sGl As String, sCurrency As String, sGLN As String, _
                                           dIniDate As Date, dEndDate As Date, _
                                           sTittleReport As String, sDateTime As String, sAccType As String)
  On Error GoTo ErrorHandler
  nLine = nLine + 1
   If (nLine >= cnPageLines) Then
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintNewLine
      oTextFile.PrintLine sDateTime
      oTextFile.PrintLine "                                                                    BANCO NACIONAL DE OBRAS Y SERVICIOS PUBLICOS, S.N.C                                     FECHA  : " & Date
      oTextFile.PrintLine "TECOYOTITLA NUM. 100                                                    INSTITUCION BANCA DE DESARROLLO                                                     PAGINA : " & nPage
      oTextFile.PrintLine FormatLine(180, "C.P. 01030 MEXICO D.F.   ", sTittleReport, " ")
      oTextFile.PrintLine "BNO-670315-CDO      "
      oTextFile.PrintLine "CONTABILIDAD         : " & sGl & " " & sGLN
      oTextFile.PrintLine "TIPO DE CONTABILIDAD : "
      oTextFile.PrintNewLine

      oTextFile.PrintLine "NUMERO DE                     NOMBRE DE                       NUMERO CUENTA                     CLAVE SECTOR    ABBREV MONEDA         "
      oTextFile.PrintLine "  GRUPO                         GRUPO                           ORIGEN                             ORIGEN          ORIGEN             "
      PrintQtyLines oTextFile
      nPage = nPage + 1
      nLine = 16
   End If
   Exit Sub
ErrorHandler:
  RaiseError ClassId, "PrintHeaderParamWithoutDestiny", Err.Number
End Sub

'********************************************************************************
'
'      Útilerias Generales
'
'********************************************************************************

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

Private Sub MinusMoney(tRes As TMoney, tMinuendo As TMoney, tSustraendo As TMoney)
  On Error GoTo ErrHandler
    tRes.SI = tMinuendo.SI - tSustraendo.SI
    tRes.Cargos = tMinuendo.Cargos - tSustraendo.Cargos
    tRes.Abonos = tMinuendo.Abonos - tSustraendo.Abonos
    tRes.SA = tMinuendo.SA - tSustraendo.SA
  Exit Sub
ErrHandler:
  RaiseError ClassId, "MinusMoney", Err.Number
End Sub

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

Private Function SpcFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
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

Private Function SpcFillAmount1(ByVal cAmount As String, nLen As Long, Optional nHowManyDec As Long = 2) As String
  Dim sFmtAmount As String
  Dim sSign      As String
  '************************************************************************
  On Error GoTo ErrorHandler
    sSign = " "
    If (cAmount < 0) Then
          cAmount = cAmount * -1
          nLen = nLen - 1
           sFmtAmount = FormatNumber(cAmount, nHowManyDec)
          If (nLen - Len(sFmtAmount)) > 0 Then
          SpcFillAmount1 = Space(nLen - Len(sFmtAmount)) & "(" & sFmtAmount & ")"
           
          Else
          nLen = nLen - 1
          SpcFillAmount1 = "(" & sFmtAmount & ")"
         End If
         Exit Function
       End If
    sFmtAmount = FormatNumber(cAmount, nHowManyDec)
    If (nLen - Len(sFmtAmount)) > 0 Then
      SpcFillAmount1 = Space(nLen - Len(sFmtAmount)) & sFmtAmount & " "
    Else
      SpcFillAmount1 = sFmtAmount & " "
    End If
  Exit Function
ErrorHandler:
  RaiseError ClassId, "SpcFillAmount1", Err.Number
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

Private Function ZeroFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
    ZeroFill = Mid("00000000000000000000000", 1, nLen - Len(sStr)) & sStr
  Exit Function
ErrorHandler:
  RaiseError ClassId, "ZeroFill", Err.Number
End Function

Private Function GetFmtDate() As String
  On Error GoTo ErrHandler
    GetFmtDate = Format(Date, "dd/mmm/yyyy") & "  " & Format(Time, "hh:mm:ss")
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetFmtDate", Err.Number
End Function

Private Function GetGroupsNames() As Collection
  Dim cCol  As New Collection
  '********************************************
  On Error GoTo ErrHandler
    ' Grupos de cuentas deudoras
    cCol.Add "CAJA BANCOS Y CORRESPONSALES", "11"
    cCol.Add "CARTERA DE VALORES", "12"
    cCol.Add "CARTERA DE CREDITOS", "13"
    cCol.Add "FUTUROS Y REPORTOS", "14"
    cCol.Add "DEUDORES DIVERSOS", "15"
    cCol.Add "OTRAS INVERSIONES", "16"
    cCol.Add "MUEBLES E INMUEBLES", "18"
    cCol.Add "PAGOS ANTIC.Y CARGOS DIFERIDOS", "19"
    cCol.Add "CAP NO SUS NO EX Y PERD L E AN", "41"
    cCol.Add "CUENTAS DE RESULTADOS DEUDORAS", "51"
    cCol.Add "CUENTAS DEUDORAS CONTIGENTES", "61"
    cCol.Add "BIENES E DEPOSITO COSTO O ADMN", "62"
    cCol.Add "CUENTAS DEUDORAS DE REGISTRO", "63"
    ' Grupos de cuentas acreedoras
    cCol.Add "CAPTACION DE RECURSOS DEL PUBL", "21"
    cCol.Add "DEPOSITOS Y PRESTAMOS DE BANCO", "22"
    cCol.Add "OTROS DEPOSITOS Y OBLIGACIONES", "23"
    cCol.Add "FUTUROS Y REPORTOS", "24"
    cCol.Add "RVAS Y PROV P. OBLIG DIVERSAS", "25"
    cCol.Add "CREDITOS DIFERIDOS", "26"
    cCol.Add "ESTIMAC P/DEPREC Y CASTIGOS", "31"
    cCol.Add "CAPITAL SOCIAL", "42"
    cCol.Add "RESERVAS", "43"
    cCol.Add "UTILIDADES POR APLICAR", "44"
    cCol.Add "SUPERAVITS X REVAL DE ACTIVOS", "45"
    cCol.Add "CUENTAS DE RESULTADOS ACREEDOR", "52"
    cCol.Add "CUENTAS ACREE DE CONTINGENCIA", "64"
    cCol.Add "VAL EN DEP CUSTODIA O ADMON", "65"
    cCol.Add "CUENTAS ACREEDORAS DE REGISTRO", "66"
    cCol.Add "CONTROL INTERNO DE CARTERA", "99"
    
    Set GetGroupsNames = cCol
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGroupsNames", Err.Number
End Function

