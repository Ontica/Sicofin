VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : GEM And PyC Interfaces (SCFIGemPyC)             Parte   : Business services                 *
'* Clase      : CInterface                                      Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Importa movimientos desde el sistema GEM y exporta los movimientos al sistema PyC.          *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CInterface"

Private Const cnPostfixBalanceExpDate  As String = "_BLED"
Private Const cnPostfixTransEstimatedTime = "_IMP_ESTIM_TIME"
Private Const cnPostfixTransExpDate As String = "_TLED"
Private Const cnPostfixTransImpDate As String = "_TLID"
Private Const cnPostfixTransImpTimer As String = "_IMP_TIMER"
Private Const cnVoucherTypeId As Long = 57

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

Private Property Get GetTransactionType(nSystemId As Long) As Long
  Select Case nSystemId
    Case 1
      GetTransactionType = 84
    Case 4
      GetTransactionType = 85
    Case 5
      GetTransactionType = 86
    Case 7
      GetTransactionType = 87
    Case 10
      GetTransactionType = 88
    Case Else
      GetTransactionType = 20
  End Select
End Property

Public Property Get BalancesLastExportationDate(sTargetAppServer As String) As Variant
  Dim sLastDateConstantName As String, dLastExportationDate As Date
  '***********************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixBalanceExpDate
    dLastExportationDate = CDate(GetConstant(sConstantName:=sLastDateConstantName))
    BalancesLastExportationDate = Format$(dLastExportationDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  If (Err.Number Xor vbObjectError) = TEnumErrors.ErrConstantNotFound Then
    BalancesLastExportationDate = "Nunca"
  Else
    RaiseError ClassId, "BalancesLastExportationDate", Err.Number
  End If
End Property

Public Property Get EstimatedImportTime(sTargetAppServer As String, nVouchers As Long) As String
  Dim sLastDateConstantName As String, sEstimatedImportTime As String, nSecondsPerVoucher  As Long
  Dim nEstimatedTime As Long
  '***********************************************************************************************
  On Error GoTo ErrHandler
    sEstimatedImportTime = sTargetAppServer & cnPostfixTransEstimatedTime
    nSecondsPerVoucher = CDate(GetConstant(sConstantName:=sEstimatedImportTime))
    nEstimatedTime = (nSecondsPerVoucher * nVouchers) \ 60
    If (nEstimatedTime = 0) Then
      nEstimatedTime = 1
    End If
    EstimatedImportTime = nEstimatedTime
  Exit Property
ErrHandler:
  RaiseError ClassId, "EstimatedImportTime", Err.Number
End Property

Public Property Get GEMErrorTransactionsCount(sSourceAppServer As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COUNT(*) Counter FROM MC_Encabezados_Err"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sSourceAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GEMErrorTransactionsCount = !Counter
      Else
        GEMErrorTransactionsCount = -1
      End If
      .Close
    End With
  Exit Property
ErrHandler:
  RaiseError ClassId, "GEMPendingTransactionsCount", Err.Number
End Property

Public Property Get GEMPendingTransactionsCount(sSourceAppServer As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COUNT(*) Counter FROM MC_Encabezados"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sSourceAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GEMPendingTransactionsCount = !Counter
      Else
        GEMPendingTransactionsCount = -1
      End If
      .Close
    End With
  Exit Property
ErrHandler:
  RaiseError ClassId, "GEMPendingTransactionsCount", Err.Number
End Property

Public Property Get ImportTimer(sTargetAppServer As String) As String
  Dim sTemp As String, sImportTimer As String
  '******************************************************************
  On Error GoTo ErrHandler
    sImportTimer = sTargetAppServer & cnPostfixTransImpTimer
    ImportTimer = CLng(GetConstant(sConstantName:=sImportTimer)) & " minutos"
  Exit Property
ErrHandler:
  RaiseError ClassId, "ImportTimer", Err.Number
End Property

Public Property Get TransactionsLastExportationDate(sTargetAppServer As String) As Variant
  Dim sLastDateConstantName As String, dLastExportationDate As Date
  '***************************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixTransExpDate
    dLastExportationDate = CDate(GetConstant(sConstantName:=sLastDateConstantName))
    TransactionsLastExportationDate = Format$(dLastExportationDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  If (Err.Number Xor vbObjectError) = TEnumErrors.ErrConstantNotFound Then
    TransactionsLastExportationDate = "Nunca"
  Else
    RaiseError ClassId, "TransactionsLastExportationDate", Err.Number
  End If
End Property

Public Property Get TransactionsLastImportationDate(sTargetAppServer As String) As Variant
  Dim sLastDateConstantName As String, dLastImportationDate As Date
  '***************************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixTransImpDate
    dLastImportationDate = CDate(GetConstant(sConstantName:=sLastDateConstantName))
    TransactionsLastImportationDate = Format$(dLastImportationDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  If (Err.Number Xor vbObjectError) = TEnumErrors.ErrConstantNotFound Then
    TransactionsLastImportationDate = "Nunca"
  Else
    RaiseError ClassId, "TransactionsLastImportationDate", Err.Number
  End If
End Property

Public Property Get URLFilesPath() As String
  On Error GoTo ErrHandler
    URLFilesPath = GetConstant(cnURLFilesPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "URLFilesPath", Err.Number
End Property

'************************************************************************************************************
'* PROPIEDADES PRIVADAS                                                                                     *
'************************************************************************************************************

Private Sub SetEstimatedImportTime(sTargetAppServer As String, nSecondsPerVoucher As Long)
  Dim sEstimatedImportTime As String
  '***************************************************************************************
  On Error GoTo ErrHandler
    sEstimatedImportTime = sTargetAppServer & cnPostfixTransEstimatedTime
    If nSecondsPerVoucher < 1 Then
      nSecondsPerVoucher = 1
    End If
    WriteConstant sEstimatedImportTime, CStr(nSecondsPerVoucher)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SetEstimatedImportTime", Err.Number
End Sub

Private Property Let BalancesLastExportationDate(sTargetAppServer As String, dNewDate As Variant)
  Dim sLastDateConstantName As String
  '**********************************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixBalanceExpDate
    WriteConstant sLastDateConstantName, Format$(dNewDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  RaiseError ClassId, "BalancesLastExportationDate", Err.Number
End Property

Private Property Let TransactionsLastExportationDate(sTargetAppServer As String, dNewDate As Variant)
  Dim sLastDateConstantName As String
  '**************************************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixTransExpDate
    WriteConstant sLastDateConstantName, Format$(dNewDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  RaiseError ClassId, "TransactionsLastExportationDate", Err.Number
End Property

Private Property Let TransactionsLastImportationDate(sTargetAppServer As String, dNewDate As Variant)
  Dim sLastDateConstantName As String
  '***************************************************************************************
  On Error GoTo ErrHandler
    sLastDateConstantName = sTargetAppServer & cnPostfixTransImpDate
    WriteConstant sLastDateConstantName, Format$(dNewDate, "dd/mmm/yyyy hh:nn:ss")
  Exit Property
ErrHandler:
  RaiseError ClassId, "TransactionsLastImportationDate", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

'Public Function ExportBalances(sAppServer As String, sTargetAppServer As String, _
'                               Optional dToDate As Variant) As Recordset
'  Dim oBalancesRS As New Recordset, oBalances As New EFABalanceEngine.CEngine
'  Dim dTemp As Date
'  '**********************************************************************************************************
'  On Error GoTo ErrHandler
'    If IsMissing(dToDate) Then
'      dTemp = Date
'    Else
'      dTemp = dToDate
'    End If
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "ExportBalances", Err.Number
'End Function

Public Function ExportTransactions(sAppServer As String, sTargetAppServer As String, _
                                   vGralLedgers As Variant) As String
  'Dim oBalances As New AOGLBalances.CBalances
  Dim oSourceTransactions As Recordset, oTargetVouchersRS As New Recordset
  Dim oTargetPostingsRS As New Recordset, oFile As New CTextFile
  Dim dFromDate As Date, dToDate As Date
  Dim sLine As String, nCounter, j As Long, sDecAmmount As String, nPosPoint As Long
  Dim sDivAmmount As String, sSubNumber As String, sOrabanksArea As String
  Dim nTransactionsCounter As Long, nCurrentTransactionId As Long
  Dim nStdAccountTypeId As Long, dApplicationDate As Date
  Dim nDebitSum As Double, nCreditSum As Double
  Dim sResponsibilityArea As String, dElaborationDate As Date, sVoucherConcept As String
  Dim nCurrentGralLedgerId As Long, sGralLedger As String, sSectionNumber As String, sUser As String
  Dim sVoucherSource As String, sExtVoucherNumber As String
  '*************************************************************************************************
  On Error GoTo ErrHandler
    dFromDate = TransactionsLastExportationDate(sTargetAppServer)
    dToDate = Now()
    Sleep 10000       'Espera 10 segs. para evitar que se estén ingresando pólizas con la misma fecha
    oFile.Open_ GetConstant(cnGeneratedFilesPath)
    Set oSourceTransactions = TransactionsForExport(sAppServer, vGralLedgers, dFromDate, dToDate, True)
    With oTargetVouchersRS
      .ActiveConnection = GetConstant(sConstantName:=sTargetAppServer)
      .CursorLocation = adUseClient
      .Open "SELECT * FROM BGM_Encabezados WHERE (bgme_num_vol = 0)", , adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    With oTargetPostingsRS
      .ActiveConnection = GetConstant(sConstantName:=sTargetAppServer)
      .CursorLocation = adUseClient
      .Open "SELECT * FROM BGM_Movimientos WHERE (bgm_num_vol = 0)", , adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    With oSourceTransactions
      Do While Not .EOF
        If (nCurrentTransactionId <> !id_transaccion) Then
          If (nCurrentTransactionId <> 0) Then
            PutExternalTransaction sAppServer, oTargetVouchersRS, nStdAccountTypeId, dApplicationDate, _
                                   sExtVoucherNumber, sVoucherConcept, sUser, nDebitSum, nCreditSum
          End If
          nCurrentTransactionId = !id_transaccion
          nTransactionsCounter = nTransactionsCounter + 1
          nStdAccountTypeId = !id_tipo_cuentas_std
          If (nStdAccountTypeId = 1) Then
            sExtVoucherNumber = Format$(Right$(!numero_mayor, 2), "00") & Mid$(!numero_transaccion, 4, 1) & _
                                Mid$(!numero_transaccion, 6, 2) & Mid$(!numero_transaccion, 9, 5)
          Else
            '¿No se sabe cómo integrar el número de póliza para fideicomisos
            sExtVoucherNumber = Format$(Right$(!numero_mayor, 2), "00") & Mid$(!numero_transaccion, 4, 1) & _
                                Mid$(!numero_transaccion, 6, 2) & Mid$(!numero_transaccion, 9, 5)
          End If
          dApplicationDate = !fecha_afectacion
          dElaborationDate = !fecha_registro
          sVoucherConcept = "(" & !numero_transaccion & ") " & !concepto_transaccion
          'sVoucherSource = !fuente
          nDebitSum = 0
          nCreditSum = 0
          If (nStdAccountTypeId = 1) Then
            sResponsibilityArea = GetGralLedgerResponsibilityArea(sAppServer, !id_mayor)
            sResponsibilityArea = "0"
          Else
            sResponsibilityArea = "0000"
          End If
          sUser = !elaborada_por
        End If
        PutExternalPosting sAppServer, oTargetPostingsRS, nStdAccountTypeId, _
                           dApplicationDate, dElaborationDate, sExtVoucherNumber, _
                          !id_movimiento, !numero_cuenta_estandar, IIf(IsNull(!clave_sector), "00", !clave_sector), _
                           IIf(IsNull(!NUMERO_CUENTA_AUXILIAR), "", !NUMERO_CUENTA_AUXILIAR), !id_area_responsabilidad, _
                           IIf(IsNull(!concepto_movimiento), "", !concepto_movimiento), !clave_disponibilidad, !clave_presupuestal, _
                           !numero_verificacion, !tipo_movimiento, !o_id_moneda, !tipo_cambio, !monto
        If (!tipo_movimiento = "D") Then
          nDebitSum = nDebitSum + !monto
        Else
          nCreditSum = nCreditSum + !monto
        End If
        .MoveNext
      Loop
      If (nCurrentTransactionId <> 0) Then
        PutExternalTransaction sAppServer, oTargetVouchersRS, nStdAccountTypeId, dApplicationDate, _
                               sExtVoucherNumber, sVoucherConcept, sUser, nDebitSum, nCreditSum
      End If
      .Close
    End With
    If nTransactionsCounter <> 0 Then
      oTargetVouchersRS.MoveFirst
      Do While Not oTargetVouchersRS.EOF
        oTargetVouchersRS.MoveNext
      Loop
      oTargetVouchersRS.MoveFirst
      SaveRecordset sTargetAppServer, oTargetVouchersRS
      SaveRecordset sTargetAppServer, oTargetPostingsRS
    End If
    TransactionsLastExportationDate(sTargetAppServer) = dToDate
    ExportTransactions = oFile.FileName
    oFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExportTransactions", Err.Number
End Function

Public Function ImportedFilesTable() As String
  Const cRow = "<tr><td><INPUT type='checkbox' name=chkItem></td><td><a href='<@FILE@>'><@DATE@></a></td></tr>" & vbCrLf
  Dim oFileMgr As New EGEFileManager.CDirectory, aFiles As Variant
  Dim sDir As String, sFile As String, dFileDate As Date, sHTML As String, sTemp As String, i As Long
  '***********************************
  On Error GoTo ErrHandler
    sDir = GetConstant(cnGeneratedFilesPath) & "\"
    aFiles = oFileMgr.Contents(sDir, "gem*.*", True)
    If IsArray(aFiles) Then
      For i = LBound(aFiles) To UBound(aFiles)
        sFile = aFiles(i)
        sTemp = Replace(cRow, "<@FILE@>", sDir & sFile)
        dFileDate = DateSerial(CInt(Mid$(sFile, 4, 4)), CInt(Mid$(sFile, 8, 2)), CInt(Mid$(sFile, 10, 2)))
        sTemp = Replace(sTemp, "<@DATE@>", Format$(dFileDate, "dddd dd"" de ""mmmm"" del ""yyyy"))
        sHTML = sHTML & sTemp
      Next i
    Else
      sHTML = "<tr><td>&nbsp;</td><td><b>No encontré archivos con resultados de importación.</b></td></tr>"
    End If
    ImportedFilesTable = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "ImportedFilesTable", Err.Number
End Function

Public Function ImportTransactions(sAppServer As String, sSourceAppServer As String, nUserId As Long, _
                                   Optional bForwardToUsers As Boolean = True, _
                                   Optional bProtectPostings As Boolean = False, _
                                   Optional bAutoGenerateSubsidiaryAccounts As Boolean = True) As String
  Dim oSourceRS As Recordset, oTransactionsRS As Recordset, oPostingsRS As Recordset, oLogFile As New CTextFile
  Dim sTransaction As String, sTransactionConcept As String, dAfectationDate As Date
  Dim sFilePath As String, sFileName As String, dImportDate As Date, nSecondsPerVoucher As Long
  Dim sAccountNumber As String, sCurrencyKey As String, sSectorKey As String, sSubsidiaryAccount As String
  Dim nAmount As Double, nBaseAmount As Double, nCurrencyId As Long
  Dim sStdAccountNumber As String, sSector As String, nTransactionId As Long, nStdAccountId As Long
  Dim nAccountId As Long, nGralLedgerId As Long, nSectorId As Long, nSubsidiaryAccountId As Long
  Dim nCounter As Long, nLine As Long, nUserToSend As Long, nSourceId As Long, dElaborationDate As Date
  Dim nResult As Long, nExchangeRate As Double, bTransactionHasErrors As Boolean, nStdAccountTypeId As Long
  Dim sStdAccountRole As String, sResponsibilityArea As String, nResponsibilityAreaId As Long, sDisponibilityKey As String
  Dim sVerificationNumber As String, sPostingConcept As String, sBudgetKey As String, sPostingType As String
  Dim nTransactionTypeId As Long, nVoucherTypeId As Long, nEncTipoCont As Long, dFechaVol As Date, nNumVol As Long
  Dim nOkCounter As Long, nErrorCounter As Long
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    dImportDate = Now()
    sFileName = "gem" & Format$(dImportDate, "yyyymmdd") & ".txt"
    If oLogFile.Exists(GetConstant(cnGeneratedFilesPath) & "\" & sFileName) Then
      oLogFile.Open_ GetConstant(cnGeneratedFilesPath), sFileName, False, True
      PrintLogFileSubHeader oLogFile, Format$(dImportDate, "dd/mmm/yyyy hh:nn:ss")
    Else
      oLogFile.Open_ GetConstant(cnGeneratedFilesPath), sFileName
      PrintLogFileHeader oLogFile, Format$(dImportDate, "dd/mmm/yyyy hh:nn:ss")
      PrintLogFileSubHeader oLogFile, Format$(dImportDate, "dd/mmm/yyyy hh:nn:ss")
    End If
    Set oTransactionsRS = ConstructTransactionsRS(sAppServer)
    Set oPostingsRS = ConstructPostingsRS(sAppServer)
    Set oSourceRS = GetGemTransactions(sSourceAppServer)
    With oSourceRS
      Do While Not .EOF
        If (sTransaction <> (!ENC_TIPO_CONT & "-" & !ENC_FECHA_VOL & "-" & Format$(!ENC_NUM_VOL, "000000"))) Then
          If (Len(sTransaction) <> 0) And (Not bTransactionHasErrors) Then
            nTransactionId = PutTransaction(sAppServer, oTransactionsRS, "No actualizada", nGralLedgerId, _
                                             nSourceId, nTransactionTypeId, nVoucherTypeId, sTransactionConcept, _
                                            dAfectationDate, dElaborationDate, nUserId, nUserToSend)
            DumpGemTransaction sAppServer, sSourceAppServer, nTransactionId, oTransactionsRS, oPostingsRS, _
                               nEncTipoCont, dFechaVol, nNumVol
            PrintLogFileLine oLogFile, sTransaction, False, "¡La póliza fue importada satisfactoriamente!"
            nOkCounter = nOkCounter + 1
            Set oTransactionsRS = ConstructTransactionsRS(sAppServer)
            Set oPostingsRS = ConstructPostingsRS(sAppServer)
          ElseIf bTransactionHasErrors Then
            bTransactionHasErrors = False
            Set oPostingsRS = ConstructPostingsRS(sAppServer)
          End If
          nCounter = nCounter + 1
          nEncTipoCont = !ENC_TIPO_CONT
          dFechaVol = !ENC_FECHA_VOL
          nNumVol = !ENC_NUM_VOL
          sTransaction = nEncTipoCont & "-" & dFechaVol & "-" & Format$(nNumVol, "000000")
          nStdAccountTypeId = nEncTipoCont
          dAfectationDate = dFechaVol
          dElaborationDate = !ENC_FECHA_CAP
          nTransactionTypeId = GetTransactionType(!MCO_SISTEMA)
          nVoucherTypeId = cnVoucherTypeId
          If nStdAccountTypeId = 1 Then
            nGralLedgerId = MapOrabanksSource(sAppServer, !ENC_AREA_CAP)
            nSourceId = GetResponsibilityAreaId(sAppServer, !ENC_AREA_CAP)
          ElseIf nStdAccountTypeId = 2 And (Not IsNull(!ENC_FIDEICOMISO)) Then
            nGralLedgerId = GralLedgerId(sAppServer, nStdAccountTypeId, !ENC_FIDEICOMISO, "")
            nSourceId = GetResponsibilityAreaId(sAppServer, !ENC_AREA_CAP)
          Else
            bTransactionHasErrors = True
          End If
          sTransactionConcept = !ENC_DESCRIP & " [Volante origen:" & Format$(!ENC_NUM_VOL, "000000") & _
                                ", Sistema: " & IIf(IsNull(!MCO_SISTEMA), "", !MCO_SISTEMA) & _
                                ", Operación: " & IIf(IsNull(!MCO_NO_OPERACION), "", !MCO_NO_OPERACION) & "]"
          If Not ValidateTransaction(sTransaction, nLine, nGralLedgerId, nSourceId, _
                                     sTransactionConcept, dAfectationDate, dElaborationDate, oLogFile) Then
            bTransactionHasErrors = True
            GoTo LblProcessNextRecord
          End If
          If bForwardToUsers Then
            nUserToSend = GetUserToSend(sAppServer, nGralLedgerId, nUserId)
          End If
        Else
          If bTransactionHasErrors Then
            GoTo LblProcessNextRecord
          End If
        End If
        nAmount = !MCO_IMPORTE
        nBaseAmount = (!MCO_T_CAMBIO * !MCO_IMPORTE)
        sPostingType = Switch(!MCO_CVE_MOV = 1, "D", !MCO_CVE_MOV = 2, "H")
        sAccountNumber = FormatAccountNumber(!MCO_REG_CONTABLE)
        sSector = IIf(IsNull(!MCO_SECTOR), "", !MCO_SECTOR)
        sSubsidiaryAccount = IIf(IsNull(!MCO_NUM_AUX), "", !MCO_NUM_AUX)
        sCurrencyKey = Format$(!MCO_MONEDA, "00")
        sResponsibilityArea = !MCO_AREA
        sBudgetKey = IIf(IsNull(!MCO_CONCEPTO), "", !MCO_CONCEPTO)
        sVerificationNumber = ""
        sDisponibilityKey = IIf(IsNull(!MCO_DISPONIB), "", !MCO_DISPONIB)
        sPostingConcept = IIf(IsNull(!MCO_DESCRIP), "", !MCO_DESCRIP)
        If (nAmount = 0) Or (nBaseAmount = 0) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "El movimiento no puede tener un monto igual a cero."
          GoTo LblProcessNextRecord
        End If
        nStdAccountId = GetStdAccountId(sAppServer, nStdAccountTypeId, sAccountNumber, dAfectationDate)
        If (nStdAccountId = 0) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, _
                           "La cuenta '" & sAccountNumber & "' no existe en el catálogo estándar " & _
                           "(al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & ")"
          GoTo LblProcessNextRecord
        End If
        If sSector <> "00" Then
          nSectorId = GetSectorId(sAppServer, sSector)
        Else
          nSectorId = 0
        End If
        If (nSectorId = -1) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "El sector '" & sSector & "' no ha sido dado de alta en el sistema."
          GoTo LblProcessNextRecord
        End If
        If Not ValidateStdAccountSector(sAppServer, nStdAccountId, nSectorId, dAfectationDate) Then
          bTransactionHasErrors = True
          If (nSectorId = 0) Then
            PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' maneja sectores pero éste no se proporcionó."
          Else
            PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' no maneja el sector '" & sSector & "' al día " & _
                               Format$(dAfectationDate, "dd/mmm/yyyy") & "."
          End If
          GoTo LblProcessNextRecord
        End If
        sStdAccountRole = GetStdAccountRole(sAppServer, nStdAccountId, nSectorId, dAfectationDate)
        If (sStdAccountRole = "S") Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' es sumaria al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & "."
          GoTo LblProcessNextRecord
        End If
        If (sStdAccountRole = "C") And (Len(sSubsidiaryAccount) = 0) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' (sector " & sSector & ") maneja auxiliar (al día " & _
                                                         Format$(dAfectationDate, "dd/mmm/yyyy") & ") pero éste no se proporcionó."
          GoTo LblProcessNextRecord
        End If
        If (sStdAccountRole = "P") And (Len(sSubsidiaryAccount) <> 0) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' (sector " & sSector & ") no maneja auxiliares (al día " & _
                                                         Format$(dAfectationDate, "dd/mmm/yyyy") & ") pero se proporcionó el auxiliar '" & _
                                                         sSubsidiaryAccount & "'."
          GoTo LblProcessNextRecord
        End If
        nCurrencyId = GetCurrencyId(sAppServer, sCurrencyKey)
        If (nCurrencyId = 0) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "No reconozco la moneda '" & sCurrencyKey & "'."
          GoTo LblProcessNextRecord
        End If
        If Not ValidateStdAccountCurrency(sAppServer, nStdAccountId, nCurrencyId, dAfectationDate) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "La cuenta '" & sAccountNumber & "' no maneja la moneda '" & sCurrencyKey & "'" & _
                                                         "(al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & ")."
          GoTo LblProcessNextRecord
        End If
        If (nAmount <> nBaseAmount) And (nCurrencyId = 1) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "El monto y el monto en moneda base deben ser iguales para la moneda '" & sCurrencyKey & "'."
          GoTo LblProcessNextRecord
        End If
        nSubsidiaryAccountId = 0
        If (Len(sSubsidiaryAccount) <> 0) Then
          If Not ValidateSubsidiaryAccount(sAppServer, nGralLedgerId, sSubsidiaryAccount) Then
            bTransactionHasErrors = True
            PrintLogFileLine oLogFile, sTransaction, True, "El auxiliar '" & sSubsidiaryAccount & "' tiene un formato que no reconozco o no puede ser empleado en el mayor."
            GoTo LblProcessNextRecord
          End If
          nSubsidiaryAccountId = GetSubsidiaryAccountId(sAppServer, nGralLedgerId, sSubsidiaryAccount)
          If (nSubsidiaryAccountId = 0) And bAutoGenerateSubsidiaryAccounts Then
            If CreateSubsidiaryAccount(sAppServer, nGralLedgerId, sSubsidiaryAccount) Then
              nSubsidiaryAccountId = GetSubsidiaryAccountId(sAppServer, nGralLedgerId, sSubsidiaryAccount)
              If (nSubsidiaryAccountId <> 0) Then
                PrintLogFileLine oLogFile, sTransaction, False, "Advertencia: Se agregó automáticamente el auxiliar '" & sSubsidiaryAccount & "' al sistema."
              Else
                bTransactionHasErrors = True
                PrintLogFileLine oLogFile, sTransaction, True, "No pude agregar automáticamente el auxiliar '" & sSubsidiaryAccount & "' al sistema."
                GoTo LblProcessNextRecord
              End If
            End If
          End If
        End If
        If Len(sResponsibilityArea) <> 0 Then
          nResponsibilityAreaId = GetResponsibilityAreaId(sAppServer, sResponsibilityArea)
        End If
        nAccountId = GetAccountId(sAppServer, nGralLedgerId, sAccountNumber, nStdAccountTypeId)
        If (nAccountId = -1) Then
          bTransactionHasErrors = True
          PrintLogFileLine oLogFile, sTransaction, True, "No pude agregar al mayor la cuenta '" & sAccountNumber & "' en forma automática."
          GoTo LblProcessNextRecord
        End If
        If Not bTransactionHasErrors Then
          PutPosting sAppServer, oPostingsRS, nTransactionId, nAccountId, Abs(nSubsidiaryAccountId), _
                     nSectorId, 0, nResponsibilityAreaId, sBudgetKey, sDisponibilityKey, sVerificationNumber, Null, sPostingConcept, _
                     sPostingType, nCurrencyId, Abs(nAmount), Abs(nBaseAmount), bProtectPostings
        End If
LblProcessNextRecord:
        If bTransactionHasErrors And (oPostingsRS.State = adStateOpen) Then
          DumpGemErrorTransaction sSourceAppServer, nEncTipoCont, dFechaVol, nNumVol
          oPostingsRS.Close
          nErrorCounter = nErrorCounter + 1
        End If
        .MoveNext
      Loop
      If (Len(sTransaction) <> 0) And (Not bTransactionHasErrors) Then
        nTransactionId = PutTransaction(sAppServer, oTransactionsRS, "No actualizada", nGralLedgerId, _
                                         nSourceId, nTransactionTypeId, nVoucherTypeId, sTransactionConcept, _
                                        dAfectationDate, dElaborationDate, nUserId, nUserToSend)
        DumpGemTransaction sAppServer, sSourceAppServer, nTransactionId, oTransactionsRS, oPostingsRS, _
                          nEncTipoCont, dFechaVol, nNumVol
        PrintLogFileLine oLogFile, sTransaction, False, "¡La póliza fue importada satisfactoriamente!"
        nOkCounter = nOkCounter + 1
      End If
    End With
    If nCounter <> 0 Then
      nSecondsPerVoucher = CLng(DateDiff("s", dImportDate, Now())) \ nCounter
      oLogFile.PrintNewLine
      oLogFile.PrintLine "Total de pólizas procesadas: " & nCounter & ". " & _
                         "Pólizas importadas: " & nOkCounter & ". " & _
                         "Pólizas con errores: " & nErrorCounter & "."
    Else
      oLogFile.PrintLine "!No se encontraron pólizas pendientes de importar!"
    End If
    oLogFile.Close_
    ImportTransactions = sFileName
    TransactionsLastImportationDate(sSourceAppServer) = dImportDate
    SetEstimatedImportTime sSourceAppServer, nSecondsPerVoucher
  Exit Function
ErrHandler:
  If oLogFile.IsOpen() Then
    oLogFile.Close_
    ImportTransactions = sFileName
  End If
  TransactionsLastImportationDate(sSourceAppServer) = dImportDate
  RaiseError ClassId, "ImportTransactions", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function TipoPoliza(nTipoPolizaGem As Long) As Long
  On Error GoTo ErrHandler
    Select Case nTipoPolizaGem
      Case 1    'GEM
        TipoPoliza = 40
      Case 2    'Fideicomisos
        TipoPoliza = 41
      Case 3    'Ventanilla doméstica
        TipoPoliza = 42
      Case 4    'Créditos de exempleados
        TipoPoliza = 43
      Case 5    'Créditos hipotecarios de exempleados
        TipoPoliza = 44
      Case 6    'Seguros
        TipoPoliza = 45
      Case 7    'Sector público, contratista, FOAEM y tramos
        TipoPoliza = 46
    End Select
  Exit Function
ErrHandler:
  RaiseError ClassId, "TipoPoliza", Err.Number
End Function

Private Sub DumpGemTransaction(sAppServer As String, sSourceAppServer As String, _
                               nTransactionId As Long, oTransactionsRS As Recordset, oPostingsRS As Recordset, _
                               nEncTipoCont As Long, dEncFechaVol As Date, nEncNumVol As Long)
  Dim oConnection As New Connection, oRecordset As New Recordset, sSQL As String
  Dim oInboxes As New MHInboxExplorer.CExplorer
  Dim nResult As Long
  '********************************************************************************************
  On Error GoTo ErrHandler
    oInboxes.CreateWorkItem sAppServer, _
                            "<b>Revisar póliza:</b> " & GetTransactionTypeName(sAppServer, oTransactionsRS!id_tipo_transaccion), _
                            "<b>" & GLName(sAppServer, oTransactionsRS!id_mayor) & "</b><br>" & _
                            oTransactionsRS!concepto_transaccion, nTransactionId, 694, 687, Now(), True, 2, "TASK_TYPE = 1024;"
    nResult = ExecuteCommand(sSourceAppServer, cnDelGEMTransaction, Array(nEncTipoCont, dEncFechaVol, nEncNumVol))
    If (nResult > 0) Then
      With oTransactionsRS
        .ActiveConnection = GetConstant(sConstantName:=sAppServer)
        .UpdateBatch
        oTransactionsRS.Close
        Set .ActiveConnection = Nothing
      End With
    End If
    With oPostingsRS
      If Not (.BOF And .EOF) Then
        .MoveFirst
        Do While Not .EOF
          !id_transaccion = nTransactionId
          .MoveNext
        Loop
      End If
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      oPostingsRS.Close
      Set .ActiveConnection = Nothing
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DumpGemTransaction", Err.Number
End Sub

Private Sub DumpGemErrorTransaction(sSourceAppServer As String, nEncTipoCont As Long, dEncFechaVol As Date, nEncNumVol As Long)
  On Error GoTo ErrHandler
    ExecuteCommand sSourceAppServer, cnUpdGEMErrTransaction, Array(nEncTipoCont, dEncFechaVol, nEncNumVol)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DumpGemErrorTransaction", Err.Number
End Sub

Private Function GetGemTransactions(sAppServer As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '*******************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM MC_ENCABEZADOS, MC_MOVIMIENTOS " & _
           "WHERE (ENC_TIPO_CONT = MCO_TIPO_CONT) AND " & _
           "(ENC_FECHA_VOL = MCO_FECHA_VOL) AND (ENC_NUM_VOL = MCO_NUM_VOL) " & _
           "ORDER BY ENC_TIPO_CONT, ENC_FECHA_VOL, ENC_NUM_VOL, MCO_FOLIO"
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    Set GetGemTransactions = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGemTransactions", Err.Number
End Function

Private Function TransactionsForExport(sAppServer As String, vGralLedgers As Variant, dFromDate As Date, _
                                       dToDate As Date, Optional bUseElaborationDates As Boolean = False, _
                                       Optional sSQLWhere As String, _
                                       Optional sSQLOrderBy As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sSQLStdAccountType As String, sSQLDates As String
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim sTempSQLWhere, sTempOrderBy As String
  '********************************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    
    If IsNull(vGralLedgers) Then
      sTempSQLWhere = ""
    ElseIf IsArray(vGralLedgers) Then
      sTempSQLWhere = "(COF_Mayor.id_mayor IN (" & Join(vGralLedgers, ",") & ") ) AND "
    ElseIf IsNumeric(vGralLedgers) Then
      sTempSQLWhere = "(COF_Mayor.id_mayor = " & vGralLedgers & ") AND "
    End If
    If bUseElaborationDates Then
      sTempSQLWhere = sTempSQLWhere & _
                      "(to_date('" & Format$(dFromDate, "dd/mm/yyyy hh:nn:ss") & "', 'DD/MM/YYYY HH24:MI:SS') <= COF_transaccion.fecha_registro) AND " & _
                      "(COF_transaccion.fecha_registro <= to_date('" & Format$(dToDate, "dd/mm/yyyy hh:nn:ss") & "', 'DD/MM/YYYY HH24:MI:SS')) AND "
    ElseIf Not bUseElaborationDates Then
      sTempSQLWhere = sTempSQLWhere & "(('" & dFromDate & "' <= COF_transaccion.fecha_afectacion) AND " & _
                                      "(COF_transaccion.fecha_afectacion <= '" & dToDate & "')) AND "
    End If
    If Len(sSQLWhere) <> 0 Then
      sTempSQLWhere = sTempSQLWhere & "(" & sSQLWhere & ") AND "
    End If
    If Len(sSQLOrderBy) <> 0 Then
      sTempOrderBy = sSQLOrderBy
    ElseIf bUseElaborationDates Then
      sTempOrderBy = "COF_Transaccion.id_transaccion, COF_Transaccion.id_mayor, numero_transaccion, fecha_registro, fecha_afectacion, id_movimiento"
    Else
      sTempOrderBy = "COF_Transaccion.id_transaccion, COF_Transaccion.id_mayor, numero_transaccion, fecha_afectacion, fecha_registro, id_movimiento"
    End If
             
    '"Areas.participantKey area_mayor, " & _
    'MHParticipants Areas, AOObjectAttrs'
    '"(Areas.participantId = AOObjectAttrs.ValNumeric) AND (AOObjectAttrs.EntityAttrDefId = 1) AND " & _
    '"(COF_Mayor.id_mayor = AOObjectAttrs.objectId (+)) AND " &
    
    sSQL = "SELECT DISTINCT COF_Transaccion.*, COF_movimiento.*, COF_mayor.numero_mayor, " & _
           "(COF_Movimiento.Monto_Moneda_Base / COF_Movimiento.Monto) Tipo_Cambio, " & _
           "COF_cuenta_estandar.id_tipo_cuentas_std, " & _
           "COF_cuenta_estandar.numero_cuenta_estandar, " & _
           "COF_cuenta_auxiliar.numero_cuenta_auxiliar, " & _
           "COF_cuenta_auxiliar.nombre_cuenta_auxiliar, COF_sector.clave_sector, " & _
           "MHParticipants.participantKey Clave_Area_Responsabilidad, AO_Currencies.o_id_moneda, " & _
           "Users.participantKey elaborada_por, " & _
           "COF_Transaccion.id_fuente fuente_transaccion " & _
           "FROM COF_Transaccion, COF_Movimiento, COF_Mayor, COF_cuenta, COF_cuenta_estandar, COF_cuenta_auxiliar, " & _
           "COF_sector, AO_Currencies, MHParticipants, MHParticipants Users " & _
           "WHERE " & sTempSQLWhere & _
           "(COF_movimiento.id_transaccion = COF_transaccion.id_transaccion) AND " & _
           "(COF_transaccion.id_mayor = COF_Mayor.id_mayor) AND " & _
           "(COF_transaccion.id_elaborada_por = Users.participantId) AND " & _
           "(COF_movimiento.id_cuenta = COF_cuenta.id_cuenta) AND " & _
           "(COF_cuenta.id_cuenta_estandar = COF_cuenta_estandar.id_cuenta_estandar) AND " & _
           "(COF_movimiento.id_cuenta_auxiliar = COF_cuenta_auxiliar.id_cuenta_auxiliar (+)) AND " & _
           "(COF_movimiento.id_sector = COF_sector.id_sector (+)) AND " & _
           "(COF_movimiento.id_moneda = AO_Currencies.currency_id) AND " & _
           "(COF_movimiento.id_area_responsabilidad = MHParticipants.participantId (+) ) " & _
           "ORDER BY " & sTempOrderBy
    With oRecordset
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CursorLocation = adUseClient
      .Open sSQL, , adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set TransactionsForExport = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "TransactionsForExport", Err.Number
End Function

Private Function GetSubsidiaryAccountId(sAppServer As String, nGralLedgerId As Long, _
                                        sSubsidiaryAccount As String) As Long
  Dim oRecordset As New Recordset, sSQL As String, nStdAccountId As Long
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    Do
      If Left$(sSubsidiaryAccount, 1) = "0" Then
        sSubsidiaryAccount = Mid$(sSubsidiaryAccount, 2)
      Else
        Exit Do
      End If
    Loop
    sSQL = "SELECT id_cuenta_auxiliar FROM COF_Cuenta_Auxiliar, COF_Mayor_Auxiliar " & _
           "WHERE (COF_Cuenta_Auxiliar.id_mayor_auxiliar = COF_Mayor_Auxiliar.id_mayor_auxiliar) AND " & _
           "(COF_Mayor_Auxiliar.id_mayor = " & nGralLedgerId & ") AND " & _
           "(numero_cuenta_auxiliar = '" & sSubsidiaryAccount & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetSubsidiaryAccountId = !id_cuenta_auxiliar
      Else
        GetSubsidiaryAccountId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetSubsidiaryAccountId", Err.Number
End Function


Private Function GetAccountId(sAppServer As String, nGeneralLedgerId As Long, sStdAccountNumber As String, _
                              nStdAccountTypeId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String, nStdAccountId As Long, nAccountId As Long
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar FROM COF_Cuenta_Estandar " & _
           "WHERE (numero_cuenta_estandar = '" & sStdAccountNumber & "') AND " & _
           "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        nStdAccountId = !id_cuenta_estandar
      Else
        nStdAccountId = 0
      End If
      .Close
    End With
    If nStdAccountId = 0 Then
      GetAccountId = 0
      Exit Function
    End If
    sSQL = "SELECT id_cuenta FROM COF_cuenta " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(id_mayor = " & nGeneralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        nAccountId = !id_cuenta
      Else
        nAccountId = 0
      End If
      .Close
    End With
    If nAccountId = 0 Then
      nAccountId = AssignStdAccountToGralLedger(sAppServer, nGeneralLedgerId, nStdAccountId)
    End If
    GetAccountId = nAccountId
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountId", Err.Number
End Function

Private Function AssignStdAccountToGralLedger(sAppServer As String, nGralLedgerId As Long, _
                                              nStdAccountId As Long) As Long
  Dim oGralLedger As New AOGralLedger.CGralLedger, sRole As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    AssignStdAccountToGralLedger = oGralLedger.AssignStdAccount(sAppServer, nGralLedgerId, nStdAccountId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "AssignStdAccountToGralLedger", Err.Number
End Function

Private Function GetCurrencyId(sAppServer As String, sCurrencyKey As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT currency_id FROM AO_Currencies " & _
           "WHERE (o_id_moneda = '" & sCurrencyKey & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetCurrencyId = !currency_id
      Else
        GetCurrencyId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetCurrencyId", Err.Number
End Function

Private Function GetSectorId(sAppServer As String, sSector As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '**************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_sector FROM COF_Sector " & _
           "WHERE (clave_sector = '" & sSector & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetSectorId = !id_sector
      Else
        GetSectorId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetSectorId", Err.Number
End Function

Private Function UsesSection(sAppServer As String, nGralLedgerId As Long) As Boolean
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT Id_Mayor " & _
           "FROM COF_ELEMENTO_GRUPO_MAYOR " & _
           "WHERE (Id_Grupo_Mayor = 11 OR Id_Grupo_Mayor = 12) AND (Id_Mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      UsesSection = Not (.BOF And .EOF)
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "UsesSection", Err.Number
End Function

Private Function SectionNumber(sAppServer As String, nGralLedgerId As Long) As String
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  On Error GoTo ErrHandler
    sSQL = "SELECT Id_Grupo_Mayor " & _
           "FROM COF_ELEMENTO_GRUPO_MAYOR " & _
           "WHERE (Id_Grupo_Mayor = 11 OR Id_Grupo_Mayor = 12) AND (Id_Mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If (!Id_Grupo_Mayor = 11) Then
          SectionNumber = "1900"
        Else
          SectionNumber = "1936"
        End If
      Else
        SectionNumber = "0000"
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "SectionNumber", Err.Number
End Function

Private Function GralLedgerId(sAppServer As String, nStdAccountTypeId As Long, _
                              sOBksGralLedger As String, sOBksSection As String) As Long
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  On Error GoTo ErrHandler
    If nStdAccountTypeId = 1 Then
      If (Len(sOBksSection) = 0) Or (sOBksSection = "0000") Then
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (numero_mayor = '" & Right$(sOBksGralLedger, 2) & "') AND " & _
               "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
      Else
        GralLedgerId = -1   'No debe traer sección ya que es un archivo de cont. bancaria
        Exit Function
      End If
    ElseIf nStdAccountTypeId = 2 Then
      If (Len(sOBksSection) = 0) Or (sOBksSection = "0000") Then        'Fideicomiso que no maneja secciones
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (numero_mayor = '" & sOBksGralLedger & "') AND " & _
               "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
      Else                                  'Maneja sección => Mayor = 1900 o 1936
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (sub_numero_mayor = '" & sOBksGralLedger & "') AND " & _
               "(numero_mayor = '" & sOBksSection & "') AND " & _
               "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
      End If
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GralLedgerId = !id_mayor
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GralLedgerId", Err.Number
End Function

Private Function MapOrabanksSource(sAppServer As String, sOBksSource As String) As Long
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  '************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT AOObjectAttrs.ObjectId " & _
           "FROM AOObjectAttrs, MHParticipants " & _
           "WHERE (AOObjectAttrs.ValNumeric = MHParticipants.participantId) AND " & _
           "(MHParticipants.participantKey = '" & sOBksSource & "') AND " & _
           "(AOObjectAttrs.entityAttrDefId = 1)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        MapOrabanksSource = !objectId
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "MapOrabanksSource", Err.Number
End Function

Private Function ValidateSectorInAccount(sAppServer As String, nAccountId As Long, _
                                         nSectorId As Long) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_CUENTA.ID_CUENTA_ESTANDAR " & _
           "FROM COF_CUENTA, COF_MAPEO_SECTOR " & _
           "WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR) AND " & _
           "(COF_CUENTA.ID_CUENTA = " & nAccountId & ") AND (COF_MAPEO_SECTOR.ID_SECTOR = " & nSectorId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        ValidateSectorInAccount = True
      Else
        ValidateSectorInAccount = False
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateSectorInAccount", Err.Number
End Function

Private Function ValidateCurrencyInAccount(sAppServer As String, nAccountId As Long, _
                                           nCurrencyId As Long) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '***********************************************************************************
  On Error GoTo ErrHandler
     sSQL = "SELECT COF_CUENTA.ID_CUENTA_ESTANDAR " & _
            "FROM COF_CUENTA, COF_MAPEO_MONEDA " & _
            "WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_MAPEO_MONEDA.ID_CUENTA_ESTANDAR) AND " & _
            "(COF_CUENTA.ID_CUENTA = " & nAccountId & ") AND (COF_MAPEO_MONEDA.ID_MONEDA = " & nCurrencyId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        ValidateCurrencyInAccount = True
      Else
        ValidateCurrencyInAccount = False
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateCurrencyInAccount", Err.Number
End Function
                       
Private Function GetUserToSend(sAppServer As String, nGralLedgerId As Long, nNotUserId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************
  On Error GoTo ErrHandler
    GetUserToSend = 687
    Exit Function
    sSQL = "SELECT Participant_Id FROM MH_Participant_Objects " & _
           "WHERE (Object_Id = " & nGralLedgerId & ") AND (Entity_Id = 9) AND " & _
           "(Participant_Id <> " & nNotUserId & ") " & _
           "ORDER BY Participant_Id"
   GetUserToSend = 135
   Exit Function
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetUserToSend = !Participant_Id
      Else
        GetUserToSend = nNotUserId
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetUserToSend", Err.Number
End Function

Private Sub PrintLogFileHeader(oLogFile As CTextFile, sDate As String)
  With oLogFile
    .PrintLine String(120, "*")
    .PrintLine "Importación de pólizas del sistema GEM hacia el SICOFIN del día " & _
               Format$(sDate, "dd"" de ""mmmm"" del ""yyyy")
    .PrintLine String(120, "*")
  End With
End Sub

Private Sub PrintLogFileSubHeader(oLogFile As CTextFile, sDate As String)
  With oLogFile
    .PrintNewLine
    .PrintLine String(120, "-")
    .PrintLine "Resultado de la ejecución del proceso efectuado a las " & _
                Format$(sDate, "hh:nn:ss") & " horas."
    .PrintLine String(120, "-")
  End With
End Sub

Private Sub PrintLogFileLine(oLogFile As CTextFile, sTransaction As String, _
                             bIsError As Boolean, sMsg As String)
  Dim sTemp As String
  '**************************************************************************
  On Error GoTo ErrHandler
    With oLogFile
      sTemp = "Póliza: " & sTransaction & "   "
      If bIsError Then
        sTemp = sTemp & "¡Error!  " & sMsg
      Else
        sTemp = sTemp & String(9, " ") & sMsg
      End If
      .PrintLine sTemp
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintLogFileLine", Err.Number
End Sub


Private Function PutPosting(sAppServer As String, oRS As Recordset, nTransactionId As Long, nAccountId As Long, _
                           nSubsidiaryAccountId As Long, nSectorId As Long, nReferencePostingId As Long, _
                           nResponsibilityAreaId As Long, sBudgetKey As String, sDisponibilityKey As String, _
                           sVerificationNumber As String, dPostingDate As Variant, sPostingConcept As String, _
                           sPostingType As String, nCurrencyId As Long, _
                           nAmount As Double, nBaseAmount As Double, bProtected As Boolean) As Long
  On Error GoTo ErrHandler
    With oRS
      .AddNew
        !id_movimiento_tmp = NewRecordId(sAppServer, "SEC_ID_MOVIMIENTO_TMP")
        !id_transaccion = nTransactionId
        !id_cuenta = nAccountId
        !id_cuenta_auxiliar = nSubsidiaryAccountId
        !id_sector = nSectorId
        !id_movimiento_referencia = nReferencePostingId
        !id_area_responsabilidad = nResponsibilityAreaId
        !clave_presupuestal = sBudgetKey
        !clave_disponibilidad = sDisponibilityKey
        !numero_verificacion = sVerificationNumber
        !tipo_movimiento = sPostingType
        !fecha_movimiento = dPostingDate
        !concepto_movimiento = sPostingConcept
        !id_moneda = nCurrencyId
        !monto = nAmount
        !monto_moneda_base = nBaseAmount
        !protegido = IIf(bProtected, 1, 0)
    End With
  Exit Function
ErrHandler:
  PutPosting = 0
  RaiseError ClassId, "PutPosting", Err.Number
End Function

Private Sub PutExternalPosting(sAppServer As String, oRecordset As Recordset, nStdAccountTypeId As Long, _
                              dApplicationDate As Date, dElaborationDate As Date, sExtVoucherNumber As String, _
                              nPostingId As Long, sAccountNumber As String, sSector As String, _
                              sSubsidiaryAccount As String, sResponsibilityArea As String, _
                              sPostingConcept As String, nDisponibilityKey As Variant, nBudgetKey As Variant, _
                              nVerificationNumber As Variant, sPostingType As String, sCurrencyKey As String, _
                              nExchangeRate As Double, nAmount As Double)
  On Error GoTo ErrHandler
    With oRecordset
      .AddNew
        !BGM_AREA = sResponsibilityArea
        !BGM_NUM_VOL = sExtVoucherNumber
        !BGM_FECHA_VOL = dApplicationDate
        !BGM_FOLIO_VOL = nPostingId
        !BGM_MONEDA = sCurrencyKey
        !BGM_REG_CONTABLE = FormatAccountNumber(sAccountNumber)
        !BGM_SECTOR = IIf(Len(sSector) = 0, "00", sSector)
        !BGM_CVE_MOV = IIf(sPostingType = "D", 1, 2)
        !BGM_NUM_AUX = sSubsidiaryAccount
        !BGM_DISPONIB = nDisponibilityKey
        !BGM_T_CAMBIO = nExchangeRate
        !BGM_IMPORTE = nAmount
        If Not IsNull(nBudgetKey) Then
          nBudgetKey = Trim$(nBudgetKey)
          If nBudgetKey = "" Then
            nBudgetKey = Null
          End If
        Else
          nBudgetKey = Null
        End If
        !BGM_CONCEPTO = nBudgetKey
        !BGM_NO_REG = Null
        !BGM_STATUS_OP = Null
        !BGM_STATUS_FINAN = Null
        !BGM_FECHA_ACT = dElaborationDate
        !BGM_DESCRIP = sPostingConcept
        !BGM_MARCA = Null
        !BGM_VERIFICACION = nVerificationNumber
      End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutExternalPosting", Err.Number
End Sub

Private Sub PutExternalTransaction(sAppServer As String, oRecordset As Recordset, nStdAccountTypeId As Long, _
                                  dApplicationDate As Date, sExtVoucherNumber As String, sVoucherConcept As String, _
                                  sUser As String, nDebitSum As Double, nCreditSum As Double)
  On Error GoTo ErrHandler
    With oRecordset
      .AddNew
        !BGME_TIPO_CONT = nStdAccountTypeId
        !bgme_fecha_vol = dApplicationDate
        !bgme_num_vol = sExtVoucherNumber
        !BGME_DESCRIP = Left$(sVoucherConcept, 240)
        !BGME_USUARIO = sUser
        !bgme_tot_cargos = nDebitSum
        !BGME_TOT_ABONOS = nCreditSum
        !BGME_DISP_C = Null
        !BGME_DISP_A = Null
        !BGME_STATUS = Null
        !BGME_MARCA = Null
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutExternalTransaction", Err.Number
End Sub

Private Sub PutExternalTransaction2(sAppServer As String, oRecordset As Recordset, nTransactionId As Long, _
                                  nStdAccountTypeId As Long, dApplicationDate As Date, _
                                  dElaborationDate As Date, sConcept As String, nDebitSum As Double, _
                                  nCreditSum As Double, sResponsibilityArea As String, _
                                  sSectionNumber As String, sUser As String)
  On Error GoTo ErrHandler
    With oRecordset
      .AddNew
        !ENC_TIPO_CONT = nStdAccountTypeId
        !ENC_FECHA_VOL = dApplicationDate
        !ENC_NUM_VOL = nTransactionId
        !ENC_DESCRIP = sConcept
        !ENC_TOT_CARGOS = nDebitSum
        !ENC_TOT_ABONOS = nCreditSum
        !ENC_TIPO_POLIZA = Null
        !ENC_FIDEICOMISO = CLng(sSectionNumber)
        !ENC_FECHA_CAP = dElaborationDate
        !ENC_USUARIO = sUser
        !ENC_AREA_CAP = sResponsibilityArea
        !ENC_SISTEMA = 1
        !ENC_NO_OPERACION = 1
        !ENC_STATUS = 0
        !ENC_GRUPO_CTL = nTransactionId
        !ENC_MARCA = Null
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutExternalTransaction", Err.Number
End Sub

Private Function PutTransaction(sAppServer As String, oRS As Recordset, sTransactionNumber As String, nGralLedgerId As Long, nSourceId As Long, _
                               nTransactionTypeId As Long, nVoucherTypeId As Long, _
                               sTransactionConcept As String, dAfectationDate As Date, _
                               dElaborationDate As Date, nElaborationUserId As Long, _
                               nAuthorizationUserId As Long) As Long
  Dim nTransactionId As Long
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    With oRS
      nTransactionId = NewRecordId(sAppServer, "SEC_ID_TRANSACCION")
      .AddNew
        !id_transaccion = nTransactionId
        !numero_transaccion = sTransactionNumber
        !id_mayor = nGralLedgerId
        !ID_FUENTE = nSourceId
        !id_tipo_transaccion = nTransactionTypeId
        !ID_TIPO_POLIZA = nVoucherTypeId
        !concepto_transaccion = TrimAll(sTransactionConcept)
        !fecha_afectacion = dAfectationDate
        !fecha_registro = dElaborationDate
        !id_elaborada_por = nElaborationUserId
        !ID_AUTORIZADA_POR = nAuthorizationUserId
        !ESTA_ABIERTA = 1
    End With
    PutTransaction = nTransactionId
  Exit Function
ErrHandler:
  RaiseError ClassId, "PutTransaction", Err.Number
End Function

Private Sub SaveRecordset(sAppServer As String, oRecordset As Recordset)
  Dim oConnection As New Connection
  '********************************************************************
  On Error GoTo ErrHandler
    oConnection.Open GetConstant(sConstantName:=sAppServer)
    oConnection.BeginTrans
    With oRecordset
      Set .ActiveConnection = oConnection
      .UpdateBatch adAffectAll
      Set .ActiveConnection = Nothing
    End With
    oConnection.CommitTrans
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveRecordset", Err.Number
End Sub

Private Function ConstructPostingsRS(sAppServer As String) As Recordset
  Dim sSQL As String, oRecordset As New Recordset
  '************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Movimiento_Tmp WHERE id_movimiento_tmp = -1"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    Set ConstructPostingsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "ConstructPostingsRS", Err.Number
End Function

Private Function ConstructTransactionsRS(sAppServer As String) As Recordset
  Dim sSQL As String, oRecordset As New Recordset
  '****************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Transaccion WHERE (id_transaccion = -1)"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    Set ConstructTransactionsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "ConstructTransactionsRS", Err.Number
End Function

Private Function NewRecordId(sAppServer As String, sSequence As String) As Long
  Dim oRecordset As New Recordset
  '***************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      .Open "SELECT " & sSequence & ".NEXTVAL AS ID FROM DUAL", _
            GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      NewRecordId = !Id
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "NewRecordId", Err.Number
End Function

Private Function GetGralLedgerResponsibilityArea(sAppServer As String, nGralLedgerId As Long) As String
  Static oRecordset As Recordset, sSQL As String
  '****************************************************************************************************
  On Error GoTo ErrHandler
    If oRecordset Is Nothing Then
      sSQL = "SELECT participantKey, objectId FROM AOObjectAttrs, MHParticipants " & _
             "WHERE (AOObjectAttrs.ValNumeric = MHParticipants.participantId) AND " & _
             "(EntityAttrDefId = 1)"
      Set oRecordset = New Recordset
      With oRecordset
        .CursorLocation = adUseClient
        oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    End If
    With oRecordset
      .Filter = "(objectId = " & nGralLedgerId & ")"
      If Not (.BOF And .EOF) Then
        GetGralLedgerResponsibilityArea = !participantKey
      Else
        GetGralLedgerResponsibilityArea = ""
      End If
      .Filter = adFilterNone
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGralLedgerResponsibilityArea", Err.Number
End Function

Private Function GetUserKey(sAppServer As String, nUserId As Long) As String
  Dim sSQL As String, oRecordset As New Recordset
  '*************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT participantKey FROM MHParticipants WHERE (participantId = " & nUserId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetUserKey = !participantKey
      Else
        GetUserKey = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetUserKey", Err.Number
End Function

Private Function FormatAccountNumber(sAccountNumber As String) As String
  Dim sTemp As String
  '*********************************************************************
  On Error GoTo ErrHandler
    If InStr(sAccountNumber, "-") = 0 Then
      sTemp = sAccountNumber
      sTemp = Left$(sTemp, 4) & "-" & Mid$(sTemp, 5, 2) & "-" & Mid$(sTemp, 7, 2) & "-" & _
              Mid$(sTemp, 9, 2) & "-" & Mid$(sTemp, 11, 2) & "-" & Mid$(sTemp, 13, 2) & "-" & Mid$(sTemp, 15, 2)
      Do
        If Right$(sTemp, 3) = "-00" Then
          sTemp = Left$(sTemp, Len(sTemp) - 3)
        Else
          Exit Do
        End If
      Loop
    Else
      sTemp = sAccountNumber
      sTemp = Replace(sTemp, "-", "")
      sTemp = Left$(sTemp & String(16, "0"), 16)
    End If
    FormatAccountNumber = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "FormatAccountNumber", Err.Number
End Function


Private Function CreateSubsidiaryAccount(sAppServer As String, nGralLedgerId As Long, _
                                         sSubsidiaryAccount As String) As Boolean
  Dim oSubsidiaryLedger As New AOGralLedger.CSubsidiaryLedger
  Dim oRecordset As Recordset, nSubsidiaryLedgerId As Long
  '************************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = oSubsidiaryLedger.GetSubsidiaryAccountRS(sAppServer)
    nSubsidiaryLedgerId = GetDefaultSubsidiaryLedgerId(sAppServer, nGralLedgerId)
    If (nSubsidiaryLedgerId <> 0) Then
      With oRecordset
        !id_mayor_auxiliar = nSubsidiaryLedgerId
        !NUMERO_CUENTA_AUXILIAR = sSubsidiaryAccount
        !NOMBRE_CUENTA_AUXILIAR = "Sin nombre (Generado con el importador GEM)"
        !DESCRIPCION = ""
        !ELIMINADA = 0
      End With
      oSubsidiaryLedger.SaveSubsidiaryAccount sAppServer, oRecordset
      CreateSubsidiaryAccount = True
    Else
      CreateSubsidiaryAccount = False
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateSubsidiaryAccount", Err.Number
End Function

Private Function GetDefaultSubsidiaryLedgerId(sAppServer As String, nGralLedgerId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_mayor_auxiliar, id_tipo_mayor_auxiliar FROM COF_Mayor_Auxiliar " & _
           "WHERE (id_mayor = " & nGralLedgerId & ") AND (eliminado = 0)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        .Filter = "id_tipo_mayor_auxiliar = 18"
        If Not (.BOF And .EOF) Then
          GetDefaultSubsidiaryLedgerId = !id_mayor_auxiliar
        Else
          .Filter = adFilterNone
          .Filter = "id_tipo_mayor_auxiliar = 19"
          If Not (.BOF And .EOF) Then
            GetDefaultSubsidiaryLedgerId = !id_mayor_auxiliar
          Else
            .Filter = adFilterNone
            GetDefaultSubsidiaryLedgerId = !id_mayor_auxiliar
          End If
        End If
      Else
        GetDefaultSubsidiaryLedgerId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetDefaultSubsidiaryLedgerId", Err.Number
End Function

Private Function GetResponsibilityAreaId(sAppServer As String, sResponsibilityArea As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  On Error GoTo ErrHandler
    sSQL = "SELECT participantId FROM MHParticipants " & _
           "WHERE (participantType = 'O') AND (participantKey = '" & sResponsibilityArea & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetResponsibilityAreaId = !participantId
      Else
        GetResponsibilityAreaId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetResponsibilityAreaId", Err.Number
End Function

Private Function GetStdAccountId(sAppServer As String, nStdAccountTypeId As Long, _
                                 sAccountNumber As String, dDate As Date) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (id_tipo_cuentas_std = " & nStdAccountTypeId & ") AND " & _
           "(numero_cuenta_estandar = '" & sAccountNumber & "') AND " & _
           "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetStdAccountId = !id_cuenta_estandar
      Else
        GetStdAccountId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountId", Err.Number
End Function

Private Function GetStdAccountRole(sAppServer As String, nStdAccountId As Long, _
                                   nSectorId As Long, dDate As Date) As String
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************************
  On Error GoTo ErrHandler
    If (nSectorId = 0) Then
      sSQL = "SELECT rol_cuenta rol FROM COF_Cuenta_Estandar_Hist " & _
             "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
             "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    Else
      sSQL = "SELECT rol_sector rol FROM COF_Mapeo_Sector " & _
             "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
             "(id_sector = " & nSectorId & ") AND " & _
             "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        GetStdAccountRole = !rol
      Else
        GetStdAccountRole = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountRole", Err.Number
End Function

Private Function ValidateStdAccountCurrency(sAppServer As String, nStdAccountId As Long, _
                                            nCurrencyId As Long, dDate As Date) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '***************************************************************************************
  On Error GoTo ErrHandler
     sSQL = "SELECT id_cuenta_estandar FROM COF_Mapeo_Moneda " & _
            "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
            "(id_moneda = " & nCurrencyId & ") AND " & _
            "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        ValidateStdAccountCurrency = True
      Else
        ValidateStdAccountCurrency = False
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateStdAccountCurrency", Err.Number
End Function

Private Function ValidateStdAccountSector(sAppServer As String, nStdAccountId As Long, _
                                          nSectorId As Long, dDate As Date) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_sector FROM COF_Mapeo_Sector " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        If (nSectorId = 0) Then
          ValidateStdAccountSector = False
        Else
          .Filter = "id_sector = " & nSectorId
          If Not (.BOF And .EOF) Then
            ValidateStdAccountSector = True
          Else
            ValidateStdAccountSector = False
          End If
          .Filter = adFilterNone
        End If
      Else  'No hay registros
        If (nSectorId = 0) Then
          ValidateStdAccountSector = True
        Else
          ValidateStdAccountSector = False
        End If
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateStdAccountSector", Err.Number
End Function

Private Function ValidateSubsidiaryAccount(sAppServer As String, nGralLedgerId As Long, _
                                           sSubsidiaryAccount As String) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '***************************************************************************************
  On Error GoTo ErrHandler
     sSQL = "SELECT prefijo_cuentas_auxiliares FROM COF_Mayor " & _
            "WHERE (id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        If !prefijo_cuentas_auxiliares = Left$(sSubsidiaryAccount, Len(!prefijo_cuentas_auxiliares)) Then
          ValidateSubsidiaryAccount = True
        Else
          ValidateSubsidiaryAccount = False
        End If
      Else
        ValidateSubsidiaryAccount = False
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateSubsidiaryAccount", Err.Number
End Function

Private Function ValidateTransaction(sTransactionNumber As String, nLine As Long, nGralLedgerId As Long, _
                                     nIdSource As Long, sTransactionConcept As String, _
                                     dAfectationDate As Date, dElaborationDate As Date, _
                                     oLogFile As CTextFile) As Boolean
  On Error GoTo ErrHandler
    ValidateTransaction = True
    With oLogFile
      If nGralLedgerId = -1 Then
        PrintLogFileLine oLogFile, sTransactionNumber, True, "Las contabilidades bancarias no pueden tener número de sección."
        ValidateTransaction = False
      End If
      If nGralLedgerId = 0 Then
        PrintLogFileLine oLogFile, sTransactionNumber, True, "No tengo registrado el número de mayor o la sección."
        ValidateTransaction = False
      End If
      If nIdSource = 0 Then
        PrintLogFileLine oLogFile, sTransactionNumber, True, "No tengo registrada el área proporcionada."
        ValidateTransaction = False
      End If
      If Len(sTransactionConcept) = 0 Then
        PrintLogFileLine oLogFile, sTransactionNumber, True, "Necesito el concepto de la póliza."
        ValidateTransaction = False
      End If
      If dAfectationDate > dElaborationDate Then
        PrintLogFileLine oLogFile, sTransactionNumber, True, "La fecha de elaboración de la póliza (" & _
                                                             Format$(dElaborationDate, "dd/mmm/yyyy") & ") " & _
                                                             "no puede ser anterior a la fecha de afectación de la misma (" & _
                                                             Format$(dAfectationDate, "dd/mmm/yyyy") & ")."
        ValidateTransaction = False
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateTransaction", Err.Number
End Function

Private Function GLName(sAppServer As String, nGralLedgerId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT numero_mayor, nombre_mayor FROM COF_Mayor " & _
           "WHERE (id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        GLName = "(" & !numero_mayor & ") " & !nombre_mayor
      Else  'No hay registros
        GLName = "Mayor no reconocido"
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GLName", Err.Number
End Function

Private Property Get GetTransactionTypeName(sAppServer As String, nTransactionTypeId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '**************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT objectName FROM AOObjects " & _
           "WHERE (objectId = " & nTransactionTypeId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockOptimistic
      If Not (.BOF And .EOF) Then
        GetTransactionTypeName = !objectName
      Else  'No hay registros
        GetTransactionTypeName = "Origen desconocido"
      End If
      .Close
    End With
  Exit Property
ErrHandler:
  RaiseError ClassId, "GetTransactionTypeName", Err.Number
End Property
