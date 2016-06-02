VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                                 Sistema : Financial Accounting          *
'* Componente : Vouchers Text File Interface (SCFIVouchersTextFile) Parte   : Business services             *
'* Clase      : CServer                                             Patrón  : Stateless Class               *
'* Fecha      : 28/Febrero/2002                                     Versión : 2.0       Versión patrón: 1.0 *
'*                                                                                                          *
'* Descripción: Clase principal del componente "Vouchers Text File Interface".                              *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CServer"

Private Const cnMaxLineLength As Long = 441
Private mCodigo_Tipo_Transaccion As String
Private mTipo_Mvto_Interno       As String

Public Property Get UsersDirectoryPath() As String
  On Error GoTo ErrHandler
    UsersDirectoryPath = GetConstant(cnUsersDirectoryPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "UsersDirectoryPath", Err.Number
End Property

Public Property Get FTPServer() As String
  On Error GoTo ErrHandler
    FTPServer = GetConstant(cnFTPServer)
  Exit Property
ErrHandler:
  RaiseError ClassId, "FTPServer", Err.Number
End Property

Public Property Get FTPUsersDirectoryPath() As String
  On Error GoTo ErrHandler
    FTPUsersDirectoryPath = GetConstant(cnFTPUsersDirectoryPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "FTPUsersDirectoryPath", Err.Number
End Property

Public Function GetUserDirectory(sAppServer As String, nUserId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '**********************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT participantKey FROM MHParticipants " & _
           "WHERE participantId = " & nUserId
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetUserDirectory = FTPUsersDirectoryPath & !participantKey & "/"
      Else
        GetUserDirectory = FTPUsersDirectoryPath
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetUserDirectory", Err.Number
End Function

Private Function UserName(sAppServer As String, nUserId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '**********************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT participantKey FROM MHParticipants " & _
           "WHERE participantId = " & nUserId
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        UserName = !participantKey
      Else
        UserName = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "UserName", Err.Number
End Function

Public Function CboUserFiles(sAppServer As String, nUserId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sPath As String
  '***************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT user_name FROM MH_Users WHERE user_id = " & nUserId
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sPath = UsersDirectoryPath & "\" & !user_name
      End If
      .Close
    End With
    CboUserFiles = GetHTMLDirectory(sPath & "\")
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboUserFiles", Err.Number
End Function

Public Sub CreateUsersDirectories(sAppServer As String)
  Dim oRecordset As New Recordset, sSQL As String, sPath As String
  '***************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT user_name FROM MH_Users ORDER BY user_name"
    sPath = UsersDirectoryPath
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        MkDir sPath & "\" & !user_name
        .MoveNext
      Loop
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CreateUsersDirectories", Err.Number
End Sub

Public Property Get URLFilesPath() As String
  On Error GoTo ErrHandler
    URLFilesPath = GetConstant(cnURLFilesPath)
  Exit Property
ErrHandler:
  RaiseError ClassId, "URLFilesPath", Err.Number
End Property

Public Function Export(sAppServer As String, nStdAccountTypeId As Long, dTransactionsDate As Date, _
                       sTransFromHour As String, sTransToHour As String, _
                       Optional nVouchersCounter As Variant, Optional sLogFile As Variant) As String
  Dim oRecordset As Recordset, oFile As New CTextFile, oLogFile As New CTextFile
  Dim sLine As String, nCounter, j As Long, sDecAmmount As String, nPointPos As Long
  Dim sDivAmmount As String, sSubNumber As String, sOrabanksArea As String
  Dim nTransactionsCounter As Long, nCurrentTransactionId As Long, dCurrentAppDate As Date
  Dim nCurrentGralLedgerId As Long, sGralLedger As String, sSectionNumber As String
  Dim nAppDateVouchersCount As Long, nAppDatePostingsSum As Double, nTotalPostingsSum As Double
  Dim nGLVouchersCount As Long, nGLPostingsSum As Double, sConcept As String
  Dim sFromDate As Date, sToDate As Date
  '*************************************************************************************************
  On Error GoTo ErrHandler
    oFile.GenerateRandomFile
    oFile.Open_
    If Len(sTransFromHour) = 0 Then
      sTransFromHour = "00:00:00"
    End If
    If Len(sTransToHour) = 0 Then
      sTransToHour = "23:59:59"
    End If
    'sFromDate = Format$(dTransactionsDate & " " & sTransFromHour, "dd/mmm/yyyy hh:nn:ss")
    'sToDate = Format$(dTransactionsDate & " " & sTransToHour, "dd/mmm/yyyy hh:nn:ss")
    'Set oRecordset = GetPosted(sAppServer, nStdAccountTypeId, sFromDate, sToDate)
    Set oRecordset = GetPosted(sAppServer, nStdAccountTypeId, dTransactionsDate, dTransactionsDate)
    oLogFile.GenerateRandomFile
    oLogFile.Open_
    PrintExportLogFileHeader oLogFile, dTransactionsDate, nStdAccountTypeId
    dCurrentAppDate = #1/1/1900#
    With oRecordset
      Do While Not .EOF
        If (dCurrentAppDate <> !fecha_afectacion) Then
          If (dCurrentAppDate <> #1/1/1900#) Then
            PrintExportLogFileGLRow oLogFile, sGralLedger, sSectionNumber, nGLVouchersCount, nGLPostingsSum
            PrintExportLogFileAppDateFooter oLogFile, dCurrentAppDate, nAppDateVouchersCount, nAppDatePostingsSum
            dCurrentAppDate = !fecha_afectacion
            PrintExportLogFileAppDateHeader oLogFile, dCurrentAppDate
          Else
            dCurrentAppDate = !fecha_afectacion
            PrintExportLogFileAppDateHeader oLogFile, dCurrentAppDate
          End If
          nAppDateVouchersCount = 0
          nAppDatePostingsSum = 0
          nCurrentGralLedgerId = 0
        End If
        If (nCurrentGralLedgerId <> !id_mayor) Then
          If nCurrentGralLedgerId <> 0 Then
            PrintExportLogFileGLRow oLogFile, sGralLedger, sSectionNumber, nGLVouchersCount, nGLPostingsSum
          End If
          nCurrentGralLedgerId = !id_mayor
          nGLVouchersCount = 0
          nGLPostingsSum = 0
        End If
        If (nCurrentTransactionId <> !id_transaccion) Then
          If (nCurrentGralLedgerId <= 32) Then
            sGralLedger = Format$(!numero_mayor, String(4, "0"))
          ElseIf Not UsesSection(sAppServer, nCurrentGralLedgerId) Then
            sGralLedger = Format$(!numero_mayor, String(4, "0"))
          Else
            sGralLedger = SectionNumber(sAppServer, nCurrentGralLedgerId)
          End If
          If nCurrentGralLedgerId <= 32 Then
            sSectionNumber = String(4, "0")
          ElseIf (nCurrentGralLedgerId > 32) Then
            If UsesSection(sAppServer, nCurrentGralLedgerId) Then
              sSectionNumber = Format$(!numero_mayor, String(4, "0"))
            Else
              sSectionNumber = String(4, "0")
            End If
          End If
          sOrabanksArea = Format$(MapUserToOrabanksSource(sAppServer, !id_elaborada_por), String(6, "0"))
          nCurrentTransactionId = !id_transaccion
          nGLVouchersCount = nGLVouchersCount + 1
          nAppDateVouchersCount = nAppDateVouchersCount + 1
          nTransactionsCounter = nTransactionsCounter + 1
        End If
        sLine = sOrabanksArea & sGralLedger
        sLine = sLine & sGralLedger & _
                "MG" & _
                Format$(Month(!fecha_afectacion), "00") & _
                Format$(nTransactionsCounter, String(4, "0"))
        sLine = sLine & String(19, "0") & _
                Format$(!fecha_afectacion, "ddmmyyyy")
        sConcept = !concepto_transaccion
        sConcept = Trim$(Replace$(sConcept, vbCrLf, ""))
        sConcept = sConcept & Space(240)
        sConcept = Replace$(sConcept, " ", "-")
        sLine = sLine & Left$(sConcept, 240)
        sLine = sLine & Left$(!numero_cuenta_estandar, 4)
        If (Len(!numero_cuenta_estandar) > 4) Then
          sSubNumber = Mid(!numero_cuenta_estandar, 5)
          sSubNumber = Replace(sSubNumber, "-", "")
          For j = Len(sSubNumber) + 1 To 12
            sSubNumber = sSubNumber & "0"
          Next j
        Else
         sSubNumber = String(12, "0")
        End If
        sSubNumber = sSubNumber & IIf(IsNull(!clave_sector), "00", !clave_sector)
        sLine = sLine & sSubNumber
        
        sLine = sLine & sSectionNumber
        sLine = sLine & sGralLedger & "0"
        
        If (!tipo_movimiento = "H") Then
          sLine = sLine & Format$(Int(!monto), "-" & String(12, "0"))
        Else
          sLine = sLine & Format$(Int(!monto), String(13, "0"))
        End If
        nPointPos = InStr(!monto, ".")
        If nPointPos <> 0 Then
           nPointPos = nPointPos + 1
           sDecAmmount = Mid(!monto, nPointPos, (Len(!monto) - nPointPos + 1))
        Else
           sDecAmmount = "0"
        End If
        If Len(sDecAmmount) > 6 Then
          sDecAmmount = Mid$(sDecAmmount, 1, 6)
        Else
          For j = Len(sDecAmmount) + 1 To 6
            sDecAmmount = sDecAmmount & "0"
          Next j
        End If
        sLine = sLine & sDecAmmount
        sLine = sLine & Format$(!o_id_moneda, "00")
        sLine = sLine & Format$(Int(!monto_moneda_base / !monto), String(7, "0"))
        sDivAmmount = CStr(!monto_moneda_base / !monto)
        nPointPos = InStr(sDivAmmount, ".")
        If nPointPos <> 0 Then
          nPointPos = nPointPos + 1
          sDecAmmount = Mid(sDivAmmount, nPointPos, (Len(sDivAmmount) - nPointPos + 1))
        Else
          sDecAmmount = "0"
        End If
        If (Len(sDecAmmount) >= 6) Then
          sDecAmmount = Left$(sDecAmmount, 6) & "00"
        Else
          sDecAmmount = sDecAmmount & String(6 - Len(sDecAmmount), "0") & "00"
        End If
        sLine = sLine & sDecAmmount
        sLine = sLine & String(19, "0")
        If IsNull(!NUMERO_CUENTA_AUXILIAR) Or (Len(!NUMERO_CUENTA_AUXILIAR) = 0) Then
          sLine = sLine & String(20, "0")
        Else
          sLine = sLine & Format$(!NUMERO_CUENTA_AUXILIAR, String(20, "0"))
        End If
        sLine = sLine & String(4, "0") & String(8, "0")
        sLine = sLine & IIf(IsNull(!clave_disponibilidad) Or Len(!clave_disponibilidad), "0", !clave_disponibilidad)
        sLine = sLine & "000000100000000"
        If IsNull(!Clave_Area_Responsabilidad) Or Len(!Clave_Area_Responsabilidad) = 0 Then
          sLine = sLine & Format$(0, String(6, "0"))
        ElseIf IsNumeric(!Clave_Area_Responsabilidad) Then
            sLine = sLine & Format$(!Clave_Area_Responsabilidad, String(6, "0"))
        Else
            sLine = sLine & Format$(0, String(6, "0"))
        End If
        'sLine = sLine & Format$(, 0, !Clave_Area_Responsabilidad, String(6, "0"))
        'sLine = sLine & Format$(IIf(IsNull(!clave_presupuestal) Or Len(!clave_presupuestal) = 0, 0, !clave_presupuestal), String(4, "0"))
        sLine = sLine & String(4, "0")          'Antes número de verificación
        sLine = sLine & mCodigo_Tipo_Transaccion & mTipo_Mvto_Interno
        sLine = sLine & Format$(IIf(IsNull(!numero_verificacion) Or Len(!numero_verificacion) = 0, 0, !numero_verificacion), String(8, "0"))
        sLine = Replace$(sLine, " ", "")
        If !tipo_movimiento = "D" Then
          nGLPostingsSum = nGLPostingsSum + !monto
          nAppDatePostingsSum = nAppDatePostingsSum + !monto
          nTotalPostingsSum = nTotalPostingsSum + !monto
        End If
        oFile.PrintLine sLine
        .MoveNext
      Loop
      .Close
      If nAppDatePostingsSum <> 0 Then
        PrintExportLogFileGLRow oLogFile, sGralLedger, sSectionNumber, nGLVouchersCount, nGLPostingsSum
        PrintExportLogFileAppDateFooter oLogFile, dCurrentAppDate, nAppDateVouchersCount, nAppDatePostingsSum
        PrintExportLogFileTotalSum oLogFile, nTransactionsCounter, nTotalPostingsSum
      End If
    End With
    Export = oFile.FileName
    If Not IsMissing(nVouchersCounter) Then
      nVouchersCounter = nTransactionsCounter
    End If
    If Not IsMissing(sLogFile) Then
      sLogFile = oLogFile.FileName
    End If
    oLogFile.Close_
    oFile.Close_
  Exit Function
ErrHandler:
  RaiseError ClassId, "Export", Err.Number
End Function

Public Function Import(sAppServer As String, sFileName As String, nStdAccountTypeId As Long, _
                       dElaborationDate As Date, nVoucherTypeId As Long, _
                       nUserId As Long, nVouchersCounter As Variant, _
                       Optional bForwardToUsers As Boolean = False, _
                       Optional bProtectPostings As Boolean = False, _
                       Optional bAutoGenerateSubsidiaryAccounts As Boolean = False) As String
  Dim oTransactionsRS As Recordset, oPostingsRS As Recordset, oFile As New CTextFile, oLogFile As New CTextFile
  Dim sGralLedger As String, sObnksSection As String, sTransactionNumber As String
  Dim sTransactionConcept As String, sSource As String, dAfectationDate As Date
  Dim nIdSource As Long, nTransactionTypeId As Long, nStdAccountId As Long, sFilePath As String
  Dim sAccountNumber As String, sCurrencyKey As String, sSectorKey As String, sSubsidiaryAccount As String
  Dim nAmount As Double, nBaseAmount As Double, nCurrencyId As Long, sDisponibilityKey As String
  Dim sStdAccountNumber As String, sStdAccountRole As String, sSector As String, nTransactionId As Long
  Dim nAccountId As Long, nGralLedgerId As Long, nSectorId As Long, nSubsidiaryAccountId As Long
  Dim bFileHasErrors As Boolean, nCounter As Long, nLine As Long, nUserToSend As Long, sLine As String
  Dim nResponsibilityAreaId As Long, sResponsibilityArea As String, nResult As Long
  '************************************************************************************************************
  On Error GoTo ErrHandler
    sFilePath = UsersDirectoryPath & "\" & UserName(sAppServer, nUserId) & "\" & sFileName
    If oFile.Exists(sFilePath) Then
      oFile.Open_ , sFilePath, True
    Else
      RaiseError ClassId, "Import", TEnumErrors.ErrFileNotExist
      Exit Function
    End If
    oLogFile.GenerateRandomFile
    oLogFile.Open_
    PrintImportLogFileHeader oLogFile, sFileName
    bFileHasErrors = False
    Set oTransactionsRS = ConstructTransactionsRS(sAppServer)
    Set oPostingsRS = ConstructPostingsRS(sAppServer)
    Do While Not oFile.EOF
      nLine = nLine + 1
      sLine = oFile.GetNextLine()
      sLine = Trim$(sLine)
      If (Len(sLine) = 0) Then
        GoTo LblProcessNextLine
      ElseIf Len(sLine) <> cnMaxLineLength Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La línea tiene un formato que no reconozco.")
        GoTo LblProcessNextLine
      End If
      If (sTransactionNumber <> Mid$(sLine, 11, 12)) Then
        nCounter = nCounter + 1
        ExtractTransactionInfo sLine, sGralLedger, sObnksSection, sTransactionNumber, sTransactionConcept, sSource, dAfectationDate
        nGralLedgerId = GralLedgerId(sAppServer, nStdAccountTypeId, sGralLedger, sObnksSection)
        nIdSource = MapOrabanksSource(sAppServer, sSource)
        nTransactionTypeId = 20
        If bForwardToUsers Then
          nUserToSend = GetUserToSend(sAppServer, nGralLedgerId, nUserId)
        End If
        If Not ValidateTransaction(sTransactionNumber, nLine, nGralLedgerId, nIdSource, _
                                   sTransactionConcept, dAfectationDate, dElaborationDate, oLogFile) Then
          bFileHasErrors = True
          GoTo LblProcessNextLine
        End If
        If Not bFileHasErrors Then
          nTransactionId = PutTransaction(sAppServer, oTransactionsRS, "No actualizada", nGralLedgerId, nIdSource, _
                                          nTransactionTypeId, nVoucherTypeId, sTransactionConcept, _
                                          dAfectationDate, dElaborationDate, nUserId, nUserToSend)
        End If
      End If
      ExtractPostingInfo sLine, sAccountNumber, sSector, sCurrencyKey, sSubsidiaryAccount, sDisponibilityKey, _
                         sResponsibilityArea, nAmount, nBaseAmount
      If (nAmount = 0) Or (nBaseAmount = 0) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "El movimiento no puede tener un monto igual a cero.")
        GoTo LblProcessNextLine
      End If
      nStdAccountId = GetStdAccountId(sAppServer, nStdAccountTypeId, sAccountNumber, dAfectationDate)
      If (nStdAccountId = 0) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & "' no existe en el catálogo estándar " & _
                                             "(al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & ")")
        GoTo LblProcessNextLine
      End If
      If sSector <> "00" Then
        nSectorId = GetSectorId(sAppServer, sSector)
      Else
        nSectorId = 0
      End If
      If (nSectorId = -1) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "El sector '" & sSector & "' no ha sido dado de alta en el sistema.")
        GoTo LblProcessNextLine
      End If
      If Not ValidateStdAccountSector(sAppServer, nStdAccountId, nSectorId, dAfectationDate) Then
        bFileHasErrors = True
        If (nSectorId = 0) Then
          oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & _
                                               "' maneja sectores pero éste no se proporcionó.")
        Else
          oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & _
                                               "' no maneja el sector '" & sSector & "' al día " & _
                                               Format$(dAfectationDate, "dd/mmm/yyyy") & ".")
        End If
        GoTo LblProcessNextLine
      End If
      sStdAccountRole = GetStdAccountRole(sAppServer, nStdAccountId, nSectorId, dAfectationDate)
      If (sStdAccountRole = "S") Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & "' es sumaria " & _
                                             "al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & ".")
        GoTo LblProcessNextLine
      End If
      If (sStdAccountRole = "C") And (Len(sSubsidiaryAccount) = 0) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & "' (sector " & sSector & ") maneja auxiliar (al día " & _
                                             Format$(dAfectationDate, "dd/mmm/yyyy") & ") pero éste no se proporcionó.")
        GoTo LblProcessNextLine
      End If
      If (sStdAccountRole = "P") And (Len(sSubsidiaryAccount) <> 0) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & "' (sector " & sSector & ") no maneja auxiliares (al día " & _
                                             Format$(dAfectationDate, "dd/mmm/yyyy") & ") pero se proporcionó el auxiliar '" & _
                                             sSubsidiaryAccount & "'.")
        GoTo LblProcessNextLine
      End If
      nCurrencyId = GetCurrencyId(sAppServer, sCurrencyKey)
      If (nCurrencyId = 0) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "No reconozco la moneda '" & sCurrencyKey & "'.")
        GoTo LblProcessNextLine
      End If
      If Not ValidateStdAccountCurrency(sAppServer, nStdAccountId, nCurrencyId, dAfectationDate) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "La cuenta '" & sAccountNumber & "' no maneja la moneda '" & sCurrencyKey & "'" & _
                                             "(al día " & Format$(dAfectationDate, "dd/mmm/yyyy") & ").")
        GoTo LblProcessNextLine
      End If
      If (nAmount <> nBaseAmount) And (nCurrencyId = 1) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "El monto y el monto en moneda base deben ser iguales para la moneda '" & sCurrencyKey & "'.")
        GoTo LblProcessNextLine
      End If
      nSubsidiaryAccountId = 0
      If (Len(sSubsidiaryAccount) <> 0) Then
        If Not ValidateSubsidiaryAccount(sAppServer, nGralLedgerId, sSubsidiaryAccount) Then
          bFileHasErrors = True
          oLogFile.PrintLine GetLogLine(nLine, "El auxiliar '" & sSubsidiaryAccount & "' tiene un formato que no reconozco o no puede ser empleado en el mayor .")
          GoTo LblProcessNextLine
        End If
        nSubsidiaryAccountId = GetSubsidiaryAccountId(sAppServer, nGralLedgerId, sSubsidiaryAccount)
        If (nSubsidiaryAccountId = 0) And bAutoGenerateSubsidiaryAccounts Then
          If CreateSubsidiaryAccount(sAppServer, nGralLedgerId, sSubsidiaryAccount) Then
            nSubsidiaryAccountId = GetSubsidiaryAccountId(sAppServer, nGralLedgerId, sSubsidiaryAccount)
            If (nSubsidiaryAccountId <> 0) Then
              oLogFile.PrintLine GetLogLine(nLine, "Advertencia: Se agregó automáticamente el auxiliar '" & sSubsidiaryAccount & "' al sistema.")
            Else
              bFileHasErrors = True
              oLogFile.PrintLine GetLogLine(nLine, "No pude agregar automáticamente el auxiliar '" & sSubsidiaryAccount & "' al sistema.")
              GoTo LblProcessNextLine
            End If
          End If
        End If
      End If
      If Len(sResponsibilityArea) <> 0 Then
        nResponsibilityAreaId = GetResponsibilityAreaId(sAppServer, sResponsibilityArea)
      End If
      nAccountId = GetAccountId(sAppServer, nGralLedgerId, sAccountNumber, nStdAccountTypeId)
      If (nAccountId = -1) Then
        bFileHasErrors = True
        oLogFile.PrintLine GetLogLine(nLine, "No pude agregar al mayor la cuenta '" & sAccountNumber & "' en forma automática.")
        GoTo LblProcessNextLine
      End If
      If Not bFileHasErrors Then
        PutPosting sAppServer, oPostingsRS, nTransactionId, nAccountId, Abs(nSubsidiaryAccountId), _
                   nSectorId, 0, nResponsibilityAreaId, "000000", sDisponibilityKey, "000000", Null, "", _
                   IIf(nAmount > 0, "D", "H"), nCurrencyId, Abs(nAmount), Abs(nBaseAmount), bProtectPostings
      End If
LblProcessNextLine:
    Loop
    oFile.Close_
    If Not bFileHasErrors Then
      SaveRecordset sAppServer, oTransactionsRS
      SaveRecordset sAppServer, oPostingsRS
    Else
      nCounter = 0
    End If
    oLogFile.PrintNewLine
    oLogFile.PrintLine String(100, "*")
    nVouchersCounter = nCounter
    If nVouchersCounter > 1 Then
      oLogFile.PrintLine "Se importaron " & nVouchersCounter & " pólizas en forma satisfactoria."
    ElseIf nVouchersCounter = 1 Then
      oLogFile.PrintLine "Se importó una póliza en forma satisfactoria."
    ElseIf nVouchersCounter < 1 Then
      oLogFile.PrintLine "¡La importación no tuvo éxito. No se incorporó ninguna póliza al sistema.!"
    End If
    Import = oLogFile.FileName
    oLogFile.Close_
  Exit Function
ErrHandler:
  nVouchersCounter = 0
  If oFile.IsOpen() Then
    oFile.Close_
  End If
  If oLogFile.IsOpen() Then
    oLogFile.PrintNewLine
    oLogFile.PrintLine "¡Error grave en el proceso de importación!"
    oLogFile.PrintLine "Línea      : " & nLine
    oLogFile.PrintLine "Descripción: " & IIf(Len(Err.Description) = 0, "Seguramente se trata de un problema en el formato de la línea" & nLine & ".", Err.Description)
    oLogFile.PrintLine "Origen     : " & IIf(Len(Err.Source) = 0, "IOrabanks.CServer.Import", Err.Source)
    oLogFile.PrintLine "Número     : " & Err.Number
    Import = oLogFile.FileName
    oLogFile.Close_
  End If
End Function

Private Function GetSubsidiaryAccountId(sAppServer As String, nGralLedgerId As Long, _
                                        sSubsidiaryAccount As String) As Long
  Dim oRecordset As New Recordset, sSQL As String, nStdAccountId As Long
  '***********************************************************************************
  On Error GoTo ErrHandler
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

Private Function CreateSubsidiaryAccount(sAppServer As String, nGralLedgerId As Long, _
                                         sSubsidiaryAccount As String) As Boolean
  Dim oSubsidiaryLedger As New CSubsidiaryLedger, oRecordset As Recordset, nSubsidiaryLedgerId As Long
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = oSubsidiaryLedger.GetSubsidiaryAccountRS(sAppServer)
    nSubsidiaryLedgerId = GetDefaultSubsidiaryLedgerId(sAppServer, nGralLedgerId)
    If nSubsidiaryLedgerId <> 0 Then
      With oRecordset
        !id_mayor_auxiliar = nSubsidiaryLedgerId
        !NUMERO_CUENTA_AUXILIAR = sSubsidiaryAccount
        !NOMBRE_CUENTA_AUXILIAR = "Sin nombre (Generado con el importador)"
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
                                                                                                                                                                  
Private Sub Class_Initialize()
  mCodigo_Tipo_Transaccion = "002"
  mTipo_Mvto_Interno = "N"
End Sub

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
      If sOBksSection = "0000" Then
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (numero_mayor = '" & Right$(sOBksGralLedger, 2) & "')"
      Else
        GralLedgerId = -1   'No debe traer sección ya que es un archivo de cont. bancaria
        Exit Function
      End If
    ElseIf nStdAccountTypeId = 2 Then
      If sOBksSection = "0000" Then         'Fideicomiso que no maneja secciones
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (numero_mayor = '" & sOBksGralLedger & "')"
      Else                                  'Maneja sección => Mayor = 1900 o 1936
        sSQL = "SELECT id_mayor " & _
               "FROM COF_Mayor " & _
               "WHERE (sub_numero_mayor = '" & sOBksGralLedger & "') AND " & _
               "(numero_mayor = '" & sOBksSection & "')"
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

Private Function MapOrabanksSource(sAppServer As String, sOBksSource As String) As Long
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  '************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT participantId FROM MHParticipants " & _
           "WHERE (participantType = 'O') AND (description = '" & sOBksSource & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        MapOrabanksSource = !participantId
      Else
        MapOrabanksSource = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "MapOrabanksSource", Err.Number
End Function

Private Function MapUserToOrabanksSource(sAppServer As String, nUserId As Long) As String
  Dim oRecordset As New ADODB.Recordset, sSQL As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT MHParticipants.description " & _
           "FROM MHParticipantsRelations, MHParticipants " & _
           "WHERE (MHParticipants.participantId = MHParticipantsRelations.participantBId) AND " & _
           "(MHParticipantsRelations.relationType = 2) AND " & _
           "(MHParticipantsRelations.participantAId = " & nUserId & ")"
    With oRecordset
      '.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      MapUserToOrabanksSource = "114120"
      Exit Function
      If Not (.BOF And .EOF) Then
        If Not IsNull(!Description) And Len(!Description) <> 0 Then
          MapUserToOrabanksSource = !Description
        End If
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "MapUserToOrabanksSource", Err.Number
End Function

Private Function ValidateTransaction(sTransactionNumber As String, nLine As Long, nGralLedgerId As Long, _
                                     nIdSource As Long, sTransactionConcept As String, _
                                     dAfectationDate As Date, dElaborationDate As Date, _
                                     oLogFile As CTextFile) As Boolean
  On Error GoTo ErrHandler
    ValidateTransaction = True
    With oLogFile
      If nGralLedgerId = -1 Then
        .PrintLine GetLogLine(nLine, "Las contabilidades bancarias no pueden tener número de sección.")
        ValidateTransaction = False
      End If
      If nGralLedgerId = 0 Then
        .PrintLine GetLogLine(nLine, "No tengo registrado el número de mayor o la sección.")
        ValidateTransaction = False
      End If
      If nIdSource = 0 Then
        .PrintLine GetLogLine(nLine, "No tengo registrada el área proporcionada.")
        ValidateTransaction = False
      End If
      If Len(sTransactionConcept) = 0 Then
        .PrintLine GetLogLine(nLine, "Necesito el concepto de la póliza.")
        ValidateTransaction = False
      End If
      If dAfectationDate > dElaborationDate Then
        .PrintLine GetLogLine(nLine, "La fecha de elaboración de la póliza (" & _
                                     Format$(dElaborationDate, "dd/mmm/yyyy") & ") " & _
                                     "no puede ser anterior a la fecha de afectación de la misma (" & _
                                     Format$(dAfectationDate, "dd/mmm/yyyy") & ").")
        ValidateTransaction = False
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValidateTransaction", Err.Number
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
        
Private Function GetHTMLDirectory(sPath As String, Optional sSelectedItem As String = "") As String
  Dim FilesArray() As String
  Dim sTemp As String, i As Integer
  
  FilesArray = DirectoryContents(sPath)
  For i = 0 To UBound(FilesArray)
    If StrComp(FilesArray(i), sSelectedItem, vbTextCompare) <> 0 Then
      sTemp = sTemp & "<OPTION value=""" & FilesArray(i) & """>" & FilesArray(i) & vbCrLf
    Else
      sTemp = sTemp & "<OPTION value=""" & FilesArray(i) & """ selected>" & FilesArray(i) & vbCrLf
    End If
  Next i
  GetHTMLDirectory = sTemp
End Function

Public Function DirectoryContents(ByVal sPath As String) As String()
  Dim aFilesArray() As String
  Dim sFileName As String
  
  On Error GoTo ErrDirectoryContents
    sFileName = Dir(sPath, vbArchive)
    ReDim aFilesArray(0)
    Do While (Len(sFileName) <> 0)
      If (GetAttr(sPath & sFileName) And vbArchive) = vbArchive Then
        If Len(aFilesArray(0)) <> 0 Then
          ReDim Preserve aFilesArray(0 To UBound(aFilesArray) + 1)
        End If
        aFilesArray(UBound(aFilesArray)) = sFileName
      End If
      sFileName = Dir
    Loop
    If UBound(aFilesArray) > 1 Then
      QuickSort aFilesArray()
    End If
    DirectoryContents = aFilesArray()
  Exit Function
ErrDirectoryContents:
  RaiseError ClassId, "DirectoryContents", Err.Number
End Function

Private Function GetUserToSend(sAppServer As String, nGralLedgerId As Long, nNotUserId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT Participant_Id FROM MH_Participant_Objects " & _
           "WHERE (Object_Id = " & nGralLedgerId & ") AND (Entity_Id = 9) AND " & _
           "(Participant_Id <> " & nNotUserId & ") " & _
           "ORDER BY Participant_Id"
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

Private Function GetLogLine(nLine As Long, sLineText As String) As String
  GetLogLine = "Línea " & Format$(nLine, "00000") & ": " & sLineText
End Function

Private Sub PrintExportLogFileHeader(oLogFile As CTextFile, dExportDate As Date, nStdAccountTypeId As Long)
  With oLogFile
    .PrintLine String(100, "*")
    If nStdAccountTypeId = 1 Then
      .PrintLine "Resultados de la exportación de pólizas 'elaboradas' el día " & _
                  Format$(dExportDate, "dd/mmm/yyyy") & " (Contabilidad bancaria)"
    ElseIf nStdAccountTypeId = 2 Then
      .PrintLine "Resultados de la exportación de pólizas elaboradas el día " & _
                  Format$(dExportDate, "dd/mmm/yyyy") & " (Contabilidad fiduciaria)"
    End If
    .PrintLine String(100, "*")
  End With
End Sub
         
Private Sub PrintExportLogFileAppDateFooter(oLogFile As CTextFile, dTransactionsDate As Date, _
                                            nAppDateVouchersCount As Long, nAppDatePostingsSum As Double)
  With oLogFile
    .PrintLine Space(27) & String(100 - 27, "-")
    .PrintLine "Total día " & Format$(dTransactionsDate, "dd/mmm/yyyy") & ":     " & _
               "Número de pólizas: " & Right$(Space(6) & FormatNumber(nAppDateVouchersCount, 0, GroupDigits:=vbTrue), 6) & "  " & _
               "Total cargos/abonos: " & Right$(Space(20) & Format$(nAppDatePostingsSum, "Currency"), 20)
    .PrintLine String(100, "-")
    .PrintNewLine
  End With
End Sub

Private Sub PrintExportLogFileAppDateHeader(oLogFile As CTextFile, dTransactionsDate As Date)
  With oLogFile
    .PrintLine String(100, "-")
    .PrintLine "Exportación de pólizas con fecha de afectación:    " & Format$(dTransactionsDate, "dd/mmm/yyyy")
  End With
End Sub

Private Sub PrintExportLogFileGLRow(oLogFile As CTextFile, sGralLedgerNumber As String, sGralLedgerSection As String, _
                                    nGLVouchersCount As Long, nGLPostingsSum As Double)
  Dim sTemp As String
  With oLogFile
    sTemp = "Mayor: " & Right$(Space(4) & sGralLedgerNumber, 4)
    If sGralLedgerSection <> "0000" Then
      sTemp = sTemp & " Sección: " & Right$(Space(4) & sGralLedgerSection, 4) & "  "
    Else
      sTemp = sTemp & Space(14) & "  "
    End If
    sTemp = sTemp & "Número de pólizas: " & Right$(Space(6) & FormatNumber(nGLVouchersCount, 0, GroupDigits:=vbTrue), 6) & "  " & _
                    "Total cargos/abonos: " & Right$(Space(20) & Format$(nGLPostingsSum, "Currency"), 20)
    .PrintLine sTemp
  End With
End Sub

Private Sub PrintExportLogFileTotalSum(oLogFile As CTextFile, nTransactionsCounter As Long, _
                                       nTotalPostingsSum As Double)
  With oLogFile
    .PrintLine String(100, "*")
    If nTransactionsCounter <> 0 Then
      .PrintLine "Total archivo exportación: " & _
                 "Número de pólizas: " & Right$(Space(6) & FormatNumber(nTransactionsCounter, 0, GroupDigits:=vbTrue), 6) & "  " & _
                 "Total cargos/abonos: " & Right$(Space(20) & Format$(nTotalPostingsSum, "Currency"), 20)
    Else
      .PrintLine "No se encontraron pólizas con la fecha de elaboración proporcionada."
    End If
  End With
End Sub

Private Sub PrintImportLogFileHeader(oLogFile As CTextFile, sFileName As String)
  With oLogFile
    .PrintLine String(100, "*")
    .PrintLine "Resultados de la importación del archivo '" & sFileName & "'"
    .PrintLine String(100, "*")
    .PrintNewLine
  End With
End Sub

Public Function PutPosting(sAppServer As String, oRS As Recordset, nTransactionId As Long, nAccountId As Long, _
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

Public Function PutTransaction(sAppServer As String, oRS As Recordset, sTransactionNumber As String, nGralLedgerId As Long, nSourceId As Long, _
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
        !NUMERO_TRANSACCION = sTransactionNumber
        !id_mayor = nGralLedgerId
        !ID_FUENTE = nSourceId
        !ID_TIPO_TRANSACCION = nTransactionTypeId
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
    With oRecordset
      Set .ActiveConnection = oConnection
      .UpdateBatch
      Set .ActiveConnection = Nothing
    End With
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

Private Function GetPosted(sAppServer As String, nStdAccountTypeId, dFromElaborationDate As Date, _
                           dToElaborationDate As Date) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sSQLStdAccountType As String, sSQLDates As String
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '*************************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    sSQL = "SELECT DISTINCT COF_Transaccion.*, COF_movimiento.*, COF_mayor.numero_mayor, (COF_Movimiento.Monto_Moneda_Base / COF_Movimiento.Monto) Tipo_Cambio, " & _
           "COF_cuenta_estandar.numero_cuenta_estandar, " & _
           "COF_cuenta_auxiliar.numero_cuenta_auxiliar, " & _
           "COF_cuenta_auxiliar.nombre_cuenta_auxiliar, COF_sector.clave_sector, " & _
           "MHParticipants.participantKey Clave_Area_Responsabilidad, AO_Currencies.o_id_moneda " & _
           "FROM COF_Transaccion, COF_Movimiento, COF_Mayor, COF_Cuenta, COF_cuenta_estandar, COF_cuenta_auxiliar, COF_sector, AO_Currencies, MHParticipants " & _
           "WHERE (trunc(to_date(COF_transaccion.fecha_registro)) = '" & dToElaborationDate & "') AND " & _
           "(COF_movimiento.id_transaccion = COF_transaccion.id_transaccion) AND " & _
           "(COF_Mayor.id_tipo_cuentas_std = " & nStdAccountTypeId & ") AND " & _
           "(COF_transaccion.id_mayor = COF_mayor.id_mayor) AND " & _
           "(COF_movimiento.id_cuenta = COF_cuenta.id_cuenta) AND " & _
           "(COF_cuenta.id_cuenta_estandar = COF_cuenta_estandar.id_cuenta_estandar) AND " & _
           "(COF_movimiento.id_cuenta_auxiliar = COF_cuenta_auxiliar.id_cuenta_auxiliar (+)) AND " & _
           "(COF_movimiento.id_sector = COF_sector.id_sector (+) ) AND " & _
           "(COF_movimiento.id_moneda = AO_Currencies.currency_id) AND " & _
           "(COF_movimiento.id_area_responsabilidad = MHParticipants.participantId (+) ) " & _
           "ORDER BY fecha_afectacion, numero_mayor, numero_transaccion, id_movimiento"
    With oRecordset
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CursorLocation = adUseClient
      .Open sSQL, , adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetPosted = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetPostedRS", Err.Number
End Function

Private Sub ExtractPostingInfo(sLine As String, sAccountNumber As String, sSector As String, _
                               sCurrencyKey As String, sSubsidiaryAccount As String, _
                               sDisponibilityKey As String, sResponsibilityArea As String, _
                               nAmount As Double, nBaseAmount As Double)
  Dim nExchangeRate As Double
  '*******************************************************************************************
  On Error GoTo ErrHandler
    nAmount = CDbl(Mid$(sLine, 317, 13) & "." & Mid$(sLine, 330, 6))
    sCurrencyKey = Mid$(sLine, 336, 2)
    sAccountNumber = Mid$(sLine, 290, 4) & "-" & _
                     Mid$(sLine, 294, 2) & "-" & Mid$(sLine, 296, 2) & "-" & Mid$(sLine, 298, 2) & "-" & _
                     Mid$(sLine, 300, 2) & "-" & Mid$(sLine, 302, 2) & "-" & Mid$(sLine, 304, 2)
    Do
      If Right$(sAccountNumber, 3) = "-00" Then
        sAccountNumber = Left$(sAccountNumber, Len(sAccountNumber) - 3)
      Else
        Exit Do
      End If
    Loop
    sSector = Mid$(sLine, 306, 2)
    nExchangeRate = CDbl(Mid$(sLine, 338, 7) & "." & Mid$(sLine, 345, 8))
    If nExchangeRate <> 0 Then
      nBaseAmount = nAmount * nExchangeRate
    ElseIf nExchangeRate = 0 Then
      nBaseAmount = nAmount
    End If
    sSubsidiaryAccount = Mid$(sLine, 372, 20)
    If sSubsidiaryAccount = String(20, "0") Then
      sSubsidiaryAccount = ""
    Else
      Do
        If Left$(sSubsidiaryAccount, 1) = "0" Then
          sSubsidiaryAccount = Mid$(sSubsidiaryAccount, 2)
        Else
          Exit Do
        End If
      Loop
    End If
    sSubsidiaryAccount = Trim$(sSubsidiaryAccount)
    sDisponibilityKey = Mid$(sLine, 404, 1)
    sResponsibilityArea = Mid$(sLine, 420, 6)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ExtractPostingInfo", Err.Number
End Sub

Private Sub ExtractTransactionInfo(sLine As String, sGralLedger As String, sObnksSection As String, _
                                   sTransactionNumber As String, sTransactionConcept As String, _
                                   sSource As String, dAfectationDate As Date)
  On Error GoTo ErrHandler
    sGralLedger = Mid$(sLine, 7, 4)
    sObnksSection = Mid$(sLine, 308, 4)
    sTransactionNumber = Mid$(sLine, 11, 12)
    sTransactionConcept = Trim$(Replace(Mid$(sLine, 50, 240), "-", " "))
    sSource = Mid$(sLine, 1, 6)
    dAfectationDate = CDate(Mid$(sLine, 42, 2) & "/" & Mid$(sLine, 44, 2) & "/" & Mid$(sLine, 46, 4))
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ExtractTransactionInfo", Err.Number
End Sub

Private Function GetCurrencyId(sAppServer As String, sCurrencyKey As String) As Long
  Static oRecordset As Recordset, sSQL As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    If oRecordset Is Nothing Then
      sSQL = "SELECT currency_id, o_id_moneda FROM AO_Currencies"
      Set oRecordset = New Recordset
      oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
    End If
    With oRecordset
      .Filter = "o_id_moneda = '" & sCurrencyKey & "'"
      If Not (.BOF And .EOF) Then
        GetCurrencyId = !currency_id
      Else
        GetCurrencyId = 0
      End If
      .Filter = adFilterNone
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetCurrencyId", Err.Number
End Function

Private Function GetSectorId(sAppServer As String, sSector As String) As Long
  Static oRecordset As Recordset, sSQL As String
  '**************************************************************************
  On Error GoTo ErrHandler
    If oRecordset Is Nothing Then
      sSQL = "SELECT id_sector, clave_sector FROM COF_Sector"
      Set oRecordset = New Recordset
      oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
    End If
    With oRecordset
      .Filter = "clave_sector = '" & sSector & "'"
      If Not (.BOF And .EOF) Then
        GetSectorId = !id_sector
      Else
        GetSectorId = -1
      End If
      .Filter = adFilterNone
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetSectorId", Err.Number
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

