VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Engine (EFARulesEngine)                   Parte   : Business services                 *
'* Clase      : CTransactions                                   Patrón  : Stateless class                   *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Genera una o más pólizas siguiendo los lineamientos de una regla contable.                  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CTransactions"

Private Const cStdAccountEntityId = 1
Private Const cTransactionTypeId = 23
Private Const cGralLedgerEntity = 9
Private Const cGralLedgerGroupEntity = 16

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function CancelTransaction(sAppServer As String, nUserId As Long, nAuthorizationUserId As Long, _
                                  nGralLedgerId As Long, sVoucherToCancelNumber As String, _
                                  sConcept As String, nSourceId As Long, nVoucherTypeId As Long, _
                                  dApplicationDate As Date, dElaborationDate As Date) As Long
  Dim oTransactionRS As Recordset, oPostingsRS As Recordset
  Dim oVoucher As New AOGLVoucher.CServer, oTransactionPostingsRS As Recordset
  Dim nSourceTransactionId As Long, nTransactionId As Long, sPostingType As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    nSourceTransactionId = oVoucher.GetTransactionId(sAppServer, nGralLedgerId, sVoucherToCancelNumber)
    If (nSourceTransactionId <> 0) Then
      Set oTransactionPostingsRS = oVoucher.GetOKPostingsRS(sAppServer, nSourceTransactionId)
      Set oVoucher = Nothing
    Else
      Set oVoucher = Nothing
      Exit Function
    End If
    Set oTransactionRS = GetTransactionsRS(sAppServer)
    Set oPostingsRS = GetPostingsRS(sAppServer)
    nTransactionId = PutTransaction(sAppServer, oTransactionRS, nGralLedgerId, nSourceId, _
                                    cTransactionTypeId, nVoucherTypeId, sConcept, _
                                    dApplicationDate, dElaborationDate, nUserId, nAuthorizationUserId)
    With oTransactionPostingsRS
      Do While Not .EOF
        sPostingType = IIf(!tipo_movimiento = "D", "H", "D")
        PutPosting sAppServer, oPostingsRS, nTransactionId, !id_cuenta, !id_cuenta_auxiliar, _
                   !id_sector, !id_movimiento_referencia, !id_area_responsabilidad, _
                   !clave_presupuestal, !clave_disponibilidad, !numero_verificacion, _
                   !fecha_movimiento, !concepto_movimiento, sPostingType, !id_moneda, _
                   !monto, !monto_moneda_base, False
        .MoveNext
      Loop
    End With
    SaveRecordset sAppServer, oTransactionRS
    SaveRecordset sAppServer, oPostingsRS
    CancelTransaction = nTransactionId
  Exit Function
ErrHandler:
  RaiseError ClassId, "CancelTransaction", Err.Number
End Function

Public Function Execute(sAppServer As String, nOperationId As Long, nUserId As Long, nSourceId As Long, _
                        sTransactionConcept As String, dApplicationDate As Date, _
                        Optional dElaborationDate As Date, _
                        Optional nGralLedgerAId As Long, Optional nGralLedgerBId As Long, _
                        Optional nAccountId As Long, Optional sPostingType As String, Optional nCurrencyId As Long, _
                        Optional nAmount As Variant, Optional nBaseAmount As Variant, _
                        Optional nCheckerUserId As Long = 0) As Long
  Dim oTransactionRS As Recordset, oPostingsRS As Recordset, nTransactionId As Long
  Dim oRuleDefRS As Recordset, nMethodId As Long, nRuleId As Long, nStdAccountTypeId As Long
  '******************************************************************************************************
  On Error GoTo ErrHandler
    nStdAccountTypeId = GetGLStdAccountTypeId(sAppServer, CLng(nGralLedgerAId))
    Set oRuleDefRS = GetRuleDefinitionRS(sAppServer, nOperationId, nStdAccountTypeId, dApplicationDate)
    With oRuleDefRS
      If Not (.BOF And .EOF) Then
        nMethodId = !id_metodo
        nRuleId = !id_regla_contable
        If nGralLedgerAId = 0 Then
          If !id_mayor_destino <> 0 Then
            nGralLedgerAId = !id_mayor_destino
          Else
            Err.Raise TEnumErrors.ErrRuleNotDefinedForThisGL
          End If
        End If
      Else
        Err.Raise TEnumErrors.ErrRuleNotFound
      End If
      .Close
    End With
    Set oRuleDefRS = Nothing
    Set oTransactionRS = GetTransactionsRS(sAppServer)
    Set oPostingsRS = GetPostingsRS(sAppServer)
    nTransactionId = PutTransaction(sAppServer, oTransactionRS, nGralLedgerAId, nSourceId, _
                                    cTransactionTypeId, nOperationId, sTransactionConcept, _
                                    dApplicationDate, dElaborationDate, nUserId, nCheckerUserId)
    Select Case nMethodId
      Case 66
        CancelaCuentasResultados sAppServer, nRuleId, oPostingsRS, nTransactionId, _
                                 nGralLedgerAId, dApplicationDate
      Case 67
        dElaborationDate = GetLastElaborationDate(sAppServer, nGralLedgerAId, nOperationId, dApplicationDate)
        ConcentracionFiduciaria sAppServer, nRuleId, oPostingsRS, nTransactionId, nGralLedgerAId, _
                                dApplicationDate, dElaborationDate
      Case 68
        TraspasoRemanentes sAppServer, nRuleId, oPostingsRS, nTransactionId, _
                           nGralLedgerAId, dApplicationDate
'      Case 69
        'CompraventaDivisas sAppServer, nUserId, nGralLedgerAId, nSourceId, sTransactionConcept, _
                           dApplicationDate, dElaborationDate, nCurrencyId, nAmount, nBaseAmount
'      Case 70
        'CompraventaDivisas sAppServer, nUserId, nGralLedgerAId, nSourceId, sTransactionConcept, _
                           dApplicationDate, dElaborationDate, nCurrencyId, nAmount, nBaseAmount
                                   
'      Case 72
        'MatrizDelegaciones sAppServer, nUserId, nGralLedgerAId, nGralLedger2Id, _
                           nSourceId, sTransactionConcept, nAccountId, _
                           sPostingType, nCurrencyId, nAmount, nBaseAmount
      Case Else
        Err.Raise TEnumErrors.ErrMethodWithoutImplementation
    End Select
    SaveRecordset sAppServer, oTransactionRS
    SaveRecordset sAppServer, oPostingsRS
    Execute = nTransactionId
  Exit Function
ErrHandler:
  Execute = -1
  If Err.Number = TEnumErrors.ErrMethodWithoutImplementation Then
    RaiseError ClassId, "Execute", Err.Number, nMethodId
  Else
    RaiseError ClassId, "Execute", Err.Number
  End If
End Function

Public Function TransferBalances(sAppServer As String, nUserId As Long, nAuthorizationUserId As Long, _
                                 vGralLedgers As Variant, sConcept As String, nSourceId As Long, _
                                 nVoucherTypeId As Long, dApplicationDate As Date, dElaborationDate As Date, _
                                 sFromAccount As String, sFromSubsAccount As String, _
                                 nFromSectorId As Long, nFromCurrencyId As Long, _
                                 sToAccount As String, sToSubsAccount As String, _
                                 nToSectorId As Long, nToCurrencyId As Long, nExchangeRate As Variant) As Long
  Dim oTransactionRS As Recordset, oPostingsRS As Recordset
  Dim oBalances As New EFABalanceEngine.CEngine, oBalancesRS As Recordset
  Dim nTransactionId As Long, sPostingType As String, nAccountId As Long
  Dim nCurrentGL As Long, nCurrentStdAccount As Long, sFilter As String
  Dim bTargetAccountUseSectors As Boolean, sTargetAccountRole As String
  Dim nTargetStdAccountId As Long, nTargetAccountId As Long, nTargetSectorId As Long
  Dim nTargetCurrencyId As Long, nTargetSubsAccountId As Long, nStdAccountTypeId As Long
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    Set oTransactionRS = GetTransactionsRS(sAppServer)
    Set oPostingsRS = GetPostingsRS(sAppServer)
    
    If IsArray(vGralLedgers) Then
      nStdAccountTypeId = GetGLStdAccountTypeId(sAppServer, CLng(vGralLedgers(LBound(vGralLedgers))))
    Else
      nStdAccountTypeId = GetGLStdAccountTypeId(sAppServer, CLng(vGralLedgers))
    End If
    If (InStr(1, sFromAccount, "*") > 0) Or (InStr(1, sFromAccount, "?") > 0) Then
      Set oBalancesRS = oBalances.BalancesRS(sAppServer, "####-##-##-##-##-##-##", vGralLedgers, _
                                             False, True, True, dApplicationDate, dApplicationDate)
    Else
      Set oBalancesRS = oBalances.BalancesRS(sAppServer, "####-##-##-##-##-##-##", vGralLedgers, _
                                             False, True, True, dApplicationDate, dApplicationDate, _
                                             sFromAccount, sFromAccount)
    End If
    
    With oBalancesRS
      sFilter = BuildBalanceFilter(sAppServer, sFromAccount, sFromSubsAccount, nFromSectorId, nFromCurrencyId)
      If Len(sFilter) <> 0 Then
        .Filter = sFilter
      End If
      Do While Not .EOF
        If nCurrentGL <> !id_mayor Then
          nCurrentGL = !id_mayor
          nTransactionId = PutTransaction(sAppServer, oTransactionRS, nCurrentGL, nSourceId, _
                                          cTransactionTypeId, nVoucherTypeId, sConcept, _
                                          dApplicationDate, dElaborationDate, nUserId, nAuthorizationUserId)
          nCurrentStdAccount = 0
        End If
        If (!naturaleza = "D") And (!saldo_actual > 0) Then
          sPostingType = "H"
        ElseIf (!naturaleza = "D") And (!saldo_actual < 0) Then
          sPostingType = "D"
        ElseIf (!naturaleza = "A") And (!saldo_actual > 0) Then
          sPostingType = "D"
        ElseIf (!naturaleza = "A") And (!saldo_actual < 0) Then
          sPostingType = "H"
        End If
        If (nCurrentStdAccount <> !id_cuenta_estandar) Then
          nCurrentStdAccount = !id_cuenta_estandar
          nAccountId = GetAccountId(sAppServer, nCurrentGL, nCurrentStdAccount)
        End If
        PutPosting sAppServer, oPostingsRS, nTransactionId, nAccountId, !id_cuenta_auxiliar, _
                  !id_sector, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
                  Abs(!saldo_actual), Abs(!saldo_actual), 0
        If Len(sToAccount) = 0 Then
          nTargetStdAccountId = GetStdAccountProps(sAppServer, nStdAccountTypeId, !numero_cuenta_estandar, _
                                                   dApplicationDate, bTargetAccountUseSectors)
          bTargetAccountUseSectors = (!id_sector <> 0)
          nTargetAccountId = nAccountId
        Else
          nTargetStdAccountId = GetStdAccountProps(sAppServer, nStdAccountTypeId, !numero_cuenta_estandar, _
                                                   dApplicationDate, bTargetAccountUseSectors)
          nTargetAccountId = GetAccountId(sAppServer, nCurrentGL, nTargetStdAccountId)
        End If
        If (nToSectorId = 0) And bTargetAccountUseSectors Then
          nTargetSectorId = !id_sector
        ElseIf (nToSectorId = 0) And (Not bTargetAccountUseSectors) Then
          nTargetSectorId = 0
        ElseIf (nToSectorId <> 0) Then
          nTargetSectorId = nToSectorId
        End If
        If (nToCurrencyId = 0) Then
          nTargetCurrencyId = !id_moneda
        ElseIf (nToCurrencyId <> 0) Then
          nTargetCurrencyId = nToCurrencyId
        End If
        sTargetAccountRole = GetStdAccountRole(sAppServer, nTargetStdAccountId, _
                                               dApplicationDate, nTargetSectorId)
        If (Len(sToSubsAccount) = 0) And (sTargetAccountRole = "C") Then
          nTargetSubsAccountId = !id_cuenta_auxiliar
        ElseIf (Len(sToSubsAccount) = 0) And (sTargetAccountRole = "P") Then
          nTargetSubsAccountId = 0
        ElseIf (Len(sToSubsAccount) <> 0) Then
          nTargetSubsAccountId = GetSubsidiaryAccountId(sAppServer, nCurrentGL, sToSubsAccount)
        End If
        sPostingType = IIf(sPostingType = "D", "H", "D")
        PutGroupedPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, nTargetSubsAccountId, _
                          nTargetCurrencyId, nTargetSectorId, sPostingType, _
                          Abs(!saldo_actual), Abs(!saldo_actual) * nExchangeRate
        .MoveNext
      Loop
    End With
    TransferBalances = 0
    If (nTransactionId <> 0) Then
      SaveRecordset sAppServer, oTransactionRS
      SaveRecordset sAppServer, oPostingsRS
      If Not IsArray(vGralLedgers) Then
        TransferBalances = nTransactionId
      End If
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "TransferBalances", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Sub CancelaCuentasResultados(sAppServer As String, nRuleId As Long, oPostingsRS As Recordset, _
                                     nTransactionId As Long, nGralLedgerId As Long, dCancelationDate As Date)
  Dim oBalances As New EFABalanceEngine.CEngine, oRulesRS As New Recordset, oBalancesRS As Recordset
  Dim nDebitSum As Double, nCreditSum As Double, sPostingType As String
  Dim nSourceAccountId As Long, nTargetAccountId As Long
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    Set oRulesRS = GetRulesRS(sAppServer, nRuleId, dCancelationDate)
    Do While Not oRulesRS.EOF
      Set oBalancesRS = oBalances.BalancesRS(sAppServer, "####-##-##-##-##-##-##", nGralLedgerId, _
                                             False, True, True, dCancelationDate, dCancelationDate, _
                                             oRulesRS!cuenta_origen_inicial, _
                                             oRulesRS!cuenta_origen_final)
      With oBalancesRS
        If (Not IsNull(oRulesRS!filtro_cuentas_origen)) And (Len(oRulesRS!filtro_cuentas_origen) <> 0) Then
          .Filter = CStr(oRulesRS!filtro_cuentas_origen)
        End If
        Do While Not .EOF
          If (!id_cuenta_estandar <> oRulesRS!id_cuenta_destino) Then
            If (!saldo_actual <> 0) Then
              If (!naturaleza = "D") And (!saldo_actual > 0) Then
                sPostingType = "H"
                nCreditSum = nCreditSum + Abs(!saldo_actual)
              ElseIf (!naturaleza = "D") And (!saldo_actual < 0) Then
                sPostingType = "D"
                nDebitSum = nDebitSum + Abs(!saldo_actual)
              ElseIf (!naturaleza = "A") And (!saldo_actual > 0) Then
                sPostingType = "D"
                nDebitSum = nDebitSum + Abs(!saldo_actual)
              ElseIf (!naturaleza = "A") And (!saldo_actual < 0) Then
                sPostingType = "H"
                nCreditSum = nCreditSum + Abs(!saldo_actual)
              End If
              nSourceAccountId = GetAccountId(sAppServer, nGralLedgerId, !id_cuenta_estandar)
              PutPosting sAppServer, oPostingsRS, nTransactionId, nSourceAccountId, !id_cuenta_auxiliar, _
                        !id_sector, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
                        Abs(!saldo_actual), Abs(!saldo_actual), 0
            End If
          End If
          .MoveNext
        Loop
        .Close
      End With
      nTargetAccountId = GetAccountId(sAppServer, nGralLedgerId, oRulesRS!id_cuenta_destino)
      If nCreditSum <> 0 Then
        PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, 0, _
                   0, 0, 0, "", "", "", Null, "", "D", 1, nCreditSum, nCreditSum, 0
      End If
      If nDebitSum <> 0 Then
        PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, 0, _
                   0, 0, 0, "", "", "", Null, "", "H", 1, nDebitSum, nDebitSum, 0
      End If
      nDebitSum = 0
      nCreditSum = 0
      oRulesRS.MoveNext
    Loop
    oRulesRS.Close
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CancelaCuentasResultados", Err.Number
End Sub

Private Sub ConcentracionFiduciaria(sAppServer As String, nRuleId As Long, oPostingsRS As Recordset, _
                                    nTransactionId As Long, nTargetGralLedgerId As Long, dApplicationDate As Date, _
                                    dElaborationDate As Date)
  Dim oRulesRS As Recordset, oDayPostingsRS As Recordset
  Dim nDebitSum As Double, nCreditSum As Double, aGralLedgers As Variant
  Dim sFilter As String, nLastEntity As Long, nLastGroup As Long, nTargetAccountId As Long, nSubsAccountId As Long
  '**************************************************************************************************
  On Error GoTo ErrHandler
    Set oRulesRS = GetRulesRS(sAppServer, nRuleId, dApplicationDate)
    Do While Not oRulesRS.EOF
      With oRulesRS
        If (nLastEntity <> !ID_ENTIDAD_AGRUPADOR) Or (nLastGroup <> !id_agrupador_origen) Then
          nLastEntity = !ID_ENTIDAD_AGRUPADOR
          nLastGroup = !id_agrupador_origen
          If (!ID_ENTIDAD_AGRUPADOR = cGralLedgerGroupEntity) Then
            aGralLedgers = GetGLGroupArray(sAppServer, !id_agrupador_origen)
          ElseIf (!ID_ENTIDAD_AGRUPADOR = cGralLedgerEntity) Then
            aGralLedgers = !id_agrupador_origen
          End If
          Set oDayPostingsRS = GetPostings(sAppServer, aGralLedgers, dApplicationDate, dElaborationDate)
        End If
      End With
      With oDayPostingsRS
        sFilter = "(numero_cuenta_estandar >= '" & oRulesRS!cuenta_origen_inicial & "') AND (numero_cuenta_estandar <= '" & oRulesRS!cuenta_origen_final & "')"
        If (Not IsNull(oRulesRS!filtro_cuentas_origen)) And (Len(oRulesRS!filtro_cuentas_origen) <> 0) Then
          sFilter = sFilter & " AND " & oRulesRS!filtro_cuentas_origen
        End If
        .Filter = sFilter
        Do While Not .EOF
          If (Not IsNull(oRulesRS!auxiliar_destino)) And (Len(oRulesRS!auxiliar_destino) <> 0) Then
            nSubsAccountId = GetSubsidiaryAccountId(sAppServer, nTargetGralLedgerId, oRulesRS!auxiliar_destino)
          Else
            nSubsAccountId = 0
          End If
          If (!tipo_movimiento = "D") Then
            nDebitSum = nDebitSum + !monto
            nTargetAccountId = GetAccountId(sAppServer, nTargetGralLedgerId, oRulesRS!id_cuenta_destino)
            PutGroupedPosting sAppServer, oPostingsRS, nTransactionId, _
                              nTargetAccountId, nSubsAccountId, _
                              !id_moneda, oRulesRS!id_sector_destino, "D", !monto, !monto_moneda_base
          End If
          If (!tipo_movimiento = "H") Then
            nCreditSum = nCreditSum + !monto
            nTargetAccountId = GetAccountId(sAppServer, nTargetGralLedgerId, oRulesRS!id_cuenta_destino)
            PutGroupedPosting sAppServer, oPostingsRS, nTransactionId, _
                              nTargetAccountId, nSubsAccountId, _
                              !id_moneda, oRulesRS!id_sector_destino, "H", !monto, !monto_moneda_base
          End If
          .MoveNext
        Loop
        .Filter = adFilterNone
      End With
      nDebitSum = 0
      nCreditSum = 0
      oRulesRS.MoveNext
    Loop
    oRulesRS.Close
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ConcentracionFiduciaria", Err.Number
End Sub

Private Sub TraspasoRemanentes(sAppServer As String, nRuleId As Long, oPostingsRS As Recordset, _
                               nTransactionId As Long, nGralLedgerId As Long, dCancelationDate As Date)
  Dim oBalances As New EFABalanceEngine.CEngine, oRulesRS As New Recordset, oBalancesRS As Recordset
  Dim nDebitSum As Double, nCreditSum As Double, bOverBalanced As Boolean, nSubsAccountId As Long
  Dim sPostingType As String, nTargetAccountId As Long
  '****************************************************************************************************
  On Error GoTo ErrHandler
'    Set oRulesRS = GetRulesRS(sAppServer, nRuleId, dCancelationDate)
'    Do While Not oRulesRS.EOF
'      Set oBalancesRS = oBalances.LastLevels(sAppServer, nGralLedgerId, _
'                                             dCancelationDate, dCancelationDate, _
'                                             oRulesRS!cuenta_origen_inicial, _
'                                             oRulesRS!cuenta_origen_final)
'      With oBalancesRS
'        If (Not IsNull(oRulesRS!filtro_cuentas_origen)) And (Len(oRulesRS!filtro_cuentas_origen) <> 0) Then
'          .Filter = CStr(oRulesRS!filtro_cuentas_origen)
'        End If
'        Do While Not .EOF
'          If (!saldo_actual <> 0) Then
'            If (!naturaleza = "D") And (!saldo_actual > 0) Then
'              sPostingType = "H"
'              nCreditSum = nCreditSum + Abs(!saldo_actual)
'              bOverBalanced = False
'            ElseIf (!naturaleza = "D") And (!saldo_actual < 0) Then
'              sPostingType = "D"
'              nDebitSum = nDebitSum + Abs(!saldo_actual)
'              bOverBalanced = True
'            ElseIf (!naturaleza = "A") And (!saldo_actual > 0) Then
'              sPostingType = "D"
'              nDebitSum = nDebitSum + Abs(!saldo_actual)
'              bOverBalanced = False
'            ElseIf (!naturaleza = "A") And (!saldo_actual < 0) Then
'              sPostingType = "H"
'              nCreditSum = nCreditSum + Abs(!saldo_actual)
'              bOverBalanced = True
'            End If
'            PutPosting sAppServer, oPostingsRS, nTransactionId, !id_cuenta, !id_cuenta_auxiliar, _
'                      !id_sector, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
'                      Abs(!saldo_actual), Abs(!saldo_actual), False
'            If (Not IsNull(oRulesRS!auxiliar_destino)) And (Len(oRulesRS!auxiliar_destino) <> 0) Then
'              nSubsAccountId = GetSubsidiaryAccountId(sAppServer, nTargetGralLedgerId, oRulesRS!auxiliar_destino)
'            Else
'              nSubsAccountId = 0
'            End If
'            If (Not bOverBalanced) Then
'              sPostingType = IIf(sPostingType = "D", "H", "D")
'              nTargetAccountId = GetAccountId(sAppServer, nGralLedgerId, oRulesRS!id_cuenta_destino)
'              PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, oRulesRS!id_auxiliar_destino, _
'                         oRulesRS!id_sector_destino, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
'                         Abs(!saldo_actual), Abs(!saldo_actual), False
'            ElseIf bOverBalanced Then
'              sPostingType = IIf(sPostingType = "D", "H", "D")
'              nTargetAccountId = GetAccountId(sAppServer, nGralLedgerId, oRulesRS!id_cuenta_destino_sobregiro)
'              PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, oRulesRS!id_auxiliar_destino_sobregiro, _
'                         oRulesRS!id_sector_destino_sobregiro, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
'                         Abs(!saldo_actual), Abs(!saldo_actual), False
'            End If
'          End If
'          .MoveNext
'        Loop
'        .Close
'      End With
'      nDebitSum = 0
'      nCreditSum = 0
'      oRulesRS.MoveNext
'    Loop
'    oRulesRS.Close
  Exit Sub
ErrHandler:
  RaiseError ClassId, "TraspasoRemanentes", Err.Number
End Sub

'Private Sub CompraventaDivisas(sAppServer As String, nGralLedgerId As Long, _
'                              nUserId As Long, nSourceId As Long, sTransactionConcept As String, _
'                              dApplicationDate As Date, dElaborationDate As Date, _
'                              nCurrencyId As Long, nAmount As Variant, nBaseAmount As Variant)
'  Dim nCuenta1 As Long, nCuenta2 As Long
'  '***********************************************************************************************
'  On Error GoTo ErrHandler
'    nCuenta1 = GetAccountId(sAppServer, nGralLedgerId, "5011", 1)
'    nCuenta2 = GetAccountId(sAppServer, nGralLedgerId, "5010", 1)
'    InitializePostingsRS sAppServer
'    AddPosting nCuenta1, 0, 0, 0, 0, "D", "", "", nMonedaExtranjera, nMonto, nMontoBase
'    AddPosting nCuentaExtranjera, nCuentaAuxExtranjera, 0, 0, 0, "H", "", "", nMonedaExtranjera, nMonto, nMontoBase
'    AddPosting nCuentaMonedaBase, 0, 0, 0, 0, "D", "", "", 1, nMontoBase, nMontoBase
'    AddPosting nCuenta2, 0, 0, 0, 0, "H", "", "", 1, nMontoBase, nMontoBase
'    SavePostingsRS sAppServer
'  Exit Sub
'ErrHandler:
'  RaiseError ClassId, "CompraventaDivisas", Err.Number
'End Sub
'
'Private Sub CompraventaDolares(sAppServer As String, nGralLedgerId As Long, _
'                              nCuentaDolares As Long, nCuentaAuxDolares As Long, _
'                              nCuentaMonedaBase As Long, _
'                              nMontoDlls As Variant, nMontoBase As Variant)
'  Dim nCuenta1 As Long, nCuenta2 As Long
'  On Error GoTo ErrHandler
'    nCuenta1 = GetAccountId(sAppServer, nGralLedgerId, "5011", 1)
'    nCuenta2 = GetAccountId(sAppServer, nGralLedgerId, "5010", 1)
'    InitializePostingsRS sAppServer
'    AddPosting nCuentaDolares, nCuentaAuxDolares, 0, 0, 0, "D", "", "", 2, nMontoDlls, nMontoBase
'    AddPosting nCuenta1, 0, 0, 0, 0, "H", "", "", 2, nMontoDlls, nMontoBase
'    AddPosting nCuenta2, 0, 0, 0, 0, "D", "", "", 1, nMontoBase, nMontoBase
'    AddPosting nCuentaMonedaBase, 0, 0, 0, 0, "H", "", "", 1, nMontoBase, nMontoBase
'    SavePostingsRS sAppServer
'  Exit Sub
'ErrHandler:
'  RaiseError ClassId, "CompraventaDolares", Err.Number
'End Sub

Private Function AddStdAccountToGralLedger(sAppServer As String, nGralLedgerId As Long, _
                                           nStdAccountId As Long) As Long
  Dim oGralLedger As New AOGralLedger.CGralLedger, sRole As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    AddStdAccountToGralLedger = oGralLedger.AssignStdAccount(sAppServer, nGralLedgerId, nStdAccountId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "AddStdAccountToGralLedger", Err.Number
End Function

Private Function CreateSubsidiaryAccount(sAppServer As String, nGralLedgerId As Long, _
                                         sSubsidiaryAccount As String) As Long
  Dim oSubsidiaryLedger As New CSubsidiaryLedger, oRecordset As Recordset, nSubsidiaryLedgerId As Long
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = oSubsidiaryLedger.GetSubsidiaryAccountRS(sAppServer)
    nSubsidiaryLedgerId = GetDefaultSubsidiaryLedgerId(sAppServer, nGralLedgerId)
    If nSubsidiaryLedgerId <> 0 Then
      With oRecordset
        !id_mayor_auxiliar = nSubsidiaryLedgerId
        !numero_cuenta_auxiliar = sSubsidiaryAccount
        !NOMBRE_CUENTA_AUXILIAR = "Sin nombre (Generado con el importador)"
        !DESCRIPCION = ""
        !ELIMINADA = 0
      End With
      CreateSubsidiaryAccount = oSubsidiaryLedger.SaveSubsidiaryAccount(sAppServer, oRecordset)
    Else
      CreateSubsidiaryAccount = 0
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateSubsidiaryAccount", Err.Number
End Function

Private Function GetAccountId(sAppServer As String, nGralLedgerId As Long, nStdAccountId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String, nAccountId As Long
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta FROM COF_Cuenta " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(id_mayor = " & nGralLedgerId & ")"
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
      nAccountId = AddStdAccountToGralLedger(sAppServer, nGralLedgerId, nStdAccountId)
    End If
    GetAccountId = nAccountId
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountId", Err.Number
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

Private Function GetGLGroupArray(sAppServer As String, nGralLedgerParentGroupId As Long, _
                                Optional sDelimiterChar As Variant) As Variant
  Dim oRecordset As New Recordset, sSQL As String, sAllGroups As String, sTemp As String, sChar As String
  '******************************************************************************************************
  On Error GoTo ErrHandler

    sAllGroups = GetGLGroupArray2(sAppServer, nGralLedgerParentGroupId)
    If Len(sAllGroups) <> 0 Then
      sAllGroups = nGralLedgerParentGroupId & ", " & sAllGroups
    Else
      sAllGroups = nGralLedgerParentGroupId
    End If
    sSQL = "SELECT DISTINCT id_mayor " & _
           "FROM COF_Elemento_Grupo_Mayor " & _
           "WHERE (id_grupo_mayor IN (" & sAllGroups & ")) " & _
           "ORDER BY id_mayor"
    If IsMissing(sDelimiterChar) Then
      sChar = ","
    Else
      sChar = sDelimiterChar
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sTemp = !id_mayor
        .MoveNext
        Do While Not .EOF
          sTemp = sTemp & sChar & !id_mayor
          .MoveNext
        Loop
        .Close
        If IsMissing(sDelimiterChar) Then
          GetGLGroupArray = Split(sTemp, sChar)
        Else
          GetGLGroupArray = sTemp
        End If
      Else
        GetGLGroupArray = Null
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLGroupArray", Err.Number
End Function

Private Function GetGLGroupArray2(sAppServer As String, nGralLedgerParentId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String, sTemp2 As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_grupo_mayor " & _
           "FROM COF_Grupo_Mayor " & _
           "WHERE (id_grupo_mayor_padre = " & nGralLedgerParentId & ") "
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        Do While Not .EOF
          sTemp = sTemp & ", " & !id_grupo_mayor
          sTemp2 = GetGLGroupArray2(sAppServer, !id_grupo_mayor)
          If Len(sTemp2) <> 0 Then
            sTemp = sTemp & ", " & GetGLGroupArray2(sAppServer, !id_grupo_mayor)
          End If
          .MoveNext
        Loop
        sTemp = Mid$(sTemp, 3)
      Else
        sTemp = ""
      End If
      .Close
    End With
    GetGLGroupArray2 = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLGroupArray2", Err.Number
End Function

Private Function GetGLStdAccountTypeId(sAppServer As String, nGralLedgerId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_tipo_cuentas_std FROM COF_Mayor " & _
           "WHERE (id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetGLStdAccountTypeId = !id_tipo_cuentas_std
      Else
        GetGLStdAccountTypeId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLStdAccountTypeId", Err.Number
End Function

Private Function GetObjectId(sAppServer As String, sSequence As String) As Long
  Dim oRecordset As New Recordset
  '****************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      .Open "SELECT " & sSequence & ".NEXTVAL AS ID FROM DUAL", _
            GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      GetObjectId = !Id
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetObjectId", Err.Number
End Function

Private Function GetPostings(sAppServer As String, aGralLedgers As Variant, dFromDate As Date, _
                             Optional dElaborationDate As Variant) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sGralLedgerCondition As String, sWhere As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    If IsArray(aGralLedgers) Then
      sGralLedgerCondition = " IN (" & Join(aGralLedgers, ",") & ")"
    Else
      sGralLedgerCondition = " = " & aGralLedgers
    End If
    If Not IsMissing(dElaborationDate) Then
      sWhere = "(COF_TRANSACCION.fecha_registro >= to_date('" & Format$(dElaborationDate, "dd/mm/yyyy hh:nn:ss") & "', 'DD/MM/YYYY HH24:MI:SS'))"
    End If
    
    sSQL = "SELECT COF_Cuenta_Estandar.numero_cuenta_estandar, " & _
           "COF_Cuenta_Auxiliar.numero_cuenta_auxiliar, " & _
           "COF_Movimiento.id_sector, COF_Movimiento.id_moneda, COF_Movimiento.tipo_movimiento, " & _
           "COF_Movimiento.monto, COF_Movimiento.monto_moneda_base " & _
           "FROM COF_Transaccion, COF_Movimiento, COF_Cuenta_Estandar, COF_Cuenta_Auxiliar, COF_Cuenta " & _
           "WHERE (COF_Transaccion.id_transaccion = COF_Movimiento.id_transaccion) AND " & _
           "(COF_Movimiento.id_cuenta = COF_Cuenta.id_cuenta) AND " & _
           "(COF_Cuenta.id_cuenta_estandar = COF_Cuenta_Estandar.id_cuenta_estandar) AND " & _
           "(COF_Movimiento.id_cuenta_auxiliar = COF_Cuenta_Auxiliar.id_cuenta_auxiliar (+)) AND " & _
           "(COF_TRANSACCION.id_mayor " & sGralLedgerCondition & ") AND " & _
           "(COF_TRANSACCION.fecha_afectacion = '" & dFromDate & "')"
    If Len(sWhere) <> 0 Then
      sSQL = sSQL & " AND " & sWhere
    End If
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set GetPostings = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetPostings", Err.Number
End Function

Private Function GetRuleDefinitionRS(sAppServer As String, nOperationId As Long, _
                                     ByRef nStdAccountTypeId As Long, dDate As Date) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    If nStdAccountTypeId <> 0 Then
      sSQL = "SELECT * FROM COF_REGLA_CONTABLE " & _
             "WHERE (id_objeto = " & nOperationId & ") AND " & _
             "(id_tipo_cuentas_std = " & nStdAccountTypeId & ") AND " & _
             "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    Else
      sSQL = "SELECT * FROM COF_REGLA_CONTABLE " & _
             "WHERE (id_objeto = " & nOperationId & ") AND " & _
             "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    End If
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
       If Not (.BOF And .EOF) Then
        Set GetRuleDefinitionRS = oRecordset
        If nStdAccountTypeId = 0 Then
          nStdAccountTypeId = !id_tipo_cuentas_std
        End If
        Exit Function
       End If
       .Close
    End With
    sSQL = "SELECT * FROM COF_REGLA_CONTABLE " & _
           "WHERE (id_objeto = " & nOperationId & ") AND " & _
           "(id_tipo_cuentas_std = 0) AND " & _
           "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin)"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set GetRuleDefinitionRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRuleDefinitionRS", Err.Number
End Function

Private Function GetRulesRS(sAppServer As String, nRuleId As Long, dRulesDate As Date) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.* FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_regla_contable = " & nRuleId & ") AND " & _
           "(fecha_inicio <= '" & dRulesDate & "') AND ('" & dRulesDate & "' <= fecha_fin) AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) " & _
           "ORDER BY id_entidad_agrupador, id_agrupador_origen, posicion, cuenta_origen_inicial, cuenta_origen_final"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set GetRulesRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRulesRS", Err.Number
End Function

Private Function GetSubsidiaryAccountId(sAppServer As String, nGralLedgerId As Long, _
                                        sSubsidiaryAccount As String, _
                                        Optional bAutoCreate As Boolean = True) As Long
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
        If bAutoCreate Then
          GetSubsidiaryAccountId = CreateSubsidiaryAccount(sAppServer, nGralLedgerId, sSubsidiaryAccount)
        Else
          GetSubsidiaryAccountId = 0
        End If
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetSubsidiaryAccountId", Err.Number
End Function

Private Function GetTransactionsRS(sAppServer As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Transaccion WHERE (id_transaccion = -1)"
    oRecordset.CursorLocation = adUseClient
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set GetTransactionsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetTransactionsRS", Err.Number
End Function

Private Function GetPostingsRS(sAppServer As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '**************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Movimiento_Tmp WHERE id_movimiento_tmp = -1"
    oRecordset.CursorLocation = adUseClient
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set GetPostingsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetPostingsRS", Err.Number
End Function

Private Sub PutGroupedPosting(sAppServer As String, oPostingsRS As Recordset, nTransactionId As Long, _
                              nAccountId As Long, nSubsidiaryAccountId As Long, nCurrencyId As Long, _
                              nSectorId As Long, sPostingType As String, nAmount As Double, nBaseAmount As Double)
  On Error GoTo ErrHandler
    With oPostingsRS
      .Filter = "(id_transaccion = " & nTransactionId & ") AND (id_cuenta = " & nAccountId & ") AND " & _
                "(id_cuenta_auxiliar = " & nSubsidiaryAccountId & ") AND (id_moneda = " & nCurrencyId & ") AND " & _
                "(id_sector = " & nSectorId & ") AND (tipo_movimiento = '" & sPostingType & "')"
      If Not (.BOF And .EOF) Then
        !monto = (!monto + nAmount)
        !monto_moneda_base = (!monto_moneda_base + nBaseAmount)
      Else
        PutPosting sAppServer, oPostingsRS, nTransactionId, nAccountId, nSubsidiaryAccountId, _
                   nSectorId, 0, 0, "", "", "", Null, "", sPostingType, _
                   nCurrencyId, nAmount, nBaseAmount, False
      End If
      .Filter = adFilterNone
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutGroupedPosting", Err.Number
End Sub

Private Sub PutPosting(sAppServer As String, oPostingsRS As Recordset, nTransactionId As Long, _
                       nAccountId As Long, nSubsidiaryAccountId As Long, nSectorId As Long, _
                       nReferencePostingId As Long, nResponsibilityAreaId As Long, sBudgetKey As Variant, _
                       sDisponibilityKey As Variant, sVerificationNumber As Variant, dPostingDate As Variant, _
                       sPostingConcept As Variant, sPostingType As String, nCurrencyId As Long, _
                       nAmount As Double, nBaseAmount As Double, bProtected As Boolean)
  On Error GoTo ErrHandler
    With oPostingsRS
      .AddNew
        !ID_MOVIMIENTO_TMP = GetObjectId(sAppServer, "SEC_ID_MOVIMIENTO_TMP")
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
        !PROTEGIDO = IIf(bProtected, 1, 0)
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutPosting", Err.Number
End Sub

Private Function PutTransaction(sAppServer As String, oTransactionRS As Recordset, _
                                nGralLedgerId As Long, nSourceId As Long, _
                                nTransactionTypeId As Long, nVoucherTypeId As Long, _
                                sTransactionConcept As String, dApplicationDate As Variant, _
                                dElaborationDate As Variant, _
                                nUserId As Long, nCheckerUserId As Long) As Long
  Dim nTransactionId As Long
  '******************************************************************************************
  On Error GoTo ErrHandler
    With oTransactionRS
      nTransactionId = GetObjectId(sAppServer, "SEC_ID_TRANSACCION")
      .AddNew
        !id_transaccion = nTransactionId
        !NUMERO_TRANSACCION = "No actualizada"
        !id_mayor = nGralLedgerId
        !ID_FUENTE = nSourceId
        !ID_TIPO_TRANSACCION = nTransactionTypeId
        !ID_TIPO_POLIZA = nVoucherTypeId
        !CONCEPTO_TRANSACCION = TrimAll(sTransactionConcept)
        !FECHA_AFECTACION = dApplicationDate
        !fecha_registro = dElaborationDate
        !ID_ELABORADA_POR = nUserId
        !ID_AUTORIZADA_POR = nCheckerUserId
        !ESTA_ABIERTA = 1
    End With
    PutTransaction = nTransactionId
  Exit Function
ErrHandler:
  RaiseError ClassId, "PutTransaction", Err.Number
End Function

Private Sub SaveRecordset(sAppServer As String, oTransactionsRS As Recordset)
  Dim oCon As New Connection
  '**************************************************************************
  On Error GoTo ErrHandler
    oCon.Open GetConstant(sConstantName:=sAppServer)
    With oTransactionsRS
      Set .ActiveConnection = oCon
      If Not (.BOF And .EOF) Then
        .UpdateBatch
      End If
      Set .ActiveConnection = Nothing
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveRecordset", Err.Number
End Sub

Private Function BuildBalanceFilter(sAppServer As String, sFromAccount As String, _
                                    sFromSubsAccount As String, nFromSectorId As Long, _
                                    nFromCurrencyId As Long) As String
  Dim sFilter As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sFromAccount = Replace(sFromAccount, "*", "%")
    sFromAccount = Replace(sFromAccount, "?", "_")
    If Len(sFromAccount) <> 0 Then
      sFilter = "(numero_cuenta_estandar like '" & sFromAccount & "')"
    End If
    sFromSubsAccount = Replace(sFromSubsAccount, "*", "%")
    sFromSubsAccount = Replace(sFromSubsAccount, "?", "_")
    If Len(sFromSubsAccount) <> 0 Then
      If Len(sFilter) <> 0 Then
        sFilter = sFilter & " AND "
      End If
      sFilter = sFilter & "(numero_cuenta_auxiliar like '" & sFromSubsAccount & "')"
    End If
    If nFromSectorId <> 0 Then
      If Len(sFilter) <> 0 Then
        sFilter = sFilter & " AND "
      End If
      sFilter = sFilter & "(id_sector = " & nFromSectorId & ")"
    End If
    If nFromCurrencyId <> 0 Then
      If Len(sFilter) <> 0 Then
        sFilter = sFilter & " AND "
      End If
      sFilter = sFilter & "(id_moneda = " & nFromCurrencyId & ")"
    End If
    If Len(sFilter) <> 0 Then
      sFilter = sFilter & " AND "
    End If
    sFilter = sFilter & "(saldo_actual <> 0)"
    
    BuildBalanceFilter = sFilter
  Exit Function
ErrHandler:
  RaiseError ClassId, "BuildBalanceFilter", Err.Number
End Function

Private Function GetStdAccountProps(sAppServer As String, nStdAccountTypeId As Long, _
                                    sStdAccountNumber As String, dDate As Date, _
                                    bUseSector As Boolean) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '***********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar, rol_cuenta " & _
           "FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (id_tipo_cuentas_std = " & nStdAccountTypeId & ") AND " & _
           "(numero_cuenta_estandar = '" & sStdAccountNumber & "') AND " & _
           "(fecha_inicio <= '" & dDate & "' AND fecha_fin >= '" & dDate & "')"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetStdAccountProps = !id_cuenta_estandar
        bUseSector = IIf(!rol_cuenta = "X", True, False)
      Else
        GetStdAccountProps = 0
        bUseSector = False
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountProps", Err.Number
End Function

Private Function GetStdAccountRole(sAppServer As String, nStdAccountId As Long, _
                                   dDate As Date, nSectorId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************************
  On Error GoTo ErrHandler
    If nSectorId <> 0 Then
      sSQL = "SELECT rol_sector rol " & _
             "FROM COF_Mapeo_Sector " & _
             "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
             "(id_sector = " & nSectorId & ") AND " & _
             "(fecha_inicio <= '" & dDate & "' AND fecha_fin >= '" & dDate & "')"
    Else
    sSQL = "SELECT rol_cuenta rol " & _
           "FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(fecha_inicio <= '" & dDate & "' AND fecha_fin >= '" & dDate & "')"
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
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

Private Function GetLastElaborationDate(sAppServer As String, nGralLedgerId As Long, _
                                        nVoucherTypeId As Long, dApplicationDate As Date) As Date
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT MAX(fecha_registro) fecha_registro " & _
           "FROM COF_Transaccion " & _
           "WHERE (id_mayor = " & nGralLedgerId & ") AND " & _
           "(fecha_afectacion = '" & dApplicationDate & "') AND " & _
           "(id_tipo_poliza = " & nVoucherTypeId & ") AND " & _
           "(esta_abierta = 0)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If Not IsNull(!fecha_registro) Then
          GetLastElaborationDate = !fecha_registro
        Else
          GetLastElaborationDate = CDate("30/11/2000")
        End If
      Else
        GetLastElaborationDate = CDate("30/11/2000")
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetLastElaborationDate", Err.Number
End Function
                                        

