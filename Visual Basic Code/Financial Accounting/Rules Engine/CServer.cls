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
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Engine (EFARulesEngine)                   Parte   : Business services                 *
'* Clase      : CServer                                         Patrón  : Stateless class                   *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase principal del componente 'Financial Accounting: Rules Engine'.                        *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CServer"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function InterpretRules(sAppServer As String, vaRules As Variant, vaGralLedgers As Variant, _
                               vnBalancesTypes As Variant, vdBalancesPeriods As Variant, _
                               nTransactionTypeId As Long, nVoucherTypeId As Long, _
                               vsBalancesCurrencies As Variant, _
                               Optional vnExchangeRateTypes As Variant, _
                               Optional vnExchangeRateCurrencies As Variant, _
                               Optional vdExchangeRateDates As Variant, _
                               Optional vbRoundBalancesTo As Variant, _
                               Optional sInterpretationPars As String, _
                               Optional dRuleDate As Variant) As Recordset
  Dim oDataBuilder As New CDataBuilder, oDataRS As Recordset, oTempRS As Recordset, nRuleId As Long
  Dim nFieldsCount As Long, nField As Long, i As Long
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dRuleDate) Then
      dRuleDate = Date
    End If
    If IsArray(vaRules) Then
      For i = LBound(vaRules) To UBound(vaRules)
        nRuleId = vaRules(i)
        If i = LBound(vaRules) Then
          Set oDataRS = oDataBuilder.Build(sAppServer, nRuleId, vaGralLedgers(i), vdBalancesPeriods, _
                                           nTransactionTypeId, nVoucherTypeId, vdExchangeRateDates, _
                                           CLng(vnExchangeRateTypes(0)), CLng(vnExchangeRateCurrencies(0)), False, _
                                           0, CLng(vbRoundBalancesTo), True, dRuleDate)
        Else
          Set oTempRS = oDataBuilder.Build(sAppServer, nRuleId, vaGralLedgers(i), vdBalancesPeriods, _
                                           nTransactionTypeId, nVoucherTypeId, vdExchangeRateDates, _
                                           CLng(vnExchangeRateTypes(0)), CLng(vnExchangeRateCurrencies(0)), False, _
                                           0, CLng(vbRoundBalancesTo), True, dRuleDate)
          With oTempRS
            If Not (.BOF And .EOF) Then
              .MoveFirst
              nFieldsCount = .Fields.Count
              Do While Not .EOF
                oDataRS.AddNew
                For nField = 0 To nFieldsCount - 1
                  oDataRS.Fields(nField).Value = .Fields(nField).Value
                Next nField
                .MoveNext
              Loop
            End If
            .Close
          End With
          Set oTempRS = Nothing
        End If
      Next i
    Else
      nRuleId = CLng(vaRules)
      If nRuleId = 11 Then    'Rentabilidad
        Set oDataRS = oDataBuilder.Build(sAppServer, nRuleId, vaGralLedgers, vdBalancesPeriods, _
                                         nTransactionTypeId, nVoucherTypeId, vdExchangeRateDates, _
                                         CLng(vnExchangeRateTypes(0)), CLng(vnExchangeRateCurrencies(0)), True, _
                                         0, CLng(vbRoundBalancesTo), True, dRuleDate)
      Else
        Set oDataRS = oDataBuilder.Build(sAppServer, nRuleId, vaGralLedgers, vdBalancesPeriods, _
                                         nTransactionTypeId, nVoucherTypeId, vdExchangeRateDates, _
                                         CLng(vnExchangeRateTypes(0)), CLng(vnExchangeRateCurrencies(0)), False, _
                                         0, CLng(vbRoundBalancesTo), True, dRuleDate)
      End If
    End If
    Set InterpretRules = oDataRS
  Exit Function
ErrHandler:
  RaiseError ClassId, "InterpretRules", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function GetReportRulesRS(sAppServer As String, nRuleId As Long, dRulesDate As Date) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.*, " & _
           "NVL(AO_Currencies.o_id_moneda, '00') clave_moneda_origen, " & _
           "NVL(COF_Sector.clave_sector, '00') clave_sector_origen " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis, COF_Sector, AO_Currencies " & _
           "WHERE (id_regla_contable = " & nRuleId & ") AND " & _
           "(fecha_inicio <= '" & dRulesDate & "') AND ('" & dRulesDate & "' <= fecha_fin) AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) AND " & _
           "(ABS(COF_Grupo_Cuenta.id_sector_origen) = COF_Sector.id_sector (+)) AND " & _
           "(ABS(COF_Grupo_Cuenta.id_moneda_origen) = AO_Currencies.currency_id (+)) " & _
           "ORDER BY posicion, cuenta_origen_inicial, cuenta_origen_final"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set GetReportRulesRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetReportRulesRS", Err.Number
End Function

Private Function RuleGroupedByStdAccountType(sAppServer As String, nRuleId As Long, _
                                             dRulesDate As Date) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_entidad_agrupador FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_regla_contable = " & nRuleId & ") AND " & _
           "(id_entidad_agrupador = 1) AND " & _
           "(fecha_inicio <= '" & dRulesDate & "') AND ('" & dRulesDate & "' <= fecha_fin) AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        RuleGroupedByStdAccountType = True
      Else
        RuleGroupedByStdAccountType = False
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleGroupedByStdAccountType", Err.Number
End Function

Private Function FormatSubsidiaryAccount(sGralLedgerNumber As String, nGralLedgerId As Long, _
                                         sSubsidiaryAccount As String) As String
  Dim sTemp As String, i As Long, sChar As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    For i = 1 To Len(sSubsidiaryAccount)
      If InStr(1, "0123456789", Mid$(sSubsidiaryAccount, i, 1)) = 0 Then
        FormatSubsidiaryAccount = ""
        Exit Function
      End If
    Next i
    If (Len(sSubsidiaryAccount) <= 16) Then
      If Left$(sSubsidiaryAccount, "*") > 0 Then
        FormatSubsidiaryAccount = Format(sSubsidiaryAccount, String(16, "0"))
      Else
        FormatSubsidiaryAccount = Format(sSubsidiaryAccount, String(16, "0"))
      End If
    Else
      FormatSubsidiaryAccount = ""
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "FormatSubsidiaryAccount", Err.Number
End Function
