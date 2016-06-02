VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDataBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Engine (EFARulesEngine)                   Parte   : Business services                 *
'* Clase      : CDataBuilder                                    Patrón  : Statefull class                   *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Calcula los saldos de cada uno de los elementos de una regla contable".                     *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CDataBuilder"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function Build(sAppServer As String, nRuleId As Long, _
                      vGralLedgers As Variant, vBalancesDates As Variant, _
                      nTransactionTypeId As Long, nVoucherTypeId As Long, _
                      Optional vExchangeRateDates As Variant, Optional nExchangeRateType As Long, _
                      Optional nExchangeRateCurrency As Long, Optional bComputeAverageFields As Boolean, _
                      Optional nGroupLevel As Long, Optional nRoundLevel As Long = 0, _
                      Optional bShowZeros As Boolean, Optional dRuleDate As Variant) As Recordset
  Dim oRecordset As New Recordset, nPeriodsCount As Long, dTempDate As Date
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    nPeriodsCount = (UBound(vBalancesDates) - LBound(vBalancesDates) + 1) \ 2
    If IsMissing(dRuleDate) Then
      dTempDate = Date
    ElseIf Not IsDate(dRuleDate) Then
      dTempDate = Date
    Else
      dTempDate = dRuleDate
    End If
    Set oRecordset = FillDataStructure(sAppServer, nRuleId, nPeriodsCount, bComputeAverageFields, dTempDate)
    CalculateLeafBalances sAppServer, oRecordset, vGralLedgers, vBalancesDates, vExchangeRateDates, _
                          nTransactionTypeId, nVoucherTypeId, nExchangeRateType, _
                          nExchangeRateCurrency, bComputeAverageFields, nRoundLevel
    CalculateParentBalances oRecordset, nPeriodsCount, bComputeAverageFields
    CalculateReferenceBalances oRecordset, nPeriodsCount, bComputeAverageFields
    UpdateFilteredBalances oRecordset, nPeriodsCount
    If nGroupLevel <> 0 Then
      DeleteLevels oRecordset, nGroupLevel
    End If
    If bShowZeros Then
      DeleteZeros oRecordset
    End If
    oRecordset.MoveFirst
    Set Build = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "Build", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Sub CalculateLeafBalances(sAppServer As String, oRecordset As Recordset, vGralLedgers As Variant, _
                                  vBalancesDates As Variant, vExchangeRateDates As Variant, _
                                  nTransactionTypeId As Long, nVoucherTypeId As Long, _
                                  nExcRateType As Long, nExcRateCurrency As Long, _
                                  bComputeAverageFields As Boolean, nRoundLevel As Long, _
                                  Optional bLargeVersion As Boolean = False)
  Dim oBalances As New EFABalanceEngine.CEngine, oBalancesRS As Recordset, oBalancesAuxRS As Recordset
  Dim nBalance As Double, nAverageBalance As Double
  Dim i As Long, dFromDate As Date, dToDate As Date, dExcRateDate As Date, nPeriodsCount As Long
  Dim sBalanceField As String, sNVBalanceField As String, sAverageField As String, sFilter As String
  '********************************************************************************************************
  On Error GoTo ErrHandler
    If Not IsArray(vBalancesDates) Then
      vBalancesDates = Array(vBalancesDates)
    End If
    If Not IsArray(vExchangeRateDates) Then
      vExchangeRateDates = Array(vExchangeRateDates)
    End If
    nPeriodsCount = ((UBound(vBalancesDates) - LBound(vBalancesDates)) + 1) \ 2
    For i = 1 To nPeriodsCount
      dFromDate = vBalancesDates(2 * (i - 1))
      dToDate = vBalancesDates((2 * (i - 1)) + 1)
      dExcRateDate = vExchangeRateDates(i - 1)
      sBalanceField = "SALDO_" & CStr(i)
      sNVBalanceField = "SALDO_" & CStr(i) & "_NV"
      sAverageField = "SALDO_PROMEDIO_" & CStr(i)

      If (Not bComputeAverageFields) Then
        Set oBalancesRS = oBalances.BalancesRS(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGralLedgers, True, True, _
                                               False, dFromDate, dToDate, nTransactionTypeId:=nTransactionTypeId, _
                                               nVoucherTypeId:=nVoucherTypeId, _
                                               nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency, _
                                               dExcRateDate:=dExcRateDate, nBalancesToDisplay:=4)
      Else
        Set oBalancesRS = oBalances.AverageBalancesRS(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGralLedgers, True, True, _
                                                      False, dFromDate, dToDate, nTransactionTypeId:=nTransactionTypeId, _
                                                      nVoucherTypeId:=nVoucherTypeId, _
                                                      nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency, _
                                                      dExcRateDate:=dExcRateDate, nBalancesToDisplay:=4)
      End If
      
      With oRecordset
        .Filter = "(tipo_grupo_cuenta <> 0)"
        Do While Not .EOF
          sFilter = GetFilterString(!cuenta_origen_inicial, !cuenta_origen_final, !filtro_cuentas_origen, _
                                    !auxiliar_origen, !id_sector_origen, !id_moneda_origen)
          If (IsNull(!auxiliar_origen) Or (Len(!auxiliar_origen) = 0)) And (!id_sector_origen = 0) Then
            sFilter = sFilter & IIf(Len(sFilter) = 0, "(", " AND (") & "rol_cuenta <> 'S')"
          End If
          If (Not bComputeAverageFields) Then
            If (IsNull(!auxiliar_origen) Or (Len(!auxiliar_origen) = 0)) Then
              GetBalance oBalancesRS, sFilter, !operador, !factor, nBalance, , nRoundLevel
            Else
              If oBalancesAuxRS Is Nothing Then
                Set oBalancesAuxRS = oBalances.BalancesRS(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGralLedgers, True, True, _
                                                          True, dFromDate, dToDate, nTransactionTypeId:=nTransactionTypeId, _
                                                          nVoucherTypeId:=nVoucherTypeId, _
                                                          nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency, _
                                                          dExcRateDate:=dExcRateDate, nBalancesToDisplay:=4)
              End If
              GetBalance oBalancesAuxRS, sFilter, !operador, !factor, nBalance, , nRoundLevel
            End If
            .Fields(sBalanceField) = nBalance
          Else
            If (IsNull(!auxiliar_origen) Or (Len(!auxiliar_origen) = 0)) Then
              GetBalance oBalancesRS, sFilter, !operador, !factor, nBalance, nAverageBalance, nRoundLevel
            Else
              If oBalancesAuxRS Is Nothing Then
                Set oBalancesAuxRS = oBalances.AverageBalancesRS(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGralLedgers, True, True, _
                                                                 True, dFromDate, dToDate, nTransactionTypeId:=nTransactionTypeId, _
                                                                 nVoucherTypeId:=nVoucherTypeId, _
                                                                 nExcRateType:=nExcRateType, nExcRateCurrency:=nExcRateCurrency, _
                                                                 dExcRateDate:=dExcRateDate, nBalancesToDisplay:=4)
              End If
              GetBalance oBalancesAuxRS, sFilter, !operador, !factor, nBalance, nAverageBalance, nRoundLevel
            End If
            .Fields(sBalanceField) = nBalance
            .Fields(sAverageField) = nAverageBalance
          End If
          .MoveNext
        Loop
        .Filter = adFilterNone
      End With
    Next i
    oRecordset.MoveFirst
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CalculateLeafBalances", Err.Number
End Sub

Private Sub CalculateParentBalances(oRecordset As Recordset, nPeriodsCount As Long, _
                                    bComputeAverageFields As Boolean)
  Dim oCurrentLevelRS As Recordset, nBalance As Double, nCurrentLevel As Long, i As Long
  '*************************************************************************************
  On Error GoTo ErrHandler
    Set oCurrentLevelRS = oRecordset.Clone()
    For nCurrentLevel = GetMaxEmptyLevel(oRecordset) To 1 Step -1
      With oCurrentLevelRS
        .Filter = "(tipo_grupo_cuenta = 0) AND (nivel = " & nCurrentLevel & ")"
        Do While Not .EOF
          PutBalancesWithRulesTable !id_grupo_cuenta, oRecordset, nPeriodsCount, bComputeAverageFields
          .MoveNext
        Loop
        .Filter = adFilterNone
      End With
    Next nCurrentLevel
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CalculateParentBalances", Err.Number
End Sub

Private Sub CalculateParentReferenceBalances(oRecordset As Recordset, nRuleId As Long, _
                                             aBalances() As Double, aAverageBalances() As Double)
  Dim oRefrencedBalancesRS As Recordset, nBalance As Double, nCurrentLevel As Long, i As Long
  Dim nParentId As Long, sFieldName As String
  '******************************************************************************************
  On Error GoTo ErrHandler
    Set oRefrencedBalancesRS = oRecordset.Clone()
    nParentId = GetParentId(oRecordset, nRuleId)
    Do
      If (nParentId <> 0) Then
        With oRefrencedBalancesRS
          .Filter = "(id_grupo_cuenta = " & nParentId & ")"
          If Not (.BOF And .EOF) Then
            For i = LBound(aBalances) To UBound(aBalances)
              sFieldName = "SALDO_" & CStr(i)
              .Fields(sFieldName) = .Fields(sFieldName) + aBalances(i)
            Next i
            If False Then
              For i = LBound(aAverageBalances) To UBound(aAverageBalances)
                sFieldName = "SALDO_PROMEDIO_" & CStr(i)
                .Fields(sFieldName) = .Fields(sFieldName) + aAverageBalances(i)
              Next i
            End If
          End If
          .Filter = adFilterNone
        End With
        nParentId = GetParentId(oRecordset, nParentId)
      Else
        Exit Do
      End If
    Loop
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CalculateReferenceBalances", Err.Number
End Sub

Private Sub CalculateReferenceBalances(oRecordset As Recordset, nPeriodsCount As Long, _
                                       bComputeAverageFields As Boolean)
  Dim oRefrencedBalancesRS As Recordset, nBalance As Double, nCurrentLevel As Long, i As Long
  Dim aBalances() As Double, aAverageBalances() As Double
  '******************************************************************************************
  On Error GoTo ErrHandler
    Set oRefrencedBalancesRS = oRecordset.Clone()
    With oRefrencedBalancesRS
      .Filter = "(tipo_grupo_cuenta = 5)"
      Do While Not .EOF
        ReDim aBalances(1 To nPeriodsCount)
        If (bComputeAverageFields) Then
          ReDim aAverageBalances(1 To nPeriodsCount)
        End If
        oRecordset.Filter = "(id_grupo_cuenta = " & !id_agrupador_origen & ")"
        If Not (oRecordset.BOF And oRecordset.EOF) Then
          For i = LBound(aBalances) To UBound(aBalances)
            aBalances(i) = aBalances(i) + oRecordset.Fields("SALDO_" & CStr(i))
            If bComputeAverageFields Then
              aAverageBalances(i) = aAverageBalances(i) + oRecordset.Fields("SALDO_PROMEDIO" & CStr(i))
            End If
          Next i
          For i = LBound(aBalances) To UBound(aBalances)
            If (!operador = "+") Then
              .Fields("SALDO_" & CStr(i)) = (aBalances(i) * !factor)
            ElseIf (!operador = "-") Then
              .Fields("SALDO_" & CStr(i)) = ((-1 * aBalances(i)) * !factor)
            ElseIf (!operador = "@") Then
              .Fields("SALDO_" & CStr(i)) = (Abs(aBalances(i)) * !factor)
            End If
            aBalances(i) = .Fields("SALDO_" & CStr(i))
            If Not IsNull(!filtro_cuentas_origen) Then
              If (!filtro_cuentas_origen = "(saldo_actual > 0)") And (.Fields("SALDO_" & CStr(i)) < 0) Then
                .Fields("SALDO_" & CStr(i)) = 0
              ElseIf (!filtro_cuentas_origen = "(saldo_actual < 0)") And (.Fields("SALDO_" & CStr(i)) > 0) Then
                .Fields("SALDO_" & CStr(i)) = 0
              End If
            End If
            aBalances(i) = .Fields("SALDO_" & CStr(i))
          Next i
          If bComputeAverageFields Then
            For i = LBound(aAverageBalances) To UBound(aAverageBalances)
              If (!operador = "+") Then
                .Fields("SALDO_PROMEDIO_" & CStr(i)) = (aAverageBalances(i) * !factor)
              ElseIf (!operador = "-") Then
                .Fields("SALDO_PROMEDIO_" & CStr(i)) = ((-1 * aAverageBalances(i)) * !factor)
              ElseIf (!operador = "@") Then
                .Fields("SALDO_PROMEDIO_" & CStr(i)) = (Abs(aAverageBalances(i)) * !factor)
              End If
              aAverageBalances(i) = .Fields("SALDO_PROMEDIO_" & CStr(i))
            Next i
          End If
        End If
        CalculateParentReferenceBalances oRecordset, !id_grupo_cuenta, aBalances, aAverageBalances
        oRecordset.Filter = adFilterNone
        .MoveNext
      Loop
      .Filter = adFilterNone
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CalculateReferenceBalances", Err.Number
End Sub

Private Function CreateDataStructure(Optional nPeriodsFieldsCount As Long = 1, _
                                     Optional bInsertAverageFields As Boolean = False) As Recordset
  Dim oRecordset As New Recordset, i As Long
  '************************************************************************************************
  On Error GoTo ErrHandler
    If nPeriodsFieldsCount < 1 Then
       nPeriodsFieldsCount = 1
    End If
    With oRecordset
      Set .ActiveConnection = Nothing
      .CursorLocation = adUseClient
      .LockType = adLockBatchOptimistic
      With .Fields
        .Append "ID_GRUPO_CUENTA", adInteger
        .Append "TIPO_GRUPO_CUENTA", adInteger
        .Append "NUMERO_GRUPO_CUENTA", adVarChar, 16, adFldMayBeNull
        .Append "NOMBRE_GRUPO_CUENTA", adVarChar, 512, adFldMayBeNull
        .Append "ID_ENTIDAD_AGRUPADOR", adInteger
        .Append "ID_AGRUPADOR_ORIGEN", adInteger
        .Append "CUENTA_ORIGEN_INICIAL", adVarChar, 256, adFldMayBeNull
        .Append "CUENTA_ORIGEN_FINAL", adVarChar, 256, adFldMayBeNull
        .Append "FILTRO_CUENTAS_ORIGEN", adVarChar, 1024, adFldMayBeNull
        .Append "AUXILIAR_ORIGEN", adVarChar, 256, adFldMayBeNull
        .Append "ID_SECTOR_ORIGEN", adInteger
        .Append "ID_MONEDA_ORIGEN", adInteger
        .Append "CLIP_CUENTA_ORIGEN", adVarChar, 1, adFldMayBeNull
        .Append "OPERADOR", adVarChar, 1
        .Append "FACTOR", adDouble
        .Append "POSICION", adInteger
        .Append "NIVEL", adInteger
        .Append "ID_GRUPO_CUENTA_PADRE", adInteger
        For i = 1 To nPeriodsFieldsCount
          .Append "SALDO_" & CStr(i), adDouble, adFldMayBeNull
          If bInsertAverageFields Then
            .Append "SALDO_PROMEDIO_" & CStr(i), adDouble, adFldMayBeNull
          End If
        Next i
      End With
      .Open
    End With
    Set CreateDataStructure = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateDataStructure", Err.Number
End Function

Private Sub DeleteLevels(oRecordset As Recordset, nGroupLevel As Long)
  Dim oBalancesRS As Recordset, i As Long, sFieldName As String
  '************************************************************
  On Error GoTo ErrHandler
    Set oBalancesRS = oRecordset.Clone()
    With oBalancesRS
      .Filter = "(nivel > " & nGroupLevel & ")"
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DeleteLevels", Err.Number
End Sub

Private Sub DeleteZeros(oRecordset As Recordset)
  Dim oBalancesRS As Recordset, i As Long, sFieldName As String
  '************************************************************
  On Error GoTo ErrHandler
    Set oBalancesRS = oRecordset.Clone()
    With oBalancesRS
      .Filter = "(SALDO_1 = 0)"
      Do While Not .EOF
        .Delete
        .MoveNext
      Loop
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DeleteZeros", Err.Number
End Sub

Private Function FillDataStructure(sAppServer As String, nRuleId As Long, Optional nPeriodsCount As Long, _
                                   Optional bInsertAverageFields As Boolean, _
                                   Optional dDate As Variant) As Recordset
  Dim oRulesRS As New Recordset, oRecordset As Recordset, dTempDate As Date
  '********************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dTempDate = Date
    ElseIf Not IsDate(dDate) Then
      dTempDate = Date
    Else
      dTempDate = dDate
    End If
    Set oRecordset = CreateDataStructure(nPeriodsCount, bInsertAverageFields)
    Set oRulesRS = GetRulesRS(sAppServer, nRuleId, dTempDate)
    With oRulesRS
      Do While Not .EOF
        oRecordset.AddNew
          oRecordset!id_grupo_cuenta = !id_grupo_cuenta
          oRecordset!tipo_grupo_cuenta = !tipo_grupo_cuenta
          oRecordset!NUMERO_GRUPO_CUENTA = !NUMERO_GRUPO_CUENTA
          oRecordset!NOMBRE_GRUPO_CUENTA = !NOMBRE_GRUPO_CUENTA
          oRecordset!ID_ENTIDAD_AGRUPADOR = !ID_ENTIDAD_AGRUPADOR
          oRecordset!id_agrupador_origen = !id_agrupador_origen
          oRecordset!cuenta_origen_inicial = !cuenta_origen_inicial
          oRecordset!cuenta_origen_final = !cuenta_origen_final
          oRecordset!filtro_cuentas_origen = !filtro_cuentas_origen
          oRecordset!auxiliar_origen = !auxiliar_origen
          oRecordset!id_sector_origen = !id_sector_origen
          oRecordset!id_moneda_origen = !id_moneda_origen
          oRecordset!clip_cuenta_origen = !clip_cuenta_origen
          oRecordset!operador = !operador
          oRecordset!factor = !factor
          oRecordset!posicion = !posicion
          oRecordset!nivel = !nivel
          oRecordset!id_grupo_cuenta_padre = !id_grupo_cuenta_padre
        .MoveNext
      Loop
    End With
    Set FillDataStructure = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "FillDataStructure", Err.Number
End Function

Private Sub GetBalance(oBalancesRS As Recordset, sFilter As String, sOperator As String, nFactor As Double, _
                        nBalance As Double, Optional nAverageBalance As Double = -1, _
                        Optional nRoundLevel As Long = 0)
  On Error GoTo ErrHandler
    oBalancesRS.Filter = sFilter
    If (nAverageBalance = -1) Then
      nBalance = 0
      Do While Not oBalancesRS.EOF
        nBalance = nBalance + oBalancesRS!saldo_actual
        oBalancesRS.MoveNext
      Loop
    Else
      nBalance = 0
      nAverageBalance = 0
      Do While Not oBalancesRS.EOF
        nBalance = nBalance + oBalancesRS!saldo_actual
        nAverageBalance = nAverageBalance + oBalancesRS!saldo_actual
        oBalancesRS.MoveNext
      Loop
    End If
    oBalancesRS.Filter = adFilterNone
    If (nBalance = 0) Then
      Exit Sub
    End If
    If (sOperator = "+") Then
      nBalance = nBalance * nFactor
    ElseIf (sOperator = "-") Then
      nBalance = (-1 * nBalance) * nFactor
    ElseIf (sOperator = "@") Then
      nBalance = Abs(nBalance) * nFactor
    End If
    If (nAverageBalance <> -1) Then
      If (sOperator = "+") Then
        nAverageBalance = nAverageBalance * nFactor
      ElseIf (sOperator = "-") Then
        nAverageBalance = (-1 * nAverageBalance) * nFactor
      ElseIf (sOperator = "@") Then
        nAverageBalance = Abs(nAverageBalance) * nFactor
      End If
    End If
    Select Case nRoundLevel
      Case 0
        nBalance = Round(nBalance, 2)
      Case 1
        nBalance = Fix(nBalance)
      Case 2
        nBalance = Fix(nBalance / 1000)
      Case 3
        nBalance = Fix(nBalance / 1000000)
    End Select
    If (nAverageBalance <> -1) Then
      Select Case nRoundLevel
        Case 0
          nAverageBalance = Round(nAverageBalance, 2)
        Case 1
          nAverageBalance = Fix(nAverageBalance)
        Case 2
          nAverageBalance = Fix(nAverageBalance / 1000)
        Case 3
          nAverageBalance = Fix(nAverageBalance / 1000000)
      End Select
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "GetBalance", Err.Number
End Sub

Private Function GetFilterString(sFromAccount As Variant, sToAccount As Variant, sFilter As Variant, _
                                 sSubsAccount As Variant, nSectorId As Long, nCurrencyId As Long) As String
  Dim stFromAccount As String, stToAccount As String, stSubsAccount As String, stFilter As String
  Dim sSQL As String
  '********************************************************************************************************
  On Error GoTo ErrHandler
    stFromAccount = IIf(IsNull(sFromAccount), "", sFromAccount)
    stToAccount = IIf(IsNull(sToAccount), "", sToAccount)
    stSubsAccount = IIf(IsNull(sSubsAccount), "", sSubsAccount)
    stFilter = IIf(IsNull(sFilter), "", sFilter)
    
    If (Right$(stFromAccount, 1) <> "*") And (Len(stToAccount) = 0) Then
      sSQL = "(numero_cuenta_estandar = '" & stFromAccount & "') "
    ElseIf (Right$(stFromAccount, 1) = "*") And (Len(stToAccount) = 0) Then
      sSQL = "(numero_cuenta_estandar LIKE '" & Left$(stFromAccount, Len(stFromAccount) - 1) & "%') "
    ElseIf (stFromAccount = stToAccount) Then
      sSQL = "(numero_cuenta_estandar = '" & stFromAccount & "') "
    ElseIf (stFromAccount <> stToAccount) And (Right$(stToAccount, 1) <> "*") Then
      sSQL = "(numero_cuenta_estandar >= '" & stFromAccount & "') AND " & _
             "(numero_cuenta_estandar <= '" & stToAccount & "') "
    ElseIf (stFromAccount <> stToAccount) And (Right$(stToAccount, 1) = "*") Then
      sSQL = "(numero_cuenta_estandar >= '" & stFromAccount & "') AND " & _
             "(numero_cuenta_estandar <= '" & Left$((Left$(stToAccount, Len(stToAccount) - 1) & "-99-99-99-99-99-99"), 22) & "') "
    End If
    If (Len(stFilter) <> 0) Then
      sSQL = sSQL & "AND (" & stFilter & ") "
    End If
    If Len(stSubsAccount) <> 0 Then
      sSQL = sSQL & "AND (numero_cuenta_auxiliar LIKE '" & stSubsAccount & "') "
    End If
    If (nSectorId) > 0 Then
      sSQL = sSQL & "AND (id_sector = " & nSectorId & ") "
    ElseIf (nSectorId < 0) Then
      sSQL = sSQL & "AND (id_sector <> " & Abs(nSectorId) & ") "
    End If
    If (nCurrencyId > 0) Then
      sSQL = sSQL & "AND (id_moneda = " & nCurrencyId & ") "
    ElseIf (nCurrencyId < 0) Then
      sSQL = sSQL & "AND (id_moneda <> " & Abs(nCurrencyId) & ") "
    End If
    GetFilterString = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetFilterString", Err.Number
End Function

Private Function GetMaxEmptyLevel(oRecordset As Recordset) As Long
  Dim nMaxEmptyLevel As Long
  '***************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      .Filter = "(tipo_grupo_cuenta = 0)"
      Do While Not .EOF
        If IsEmpty(!saldo_1) Or IsNull(!saldo_1) Then
          If nMaxEmptyLevel < !nivel Then
            nMaxEmptyLevel = !nivel
          End If
        End If
        .MoveNext
      Loop
      .Filter = adFilterNone
    End With
    GetMaxEmptyLevel = nMaxEmptyLevel
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetMaxEmptyLevel", Err.Number
End Function

Private Function GetParentId(oRecordset As Recordset, nRuleId As Long) As Long
  Dim oCloneRS As Recordset
  '*****************************************************************************
  On Error GoTo ErrHandler
    Set oCloneRS = oRecordset.Clone()
    With oCloneRS
      .Filter = "(id_grupo_cuenta = " & nRuleId & ")"
      If Not (.BOF And .EOF) Then
        GetParentId = !id_grupo_cuenta_padre
      Else
        GetParentId = 0
      End If
      .Filter = adFilterNone
    End With
    Set oCloneRS = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetParentId", Err.Number
End Function

Private Function GetRulesRS(sAppServer As String, nRuleId As Long, dRulesDate As Date) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.* FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_regla_contable = " & nRuleId & ") AND " & _
           "(fecha_inicio <= '" & dRulesDate & "') AND ('" & dRulesDate & "' <= fecha_fin) AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) " & _
           "ORDER BY posicion"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set GetRulesRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRulesRS", Err.Number
End Function

Private Sub PutBalancesWithRulesTable(nAccountGroupId As Long, oRecordset As Recordset, _
                                      nPeriodsCount As Long, bComputeAverageFields As Boolean)
  Dim aBalances() As Double, aAverageBalances() As Double
  Dim sFieldName As String, i As Long
  '*******************************************************************************************
  On Error GoTo ErrHandler
    ReDim aBalances(1 To nPeriodsCount)
    If (bComputeAverageFields) Then
      ReDim aAverageBalances(1 To nPeriodsCount)
    End If
    With oRecordset
      .Filter = "(id_grupo_cuenta_padre = " & nAccountGroupId & ")"
      Do While Not .EOF
        For i = LBound(aBalances) To UBound(aBalances)
          aBalances(i) = aBalances(i) + .Fields("SALDO_" & CStr(i))
          If bComputeAverageFields Then
            aAverageBalances(i) = aAverageBalances(i) + .Fields("SALDO_PROMEDIO" & CStr(i))
          End If
        Next i
        .MoveNext
      Loop
      .Filter = adFilterNone
      .Filter = "(id_grupo_cuenta = " & nAccountGroupId & ")"
      For i = LBound(aBalances) To UBound(aBalances)
        sFieldName = "SALDO_" & CStr(i)
        If (!operador = "+") Then
          .Fields(sFieldName) = (aBalances(i) * !factor)
        ElseIf (!operador = "-") Then
          .Fields(sFieldName) = ((-1 * aBalances(i)) * !factor)
        ElseIf (!operador = "@") Then
          .Fields(sFieldName) = (Abs(aBalances(i)) * !factor)
        End If
        If Not IsNull(!filtro_cuentas_origen) Then
          If (!filtro_cuentas_origen = "(saldo_actual > 0)") And (.Fields(sFieldName) < 0) Then
            .Fields(sFieldName) = 0
          ElseIf (!filtro_cuentas_origen = "(saldo_actual < 0)") And (.Fields(sFieldName) > 0) Then
            .Fields(sFieldName) = 0
          End If
        End If
      Next i
      If bComputeAverageFields Then
        For i = LBound(aAverageBalances) To UBound(aAverageBalances)
          If (!operador = "+") Then
            .Fields("SALDO_PROMEDIO_" & CStr(i)) = (aAverageBalances(i) * !factor)
          ElseIf (!operador = "-") Then
            .Fields("SALDO_PROMEDIO_" & CStr(i)) = ((-1 * aAverageBalances(i)) * !factor)
          ElseIf (!operador = "@") Then
            .Fields("SALDO_PROMEDIO_" & CStr(i)) = (Abs(aAverageBalances(i)) * !factor)
          End If
        Next i
      End If
      .Filter = adFilterNone
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutBalancesWithRulesTable", Err.Number
End Sub

Private Sub UpdateFilteredBalances(oRecordset As Recordset, nPeriodsCount As Long)
  Dim oBalancesRS As Recordset, i As Long, sFieldName As String
  '******************************************************************************
  On Error GoTo ErrHandler
    Set oBalancesRS = oRecordset.Clone()
    With oBalancesRS
      .Filter = "(filtro_cuentas_origen <> NULL)"
      Do While Not .EOF
        If (!filtro_cuentas_origen = "(saldo_actual > 0)") Then
          For i = 1 To nPeriodsCount
            sFieldName = "SALDO_" & CStr(i)
            If (.Fields(sFieldName) < 0) Then
              .Fields(sFieldName) = 0
            End If
          Next i
        ElseIf (!filtro_cuentas_origen = "(saldo_actual < 0)") Then
          For i = 1 To nPeriodsCount
            sFieldName = "SALDO_" & CStr(i)
            If (.Fields(sFieldName) > 0) Then
              .Fields(sFieldName) = 0
            End If
          Next i
        End If
        .MoveNext
      Loop
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "UpdateFilteredBalances", Err.Number
End Sub
