VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUngroupedItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Manager BS (EFARulesMgrBS)                Parte   : Business services                 *
'* Clase      : CUngroupedItems                                 Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Representa los elementos no agrupados para una regla y fecha determinadas.                  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CUngroupedItems"

Public Function UngroupedAccountsHeader(Optional nSelectedColumnId As Long = 1) As String
  Dim oRecordset As Recordset, sHTML As String, sTemp As String
  
  '**************************************************************************************
  On Error GoTo ErrHandler
    UngroupedAccountsHeader = "<TD>Cuenta</TD><TD>Sector</TD><TD>Moneda</TD><TD>Nombre</TD>"
  Exit Function
ErrHandler:
  RaiseError ClassId, "UngroupedAccountsHeader", Err.Number
End Function

Public Function UngroupedAccountsTbl(sAppServer As String, nRuleDefId As Long, _
                                     Optional dHistoricDate As Variant) As String
  Const cGroupRow = _
    "<TR class=applicationTableRowDivisor><TD colspan=4>" & _
      "<A href='' onclick=""return(callEditor('<@NUMBER@>', 0, 0));""><@NUMBER@></A>" & _
    "</TD></TR>" & vbCrLf
  Const cRow = _
    "<TR><TD>" & _
    "<A href='' onclick=""return(callEditor('<@NUMBER@>', <@SECTOR_ID@>, <@CURRENCY_ID@>));""><@NUMBER@></A>" & _
    "</TD><TD><@SECTOR@></TD><TD><@CURRENCY@></TD><TD><@NAME@></TD></TR>" & vbCrLf
  Dim oRecordset As New Recordset, sHTML As String, sTemp As String, dDate As Date, sLastGroup As String
  '************************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dHistoricDate) Then
      dDate = Date
    Else
      dDate = dHistoricDate
    End If
    Set oRecordset = UngroupedItemsRS(sAppServer, nRuleDefId, dDate)
    With oRecordset
      If Not (.BOF And .EOF) Then
        .Sort = "numero_cuenta_estandar, clave_sector, clave_moneda"
        .MoveFirst
      End If
      Do While Not .EOF
        If sLastGroup <> Left$(!numero_cuenta_estandar, 4) Then
          sLastGroup = Left$(!numero_cuenta_estandar, 4)
          sHTML = sHTML & Replace(cGroupRow, "<@NUMBER@>", sLastGroup & "*")
        End If
        sTemp = Replace(cRow, "<@NUMBER@>", !numero_cuenta_estandar)
        sTemp = Replace(sTemp, "<@NAME@>", !nombre_cuenta_estandar)
        sTemp = Replace(sTemp, "<@SECTOR@>", !clave_sector)
        sTemp = Replace(sTemp, "<@SECTOR_ID@>", !id_sector)
        sTemp = Replace(sTemp, "<@CURRENCY@>", !clave_moneda)
        sTemp = Replace(sTemp, "<@CURRENCY_ID@>", !id_moneda)
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    UngroupedAccountsTbl = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "UngroupedAccountsTbl", Err.Number
End Function

Private Function UngroupedItemsRS(sAppServer As String, nRuleDefId As Long, dDate As Date) As Recordset
  Dim oUngroupedItemsRS As Recordset, oRulesRS As New Recordset
  Dim oBalancesRS As New Recordset, sActsFilter As String, sFilter As String
  '****************************************************************************************************
  On Error GoTo ErrHandler
    Set oBalancesRS = BalancesRS(sAppServer, nRuleDefId, dDate)
    Set oRulesRS = RulesRS(sAppServer, nRuleDefId, dDate)
    Set oUngroupedItemsRS = CreateDataStructure()
    With oBalancesRS
      .Filter = "(id_moneda <> 0) AND (rol_cuenta <> 'S')"
      Do While Not .EOF
        sActsFilter = "((cuenta_origen_inicial <= '" & !numero_cuenta_estandar & "') AND " & _
                       "(cuenta_origen_final >= '" & !numero_cuenta_estandar & "'))"
        If (!id_sector <> 0) Then
          sFilter = "(" & sActsFilter & " AND " & "(id_sector_origen = 0)) OR " & _
                    "(" & sActsFilter & " AND " & "(id_sector_origen = " & !id_sector & "))"
        Else
          sFilter = "(" & sActsFilter & " AND " & "(id_sector_origen = 0))"
        End If
        With oRulesRS
          .Filter = sFilter
          If (.BOF And .EOF) Then
            With oUngroupedItemsRS
              .AddNew
                !numero_cuenta_estandar = oBalancesRS!numero_cuenta_estandar
                !nombre_cuenta_estandar = oBalancesRS!nombre_cuenta
                !id_sector = oBalancesRS!id_sector
                !clave_sector = oBalancesRS!clave_sector
                !id_moneda = oBalancesRS!id_moneda
                !clave_moneda = oBalancesRS!clave_moneda
            End With
          End If
          .Filter = adFilterNone
        End With
        .MoveNext
      Loop
    End With
    Set UngroupedItemsRS = oUngroupedItemsRS
  Exit Function
ErrHandler:
  RaiseError ClassId, "UngroupedAccountsRS", Err.Number
End Function

Private Function RulesRS(sAppServer As String, nRuleDefId As Long, dDate As Date) As Recordset
  Dim oRules As New Recordset, sSQL As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT " & _
              "DECODE(NVL(cuenta_origen_final, ''), '', " & _
                  "REPLACE(cuenta_origen_inicial, '*', ''), " & _
                  "REPLACE(cuenta_origen_inicial, '*', '00-00-00-00-00-00')) cuenta_origen_inicial, " & _
              "DECODE(NVL(cuenta_origen_final, ''), '', " & _
                  "REPLACE(cuenta_origen_inicial, '*', '-99-99-99-99-99-99'), " & _
                  "REPLACE(cuenta_origen_final, '*', '-99-99-99-99-99-99')) cuenta_origen_final, " & _
              "id_sector_origen, id_moneda_origen " & _
           "FROM COF_Grupo_Cuenta " & _
           "WHERE (id_regla_contable = " & nRuleDefId & ") AND (tipo_grupo_cuenta <> 0) " & _
           "ORDER BY cuenta_origen_inicial, cuenta_origen_final, id_sector_origen, id_moneda_origen"
    With oRules
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
       Set .ActiveConnection = Nothing
    End With
    Set RulesRS = oRules
  Exit Function
ErrHandler:
  RaiseError ClassId, "RulesRS", Err.Number
End Function

Private Function BalancesRS(sAppServer As String, nRuleDefId As Long, dDate As Date) As Recordset
  Dim oBalances As New EFABalanceEngine.CEngine, oBalancesRS As Recordset, vaGralLedgers As Variant
  '************************************************************************************************
  On Error GoTo ErrHandler
    vaGralLedgers = RuleDefGralLedgers(sAppServer, nRuleDefId)
    Set oBalancesRS = oBalances.BalancesRS(sAppServer, "####-##-##-##-##-##-##", vaGralLedgers, _
                                           True, True, False, #12/29/2000#, dDate)
    Set BalancesRS = oBalancesRS
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalancesRS", Err.Number
End Function

Private Function RuleDefGralLedgers(sAppServer As String, nRuleDefId As Long) As Variant
  Dim oRecordset As New Recordset, sSQL As String, vaGralLedgers As Variant, nStdAccountTypeId As Long
  '***************************************************************************************************
  On Error GoTo ErrHandler
    nStdAccountTypeId = GetDataValue(sAppServer, cnQryRuleDefStdAccountTypeId, nRuleDefId)
    sSQL = "SELECT id_mayor FROM COF_Mayor " & _
           "WHERE (id_tipo_cuentas_std = " & nStdAccountTypeId & ") " & _
           "ORDER BY id_mayor"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Not IsEmpty(vaGralLedgers) Then
          ReDim Preserve vaGralLedgers(UBound(vaGralLedgers) + 1)
        Else
          ReDim vaGralLedgers(0)
        End If
        vaGralLedgers(UBound(vaGralLedgers)) = !id_mayor
        .MoveNext
      Loop
    End With
    RuleDefGralLedgers = vaGralLedgers
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefGralLedgers", Err.Number
End Function

Private Function CreateDataStructure() As Recordset
  Dim oRecordset As New Recordset
  '************************************************
  On Error GoTo ErrHandler
    With oRecordset
      Set .ActiveConnection = Nothing
      .CursorLocation = adUseClient
      .LockType = adLockBatchOptimistic
      With .Fields
        .Append "numero_cuenta_estandar", adVarChar, 256, adFldMayBeNull
        .Append "nombre_cuenta_estandar", adVarChar, 512, adFldMayBeNull
        .Append "id_sector", adInteger
        .Append "clave_sector", adVarChar, 6, adFldMayBeNull
        .Append "id_moneda", adInteger
        .Append "clave_moneda", adVarChar, 6, adFldMayBeNull
      End With
      .Open
    End With
    Set CreateDataStructure = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateDataStructure", Err.Number
End Function
