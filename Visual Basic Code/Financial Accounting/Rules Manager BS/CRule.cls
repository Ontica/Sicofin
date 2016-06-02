VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Manager BS (EFARulesMgrBS)                Parte   : Business services                 *
'* Clase      : CRule                                           Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Representa un elemento de una definición de regla de la base de conocimiento contable.      *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CRule"
                 
Private Const cnRuleLineType1 = _
        "<@TAB@>|<A href='' onclick='return callEditor(1, <@ID@>);'><@FROM_ACCOUNT@></A>|" & _
        "<@TO_ACCOUNT@>|<@SUBS_ACCOUNT@>|<@CURRENCY@>|<@SECTOR@>|<@RESTRICTION@>|" & _
        "<@TARGET_ACCOUNT@>|<@TARGET_SUBS_ACCOUNT@>|" & _
        "<@TARGET_OVER_BALANCED_ACCOUNT@>|<@TARGET_OVER_BALANCED_SUBS_ACCOUNT@>"

Private Const cnRuleLineType2 = _
  "<@TAB@>|<A href='' onclick='return callEditor(1, <@ID@>);'><@LABEL@></A>|" & _
  "<A href='' onclick='return callEditor(1, <@ID@>);'><@FROM_ACCOUNT@></A>|" & _
  "<@TO_ACCOUNT@>|<@SECTOR@>|<@CURRENCY@>|<@SUBS_ACCOUNT@>|<@OPERATOR@>"
  
Private Const cnRuleLineTypeRef = _
  "<@TAB@>|<A href='' onclick='return callEditor(5, <@ID@>);'><@LABEL@> &nbsp;<b>[Referencia]</b></A>|" & _
  "|&nbsp;|&nbsp;|&nbsp;|&nbsp;|<@OPERATOR@>"
  
'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Sub AddRule(sAppServer As String, oRulesRS As Recordset, nRuleGroupId As Long, _
                   nInsertAfterItemId As Long)
  Dim oRecordset As New Recordset, sSQL1 As String, sSQL2 As String
  Dim sTemp As String, nNewItemOrder As Long, nItem As Long
  '************************************************************************************
  On Error GoTo ErrHandler
    sSQL1 = "SELECT * FROM COF_Grupo_Cuenta WHERE (id_grupo_cuenta = -1)"
    sSQL2 = "SELECT * FROM COF_Grupo_Cuenta_BIS WHERE (id_grupo_cuenta = -1)"
    nNewItemOrder = GetRulePosition(sAppServer, nInsertAfterItemId) + 1
    ReorderRuleItems sAppServer, oRulesRS!id_regla_contable, nNewItemOrder
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL1, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      .ActiveConnection = Nothing
      .AddNew
      nItem = NewRecordId(sAppServer, "SEC_ID_GRUPO_CUENTA")
      !id_grupo_cuenta = nItem
      !id_regla_contable = oRulesRS!id_regla_contable
      !tipo_grupo_cuenta = oRulesRS!tipo_grupo_cuenta
      !numero_grupo_cuenta = oRulesRS!numero_grupo_cuenta
      !nombre_grupo_cuenta = oRulesRS!nombre_grupo_cuenta
      !id_entidad_agrupador = oRulesRS!id_entidad_agrupador
      !id_agrupador_origen = oRulesRS!id_agrupador_origen
      !cuenta_origen_inicial = oRulesRS!cuenta_origen_inicial
      !cuenta_origen_final = oRulesRS!cuenta_origen_final
      !filtro_cuentas_origen = oRulesRS!filtro_cuentas_origen
      !auxiliar_origen = oRulesRS!auxiliar_origen
      !id_sector_origen = oRulesRS!id_sector_origen
      !id_moneda_origen = oRulesRS!id_moneda_origen
      !clip_cuenta_origen = oRulesRS!clip_cuenta_origen
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
      
      .CursorLocation = adUseClient
      .Open sSQL2, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      .ActiveConnection = Nothing
      .AddNew
      !id_grupo_cuenta = nItem
      !ID_CUENTA_DESTINO = oRulesRS!ID_CUENTA_DESTINO
      !ID_AUXILIAR_DESTINO = oRulesRS!ID_AUXILIAR_DESTINO
      !ID_SECTOR_DESTINO = oRulesRS!ID_SECTOR_DESTINO
      !ID_CUENTA_DESTINO_SOBREGIRO = oRulesRS!ID_CUENTA_DESTINO_SOBREGIRO
      !ID_AUXILIAR_DESTINO_SOBREGIRO = oRulesRS!ID_AUXILIAR_DESTINO_SOBREGIRO
      !ID_SECTOR_DESTINO_SOBREGIRO = oRulesRS!ID_SECTOR_DESTINO_SOBREGIRO
      !operador = oRulesRS!operador
      !FACTOR = oRulesRS!FACTOR
      !posicion = nNewItemOrder
      !nivel = GetRuleLevel(sAppServer, nRuleGroupId) + 1
      !id_grupo_cuenta_padre = nRuleGroupId
      !FECHA_INICIO = #12/31/2000#
      !FECHA_FIN = #12/31/2049#
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddVoucherRule", Err.Number
End Sub

Public Sub AddRuleGroup(sAppServer As String, oRulesRS As Recordset, _
                        nParentRuleGroupId As Long, nInsertAfterItemId As Long, bSaveAsChild As Boolean)
  Dim oRecordset As New Recordset, sSQL1 As String, sSQL2 As String
  Dim nNewItemPosition As Long, nItem As Long
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    sSQL1 = "SELECT * FROM COF_Grupo_Cuenta WHERE (id_grupo_cuenta = -1)"
    sSQL2 = "SELECT * FROM COF_Grupo_Cuenta_BIS WHERE (id_grupo_cuenta = -1)"
    If (nInsertAfterItemId <> 0) Then
      If (RuleChildsType(sAppServer, nInsertAfterItemId) <> 0) Then      'Son rangos de cuentas
        nNewItemPosition = GetRulePosition(sAppServer, nInsertAfterItemId) + RuleChildsCount(sAppServer, nInsertAfterItemId) + 1
      Else   'Los hijos son grupos de cuentas
        If (nInsertAfterItemId = nParentRuleGroupId) And (bSaveAsChild) Then
          nNewItemPosition = GetRulePosition(sAppServer, nInsertAfterItemId) + 1
        Else
          nNewItemPosition = GetLastChildPosition(sAppServer, nInsertAfterItemId) + 1
        End If
      End If
    Else
      If bSaveAsChild Then
        nNewItemPosition = GetRulePosition(sAppServer, nParentRuleGroupId) + 1
      ElseIf Not bSaveAsChild Then
        nNewItemPosition = GetRulePosition(sAppServer, GetParentRuleId(sAppServer, nParentRuleGroupId)) + 1
      End If
    End If
    ReorderRuleItems sAppServer, oRulesRS!id_regla_contable, nNewItemPosition
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL1, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      .ActiveConnection = Nothing
      .AddNew
      nItem = NewRecordId(sAppServer, "SEC_ID_GRUPO_CUENTA")
      !id_grupo_cuenta = nItem
      !id_regla_contable = oRulesRS!id_regla_contable
      !tipo_grupo_cuenta = oRulesRS!tipo_grupo_cuenta
      !numero_grupo_cuenta = oRulesRS!numero_grupo_cuenta
      !nombre_grupo_cuenta = oRulesRS!nombre_grupo_cuenta
      !id_entidad_agrupador = oRulesRS!id_entidad_agrupador
      !id_agrupador_origen = oRulesRS!id_agrupador_origen
      !cuenta_origen_inicial = ""
      !cuenta_origen_final = ""
      !filtro_cuentas_origen = oRulesRS!filtro_cuentas_origen
      !auxiliar_origen = ""
      !id_sector_origen = 0
      !id_moneda_origen = 0
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
      
      .CursorLocation = adUseClient
      .Open sSQL2, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      .ActiveConnection = Nothing
      .AddNew
      !id_grupo_cuenta = nItem
      !ID_CUENTA_DESTINO = 0
      !ID_AUXILIAR_DESTINO = 0
      !ID_SECTOR_DESTINO = 0
      !ID_CUENTA_DESTINO_SOBREGIRO = 0
      !ID_AUXILIAR_DESTINO_SOBREGIRO = 0
      !ID_SECTOR_DESTINO_SOBREGIRO = 0
      !operador = oRulesRS!operador
      !FACTOR = oRulesRS!FACTOR
      !posicion = nNewItemPosition
      If (bSaveAsChild) Then
        !nivel = GetRuleLevel(sAppServer, nParentRuleGroupId) + 1
        !id_grupo_cuenta_padre = nParentRuleGroupId
      Else
        !nivel = GetRuleLevel(sAppServer, nParentRuleGroupId)
        !id_grupo_cuenta_padre = GetParentRuleId(sAppServer, nParentRuleGroupId)
      End If
      !FECHA_INICIO = #12/31/2000#
      !FECHA_FIN = #12/31/2049#
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddRuleGroup", Err.Number
End Sub

Public Sub AddVoucherRule(sAppServer As String, oRulesRS As Recordset, _
                          sTargetAccount As String, sTargetSubsidiaryAccount As String, _
                          sTargetOBAccount, sTargetOBSubsidiaryAccount As String)
  Dim oRecordset As New Recordset, sSQL1 As String, sSQL2 As String, sTemp As String
  Dim nStdAccountTypeId As Long, nNewItemOrder As Long, nItem As Long
  '**************************************************************************************
  On Error GoTo ErrHandler
    sSQL1 = "SELECT * FROM COF_Grupo_Cuenta WHERE (id_grupo_cuenta = -1)"
    sSQL2 = "SELECT * FROM COF_Grupo_Cuenta_BIS WHERE (id_grupo_cuenta = -1)"
    'nNewItemOrder = 1
    'ReorderRuleItems sAppServer, oRulesRS!id_regla_contable, nNewItemOrder
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL1, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      .ActiveConnection = Nothing
      .AddNew
      nItem = NewRecordId(sAppServer, "SEC_ID_GRUPO_CUENTA")
      !id_grupo_cuenta = nItem
      !id_regla_contable = oRulesRS!id_regla_contable
      !tipo_grupo_cuenta = 4 'oRulesRS!tipo_grupo_cuenta
      !numero_grupo_cuenta = oRulesRS!numero_grupo_cuenta
      !nombre_grupo_cuenta = oRulesRS!nombre_grupo_cuenta
      !id_entidad_agrupador = oRulesRS!id_entidad_agrupador
      !id_agrupador_origen = oRulesRS!id_agrupador_origen
      !cuenta_origen_inicial = oRulesRS!cuenta_origen_inicial
      !cuenta_origen_final = oRulesRS!cuenta_origen_final
      !filtro_cuentas_origen = oRulesRS!filtro_cuentas_origen
      !auxiliar_origen = oRulesRS!auxiliar_origen
      !id_sector_origen = oRulesRS!id_sector_origen
      !id_moneda_origen = oRulesRS!id_moneda_origen
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
      
      .CursorLocation = adUseClient
      .Open sSQL2, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      .ActiveConnection = Nothing
      .AddNew
      !id_grupo_cuenta = nItem
      nStdAccountTypeId = 1
      !ID_CUENTA_DESTINO = GetStdAccountId(sAppServer, nStdAccountTypeId, sTargetAccount)
      !ID_AUXILIAR_DESTINO = 0 'GetSubsidiaryAccountId(sAppServer, 9, sTargetAccount)
      !auxiliar_destino = oRulesRS!auxiliar_destino
      !ID_SECTOR_DESTINO = oRulesRS!ID_SECTOR_DESTINO
      !ID_CUENTA_DESTINO_SOBREGIRO = 0 'GetStdAccountId(sAppServer, nStdAccountTypeId, sTargetAccount)
      !ID_AUXILIAR_DESTINO_SOBREGIRO = 0 'GetSubsidiaryAccountId(sAppServer, 9, sTargetAccount)
      !ID_SECTOR_DESTINO_SOBREGIRO = oRulesRS!ID_SECTOR_DESTINO_SOBREGIRO
      !operador = oRulesRS!operador
      !FACTOR = oRulesRS!FACTOR
      !posicion = 0 'nNewItemOrder
      !nivel = 1
      !id_grupo_cuenta_padre = 0
      !FECHA_INICIO = #12/31/2000#
      !FECHA_FIN = #12/31/2049#
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddVoucherRule", Err.Number
End Sub

Public Function CboClips(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim sTemp As String
  '***************************************************************************************
  On Error GoTo ErrHandler
    sTemp = sTemp & "<OPTION value=0>No calificar</OPTION>" & vbCrLf & _
                    "<OPTION value=1>Moneda nacional</OPTION>" & vbCrLf & _
                    "<OPTION value=2>Divisas</OPTION>" & vbCrLf & _
                    "<OPTION value=3>Unidades de inversión</OPTION>" & vbCrLf
    If nSelectedItemId <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItemId & ">", _
                             "<OPTION SELECTED value=" & nSelectedItemId & ">")
    End If
    CboClips = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboClips", Err.Number
End Function

Public Function CboOperators(sAppServer As String, Optional sSelectedItem As String) As String
  Dim sTemp As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    sTemp = "<OPTION value='+'>Suma</OPTION>" & vbCrLf
    sTemp = sTemp & "<OPTION value='-'>Resta</OPTION>" & vbCrLf
    sTemp = sTemp & "<OPTION value='@'>Valor absoluto</OPTION>" & vbCrLf
    If Len(sSelectedItem) <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value='" & sSelectedItem & "'>", _
                             "<OPTION SELECTED value='" & sSelectedItem & "'>")
    End If
    CboOperators = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboOperators", Err.Number
End Function

Public Function CboRuleChilds(sAppServer As String, nRuleId As Long, _
                              Optional nSelectedItem As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String, sItem As String
  '********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.* FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_grupo_cuenta_padre = " & nRuleId & ") AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) " & _
           "ORDER BY posicion"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If (!tipo_grupo_cuenta = 0) Or (!tipo_grupo_cuenta = 5) Then
          sItem = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " ") & !nombre_grupo_cuenta
        Else
          sItem = !cuenta_origen_inicial & _
                  IIf(IsNull(!cuenta_origen_final), "", " / " & !cuenta_origen_final)
        End If
        sTemp = sTemp & "<OPTION value=" & !id_grupo_cuenta & ">" & sItem & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
    End With
    If nSelectedItem <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItem & ">", _
                             "<OPTION SELECTED value=" & nSelectedItem & ">")
    End If
    CboRuleChilds = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboRuleChilds", Err.Number
End Function

Public Function CboRuleSiblings(sAppServer As String, nRuleDefId As Long, nRuleId As Long, _
                                Optional nSelectedItem As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String, sItem As String, nParentId As Long
  '***************************************************************************************************
  On Error GoTo ErrHandler
    nParentId = GetParentRuleId(sAppServer, nRuleId)
    sSQL = "SELECT COF_Grupo_Cuenta.* FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_grupo_cuenta_padre = " & nParentId & ") AND " & _
           "(id_regla_contable = " & nRuleDefId & ") AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) " & _
           "ORDER BY posicion"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
          If (!tipo_grupo_cuenta = 0) Or (!tipo_grupo_cuenta = 5) Then
            sItem = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " ") & !nombre_grupo_cuenta
          Else
            sItem = !cuenta_origen_inicial & _
                    IIf(IsNull(!cuenta_origen_final), "", " / " & !cuenta_origen_final)
          End If
          sTemp = sTemp & "<OPTION value=" & !id_grupo_cuenta & ">" & sItem & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
    End With
    If nSelectedItem <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItem & ">", _
                             "<OPTION SELECTED value=" & nSelectedItem & ">")
    End If
    CboRuleSiblings = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboRuleSiblings", Err.Number
End Function

Public Sub DeleteRule(sAppServer As String, nRuleId As Long)
  On Error GoTo ErrHandler
    ExecuteCommand sAppServer, cnDelRule, nRuleId
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DeleteRule", Err.Number
End Sub

Public Sub DeleteRuleGroup(sAppServer As String, nRuleGroupId As Long)
  Dim oRecordset As New Recordset, sSQL As String, nRuleId As Long
  '*******************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_grupo_cuenta FROM COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_grupo_cuenta_padre = " & nRuleGroupId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        ExecuteCommand sAppServer, cnDelRule, !id_grupo_cuenta
        .MoveNext
      Loop
    End With
    ExecuteCommand sAppServer, cnDelRule, nRuleGroupId
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DeleteRuleGroup", Err.Number
End Sub

Public Function GetParentRuleId(sAppServer As String, nChildRuleId As Long) As Long
  On Error GoTo ErrHandler
    GetParentRuleId = GetDataValue(sAppServer, cnQryRuleParentId, nChildRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetParentRuleId", Err.Number
End Function

Public Function RuleChildsCount(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    RuleChildsCount = GetDataValue(sAppServer, cnQryRuleChildsCount, nRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleChildsCount", Err.Number
End Function

Public Function RuleChildsType(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    RuleChildsType = GetDataValue(sAppServer, cnQryRuleChildsType, nRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleChildsType", Err.Number
End Function

Public Function RuleLabel(sAppServer As String, nRuleId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '***********************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Grupo_Cuenta " & _
           "WHERE (id_grupo_cuenta = " & nRuleId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If (!tipo_grupo_cuenta = 0) Or (!tipo_grupo_cuenta = 5) Then
          sTemp = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " ") & !nombre_grupo_cuenta
        Else
          sTemp = !cuenta_origen_inicial & " - " & _
                  IIf(IsNull(!cuenta_origen_final), !cuenta_origen_inicial, !cuenta_origen_final)
        End If
      Else
        sTemp = "No encontré la cuenta o rango seleccionados"
      End If
      .Close
    End With
    RuleLabel = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleLabel", Err.Number
End Function

Public Function RuleRS(sAppServer As String, nRuleId As Long, Optional bReadOnly As Boolean = True) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String, sHTML As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.* " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (COF_Grupo_Cuenta.id_grupo_cuenta = " & nRuleId & ") AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta)"
    With oRecordset
      .CursorLocation = adUseClient
      If bReadOnly Then
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Else
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      End If
      Set .ActiveConnection = Nothing
      If nRuleId = 0 And (Not bReadOnly) Then
        .AddNew
      End If
    End With
    Set RuleRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleRS", Err.Number
End Function

Public Function RuleType(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    RuleType = GetDataValue(sAppServer, cnQryRuleType, nRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleType", Err.Number
End Function

Public Sub SaveRule(sAppServer As String, oRulesRS As Recordset, nRuleId As Long)
  Dim oRecordset As New Recordset, sSQL1 As String, sSQL2 As String, sTemp As String, sHTML As String
  '****************************************************************************************************
  On Error GoTo ErrHandler
    sSQL1 = "SELECT * FROM COF_Grupo_Cuenta WHERE (id_grupo_cuenta = " & nRuleId & ")"
    sSQL2 = "SELECT * FROM COF_Grupo_Cuenta_BIS WHERE (id_grupo_cuenta = " & nRuleId & ")"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL1, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      .ActiveConnection = Nothing
      !id_grupo_cuenta = oRulesRS!id_grupo_cuenta
      !id_regla_contable = oRulesRS!id_regla_contable
      !tipo_grupo_cuenta = oRulesRS!tipo_grupo_cuenta
      !numero_grupo_cuenta = oRulesRS!numero_grupo_cuenta
      !nombre_grupo_cuenta = oRulesRS!nombre_grupo_cuenta
      !id_entidad_agrupador = oRulesRS!id_entidad_agrupador
      !id_agrupador_origen = oRulesRS!id_agrupador_origen
      !cuenta_origen_inicial = oRulesRS!cuenta_origen_inicial
      !cuenta_origen_final = oRulesRS!cuenta_origen_final
      !filtro_cuentas_origen = oRulesRS!filtro_cuentas_origen
      !auxiliar_origen = oRulesRS!auxiliar_origen
      !id_sector_origen = oRulesRS!id_sector_origen
      !id_moneda_origen = oRulesRS!id_moneda_origen
      !clip_cuenta_origen = oRulesRS!clip_cuenta_origen
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
      
      .CursorLocation = adUseClient
      .Open sSQL2, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      .ActiveConnection = Nothing
      !id_grupo_cuenta = oRulesRS!id_grupo_cuenta
      !ID_CUENTA_DESTINO = 7367 'oRulesRS!ID_CUENTA_DESTINO
      !ID_AUXILIAR_DESTINO = oRulesRS!ID_AUXILIAR_DESTINO
      !auxiliar_destino = oRulesRS!auxiliar_destino
      !ID_SECTOR_DESTINO = oRulesRS!ID_SECTOR_DESTINO
      !ID_CUENTA_DESTINO_SOBREGIRO = oRulesRS!ID_CUENTA_DESTINO_SOBREGIRO
      !ID_AUXILIAR_DESTINO_SOBREGIRO = oRulesRS!ID_AUXILIAR_DESTINO_SOBREGIRO
      !ID_SECTOR_DESTINO_SOBREGIRO = oRulesRS!ID_SECTOR_DESTINO_SOBREGIRO
      !operador = oRulesRS!operador
      !FACTOR = oRulesRS!FACTOR
      !posicion = oRulesRS!posicion
      !nivel = oRulesRS!nivel
      !id_grupo_cuenta_padre = oRulesRS!id_grupo_cuenta_padre
      !FECHA_INICIO = oRulesRS!FECHA_INICIO
      !FECHA_FIN = oRulesRS!FECHA_FIN
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveRule", Err.Number
End Sub

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function GetLastChildPosition(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    GetLastChildPosition = GetDataValue(sAppServer, cnQryLastChildPosition, nRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetLastChildPosition", Err.Number
End Function

Private Function GetRuleLevel(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    If nRuleId <> 0 Then
      GetRuleLevel = GetDataValue(sAppServer, cnQryRuleLevel, nRuleId)
    Else
      GetRuleLevel = 1
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRuleLevel", Err.Number
End Function

Private Function GetRulePosition(sAppServer As String, nRuleId As Long) As Long
  On Error GoTo ErrHandler
    GetRulePosition = GetDataValue(sAppServer, cnQryRulePosition, nRuleId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRulePosition", Err.Number
End Function

Private Sub ReorderRuleItems(sAppServer As String, nRuleId As Long, nStartItemOrder As Long)
  On Error GoTo ErrHandler
    ExecuteCommand sAppServer, cnUpdReorderRuleItems, Array(nRuleId, nStartItemOrder)
  Exit Sub
ErrHandler:
  RaiseError ClassId, "ReorderItems", Err.Number
End Sub

'************************************************************************************************************
'* OOJJOO                MÉTODOS QUE DEBEN SER MANEJADOS EN OTROS COMPONENTES                        OOJJOO *
'************************************************************************************************************

Public Function CboCurrencies(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '********************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT currency_id, currency_name FROM AO_Currencies " & _
           "WHERE (deleted = 0) ORDER BY currency_name"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !currency_id & ">" & !currency_name & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    sTemp = "<OPTION value=0>Todas las monedas</OPTION>" & vbCrLf & sTemp
    If Abs(nSelectedItemId) <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & Abs(nSelectedItemId) & ">", _
                             "<OPTION SELECTED value=" & Abs(nSelectedItemId) & ">")
    End If
    CboCurrencies = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboCurrencies", Err.Number
End Function

Public Function CboSectors(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_sector, clave_sector, nombre_sector FROM COF_Sector " & _
           "ORDER BY clave_sector"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !id_sector & ">(" & !clave_sector & ") " & !nombre_sector & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    sTemp = "<OPTION value=0>Todos los sectores</OPTION>" & vbCrLf & sTemp
    If Abs(nSelectedItemId) <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & Abs(nSelectedItemId) & ">", _
                             "<OPTION SELECTED value=" & Abs(nSelectedItemId) & ">")
    End If
    CboSectors = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboSectors", Err.Number
End Function

Public Function CboTargetSectors(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_sector, clave_sector, nombre_sector FROM COF_Sector " & _
           "ORDER BY clave_sector"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !id_sector & ">(" & !clave_sector & ") " & !nombre_sector & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    sTemp = "<OPTION value=0>Ningún sector</OPTION>" & vbCrLf & sTemp
    If Abs(nSelectedItemId) <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & Abs(nSelectedItemId) & ">", _
                             "<OPTION SELECTED value=" & Abs(nSelectedItemId) & ">")
    End If
    CboTargetSectors = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboTargetSectors", Err.Number
End Function

Public Function GetAccountId(sAppServer As String, sAccountNumber As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (numero_cuenta_estandar = '" & sAccountNumber & "') AND " & _
           "(id_tipo_cuentas_std = 1) AND (fecha_fin = '31/12/2049')"
           
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetAccountId = !id_cuenta_estandar
      Else
        GetAccountId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountId", Err.Number
End Function

Public Function GetStdAccountId(sAppServer As String, nStdAccountTypeId, sAccountNumber As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '***********************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (numero_cuenta_estandar = '" & sAccountNumber & "') AND " & _
           "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
           
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

Public Function GetSubsidiaryAccountId(sAppServer As String, nGralLedgerId As Long, sSubsAccountNumber As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '**********************************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_auxiliar FROM COF_Cuenta_Auxiliar, COF_Mayor_Auxiliar " & _
           "WHERE (COF_Cuenta_Auxiliar.id_mayor_auxiliar = COF_Mayor_Auxiliar.id_mayor_auxiliar) AND " & _
           "(id_mayor = " & nGralLedgerId & ") AND " & _
           "(numero_cuenta_auxiliar = '" & sSubsAccountNumber & "')"
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

Public Function LeafItemsAsRows(sAppServer As String, nRuleDefId As Long, nRuleId As Long, _
                                Optional sOrderBy As String) As String
  Dim nRuleDefType As Long
  '*****************************************************************************************
  nRuleDefType = GetDataValue(sAppServer, cnQryRuleDefType, nRuleDefId)
  On Error GoTo ErrHandler
    Select Case nRuleDefType
      Case 1
        LeafItemsAsRows = LeafItemsAsRowsType1(sAppServer, nRuleDefId, nRuleId, sOrderBy)
      Case 2
        LeafItemsAsRows = LeafItemsAsRowsType2(sAppServer, nRuleId, sOrderBy)
      Case 3
        LeafItemsAsRows = LeafItemsAsRowsType2(sAppServer, nRuleId, sOrderBy)
    End Select
  Exit Function
ErrHandler:
  RaiseError ClassId, "LeafItemsAsRows", Err.Number
End Function

'Private Function RulesTableType1(sAppServer As String, nRuleId As Long, Optional sOrderBy As String)

'
'  Dim oRecordset As New Recordset
'  Dim sSQL As String, sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
'  Dim nLastGroupEntity As Long, nLastGroup As Long
'  On Error GoTo ErrHandler
'    With oRecordset
'
'
'      'If Len(sOrderBy) = 0 Then
'      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, fecha_afectacion, concepto_transaccion"
'      'Else
'      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, " & sOrderBy
'      'End If
'      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
'      Do While Not .EOF
'        If (nLastGroupEntity <> !id_entidad_agrupador) Or (nLastGroup <> !id_agrupador_origen) Then
'          sTemp = Replace(cRow1, "<@ID@>", !id_grupo_cuenta)
'          sTemp = Replace(sTemp, "<@GROUP_NAME@>", GroupName(sAppServer, !id_entidad_agrupador, !id_agrupador_origen))
'          sHTML = sHTML & sTemp
'          nLastGroupEntity = !id_entidad_agrupador
'          nLastGroup = !id_agrupador_origen
'        End If

'        .MoveNext
'      Loop
'      .Close
'    End With
'    RulesTableType1 = sHTML
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "RulesTableType1", Err.Number
'End Function


Private Function LeafItemsAsRowsType1(sAppServer As String, nRuleDefId As Long, nSourceGrouperId As Long, _
                                      Optional sOrderBy As String) As String
  Dim oRecordset As New Recordset, sSQL As String
  Dim sHTML As String, sTemp As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.*, NVL(AO_Currencies.abbrev,'Todas') moneda_origen, " & _
           "NVL(COF_Sector.clave_sector,'Todos') sector_origen, " & _
           "COF_Cuenta_Estandar_Hist.numero_cuenta_estandar cuenta_destino, " & _
           "NVL(COF_Cuenta_Auxiliar.numero_cuenta_auxiliar,'Ninguno') cuenta_auxiliar_destino, " & _
           "NVL(COF_Cuenta_Estandar2.numero_cuenta_estandar, 'La misma') cuenta_sobregiro, " & _
           "NVL(COF_Cuenta_Auxiliar2.numero_cuenta_auxiliar, 'La misma') cuenta_auxiliar_sobregiro " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis, AO_Currencies, COF_Sector, " & _
           "COF_Cuenta_Estandar_Hist, COF_Cuenta_Estandar_Hist COF_Cuenta_Estandar2, " & _
           "COF_Cuenta_Auxiliar, COF_Cuenta_Auxiliar COF_Cuenta_Auxiliar2 " & _
           "WHERE (ABS(COF_Grupo_Cuenta.id_moneda_origen) = AO_Currencies.currency_id (+)) AND " & _
           "(ABS(COF_Grupo_Cuenta.id_sector_origen) = COF_Sector.id_sector (+)) AND " & _
           "(COF_Grupo_Cuenta.id_regla_contable = " & nRuleDefId & ") AND " & _
           "(COF_Grupo_Cuenta.id_entidad_agrupador = " & IIf(nSourceGrouperId < 0, 9, 16) & ") AND " & _
           "(COF_Grupo_Cuenta.id_agrupador_origen = " & Abs(nSourceGrouperId) & ") AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_BIS.id_grupo_cuenta) AND " & _
           "(COF_Grupo_Cuenta_Bis.id_cuenta_destino = COF_Cuenta_Estandar_Hist.id_cuenta_estandar) AND " & _
           "(COF_Grupo_Cuenta_Bis.id_auxiliar_destino = COF_Cuenta_Auxiliar.id_cuenta_auxiliar (+)) AND " & _
           "(COF_Grupo_Cuenta_Bis.id_cuenta_destino_sobregiro = COF_Cuenta_Estandar2.id_cuenta_estandar (+)) AND " & _
           "(COF_Grupo_Cuenta_Bis.id_auxiliar_destino_sobregiro = COF_Cuenta_Auxiliar2.id_cuenta_auxiliar (+)) AND " & _
           "(COF_Cuenta_Estandar_Hist.fecha_inicio <= COF_Grupo_Cuenta_Bis.fecha_inicio) AND " & _
           "(COF_Grupo_Cuenta_Bis.fecha_fin <= COF_Grupo_Cuenta_Bis.fecha_fin)" & _
           "ORDER BY cuenta_origen_inicial, cuenta_origen_final, posicion"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sHTML) <> 0 Then
          sHTML = sHTML & "¦"
        End If
        sTemp = Replace(cnRuleLineType1, "<@ID@>", !id_grupo_cuenta)
        sTemp = Replace(sTemp, "<@TAB@>", Replace(String(1, "X"), "X", " &nbsp;"))
        sTemp = Replace(sTemp, "<@FROM_ACCOUNT@>", IIf(IsNull(!cuenta_origen_inicial), "&nbsp;", !cuenta_origen_inicial))
        sTemp = Replace(sTemp, "<@TO_ACCOUNT@>", IIf(IsNull(!cuenta_origen_final), "&nbsp;", !cuenta_origen_final))
        sTemp = Replace(sTemp, "<@RESTRICTION@>", IIf(IsNull(!filtro_cuentas_origen) Or Len(!filtro_cuentas_origen) = 0, "No", "Sí"))
        sTemp = Replace(sTemp, "<@CURRENCY@>", IIf(!id_moneda_origen >= 0, !moneda_origen, "!= " & !moneda_origen))
        sTemp = Replace(sTemp, "<@SECTOR@>", IIf(!id_sector_origen >= 0, !sector_origen, "!= " & !sector_origen))
        sTemp = Replace(sTemp, "<@SUBS_ACCOUNT@>", IIf(IsNull(!auxiliar_origen), "Todos", !auxiliar_origen))
        sTemp = Replace(sTemp, "<@TARGET_ACCOUNT@>", !cuenta_destino)
        sTemp = Replace(sTemp, "<@TARGET_SUBS_ACCOUNT@>", IIf(IsNull(!auxiliar_destino), "&nbsp;", !auxiliar_destino))
        sTemp = Replace(sTemp, "<@TARGET_OVER_BALANCED_ACCOUNT@>", !cuenta_sobregiro)
        sTemp = Replace(sTemp, "<@TARGET_OVER_BALANCED_SUBS_ACCOUNT@>", !cuenta_auxiliar_sobregiro)
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    LeafItemsAsRowsType1 = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "LeafItemsAsRowsType1", Err.Number
End Function

Private Function LeafItemsAsRowsType2(sAppServer As String, nRuleId As Long, _
                                      Optional sOrderBy As String) As String
  Dim oRecordset As New Recordset, sSQL As String
  Dim sHTML As String, sTemp As String, sLabel As String
  '***************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.*, NVL(AO_Currencies.abbrev,'Todas') moneda_origen, " & _
           "NVL(COF_Sector.clave_sector,'Todos') sector_origen " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis, AO_Currencies, COF_Sector " & _
           "WHERE (ABS(COF_Grupo_Cuenta.id_moneda_origen) = AO_Currencies.currency_id (+)) AND " & _
           "(ABS(COF_Grupo_Cuenta.id_sector_origen) = COF_Sector.id_sector (+)) AND " & _
           "(COF_Grupo_Cuenta_Bis.id_grupo_cuenta_padre = " & nRuleId & ") AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta)" & _
           "ORDER BY posicion"
    With oRecordset
      'If Len(sOrderBy) = 0 Then
      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, fecha_afectacion, concepto_transaccion"
      'Else
      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, " & sOrderBy
      'End If
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sHTML) <> 0 Then
          sHTML = sHTML & "¦"
        End If
        sLabel = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " - ") & _
                 IIf(IsNull(!nombre_grupo_cuenta), "", !nombre_grupo_cuenta)
        If Len(sLabel) = 0 Then
          sLabel = "&nbsp;"
        End If
        If (!tipo_grupo_cuenta = 5) Then
          sTemp = Replace(cnRuleLineTypeRef, "<@ID@>", !id_grupo_cuenta)
          sTemp = Replace(sTemp, "<@TAB@>", Replace(String(!nivel - 1, "X"), "X", " &nbsp;"))
          sTemp = Replace(sTemp, "<@LABEL@>", sLabel)
          sTemp = Replace(sTemp, "<@OPERATOR@>", !operador)
        Else
          sTemp = Replace(cnRuleLineType2, "<@ID@>", !id_grupo_cuenta)
          sTemp = Replace(sTemp, "<@TAB@>", Replace(String(!nivel - 1, "X"), "X", " &nbsp;"))
          sTemp = Replace(sTemp, "<@LABEL@>", sLabel)
          sTemp = Replace(sTemp, "<@FROM_ACCOUNT@>", IIf(IsNull(!cuenta_origen_inicial), "&nbsp;", !cuenta_origen_inicial))
          sTemp = Replace(sTemp, "<@TO_ACCOUNT@>", IIf(IsNull(!cuenta_origen_final), "&nbsp;", !cuenta_origen_final))
          sTemp = Replace(sTemp, "<@CURRENCY@>", IIf(!id_moneda_origen >= 0, !moneda_origen, "!= " & !moneda_origen))
          sTemp = Replace(sTemp, "<@SECTOR@>", IIf(!id_sector_origen >= 0, !sector_origen, "!= " & !sector_origen))
          sTemp = Replace(sTemp, "<@SUBS_ACCOUNT@>", IIf(IsNull(!auxiliar_origen), "&nbsp;", !auxiliar_origen))
          sTemp = Replace(sTemp, "<@OPERATOR@>", !operador)
        End If
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    LeafItemsAsRowsType2 = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "LeafItemsAsRowsType2", Err.Number
End Function

Private Function RulesTableType2(sAppServer As String, nRuleId As Long, Optional sOrderBy As String)
  Dim oRecordset As New Recordset, sSQL As String, bLastItemWasRuleItem As Boolean
  Dim sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
  '*************************************************************************************************
  On Error GoTo ErrHandler

ErrHandler:
  RaiseError ClassId, "RulesTableType2", Err.Number
End Function

Public Function StdAccountNumber(sAppServer As String, nStdAccountId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '***********************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT numero_cuenta_estandar FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (id_cuenta_estandar = '" & nStdAccountId & "')"
           
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        StdAccountNumber = !numero_cuenta_estandar
      Else
        StdAccountNumber = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountNumber", Err.Number
End Function
