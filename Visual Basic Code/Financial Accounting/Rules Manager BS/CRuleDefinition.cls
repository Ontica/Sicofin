VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRuleDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Manager BS (EFARulesMgrBS)                Parte   : Business services                 *
'* Clase      : CRuleDefinition                                 Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Representa una definición de regla del experto contable.                                    *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CRuleDefinition"

Private Const cnMaxNumberOfColumns = 10
Private Const cHeadType1 = _
        "<TD nowrap><img src='/empiria/images/expanded.gif' onclick='outline();' alt='Expandir todo'></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(1);'>Cuenta inicial</A><@MARK_1@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(2);'>Cuenta final</A><@MARK_2@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Auxiliar</A><@MARK_3@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(4);'>Moneda</A><@MARK_4@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(5);'>Sector</A><@MARK_5@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(6);'>Restricción</A><@MARK_6@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(7);'>Cta destino</A><@MARK_7@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(8);'>Aux destino</A><@MARK_8@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(9);'>Cta sobregiro</A><@MARK_9@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(10);'>Aux sobregiro</A><@MARK_10@></TD>" & vbCrLf
        
Private Const cHeadType2 = _
        "<TD nowrap><img src='/empiria/images/expanded.gif' onclick='outline();' alt='Expandir todo'></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(1);'>Grupo</A><@MARK_1@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(2);'>Cuenta inicial</A><@MARK_2@> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Cuenta final</A><@MARK_3@> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(4);'>Sector</A><@MARK_4@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(5);'>Mon</A><@MARK_5@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(6);'>Auxiliar</A><@MARK_6@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(7);'>Op</A><@MARK_7@></TD>" & vbCrLf

Private Const cHeadType3 = _
        "<TD nowrap><img src='/empiria/images/expanded.gif' onclick='outline();' alt='Expandir todo'></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(1);'>Cuenta inicial &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </A><@MARK_1@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(2);'>Cuenta final  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; </A><@MARK_2@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Auxiliar</A><@MARK_3@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(4);'>Moneda</A><@MARK_4@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(5);'>Sector</A><@MARK_5@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(6);'>Restricción</A><@MARK_6@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(7);'>Cuenta destino</A><@MARK_7@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(8);'>Auxiliar destino</A><@MARK_8@></TD>" & _
        "<TD nowrap><A href='' onclick='return orderBy(9);'>Operación</A><@MARK_9@></TD>" & vbCrLf

Private Const cnRuleDefCollapsedDivLineType1 = _
          "<TBODY <@GROUP_ID@> value=<@LEVEL@>><TR class=applicationTableRowDivisor valign=center>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@TAB@><img align=absbottom src='/empiria/images/collapsed.gif' onclick='outliner();'>" & _
          "<A href='' onclick='return callEditor(3, <@ID@>);'><@LABEL@></A></TD></TR></TBODY>" & vbCrLf
 
Private Const cnRuleDefExpandedDivLineType1 = _
          "<TBODY <@GROUP_ID@> value=<@LEVEL@> style='display:<@DISPLAY@>;'><TR class=applicationTableRowDivisor>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@TAB@><img align=absbottom src='/empiria/images/expanded.gif' onclick='outliner();'>" & _
          "<A href='' onclick='return callEditor(3, <@ID@>);'><@LABEL@></TD></TR></TBODY>" & vbCrLf
          
Private Const cnRuleDefCollapsedDivLineType2 = _
          "<TBODY <@GROUP_ID@> value=<@LEVEL@>><TR class=applicationTableRowDivisor valign=center>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@TAB@><img align=absbottom src='/empiria/images/collapsed.gif' onclick='outliner();'>" & _
          "<A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></A></TD>" & _
          "<TD align=center><@OPERATOR@></TD></TR></TBODY>" & vbCrLf
 
Private Const cnRuleDefExpandedDivLineType2 = _
          "<TBODY <@GROUP_ID@> value=<@LEVEL@> style='display:<@DISPLAY@>;'><TR class=applicationTableRowDivisor>" & _
          "<TD nowrap colspan=<@COLSPAN@>><@TAB@><img align=absbottom src='/empiria/images/expanded.gif' onclick='outliner();'>" & _
          "<A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></TD>" & _
          "<TD><@OPERATOR@></TD></TR></TBODY>" & vbCrLf

Private Const cnLeafsSection = _
          "<TBODY value=<@LEVEL@> style='display:none;'></TBODY>" & vbCrLf

  Const cRow1 = "<TR><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></TD>" & vbCrLf & _
                "<TD valign=top nowrap><@OPERATOR@></TD></TR>" & vbCrLf
  Const cRow2 = "<TR><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></A></TD>" & vbCrLf & _
                "<TD><@OPERATOR@></TD></TR>" & vbCrLf
  Const cRow3 = "<TR><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></A></TD>" & vbCrLf & _
                "<TD><@OPERATOR@></TD></TR>" & vbCrLf
          
Private Const cnRuleLineType1 = _
          "<TR valign=center style='display:none;'>" & _
          "<TD nowrap> &nbsp; </TD>" & _
          "<TD nowrap><A href='rules.asp?id=<@ID@>'><@RULE_NAME@></A></TD>" & _
          "<TD nowrap><@VERSION@></TD>" & _
          "<TD nowrap><@LIFE_TIME@></TD>" & _
          "<TD nowrap><A href='edit_rule_def.asp?id=<@ID@>'>Editar encabezado</A><@NEW_VERSION@></TD></TR>" & vbCrLf

Private Const cnRuleLineType2 = _
  "<TR>" & _
  "<TD>&nbsp; &nbsp; &nbsp; &nbsp;</TD>" & _
  "<TD><A href='' onclick='return callEditor(1, <@ID@>);'><@LABEL@></A></TD>" & _
  "<TD nowrap><A href='' onclick='return callEditor(1, <@ID@>);'><@FROM_ACCOUNT@></A></TD>" & _
  "<TD nowrap><@TO_ACCOUNT@></TD>" & _
  "<TD nowrap><@SUBS_ACCOUNT@></TD>" & _
  "<TD nowrap><@CURRENCY@></TD>" & _
  "<TD nowrap><@SECTOR@></TD>" & _
  "<TD><@RESTRICTION@></TD>" & _
  "<TD><@OPERATOR@></TD></TR>" & vbCrLf
  
Private Const cnRuleLineType3 = ""

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function CboRules(sAppServer As String, nRuleDefId As Long, Optional nSelectedItem As Long, _
                         Optional nNotIncludeItem As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sNotIncludeBranch As String
  Dim sTemp As String, sItem As String
  '*************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Grupo_Cuenta.id_grupo_cuenta, numero_grupo_cuenta, nombre_grupo_cuenta, posicion " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_regla_contable = " & nRuleDefId & ") AND (tipo_grupo_cuenta = 0) AND " & _
           "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) "
    If nNotIncludeItem <> 0 Then
      sNotIncludeBranch = BranchItemsIds(sAppServer, nNotIncludeItem, True, True)
      sSQL = sSQL & "AND (COF_Grupo_Cuenta.id_grupo_cuenta NOT IN (" & sNotIncludeBranch & ") ) "
    End If
    sSQL = sSQL & "ORDER BY posicion"
    
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sItem = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " ") & !nombre_grupo_cuenta
        sTemp = sTemp & "<OPTION value=" & !id_grupo_cuenta & ">" & sItem & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
    End With
    If nSelectedItem <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItem & ">", _
                             "<OPTION SELECTED value=" & nSelectedItem & ">")
    End If
    CboRules = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboRules", Err.Number
End Function

Public Function CboGLCategories(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '**********************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT object_id, object_name " & _
           "FROM AO_Objects " & _
           "WHERE (entity_id = 1) " & _
           "ORDER BY object_name"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !object_id & ">" & !object_name & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    sTemp = sTemp & "<OPTION value=0>Todas las contabilidades</OPTION>" & vbCrLf
    If Abs(nSelectedItemId) <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & Abs(nSelectedItemId) & ">", _
                             "<OPTION SELECTED value=" & Abs(nSelectedItemId) & ">")
    End If
    CboGLCategories = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboGLCategories", Err.Number
End Function

Public Function CboMethods(sAppServer As String, Optional nSelectedItemId As Long) As String
  Dim oRecordset As New Recordset
  Dim sSQL As String, sTemp As String
  '**************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT object_id, object_name " & _
           "FROM AO_Objects " & _
           "WHERE (entity_id = 17) " & _
           "ORDER BY object_name"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !object_id & ">" & !object_name & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    If nSelectedItemId <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItemId & ">", _
                             "<OPTION SELECTED value=" & nSelectedItemId & ">")
    End If
    CboMethods = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboMethods", Err.Number
End Function

Public Function Header(sAppServer As String, nRuleTypeId As Long, Optional nSelectedColumn As Long) As String
  Dim sTemp As String, i As Long
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    Select Case nRuleTypeId
      Case 1
        sTemp = cHeadType1
      Case 2
        sTemp = cHeadType2
      Case 3
        sTemp = cHeadType3
    End Select
    If (nSelectedColumn > 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & nSelectedColumn & "@>", _
                      " <img src='/empiria/images/order_asc_red.gif'>")
    ElseIf (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & Abs(nSelectedColumn) & "@>", _
                      " <img src='/empiria/images/order_desc_red.gif'>")
    End If
    For i = 1 To cnMaxNumberOfColumns
      sTemp = Replace(sTemp, "<@MARK_" & i & "@>", "")
    Next i
    Header = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "Header", Err.Number
End Function

Public Function RuleDefName(sAppServer As String, nRuleDefId As Long) As String
  On Error GoTo ErrHandler
    RuleDefName = GetDataValue(sAppServer, cnQryRuleDefName, nRuleDefId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefName", Err.Number
End Function

Public Function RuleDefStdAccountTypeId(sAppServer As String, nRuleDefId As Long) As Long
  On Error GoTo ErrHandler
    RuleDefStdAccountTypeId = GetDataValue(sAppServer, cnQryRuleDefStdAccountTypeId, nRuleDefId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefStdAccountTypeId", Err.Number
End Function

Public Function RuleDefRS(sAppServer As String, nRuleId As Long) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sHTML As String
  '***********************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Regla_Contable WHERE (id_regla_contable = " & nRuleId & ")"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set RuleDefRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefRS", Err.Number
End Function

Public Function RuleDefType(sAppServer As String, nRuleDefId As Long) As Long
  On Error GoTo ErrHandler
    RuleDefType = GetDataValue(sAppServer, cnQryRuleDefType, nRuleDefId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefType", Err.Number
End Function

Public Function RulesTable(sAppServer As String, nRuleDefId As Long, Optional sOrderBy As String) As String
  On Error GoTo ErrHandler
    Select Case RuleDefType(sAppServer, nRuleDefId)
      Case 1
        RulesTable = RulesTableType1(sAppServer, nRuleDefId, sOrderBy)
      Case 2
        RulesTable = RulesTableType2(sAppServer, nRuleDefId, sOrderBy)
      Case 3
        RulesTable = RulesTableType3(sAppServer, nRuleDefId, sOrderBy)
    End Select
  Exit Function
ErrHandler:
  RaiseError ClassId, "RulesTable", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function BranchItemsIds(sAppServer As String, nItem As Long, _
                                Optional bIncludeSiblings As Boolean = False, _
                                Optional bIncludeItem As Boolean = False) As String
  Dim oRule As New CRule, nTempItem As Long, sTemp As String, sTempSilings As String
  '*********************************************************************************
  On Error GoTo ErrHandler
    If bIncludeItem Then
      sTemp = nItem
    End If
    nTempItem = oRule.GetParentRuleId(sAppServer, nItem)
    Do
      If nTempItem <> 0 Then
        If Len(sTemp) <> 0 Then
          sTemp = sTemp & ", " & nTempItem
        Else
          sTemp = nTempItem
        End If
        nTempItem = oRule.GetParentRuleId(sAppServer, nTempItem)
      Else
        Exit Do
      End If
    Loop
    If bIncludeSiblings Then
      sTempSilings = SiblingsIds(sAppServer, nItem, False)
      If Len(sTempSilings) <> 0 Then
        sTempSilings = sTemp & ", " & sTempSilings
      End If
    End If
    BranchItemsIds = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "BranchItemsIds", Err.Number
End Function

Private Function GroupName(sAppServer As String, nGroupEntityId As Long, nGroupId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************************************
  On Error GoTo ErrHandler
    If nGroupEntityId = 1 Then
      sSQL = "SELECT object_name group_name FROM AO_Objects " & _
             "WHERE (object_id = " & nGroupId & ")"
    ElseIf nGroupEntityId = 9 Then
      sSQL = "SELECT nombre_mayor group_name FROM COF_Mayor " & _
             "WHERE (id_mayor = " & nGroupId & ")"
    ElseIf nGroupEntityId = 16 Then
      sSQL = "SELECT nombre_grupo_mayor group_name FROM COF_Grupo_Mayor " & _
             "WHERE (id_grupo_mayor = " & nGroupId & ")"
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If nGroupEntityId = 1 Then
          GroupName = !group_name & " (Se aplican a todas sus contabilidades)"
        ElseIf nGroupEntityId = 9 Then
          GroupName = !group_name & " (Se aplican sólo a esta contabilidad)"
        ElseIf nGroupEntityId = 16 Then
          GroupName = !group_name & " (Se aplican sólo a las contabilidades de este grupo)"
        End If
      Else
        GroupName = ""
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GroupName", Err.Number
End Function

Private Function RulesTableType1(sAppServer As String, nRuleId As Long, Optional sOrderBy As String) As String
  Dim oRecordset As New Recordset
  Dim sSQL As String, sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
  Dim nLastGroupEntity As Long, nLastGroup As Long
  '************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT DISTINCT COF_Grupo_Mayor.* " & _
           "FROM COF_Grupo_Cuenta, COF_Grupo_Mayor " & _
           "WHERE (COF_Grupo_Cuenta.id_agrupador_origen = COF_Grupo_Mayor.id_grupo_mayor) AND " & _
           "(COF_Grupo_Cuenta.id_entidad_agrupador = 16) AND (COF_Grupo_Mayor.clip_grupo_mayor = 2) AND " & _
           "(COF_Grupo_Cuenta.id_regla_contable = " & nRuleId & ") " & _
           "ORDER BY COF_Grupo_Mayor.orden_presentacion"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sTemp = Replace(cnRuleDefCollapsedDivLineType1, "<@ID@>", "0")
        sTemp = Replace(sTemp, "<@DISPLAY@>", "inline")
        sTemp = Replace(sTemp, "<@GROUP_ID@>", "")
        sTemp = Replace(sTemp, "<@LEVEL@>", 1)
        sTemp = Replace(sTemp, "<@COLSPAN@>", 11)
        sTemp = Replace(sTemp, "<@TAB@>", "")
        sTemp = Replace(sTemp, "<@LABEL@>", "<b>Reglas aplicables a todas las contabilidades incluidas en el grupo</b>")
        sHTML = sHTML & sTemp
      End If
      Do While Not .EOF
        sTemp = Replace(cnRuleDefExpandedDivLineType1, "<@ID@>", !id_grupo_mayor)
        sTemp = Replace(sTemp, "<@DISPLAY@>", "inline")
        sTemp = Replace(sTemp, "<@GROUP_ID@>", "id=" & !id_grupo_mayor)
        sTemp = Replace(sTemp, "<@LEVEL@>", 2)
        sTemp = Replace(sTemp, "<@COLSPAN@>", 11)
        sTemp = Replace(sTemp, "<@TAB@>", " &nbsp;")
        sTemp = Replace(sTemp, "<@LABEL@>", !nombre_grupo_mayor)
        sHTML = sHTML & sTemp & Replace(cnLeafsSection, "<@LEVEL@>", 3)
      .MoveNext
      Loop
      .Close
    End With
    sSQL = "SELECT DISTINCT COF_Mayor.* " & _
           "FROM COF_Grupo_Cuenta, COF_Mayor " & _
           "WHERE (COF_Grupo_Cuenta.id_agrupador_origen = COF_Mayor.id_mayor) AND " & _
           "(COF_Grupo_Cuenta.id_entidad_agrupador = 9) AND " & _
           "(COF_Grupo_Cuenta.id_regla_contable = " & nRuleId & ") " & _
           "ORDER BY COF_Mayor.numero_mayor, COF_Mayor.nombre_mayor"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        sTemp = Replace(cnRuleDefCollapsedDivLineType1, "<@ID@>", 0)
        sTemp = Replace(sTemp, "<@DISPLAY@>", "inline")
        sTemp = Replace(sTemp, "<@GROUP_ID@>", "")
        sTemp = Replace(sTemp, "<@LEVEL@>", 1)
        sTemp = Replace(sTemp, "<@COLSPAN@>", 11)
        sTemp = Replace(sTemp, "<@TAB@>", "")
        sTemp = Replace(sTemp, "<@LABEL@>", "<b>Reglas aplicables por contabilidad</b>")
        sHTML = sHTML & sTemp
      End If
      Do While Not .EOF
        sTemp = Replace(cnRuleDefExpandedDivLineType1, "<@ID@>", (-1 * !id_mayor))
        sTemp = Replace(sTemp, "<@DISPLAY@>", "inline")
        sTemp = Replace(sTemp, "<@GROUP_ID@>", "id=" & (-1 * !id_mayor))
        sTemp = Replace(sTemp, "<@LEVEL@>", 2)
        sTemp = Replace(sTemp, "<@COLSPAN@>", 11)
        sTemp = Replace(sTemp, "<@TAB@>", " &nbsp;")
        sTemp = Replace(sTemp, "<@LABEL@>", "(" & !numero_mayor & ") " & !nombre_mayor)
        sHTML = sHTML & sTemp & Replace(cnLeafsSection, "<@LEVEL@>", 3)
      .MoveNext
      Loop
      .Close
    End With
    RulesTableType1 = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "RulesTableType1", Err.Number
End Function

Private Function RulesTableType2(sAppServer As String, nRuleId As Long, Optional sOrderBy As String) As String
  Dim oRecordset As New Recordset, sSQL As String, nLastLevel As Long, nLastGroupId As Long
  Dim bLastItemWasRuleItem As Boolean
  Dim sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.* " & _
             "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis " & _
             "WHERE (COF_Grupo_Cuenta.id_regla_contable = " & nRuleId & ") AND " & _
             "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta) AND " & _
             "(tipo_grupo_cuenta = 0)" & _
             "ORDER BY posicion"
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sLabel = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " - ") & _
                 IIf(IsNull(!nombre_grupo_cuenta), "", !nombre_grupo_cuenta)
        If Len(sLabel) = 0 Then
          sLabel = "&nbsp;"
        End If
        If (!nivel <= nLastLevel) Then
          sHTML = Replace(sHTML, "<@GROUP_ID@>", "id=" & nLastGroupId) & _
                  Replace(cnLeafsSection, "<@LEVEL@>", nLastLevel + 1)
        Else
          sHTML = Replace(sHTML, "<@GROUP_ID@>", "")
        End If
        sTemp = Replace(cnRuleDefExpandedDivLineType2, "<@ID@>", !id_grupo_cuenta)
        sTemp = Replace(sTemp, "<@DISPLAY@>", IIf(!nivel = 1, "inline", "none"))
        sTemp = Replace(sTemp, "<@LABEL@>", sLabel)
        sTemp = Replace(sTemp, "<@LEVEL@>", !nivel)
        sTemp = Replace(sTemp, "<@TAB@>", Replace(String(!nivel - 1, "X"), "X", " &nbsp; "))
        sTemp = Replace(sTemp, "<@COLSPAN@>", 7)
        sTemp = Replace(sTemp, "<@OPERATOR@>", !operador)
        sHTML = sHTML & sTemp
        nLastLevel = !nivel
        nLastGroupId = !id_grupo_cuenta
        .MoveNext
      Loop
      .Close
    End With
    sHTML = Replace(sHTML, "<@GROUP_ID@>", "id=" & nLastGroupId) & _
            Replace(cnLeafsSection, "<@LEVEL@>", nLastLevel + 1)
    RulesTableType2 = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "RulesTableType2", Err.Number
End Function

Private Function RulesTableType3(sAppServer As String, nRuleId As Long, Optional sOrderBy As String)
  Const cRow1 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(3, <@ID@>);'><FONT face=Arial size=3 color=maroon><b><@LABEL@></b></FONT></A></TD>" & vbCrLf & _
                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf
  Const cRow2 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><b><@LABEL@></b></A></TD>" & vbCrLf & _
                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf
  Const cRow3 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></A></TD>" & vbCrLf & _
                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf

  Dim oRecordset As New Recordset
  Dim sSQL As String, sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
  On Error GoTo ErrHandler
    With oRecordset
      sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.*, NVL(AO_Currencies.abbrev,'Todas') moneda_origen, " & _
             "NVL(COF_Sector.clave_sector,'Todos') sector_origen " & _
             "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis,  AO_Currencies, COF_Sector " & _
             "WHERE (ABS(COF_Grupo_Cuenta.id_moneda_origen) = AO_Currencies.currency_id (+)) AND " & _
             "(ABS(COF_Grupo_Cuenta.id_sector_origen) = COF_Sector.id_sector (+)) AND " & _
             "(COF_Grupo_Cuenta.id_regla_contable = " & nRuleId & ") AND " & _
             "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_BIS.id_grupo_cuenta) " & _
             "ORDER BY posicion"
      'If Len(sOrderBy) = 0 Then
      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, fecha_afectacion, concepto_transaccion"
      'Else
      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, " & sOrderBy
      'End If
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sLabel = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " - ") & _
                 IIf(IsNull(!nombre_grupo_cuenta), "", !nombre_grupo_cuenta)
        If Len(sLabel) = 0 Then
          sLabel = "&nbsp;"
        End If
        If (!tipo_grupo_cuenta = 0) And (!nivel = 1) Then
          sTemp = Replace(cRow1, "<@ID@>", !id_grupo_cuenta)
        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel = 2) Then
          sTemp = Replace(cRow2, "<@ID@>", !id_grupo_cuenta)
        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel = 3) Then
          sTemp = Replace(cRow3, "<@ID@>", !id_grupo_cuenta)
        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel > 3) Then
          sTemp = Replace(cRow3, "<@ID@>", !id_grupo_cuenta)
          sLabel = Replace(Replace("X", "X", String(!nivel - 3, "X")), "X", "&nbsp;&nbsp;") & " " & sLabel
        ElseIf (!tipo_grupo_cuenta <> 0) Then
          sTemp = Replace(cnRuleLineType2, "<@ID@>", !id_grupo_cuenta)
          sTemp = Replace(sTemp, "<@LABEL@>", sLabel)
          sTemp = Replace(sTemp, "<@FROM_ACCOUNT@>", IIf(IsNull(!cuenta_origen_inicial), "&nbsp;", !cuenta_origen_inicial))
          sTemp = Replace(sTemp, "<@TO_ACCOUNT@>", IIf(IsNull(!cuenta_origen_final), "&nbsp;", !cuenta_origen_final))
          sTemp = Replace(sTemp, "<@RESTRICTION@>", IIf(IsNull(!filtro_cuentas_origen), "&nbsp;", !filtro_cuentas_origen))
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
    RulesTableType3 = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "RulesTableType3", Err.Number
End Function

'Private Function RulesTableType3(sAppServer As String, nRuleId As Long, Optional sOrderBy As String)
'  Const cRow1 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(3, <@ID@>);'><FONT face=Arial size=3 color=maroon><b><@LABEL@></b></FONT></A></TD>" & vbCrLf & _
'                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf
'  Const cRow2 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><b><@LABEL@></b></A></TD>" & vbCrLf & _
'                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf
'  Const cRow3 = "<TR bgcolor=khaki><TD colspan=7><A href='' onclick='return callEditor(2, <@ID@>);'><@LABEL@></A></TD>" & vbCrLf & _
'                "<TD valign=top nowrap align=center><@OPERATOR@></TD></TR>" & vbCrLf
'
'  Dim oRecordset As New Recordset
'  Dim sSQL As String, sTemp As String, sHTML As String, sMayorAnterior As String, sLabel As String
'  On Error GoTo ErrHandler
'    With oRecordset
'      sSQL = "SELECT COF_Grupo_Cuenta.*, COF_Grupo_Cuenta_Bis.*, NVL(AO_Currencies.abbrev,'Todas') moneda_origen, " & _
'             "NVL(COF_Sector.clave_sector,'Todos') sector_origen " & _
'             "FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis,  AO_Currencies, COF_Sector " & _
'             "WHERE (ABS(COF_Grupo_Cuenta.id_moneda_origen) = AO_Currencies.currency_id (+)) AND " & _
'             "(ABS(COF_Grupo_Cuenta.id_sector_origen) = COF_Sector.id_sector (+)) AND " & _
'             "(COF_Grupo_Cuenta.id_regla_contable = " & nRuleId & ") AND " & _
'             "(COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_BIS.id_grupo_cuenta) " & _
'             "ORDER BY posicion"
'      'If Len(sOrderBy) = 0 Then
'      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, fecha_afectacion, concepto_transaccion"
'      'Else
'      '  sSQL = sSQL & "COF_Mayor.id_tipo_cuentas_std, numero_mayor, nombre_mayor, " & sOrderBy
'      'End If
'      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
'      Do While Not .EOF
'        sLabel = IIf(IsNull(!numero_grupo_cuenta), "", !numero_grupo_cuenta & " - ") & _
'                 IIf(IsNull(!nombre_grupo_cuenta), "", !nombre_grupo_cuenta)
'        If Len(sLabel) = 0 Then
'          sLabel = "&nbsp;"
'        End If
'        If (!tipo_grupo_cuenta = 0) And (!nivel = 1) Then
'          sTemp = Replace(cRow1, "<@ID@>", !id_grupo_cuenta)
'        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel = 2) Then
'          sTemp = Replace(cRow2, "<@ID@>", !id_grupo_cuenta)
'        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel = 3) Then
'          sTemp = Replace(cRow3, "<@ID@>", !id_grupo_cuenta)
'        ElseIf (!tipo_grupo_cuenta = 0) And (!nivel > 3) Then
'          sTemp = Replace(cRow3, "<@ID@>", !id_grupo_cuenta)
'          sLabel = Replace(Replace("X", "X", String(!nivel - 3, "X")), "X", "&nbsp;&nbsp;") & " " & sLabel
'        ElseIf (!tipo_grupo_cuenta <> 0) Then
'          sTemp = Replace(cnRuleLineType2, "<@ID@>", !id_grupo_cuenta)
'          sTemp = Replace(sTemp, "<@LABEL@>", sLabel)
'          sTemp = Replace(sTemp, "<@FROM_ACCOUNT@>", IIf(IsNull(!cuenta_origen_inicial), "&nbsp;", !cuenta_origen_inicial))
'          sTemp = Replace(sTemp, "<@TO_ACCOUNT@>", IIf(IsNull(!cuenta_origen_final), "&nbsp;", !cuenta_origen_final))
'          sTemp = Replace(sTemp, "<@RESTRICTION@>", IIf(IsNull(!filtro_cuentas_origen), "&nbsp;", !filtro_cuentas_origen))
'          sTemp = Replace(sTemp, "<@CURRENCY@>", IIf(!id_moneda_origen >= 0, !moneda_origen, "!= " & !moneda_origen))
'          sTemp = Replace(sTemp, "<@SECTOR@>", IIf(!id_sector_origen >= 0, !sector_origen, "!= " & !sector_origen))
'          sTemp = Replace(sTemp, "<@SUBS_ACCOUNT@>", IIf(IsNull(!auxiliar_origen), "&nbsp;", !auxiliar_origen))
'          sTemp = Replace(sTemp, "<@OPERATOR@>", !operador)
'        End If
'        sHTML = sHTML & sTemp
'        .MoveNext
'      Loop
'      .Close
'    End With
'    RulesTableType3 = sHTML
'  Exit Function
'ErrHandler:
'  RaiseError ClassId, "RulesTableType3", Err.Number
'End Function


Private Function SiblingsIds(sAppServer As String, nItem As Long, _
                             Optional bIncludeItem As Boolean = False) As String
  Dim oRule As New CRule, oRecordset As New Recordset
  Dim sSQL As String, sTemp As String, nParentId As Long
  '*****************************************************************************
  On Error GoTo ErrHandler
    nParentId = oRule.GetParentRuleId(sAppServer, nItem)
    sSQL = "SELECT id_grupo_cuenta " & _
           "FROM COF_Grupo_Cuenta_Bis " & _
           "WHERE (id_grupo_cuenta_padre = " & nParentId & ")"
    If (Not bIncludeItem) Then
      sSQL = sSQL & " AND (id_grupo_cuenta <> " & nItem & ")"
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sTemp) = 0 Then
          sTemp = !id_grupo_cuenta
        Else
          sTemp = sTemp & ", " & !id_grupo_cuenta
        End If
        .MoveNext
      Loop
      .Close
    End With
    SiblingsIds = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "SiblingsIds", Err.Number
End Function



