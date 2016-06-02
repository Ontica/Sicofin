VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Manager BS (EFARulesMgrBS)                Parte   : Business services                 *
'* Clase      : CManager                                        Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase principal del componente.                                                             *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CManager"
Private Const gcMaxDate As Date = #12/31/2049#

Private Const cnRuleDefHeaderColumns As Long = 3
Private Const cnRuleDefHeader = _
        "<TD nowrap><img src='/empiria/images/expanded.gif' onclick='outline();' alt='Expandir todo'></TD>" & _
        "<TD nowrap width=90%><A href='' onclick='return orderBy(1);'>Regla contabilizadora</A><@MARK_1@></TD> " & _
        "<TD nowrap><A href='' onclick='return orderBy(2);'>Versión</A><@MARK_2@></TD> " & _
        "<TD nowrap><A href='' onclick='return orderBy(3);'>Período de aplicación</A><@MARK_3@></TD> " & _
        "<TD nowrap>Opciones</TD>"

Private Const cnRuleDefCollapsedDivLine = _
          "<TBODY value=1><TR class=applicationTableRowDivisor valign=center>" & _
          "<TD nowrap colspan=5>" & _
          "<img align=absbottom src='/empiria/images/collapsed.gif' onclick='outliner();'>" & _
          "<@NAME@> (<@COUNTER@>)</TD></TR>" & vbCrLf
 
Private Const cnRuleDefExpandedDivLine = _
          "<TBODY value=1><TR class=applicationTableRowDivisor>" & _
          "<TD nowrap colspan=5 valign=top>" & _
          "<img align=absbottom src='/empiria/images/expanded.gif' onclick='outliner();'>" & _
          "<@NAME@> (<@COUNTER@>)</TD></TR>" & vbCrLf
                    
Private Const cnRuleDefItemLine = _
          "<TR valign=center style='display:none;'>" & _
          "<TD nowrap> &nbsp; </TD>" & _
          "<TD nowrap><A href='rules.asp?id=<@ID@>'><@RULE_NAME@></A></TD>" & _
          "<TD nowrap><@VERSION@></TD>" & _
          "<TD nowrap><@LIFE_TIME@></TD>" & _
          "<TD nowrap><A href='edit_rule_def.asp?id=<@ID@>'>Editar encabezado</A><@NEW_VERSION@></TD></TR>"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function RuleDefinitionsBody(sAppServer As String, nUserId As Long, _
                                    Optional nOrderColumn As Long) As String
  Dim oRecordset As Recordset, sTemp As String, sHTML As String, sDate As String
  Dim nLastAccountTypeId As Long, nCounter As Long
  '*****************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = RuleDefinitionsRS(sAppServer, nUserId, nOrderColumn)
    nLastAccountTypeId = -1
    With oRecordset
      Do While Not .EOF
        If (!id_tipo_cuentas_std <> nLastAccountTypeId) Then
          If nCounter <> 0 Then
            sHTML = Replace(sHTML, "<@COUNTER@>", nCounter) & "</TBODY>"
          End If
          sTemp = Replace(cnRuleDefExpandedDivLine, "<@NAME@>", _
                  IIf(IsNull(!tipo_contabilidad), "Todas las contabilidades", !tipo_contabilidad))
          sHTML = sHTML & sTemp
          nLastAccountTypeId = !id_tipo_cuentas_std
          nCounter = 0
        End If
        sTemp = Replace(cnRuleDefItemLine, "<@RULE_NAME@>", !nombre_regla_contable)
        If (!version_regla = 0) Then
          sTemp = Replace(sTemp, "<@VERSION@>", "Única")
          sTemp = Replace(sTemp, "<@NEW_VERSION@>", "")
        Else
          sTemp = Replace(sTemp, "<@VERSION@>", Format$(!version_regla, "#.0"))
          sTemp = Replace(sTemp, "<@NEW_VERSION@>", " / <A href='new_version.asp?id=<@ID@>'>Nueva versión</A>")
        End If
        If !FECHA_FIN = gcMaxDate Then
          sDate = "Desde el " & Format$(!FECHA_INICIO, "dd/mmm/yyyy")
        Else
          sDate = "Del " & Format$(!FECHA_INICIO, "dd/mmm/yyyy") & " al " & Format$(!FECHA_FIN, "dd/mmm/yyyy")
        End If
        sTemp = Replace(sTemp, "<@LIFE_TIME@>", sDate)
        sTemp = Replace(sTemp, "<@ID@>", !id_regla_contable)
        sHTML = sHTML & sTemp
        nCounter = nCounter + 1
        .MoveNext
      Loop
      sHTML = Replace(sHTML, "<@COUNTER@>", nCounter) & "</TBODY>"
      .Close
    End With
    RuleDefinitionsBody = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefinitionsBody", Err.Number
End Function

Public Function RuleDefinitionsHeader(sAppServer As String, Optional nSelectedColumn As Long) As String
  Dim sTemp As String, i As Long
  '****************************************************************************************************
  On Error GoTo ErrHandler
    sTemp = cnRuleDefHeader
    If (nSelectedColumn > 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & nSelectedColumn & "@>", " <img src='/empiria/images/order_asc_red.gif'>")
    ElseIf (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, "<@MARK_" & Abs(nSelectedColumn) & "@>", _
                              " <img src='/empiria/images/order_desc_red.gif'>")
    End If
    For i = 1 To cnRuleDefHeaderColumns
      sTemp = Replace(sTemp, "<@MARK_" & i & "@>", "")
    Next i
    RuleDefinitionsHeader = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefinitionsHeader", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Property Get RulesDefinitionsColOrderBy(nSelectedColumn As Long) As String
  Dim sTemp As String
  '*******************************************************************************
  On Error GoTo ErrHandler
    Select Case Abs(nSelectedColumn)
      Case 1
        sTemp = "nombre_regla_contable, version_regla DESC"
      Case 2
        sTemp = "version_regla, nombre_regla_contable"
      Case 3
        sTemp = "fecha_inicio, nombre_regla_contable, version_regla DESC"
      Case Else
        sTemp = "nombre_regla_contable, version_regla DESC"
    End Select
    If (nSelectedColumn < 0) Then
      sTemp = Replace(sTemp, ",", " DESC,")
    End If
    If Right$(sTemp, 1) = "," Then
      sTemp = Left$(sTemp, Len(sTemp) - 1)
    End If
    RulesDefinitionsColOrderBy = sTemp
  Exit Property
ErrHandler:
  RaiseError ClassId, "RulesDefinitionsColOrderBy", Err.Number
End Property

Private Function RuleDefinitionsRS(sAppServer As String, nUserId As Long, nOrderColumn As Long) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Regla_contable.*, object_name tipo_contabilidad " & _
           "FROM COF_Regla_Contable, AO_Objects " & _
           "WHERE (COF_Regla_Contable.id_tipo_cuentas_std = AO_Objects.Object_id (+)) " & _
           "ORDER BY id_tipo_cuentas_std, "
    sSQL = sSQL & RulesDefinitionsColOrderBy(nOrderColumn)
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set RuleDefinitionsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "RuleDefinitionsRS", Err.Number
End Function
