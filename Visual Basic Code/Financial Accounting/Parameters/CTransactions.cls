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
'* Componente : Parameters (EFAParameters)                      Parte   : Business services                 *
'* Clase      : CTransactions                                   Patrón  : Stateless class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Regresa los parámetros para el manejo de las transacciones contables.                       *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CTransactions"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function CreateFilter(sAppServer As String, _
                             Optional nUserId As Long, Optional bIncludeWorkGroups As Boolean, _
                             Optional nGralLedgerGroup As Long, Optional nGralLedger As Long, _
                             Optional dFromApplicationDate As String, Optional dToApplicationDate As String, _
                             Optional dFromElaborationDate As String, Optional dToElaborationDate As String, _
                             Optional sVoucherNumber As String, Optional sVoucherConcept As String, _
                             Optional sAccounts As String, Optional nTransactionTypeId As Long, _
                             Optional nVoucherTypeId As Long, Optional nBalancingType As Long) As String
  Dim oQueryBuilder As New ECEQryStrBuilder.CBuilder
  Dim sResult As String, sTemp As String
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    With oQueryBuilder
      If (nGralLedger <> 0) Then
        sTemp = .BuildIdCompareStr("COF_Mayor.id_mayor", nGralLedger)
      ElseIf (nGralLedgerGroup <> 0) And (Not bIncludeWorkGroups) Then
        sTemp = .BuildInListStr("COF_Mayor.id_mayor", GetGLGroupArray(sAppServer, nGralLedgerGroup, ","))
      ElseIf (nGralLedgerGroup <> 0) And bIncludeWorkGroups Then
        sTemp = "(" & .BuildInListStr("COF_Mayor.id_mayor", GetGLGroupArray(sAppServer, nGralLedgerGroup, ",")) & "AND" & _
                .BuildInListStr("COF_Mayor.id_mayor", UserGralLedgers(sAppServer, nUserId)) & ")"
      ElseIf (nGralLedger = 0) And (nGralLedgerGroup = 0) And bIncludeWorkGroups Then
        sTemp = .BuildInListStr("COF_Mayor.id_mayor", UserGralLedgers(sAppServer, nUserId))
      End If
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      If (nUserId <> 0) And (Not bIncludeWorkGroups) Then
        sTemp = "(" & .BuildIdCompareStr("id_elaborada_por", nUserId) & "OR" & _
                .BuildIdCompareStr("id_autorizada_por", nUserId) & ")"
      End If
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
       
      sTemp = .BuildDateIntervalStr("fecha_afectacion", dFromApplicationDate, dToApplicationDate, False)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildDateIntervalStr("fecha_registro", dFromElaborationDate, dToElaborationDate, False)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildLikeStr("numero_transaccion", sVoucherNumber)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildLikeStr("concepto_transaccion", sVoucherConcept)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildInListStr("numero_cuenta_estandar", sAccounts)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildIdCompareStr("id_tipo_transaccion", nTransactionTypeId)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      sTemp = .BuildIdCompareStr("id_tipo_poliza", nVoucherTypeId)
      If (Len(sTemp) <> 0) Then
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
      
      If (nBalancingType = 1) Then
        sTemp = "(suma_cargos <> suma_abonos)"
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      ElseIf (nBalancingType = 2) Then
        sTemp = "(suma_cargos = suma_abonos)"
        sResult = sResult & IIf(Len(sResult) <> 0, "AND", "") & sTemp
      End If
    End With
    CreateFilter = sResult
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateFilter", Err.Number
End Function

Public Function GetGLGroupArray(sAppServer As String, nGralLedgerParentGroupId As Long, _
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
    
    sSQL = "SELECT DISTINCT COF_ELEMENTO_GRUPO_MAYOR.id_mayor, numero_mayor, nombre_mayor " & _
           "FROM COF_Elemento_Grupo_Mayor, COF_MAYOR " & _
           "WHERE (COF_ELEMENTO_GRUPO_MAYOR.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND " & _
           "      (id_grupo_mayor IN (" & sAllGroups & ")) " & _
           "ORDER BY NUMERO_MAYOR, NOMBRE_MAYOR"

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

Private Function UserGralLedgers(sAppServer As String, nUserId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '******************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT objectId FROM MHParticipantObjects " & _
           "WHERE (participantId = " & nUserId & ") AND (entityId = 9)"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sTemp) <> 0 Then
          sTemp = sTemp & ", " & !objectId
        Else
          sTemp = !objectId
        End If
        .MoveNext
      Loop
      .Close
    End With
    UserGralLedgers = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "UserGralLedgers", Err.Number
End Function
