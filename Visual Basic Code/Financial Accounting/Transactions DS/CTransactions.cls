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
'* Componente : Transactions DS (EFATransactionsDS)             Parte   : Data services                     *
'* Clase      : CTransactions                                   Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Regresa un conjunto de transacciones en ADO.Recordsets o como documento XML.                *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CTransactions"

'************************************************************************************************************
'* ENUMERACIONES PÚBLICAS                                                                                   *
'************************************************************************************************************

Public Enum TEnumVoucherInboxes
  efaMyPendingVouchers = 1
  efaMyPostedVouchers = 2
  efaGroupPendingVouchers = 3
  efaGroupPostedVouchers = 4
End Enum

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function GetPending(sAppServer As String, Optional bIncludePostings As Boolean, _
                            Optional sWhere As String, Optional sOrderBy As String) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    If (Not bIncludePostings) Then
      sSQL = GetQueryString(cnQryPendingTransactions)
    Else
      sSQL = GetQueryString(cnQryPendingTransactionsWithPostingsSum)
    End If
    sSQL = Replace(sSQL, "<@SQL_WHERE@>", IIf(Len(sWhere) <> 0, " AND " & sWhere, ""))
    sSQL = Replace(sSQL, "<@SQL_ORDER_BY@>", IIf(Len(sOrderBy) <> 0, sOrderBy, ""))
    With oRecordset
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CursorLocation = adUseClient
      .Open sSQL, , adOpenForwardOnly, adLockReadOnly
    End With
    Set GetPending = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetPending", Err.Number
End Function

Public Function GetPosted(sAppServer As String, Optional bIncludePostings As Boolean, _
                          Optional sWhere As String, Optional sOrderBy As String) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    If (Not bIncludePostings) Then
      sSQL = GetQueryString(cnQryPostedTransactions)
    Else
      sSQL = GetQueryString(cnQryPostedTransactionsWithPostingsSum)
    End If
    sSQL = Replace(sSQL, "<@SQL_WHERE@>", IIf(Len(sWhere) <> 0, " AND " & sWhere, ""))
    sSQL = Replace(sSQL, "<@SQL_ORDER_BY@>", IIf(Len(sOrderBy) <> 0, sOrderBy, ""))
    With oRecordset
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CursorLocation = adUseClient
      .Open sSQL, , adOpenForwardOnly, adLockReadOnly
    End With
    Set GetPosted = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetPosted", Err.Number
End Function
