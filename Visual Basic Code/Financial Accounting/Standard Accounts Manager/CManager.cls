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
'* Componente : Standard Account Manager (EFAStdActBS)          Parte   : Business services                 *
'* Clase      : CManager                                        Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase raíz del componente "Standard Account Manager".                                       *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CManager"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function GetCategoryRS(sAppServer As String, Optional nCategoryId As Long = 0) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean, oRecordset As Recordset
  '***********************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set oRecordset = GetRecordsetDef(sAppServer, cnDefCategory, nCategoryId)
    If nCategoryId = 0 Then
      oRecordset.AddNew
    End If
    Set GetCategoryRS = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetCategoryRS", Err.Number
End Function

Public Sub DeleteStandardAccountCategory(sAppServer As String, nCategoryId As Long)
'  Dim oContext As ObjectContext, bObjectContextOK As Boolean, nResult As Long
'  '****************************************************************************
'  On Error GoTo ErrHandler
'    nResult = ExecuteCommand(sAppServer, cnDelStdAccountCategory, nCategoryId)
'    If nResult > 0 Then
'
'    ElseIf nResult = 0 Then
'      'RaiseError ClassId, "DeleteGeneralLedger", TEnumErrors.ErrGeneralLedgerNotExist, ncat
'    ElseIf nResult < 0 Then
'      'RaiseError ClassId, "DeleteGeneralLedger", TEnumErrors.ErrGeneralLedgerInUse
'    End If
'  Exit Sub
'ErrHandler:
'  RaiseError ClassId, "DeleteGeneralLedger", Err.Number
End Sub

Public Sub SaveStandardAccountCategory(sAppServer As String, oRecordset As Recordset, nCategoryId As Long, _
                                       sPattern As String)
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oStandardAccount As New CStdAccount
  '**************************************************************************
  On Error GoTo ErrHandler
    If (nCategoryId = 0) Then
      oRecordset!parent_id = 4
      oRecordset!is_item = 1
      oRecordset!deleted = 0
      oRecordset!Position = 4
      AppendRecordset sAppServer, oRecordset, "Category_Id", "SEC_CATEGORY_ID"
      oStandardAccount.SetPattern oRecordset!Category_Id, sPattern
    Else
      oRecordset!parent_id = 4
      oRecordset!is_item = 1
      oRecordset!deleted = 0
      oRecordset!Position = 4
      SaveRecordset sAppServer, oRecordset
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveStandardAccountCategory", Err.Number
End Sub


