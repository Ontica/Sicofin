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
'*** Empiria� ***********************************************************************************************
'*                                                                                                          *
'* Soluci�n   : Empiria� Software Components                    Sistema : Financial Accounting              *
'* Componente : Ledger Accounts DS (EFALedgerAcctsDS)           Parte   : Data services                     *
'* Clase      : CServer                                         Patr�n  : Stateless Class                   *
'* Fecha      : 31/Diciembre/2001                               Versi�n : 1.0       Versi�n patr�n: 1.0     *
'*                                                                                                          *
'* Descripci�n: Provee los servicios de acceso a datos para las cuentas de mayor general".                  *
'*                                                                                                          *
'****************************************************** Copyright � La V�a Ontica, S.C. M�xico, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CServer"

'************************************************************************************************************
'* M�TODOS P�BLICOS                                                                                         *
'************************************************************************************************************

Public Function LedgerAccountTypesRS(sAppServer As String) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oRecordset As New Recordset
  '********************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set LedgerAccountTypesRS = GetRecordsetWithPars(sAppServer, cnQryApplicationWithTaskId, nTaskId)
    If bObjectContextOK Then oContext.SetComplete
    
'sSQL = "SELECT * " & _
        "FROM AOObjects " & _
        "WHERE (entity_id = 1) AND (deleted = 0) " & _
        "ORDER BY presentation_order"
        
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "LedgerAccountTypesRS", Err.Number
End Function

        

'************************************************************************************************************
'* M�TODOS PRIVADOS                                                                                         *
'************************************************************************************************************
