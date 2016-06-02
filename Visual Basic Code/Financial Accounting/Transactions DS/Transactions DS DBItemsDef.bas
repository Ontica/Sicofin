Attribute VB_Name = "MDBItemsDef"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Transactions DS (EFATransactionsDS)             Parte   : Data services                     *
'* Módulo     : MDBItemsDef                                     Patrón  : Database Items Definition Module  *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Proporciona las constantes de acceso a datos del componente y los métodos que las manejan.  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MDBItemsDef"

Public Enum TEnumDataConstant
  cnQryPendingTransactions = 1000
  cnQryPendingTransactionsWithPostingsSum = 1001
  cnQryPostedTransactions = 1002
  cnQryPostedTransactionsWithPostingsSum = 1003
End Enum

Public Sub GetParameters(oCommand As Command, cDataConstant As TEnumDataConstant, ParsValues As Variant, _
                         Optional bUseOracle As Boolean = False)
  Dim nIndex As Long, c As String
  '********************************************************************************************************
  On Error GoTo ErrHandler
    nIndex = LBound(ParsValues)
    c = IIf(bUseOracle, "p", "@")
    With oCommand
      Select Case cDataConstant
        Case cnQryPendingTransactions
          .Parameters.Append .CreateParameter(c & "TaskId", adInteger, adParamInput, , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "UserId", adInteger, adParamInput, , ParsValues(nIndex + 1))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case Else
          Err.Raise TEnumErrors.ErrDataSourceWithoutPars
      End Select
    End With
  Exit Sub
ErrHandler:
  If Err.Number = TEnumErrors.ErrDataSourceWithoutPars Then
    RaiseError ClassId, "GetParameters", Err.Number, cDataConstant
  Else
    RaiseError ClassId, "GetParameters", Err.Number
  End If
End Sub


