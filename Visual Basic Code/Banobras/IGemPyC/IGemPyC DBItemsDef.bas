Attribute VB_Name = "MDBItemsDef"
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Soluci�n   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : GEM And PyC Interfaces (SCFIGemPyC)             Parte   : Business services                 *
'* M�dulo     : MDBItemsDef                                     Patr�n  : Database Items Definition Module  *
'* Fecha      : 31/Enero/2002                                   Versi�n : 1.1       Versi�n patr�n: 1.0     *
'*                                                                                                          *
'* Descripci�n: Proporciona las constantes de acceso a datos del componente y los m�todos que las manejan.  *
'*                                                                                                          *
'****************************************************** Copyright � La V�a Ontica, S.C. M�xico, 2001-2002. **
Option Explicit
Private Const ClassId As String = "MDBItemsDef"

Public Enum TEnumDataConstant
  cnDelGEMTransaction = 1000
  cnUpdGEMErrTransaction = 1001
End Enum

Public Sub GetParameters(oCommand As Command, cDataConstant As TEnumDataConstant, ParsValues As Variant, _
                         Optional bUseOracle As Boolean = False)
  Dim nIndex As Long, c As String
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    nIndex = LBound(ParsValues)
    c = IIf(bUseOracle, "p_", "@")
    With oCommand
      Select Case cDataConstant
        Case cnDelGEMTransaction
          .Parameters.Append .CreateParameter(c & "EncTipoCont", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "EncFechaVol", adDate, , , ParsValues(nIndex + 1))
          .Parameters.Append .CreateParameter(c & "EncNumVol", adInteger, , , ParsValues(nIndex + 2))
          .Parameters.Append .CreateParameter(c & "result", adInteger, adParamOutput)
        Case cnUpdGEMErrTransaction
          .Parameters.Append .CreateParameter(c & "EncTipoCont", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "EncFechaVol", adDate, , , ParsValues(nIndex + 1))
          .Parameters.Append .CreateParameter(c & "EncNumVol", adInteger, , , ParsValues(nIndex + 2))
          .Parameters.Append .CreateParameter(c & "result", adInteger, adParamOutput)
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


