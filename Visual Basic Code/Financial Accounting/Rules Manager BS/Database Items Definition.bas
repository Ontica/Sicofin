Attribute VB_Name = "MDBItemsDef"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Manager BS (EFARulesMgrBS)                Parte   : Business services                 *
'* Módulo     : MDBItemsDef                                     Patrón  : Database Items Definition Module  *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Proporciona las constantes de acceso a datos del componente y los métodos que las manejan.  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MDBItemsDef"

Public Enum TEnumDataConstant
  cnQryRuleDefName = 1000
  cnQryRuleDefStdAccountTypeId = 1001
  cnQryRuleDefType = 1002
  cnQryRuleChildsCount = 1003
  cnQryRuleChildsType = 1004
  cnQryRuleLevel = 1005
  cnQryRuleParentId = 1006
  cnQryRulePosition = 1007
  cnQryLastChildPosition = 1008
  cnQryRuleType = 1009
  cnUpdReorderRuleItems = 1010
  cnDelRule = 1011
End Enum

Public Sub GetParameters(oCommand As Command, cDataConstant As TEnumDataConstant, ParsValues As Variant, _
                         Optional bUseOracle As Boolean = False)
  Dim nIndex As Long, c As String
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    nIndex = LBound(ParsValues)
    c = IIf(bUseOracle, "p", "@")
    With oCommand
      Select Case cDataConstant
        Case cnQryRuleDefName
          .Parameters.Append .CreateParameter(c & "RuleDefId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adVarChar, adParamOutput, 256)
        Case cnQryRuleDefStdAccountTypeId
          .Parameters.Append .CreateParameter(c & "RuleDefId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleDefType
          .Parameters.Append .CreateParameter(c & "RuleDefId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleChildsCount
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleChildsType
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleLevel
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleParentId
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRulePosition
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryLastChildPosition
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnQryRuleType
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "Result", adInteger, adParamOutput)
        Case cnUpdReorderRuleItems
          .Parameters.Append .CreateParameter(c & "RuleDefId", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "StartOrder", adInteger, , , ParsValues(nIndex + 1))
        Case cnDelRule
          .Parameters.Append .CreateParameter(c & "RuleId", adInteger, , , ParsValues(nIndex))
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
