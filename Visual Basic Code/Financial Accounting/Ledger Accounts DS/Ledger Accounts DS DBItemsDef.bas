Attribute VB_Name = "MDBItemsDef"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Ledger Accounts DS (EFALedgerAcctsDS)           Parte   : Data services                     *
'* Módulo     : MDBItemsDef                                     Patrón  : Database Items Definition Module  *
'* Fecha      : 31/Diciembre/2001                               Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Proporciona las constantes de acceso a datos del componente y los métodos que las manejan.  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MDBItemsDef"

Public Enum TEnumDataConstant
  cnDefCategory = 3051
  
  cnQryResponsabilityAreas = 3006
  cnQryResponsabilityArea = 3007
  cnQrySectors = 3008
  cnQrySector = 3009
  cnQrySources = 3010
  cnQrySource = 3011
  cnQryStandardAccounts = 3012
  cnQryStandardAccount = 3013
  cnQryGeneralLedgers = 3014
  cnQryGeneralLedger = 3015
  cnQrySectorStandardAccounts = 3016
  cnQryStdAccountSectors = 3017
  
  cnQrySubsidiaryLedgers = 3018
  cnQrySubsidiaryLedger = 3019
  cnQrySubsidiaryAccounts = 3020
  cnQrySubsidiaryAccount = 3021
  cnQryMaxSubryAccountNumber = 3022
  cnQryAccount = 3023
  cnQryAccountSectors = 3024
  cnQryExistsSubsidiaryAccount = 3025
  
  cnDefAddress = 3042
  cnDefBudgetStruct = 3043
  cnDefOrganization = 3044
  cnDefResponsabilityArea = 3045
  cnDefSector = 3046
  cnDefSource = 3047
  cnDefStandardAccount = 3048
  cnDefGeneralLedger = 3049
  cnDefStandardAccountHist = 3050
  cnDefEntity = 3052
  
  cnDefSubsidiaryLedger = 3055
  cnDefSubsidiaryAccount = 3056
  cnDefSubsidiaryAccountHist = 3057
  cnDefAccount = 3058
  cnDefSubsidiaryAccountMapping = 3059
  
  cnDelAddress = 3062
  cnDelBudgetStruct = 3063
  cnDelOrganization = 3064
  cnDelResponsabilityArea = 3065
  cnDelSector = 3066
  cnDelSource = 3067
  cnDelStandardAccount = 3068
  cnDelGeneralLedger = 3069
  cnDelStdAccountCatalog = 3070
  
  cnDelSubsidiaryLedger = 3075
  cnDelSubsidiaryAccount = 3076
  cnDelEntity = 3077
  
  cnValBudgetStruct = 3082
  cnValSource = 3083
  cnValGeneralLedger = 3084
  cnValStandardAccount = 3085
  cnValAddressForUpd = 3086
  cnValBudgetStructForUpd = 3087
  cnValOrganizationForUpd = 3088
  cnValResponsabilityAreaForUpd = 3089
  cnValSectorForUpd = 3090
  cnValSourceForUpd = 3091
  cnValStandardAccountForUpd = 3092
  cnValGeneralLedgerForUpd = 3093

  cnValSubsidiaryLedgerForUpd = 3094
  cnValSubsidiaryAccount = 3095
  cnValSubsidiaryAccountForUpd = 3096
  
End Enum

Public Sub GetParameters(oCommand As Command, cDataConstant As TEnumDataConstant, ParsValues As Variant, _
                         Optional bUseOracle As Boolean = False)
  Dim nIndex As Long, c As String
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    nIndex = LBound(ParsValues)
    c = IIf(bUseOracle, "p", "@")
    With oCommand
      Select Case cDataConstant
        Case cnQryStandardAccount
          .Parameters.Append .CreateParameter(c & "_standard_account_id", adInteger, , , ParsValues(nIndex))
        Case cnQryGeneralLedger
          .Parameters.Append .CreateParameter(c & "_general_ledger_id", adInteger, , , ParsValues(nIndex))
        Case cnQrySectorStandardAccounts
          .Parameters.Append .CreateParameter(c & "_sector_id", adInteger, , , ParsValues(nIndex))
        Case cnQryStdAccountSectors
          .Parameters.Append .CreateParameter(c & "_standard_account_id", adInteger, , , ParsValues(nIndex))
         
        Case cnDelStandardAccount
          .Parameters.Append .CreateParameter(c & "_standard_account_id", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter(c & "_delete_date", adDate, , , ParsValues(nIndex + 1))
          .Parameters.Append .CreateParameter(c & "_result", adInteger, adParamOutput)
          
        Case cnValStandardAccount
          .Parameters.Append .CreateParameter("p_std_account_category_id", adInteger, , , ParsValues(nIndex))
          .Parameters.Append .CreateParameter("p_balance_group_id", adInteger, , , ParsValues(nIndex + 1))
          .Parameters.Append .CreateParameter("p_result", adInteger, adParamOutput)
         
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
