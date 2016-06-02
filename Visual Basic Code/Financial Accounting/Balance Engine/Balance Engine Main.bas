Attribute VB_Name = "MMain"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Balance Engine (EFABalanceEngine)               Parte   : Business services                 *
'* Módulo     : MMain                                           Patrón  : B/D services Main Module          *
'* Fecha      : 28/Febrero/2002                                 Versión : 2.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "Financial Accounting: Balance Engine".                     *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsEmpiriaComponent As Boolean = True
Private Const cnSystemName As String = "Financial Accounting"
Private Const cnComponentName As String = "Balance Engine"

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************
'Quitar esta constante en la versión 2.0 y leerla de la tabla COF_MAYOR
Public Const cnSourceCurrency As Long = 1

Public Enum TEnumConstants
  cnXMLQueryFilePath = 900
  cnXMLQueryFileName = 901
End Enum

Public Enum TEnumDataConstant
  cnQryBeforePeriodBalances = 1000
  cnQryBeforePeriodBalancesWithVouchers = 1028
  cnQryBeforePeriodPostings = 1001
  cnQryInitialBalances = 1002
  cnQryInPeriodBalances = 1003
  cnQryInPeriodBalancesWithVouchers = 1027
  cnQryLastLevelsBalances = 1004
  cnQryLastLevelsBalancesWithVouchers = 1024
  cnQryStdAccountsHist = 1005
  cnQryValorizateToSourceCurrency = 1006
  cnQryValorizateToTargetCurrency = 1007
  cnQryAllLevelsFromTempTable1 = 1008
  cnQryAllLevelsFromTempTable2 = 1009
  cnQryAllLevelsFromTempTable3 = 1010
  cnQryLastLevelsFromTempTable = 1011
  cnQryLastLevelsFromTempTableWithVouchers = 1025
  cnQryOutputBalances = 1012
  cnQryOutputBalancesWithVouchers = 1026
  cnQryStdAccountHistComplete = 1013
  cnQryAvgLastLevelsBalances = 1014
  cnQryAvgBeforePeriodBalances = 1015
  cnQryAvgInPeriodBalances = 1016
  cnQryAvgValorizateToSourceCurrency = 1017
  cnQryAvgValorizateToTargetCurrency = 1018
  cnQryAvgLastLevelsFromTempTable = 1019
  cnQryAvgAllLevelsFromTempTable1 = 1020
  cnQryAvgAllLevelsFromTempTable2 = 1021
  cnQryAvgAllLevelsFromTempTable3 = 1022
  cnQryAvgOutputBalancesSQLStr = 1023
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES DE EXCEPCIÓN                                                               *
'************************************************************************************************************
Private Const cnFirstError = 6000
Private Const cnLastError = 6019

Public Enum TEnumErrors
  ErrConstantNotFound = 6000
  ErrDataConstantNotFound = 6001
End Enum

'************************************************************************************************************
'* MÉTODOS PÚBLICOS Y PRIVADOS MANEJADORES DE EXCEPCIONES Y ERRORES                                         *
'************************************************************************************************************

Private Function IsAppErr(ErrNumber As Long) As Boolean
  IsAppErr = ((cnFirstError <= ErrNumber) And (ErrNumber <= cnLastError))
End Function

Public Sub RaiseError(sClassId As String, sMethod As String, ErrNumber As Long, _
                      Optional ErrPars As Variant, Optional bLogOnly As Boolean = False)
  Dim oException As Object
  '*************************************************************************************
  With Err
    Set oException = CreateObject("ECEExceptionsMgr.CException")
    If IsAppErr(ErrNumber) Then
      .Description = ErrorDescription(ErrNumber, .Description)
      If Not IsMissing(ErrPars) Then
        .Description = oException.ParseErrorDescription(.Description, ErrPars)
      End If
      .Number = vbObjectError Or ErrNumber
    End If
    .Source = cnComponentName & "." & sClassId & "." & sMethod
    If IsAppErr(ErrNumber And (Not vbObjectError)) Then
      oException.DumpErrObject Err
    Else
      oException.ConstructVBError Err, ErrPars, True
    End If
    Set oException = Nothing
    If Not bLogOnly Then
      .Raise .Number, .Source, .Description
    End If
  End With
End Sub

Private Function ErrorDescription(cErrNumber As TEnumErrors, LastErrDescription As String) As String
  Dim sTemp As String
  '*************************************************************************************************
  On Error Resume Next
  sTemp = LoadResString(cErrNumber)
  If Err.Number <> 0 Then
    sTemp = "Ocurrió la excepción número: &H" & Hex$(cErrNumber Or vbObjectError) & vbCrLf & _
            LastErrDescription
  End If
  sTemp = Replace(sTemp, "\n", vbCrLf)
  ErrorDescription = sTemp
End Function

'************************************************************************************************************
'* MÉTODOS MANEJADORES DE LAS CONSTANTES DEL COMPONENTE                                                     *
'************************************************************************************************************

Public Function GetConstant(Optional cConstantId As TEnumConstants, _
                            Optional sConstantName As String) As Variant
  Static colConstants As Collection
  '*********************************************************************
  On Error GoTo ErrHandler
    If colConstants Is Nothing Then
      Set colConstants = FillConstantsCol()
    End If
    If Len(sConstantName) <> 0 Then
      GetConstant = colConstants(sConstantName)
      Exit Function
    Else
     GetConstant = colConstants(LoadResString(cConstantId))
    End If
  Exit Function
ErrHandler:
  GetConstant = Null
  RaiseError ClassId, "GetConstant", TEnumErrors.ErrConstantNotFound, _
             IIf(Len(sConstantName) = 0, cConstantId, sConstantName)
End Function

Private Function FillConstantsCol() As Collection
  Dim oRegManager As Object
  '**********************************************
  On Error GoTo ErrHandler
    Set oRegManager = CreateObject("ECERegistryMgr.CRegistry")
    If cnIsEmpiriaComponent Then
      Set FillConstantsCol = oRegManager.ReadKeysForEmpiriaApp(cnSystemName, cnComponentName)
    Else
      Set FillConstantsCol = oRegManager.ReadKeysForOnticaApp(cnSystemName, cnComponentName)
    End If
    Set oRegManager = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "FillConstantsCol", Err.Number
End Function

