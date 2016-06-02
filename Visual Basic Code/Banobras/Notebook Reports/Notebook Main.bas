Attribute VB_Name = "MMain"
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Soluci�n   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : Notebook (SCFNotebook)                          Parte   : Business services                 *
'* M�dulo     : MMain                                           Patr�n  : B/D services Main Module          *
'* Fecha      : 15/Enero/2002                                   Versi�n : 1.1       Versi�n patr�n: 1.0     *
'*                                                                                                          *
'* Descripci�n: M�dulo principal del componente "Notebook".                                                 *
'*                                                                                                          *
'****************************************************** Copyright � La V�a Ontica, S.C. M�xico, 2001-2002. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsEmpiriaComponent As Boolean = False
Private Const cnSystemName As String = "Customer Components\Financial Accounting"
Private Const cnComponentName As String = "Notebook"

'************************************************************************************************************
'* DECLARACI�N DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************
Public Enum TEnumConstants
  cnXMLQueryFilePath = 900
  cnXMLQueryFileName = 901
  cnGeneratedFilesPath = 902
  cnPatternsPath = 903
  cnURLFilesPath = 904
  cnNotebookReportsFile = 905
End Enum

Public Enum TEnumDataConstant
  cnQryBeforePeriodBalances = 1000
  cnQryBeforePeriodPostings = 1001
  cnQryInitialBalances = 1002
  cnQryInPeriodBalances = 1003
  cnQryLastLevelsBalances = 1004
  cnQryStdAccountsHist = 1005
  cnQryValorizateToSourceCurrency = 1006
  cnQryValorizateToTargetCurrency = 1007
  cnQryAllLevelsFromTempTable1 = 1008
  cnQryAllLevelsFromTempTable2 = 1009
  cnQryAllLevelsFromTempTable3 = 1010
  cnQryLastLevelsFromTempTable = 1011
  cnQryOutputBalancesSQLStr = 1012
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
'* DECLARACI�N DE LAS CONSTANTES DE EXCEPCI�N                                                               *
'************************************************************************************************************
Private Const cnFirstError = 6000
Private Const cnLastError = 6019

Public Enum TEnumErrors
  ErrConstantNotFound = 6000
  ErrDataConstantNotFound = 6001
End Enum

'************************************************************************************************************
'* M�TODOS P�BLICOS Y PRIVADOS MANEJADORES DE EXCEPCIONES Y ERRORES                                         *
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
    sTemp = "Ocurri� la excepci�n n�mero: &H" & Hex$(cErrNumber Or vbObjectError) & vbCrLf & _
            LastErrDescription
  End If
  sTemp = Replace(sTemp, "\n", vbCrLf)
  ErrorDescription = sTemp
End Function

'************************************************************************************************************
'* M�TODOS MANEJADORES DE LAS CONSTANTES DEL COMPONENTE                                                     *
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
