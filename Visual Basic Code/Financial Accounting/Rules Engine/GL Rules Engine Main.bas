Attribute VB_Name = "MMain"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Rules Engine (EFARulesEngine)                   Parte   : Business services                 *
'* Módulo     : MMain                                           Patrón  : Business services Main Module     *
'* Fecha      : 31/Enero/2001                                   Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "Financial Accounting: Rules Engine".                       *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsEmpiriaComponent As Boolean = True
Private Const cnSystemName As String = "Financial Accounting"
Private Const cnComponentName As String = "Rules Engine"

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************

Public Enum TEnumConstants
  cnVoid = 900
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES DE EXCEPCIÓN                                                               *
'************************************************************************************************************
Private Const cnFirstError = 6060
Private Const cnLastError = 6099

Public Enum TEnumErrors
  ErrConstantNotFound = 6060
  ErrRuleNotFound = 6061
  ErrRuleNotDefinedForThisGL = 6062
  ErrMethodWithoutImplementation = 6063
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

'************************************************************************************************************
'* OTROS MÉTODOS PÚBLICOS DEL COMPONENTE                                                                    *
'************************************************************************************************************

Public Function TrimAll(sString As String) As String
  Dim sTemp As String, iPos As Long
  '*************************************************
  sTemp = Trim$(sString)
  Do
    iPos = InStr(sTemp, Space(2))
    If iPos <> 0 Then
      sTemp = Left$(sTemp, iPos) & Right$(sTemp, Len(sTemp) - (iPos + 1))
    Else
      Exit Do
    End If
  Loop
  Do
    If Left$(sTemp, 2) = vbCrLf Then
      sTemp = Mid$(sTemp, 3)
    ElseIf Right$(sTemp, 2) = vbCrLf Then
      sTemp = Left$(sTemp, Len(sTemp) - 2)
    Else
      Exit Do
    End If
  Loop
  TrimAll = sTemp
End Function
