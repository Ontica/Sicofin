Attribute VB_Name = "MMain"
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : GEM And PyC Interface (SCFIGemPyC)              Parte   : Business services                 *
'* Módulo     : MMain                                           Patrón  : Standard Main Module              *
'* Fecha      : 31/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "Customer Components: GEM And PyC Interfaces".              *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsEmpiriaComponent As Boolean = False
Private Const cnSystemName As String = "Customer Components\Banobras\Financial Accounting"
Private Const cnComponentName As String = "GEM And PyC Interface"

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************

Public Enum TEnumConstants
  cnXMLQueryFilePath = 900
  cnXMLQueryFileName = 901
  cnGeneratedFilesPath = 902
  cnURLFilesPath = 903
  cnDocumentNamePrefix = 904
  cnDocumentExtension = 905
  cnMaxNumberOfDocs = 906
  cnUsersDirectoryPath = 907
  cnFTPUsersDirectoryPath = 908
  cnFTPServer = 909
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES DE EXCEPCIÓN                                                               *
'************************************************************************************************************
Private Const cnFirstError = 30060
Private Const cnLastError = 30079

Public Enum TEnumErrors
  ErrConstantNotFound = 30060
  ErrDataConstantNotFound = 30061
  ErrDataSourceWithoutPars = 30062
  ErrFileExist = 30063
  ErrFileNotExist = 30064
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE MÉTODOS EN API'S DE WIN32                                                                 *
'************************************************************************************************************

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
   
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
  Dim colConstants As Collection
  '*********************************************************************
  On Error GoTo ErrHandler
    Set colConstants = FillConstantsCol()
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

Public Sub WriteConstant(sConstantName As String, vConstantValue As Variant)
  Dim oRegManager As Object
  '*************************************************************************
  On Error GoTo ErrHandler
    Set oRegManager = CreateObject("ECERegistryMgr.CRegistry")
    If cnIsEmpiriaComponent Then
      oRegManager.WriteEmpiriaAppKey cnSystemName, cnComponentName, sConstantName, vConstantValue
    Else
      oRegManager.WriteOnticaAppKey cnSystemName, cnComponentName, sConstantName, vConstantValue
    End If
    Set oRegManager = Nothing
  Exit Sub
ErrHandler:
  RaiseError ClassId, "WriteConstant", Err.Number
End Sub

'************************************************************************************************************
'* OTROS MÉTODOS PÚBLICOS DEL COMPONENTE                                                                    *
'************************************************************************************************************
