Attribute VB_Name = "MMain"
'*** La Aldea Ontica 1.0 ************************************************************************************
'*                                                                                                          *
'* Solución   : La Aldea Ontica®                                Sistema : Accounting                        *
'* Componente : Financial Statements (AOGLFinancialSt)          Parte   : User Services                     *
'* Módulo     : MMain                                           Patrón  : Main Module in User Services      *
'* Fecha      : 31/Agosto/2000                                  Versión : 1.0.1                             *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "AOGLFinancialSt".                                          *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2001. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsAldeaComponent As Boolean = True
Private Const cnSystemName As String = "Accounting"
Private Const cnComponentName As String = "ISigro"

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************

Public Enum TEnumConstants
  cnConnectionString = 1
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES DE EXCEPCIÓN                                                               *
'************************************************************************************************************
Private Const cnFirstError = 23181
Private Const cnLastError = 23200

Public Enum TEnumErrors
  ErrConstantNotFound = 23181
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
  With Err
    Set oException = CreateObject("AOExceptionsMgr.CException")
    If IsAppErr(ErrNumber) Then
      .Description = ErrorDescription(ErrNumber, .Description)
      If Not IsMissing(ErrPars) Then
        .Description = oException.ErrDescriptionParser(.Description, ErrPars)
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
  RaiseError ClassId, "GetConstant", ErrConstantNotFound, cConstantId
End Function

Private Function FillConstantsCol() As Collection
  Const cRegistryManagerPID As String = "AORegistryMgr.CRegistry"
  Dim oRegManager As Object
  'Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '**************************************************************
  On Error GoTo ErrHandler
    'Set oContext = GetObjectContext
    'bObjectContextOK = Not (oContext Is Nothing)
    'If bObjectContextOK Then
    '  Set oRegManager = oContext.CreateInstance(cRegistryManagerPID)
    'Else
      Set oRegManager = CreateObject(cRegistryManagerPID)
    'End If
    If cnIsAldeaComponent Then
      Set FillConstantsCol = oRegManager.ReadKeysForAldeaApp(cnSystemName, cnComponentName)
    Else
      Set FillConstantsCol = oRegManager.ReadKeysForOnticaApp(cnSystemName, cnComponentName)
    End If
    Set oRegManager = Nothing
    'If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  'If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "FillConstantsCol", Err.Number
End Function


