Attribute VB_Name = "MMain"
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                                 Sistema : Financial Accounting          *
'* Componente : Vouchers Text File Interface (SCFIVouchersTextFile) Parte   : Business services             *
'* Módulo     : MMain                                               Patrón  : Business Services Main Module *
'* Fecha      : 28/Febrero/2002                                     Versión : 2.0       Versión patrón: 1.0 *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "Vouchers Text File Interface".                             *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "MMain"

Private Const cnIsEmpiriaComponent As Boolean = False
Private Const cnSystemName As String = "Customer Components\Banobras\Financial Accounting"
Private Const cnComponentName As String = "Vouchers Text File Interface"

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES Y OBJETOS GLOBALES A TODO EL COMPONENTE                                    *
'************************************************************************************************************

Public Enum TEnumConstants
  cnGeneratedFilesPath = 2000
  cnURLFilesPath = 2001
  cnDocumentNamePrefix = 2002
  cnDocumentExtension = 2003
  cnMaxNumberOfDocs = 2004
  cnUsersDirectoryPath = 2005
  cnFTPUsersDirectoryPath = 2006
  cnFTPServer = 2007
End Enum

'************************************************************************************************************
'* DECLARACIÓN DE LAS CONSTANTES DE EXCEPCIÓN                                                               *
'************************************************************************************************************
Private Const cnFirstError = 12200
Private Const cnLastError = 12249


Public Enum TEnumErrors
  ErrConstantNotFound = 12200
  ErrDataConstantNotFound = 12201
  ErrDataSourceWithoutPars = 12202
  ErrFileExist = 12203
  ErrFileNotExist = 12204
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

Public Sub QuickSort(StringArray() As String, Optional vFirst As Variant, Optional vLast As Variant)
  Dim iFirst As Long, iLast As Long
  '*************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(vFirst) Then iFirst = LBound(StringArray) Else iFirst = vFirst
    If IsMissing(vLast) Then iLast = UBound(StringArray) Else iLast = vLast
    
    If iFirst < iLast Then
      'Only two elements in this subdivision; exchange if they are out of order, and end recursive calls
      If iLast - iFirst = 1 Then
        If Compare(StringArray(iFirst), StringArray(iLast)) > 0 Then
            Swap StringArray(iFirst), StringArray(iLast)
        End If
      Else
        Dim iLo As Long, iHi As Long
        'Pick pivot element at random and move to end
        Swap StringArray(iLast), StringArray((iFirst + iLast) \ 2)
        iLo = iFirst: iHi = iLast
        Do
          'Move in from both sides toward pivot element
           Do While (iLo < iHi) And Compare(StringArray(iLo), StringArray(iLast)) <= 0
            iLo = iLo + 1
           Loop
           Do While (iHi > iLo) And Compare(StringArray(iHi), StringArray(iLast)) >= 0
            iHi = iHi - 1
           Loop
           ' If you haven’t reached pivot element, it means that two elements on either side are out of
           'order, so swap them
           If iLo < iHi Then
             Swap StringArray(iLo), StringArray(iHi)
           End If
        Loop While iLo < iHi
        ' Move pivot element back to its proper place
        Swap StringArray(iLo), StringArray(iLast)
  
        'Recursively call SortArrayRec (pass smaller subdivision first to use less stack space)
        If (iLo - iFirst) < (iLast - iLo) Then
          QuickSort StringArray(), iFirst, iLo - 1
          QuickSort StringArray(), iLo + 1, iLast
        Else
          QuickSort StringArray(), iLo + 1, iLast
          QuickSort StringArray(), iFirst, iLo - 1
        End If
      End If
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "QuickSort", Err.Number
End Sub

Private Function Compare(Item1 As Variant, Item2 As Variant) As Integer
  If Item1 < Item2 Then
    Compare = -1
  ElseIf Item1 = Item2 Then
    Compare = 0
  ElseIf Item1 > Item2 Then
    Compare = 1
  End If
End Function

Private Sub Swap(ByRef Item1 As Variant, ByRef Item2 As Variant)
  Dim Temp As String
  
  Temp = Item1
  Item1 = Item2
  Item2 = Temp
End Sub
