VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** La Aldea Ontica 1.0 ************************************************************************************
'*                                                                                                          *
'* Solución   : La Aldea Ontica®                                Sistema : Accounting                        *
'* Componente : AOGLReports (AOGLReports)                       Parte   : User Services                     *
'* Clase      : CTextFile                                       Patrón  : Standard Class                    *
'* Fecha      : 31/Agosto/2000                                  Versión : 1.0.1                             *
'*                                                                                                          *
'* Descripción: Módulo principal del componente "AOGLReports".                                              *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2001. **
Option Explicit
Private Const ClassId As String = "CTextFile"

Private mFileName As String
Private mFilePath As String
Private mFileNumber As Long

'************************************************************************************************************
'* PROPIEDADES PÚBLICAS                                                                                     *
'************************************************************************************************************

Public Property Get FileName() As String
  On Error GoTo ErrHandler
    FileName = mFileName
  Exit Property
ErrHandler:
  RaiseError ClassId, "FileName", Err.Number
End Property

Public Property Get FilePath() As String
  On Error GoTo ErrHandler
    FilePath = mFilePath
  Exit Property
ErrHandler:
  RaiseError ClassId, "FilePath", Err.Number
End Property

Public Property Get IsOpen() As Boolean
  On Error GoTo ErrHandler
    IsOpen = (mFileNumber <> 0)
  Exit Property
ErrHandler:
  RaiseError ClassId, "IsOpen", Err.Number
End Property

'************************************************************************************************************
'* MÉTODOS PÚBLICOS MANEJADORES DE LA CREACIÓN, APERTURA Y CERRADO DEL ARCHIVO                              *
'************************************************************************************************************

Public Sub Close_()
  On Error GoTo ErrHandler
    If IsOpen Then
      Close #mFileNumber
      mFileNumber = 0
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Close_", Err.Number
End Sub

Public Function GetNextLine() As String
  Dim sLine As String
  On Error GoTo ErrHandler
    Input #mFileNumber, sLine
    GetNextLine = sLine
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetNextLine", Err.Number
End Function

Public Function EOF() As Boolean
  On Error GoTo ErrHandler
    EOF = VBA.EOF(mFileNumber)
  Exit Function
ErrHandler:
  RaiseError ClassId, "EOF", Err.Number
End Function

Public Function GenerateRandomFile() As String
  On Error GoTo ErrHandler
    GenerateRandomFile = CreateRandom(GetConstant(cnGeneratedFilesPath), _
                                      GetConstant(cnDocumentNamePrefix), _
                                      GetConstant(cnMaxNumberOfDocs), _
                                      GetConstant(cnDocumentExtension))
  Exit Function
ErrHandler:
  RaiseError ClassId, "GenerateRandomFile", Err.Number
End Function

Public Sub Open_(Optional sFilePath As String, Optional sFileName As String, Optional bReadOnly = False)
  On Error GoTo ErrHandler
    If IsOpen Then
      Close_
    End If
    If Len(sFilePath) <> 0 Then
      mFilePath = sFilePath
    End If
    If Len(sFileName) <> 0 Then
      mFileName = sFileName
    End If
    mFileNumber = FreeFile
    If Not bReadOnly Then
      Open mFilePath & "\" & mFileName For Output Lock Write As #mFileNumber
    Else
      Open mFilePath & "\" & mFileName For Input As #mFileNumber
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Open_", Err.Number
End Sub

'************************************************************************************************************
'* MÉTODOS PÚBLICOS DE ESCRITURA EN EL ARCHIVO                                                              *
'************************************************************************************************************

Public Sub PageFeed(nLine As Long)
  On Error GoTo ErrHandler
    Do
      If (nLine > cnPageLines) Then
        Exit Do
      End If
      PrintNewLine
      nLine = nLine + 1
    Loop
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PageFeed", Err.Number
End Sub

Public Sub PrintLine(sTextLine As String)
  On Error GoTo ErrHandler
    Print #mFileNumber, sTextLine
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintLine", Err.Number
End Sub

Public Sub PrintNewLine()
  On Error GoTo ErrHandler
    Print #mFileNumber, ""
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PrintNewLine", Err.Number
End Sub

'************************************************************************************************************
'* MÉTODOS PRIVADOS MANEJADORES DEL ARCHIVO                                                                 *
'************************************************************************************************************

Private Function CreateRandom(sFilePath As String, sFilePrefix As String, _
                              nMaxNumberOfDocs As Long, sFileExtension As String) As String
  Dim sFileName As String, nFileNumber As Long
  '********************************************************************************************
  On Error GoTo ErrHandler
    nFileNumber = FreeFile
    mFilePath = sFilePath
    Do
      sFileName = CreateRandomFileName(sFilePrefix, nMaxNumberOfDocs, sFileExtension)
      If Not ExistsFile(mFilePath & "\" & sFileName) Then
        Open mFilePath & "\" & sFileName For Output Lock Write As #nFileNumber
        Close #nFileNumber
        mFileName = sFileName
        CreateRandom = mFileName
        Exit Do
      End If
    Loop
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateRandom", Err.Number
End Function

Private Function CreateRandomFileName(sFilePrefix As String, nMaxNumberOfDocs As Long, _
                                      sExtension As String) As String
  Dim sTemp As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    Randomize
    sTemp = CStr(Int((nMaxNumberOfDocs * Rnd) + 1))
    CreateRandomFileName = sFilePrefix & sTemp & "." & sExtension
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateRandomFileName", Err.Number
End Function

Private Function ExistsFile(FileName As String) As Boolean
  On Error GoTo ErrHandler
    ExistsFile = (Len(Dir$(FileName))) > 0
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExistsFile", Err.Number
End Function

Private Sub Class_Initialize()
  On Error Resume Next
  Close_
End Sub
