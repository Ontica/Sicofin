VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ClassId As String = "CFile"

Private mFile As Long

Private mFileName As String
Private mFilePath As String
Private mSeparatorLineChar As String * 1
Private mSeparatorLineLength As Long

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

Public Property Get SeparatorLineChar() As String
 On Error GoTo ErrHandler
    SeparatorLineChar = mSeparatorLineChar
  Exit Property
ErrHandler:
  RaiseError ClassId, "SeparatorLineChar", Err.Number
End Property

Public Property Let SeparatorLineChar(newvalue As String)
 On Error GoTo ErrHandler
    mSeparatorLineChar = Left$(newvalue, 1)
  Exit Property
ErrHandler:
  RaiseError ClassId, "SeparatorLineLength", Err.Number
End Property

Public Property Get SeparatorLineLength() As Long
 On Error GoTo ErrHandler
    SeparatorLineLength = mSeparatorLineLength
  Exit Property
ErrHandler:
  RaiseError ClassId, "SeparatorLineLength", Err.Number
End Property

Public Property Let SeparatorLineLength(newvalue As Long)
 On Error GoTo ErrHandler
    mSeparatorLineLength = newvalue
  Exit Property
ErrHandler:
  RaiseError ClassId, "SeparatorLineLength", Err.Number
End Property


'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Sub Close_()
  On Error GoTo ErrHandler
    Close #mFile
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Close_", Err.Number
End Sub

Public Function CreateRandom(sDocumentPath As String, sDocumentPrefix As String, _
                             nMaxNumberOfDocs As Long, sDocumentExtension As String) As String
  Dim FileNumber As Integer
  Dim sFileName As String
  On Error GoTo ErrHandler
    mFilePath = sDocumentPath
    FileNumber = FreeFile
    Do
      sFileName = CreateRandomFileName(sDocumentPrefix, nMaxNumberOfDocs, sDocumentExtension)
      If Not ExistsFile(mFilePath & "\" & sFileName) Then
        Open mFilePath & "\" & sFileName For Output Lock Write As #FileNumber
        Close #FileNumber
        mFileName = sFileName
        CreateRandom = mFileName
        Exit Do
      End If
    Loop
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateRandom", Err.Number
End Function

Public Sub Open_Write(Optional sFilePath As String, Optional sFileName As String)
  Dim sTemp As String
  On Error GoTo ErrHandler
    If Len(sFilePath) <> 0 Then
      mFilePath = sFilePath
    End If
    If Len(sFileName) <> 0 Then
      mFileName = sFileName
    End If
    mFile = FreeFile
    'Open mFilePath & "\" & mFileName For Output Lock Write As #mFile
    Open mFileName For Output Lock Write As #mFile
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Open_Write", Err.Number
End Sub

Public Function Open_Read(Optional sFilePath As String, Optional sFileName As String) As Long
  Dim sTemp As String
  On Error GoTo ErrHandler
    If Len(sFilePath) <> 0 Then
      mFilePath = sFilePath
    End If
    If Len(sFileName) <> 0 Then
      mFileName = sFileName
    End If
    mFile = FreeFile
    Open mFileName For Input As #mFile
    Open_Read = mFile
  Exit Function
ErrHandler:
  RaiseError ClassId, "Open_Read", Err.Number
End Function

Public Sub WriteNewLineChar()
  On Error GoTo ErrHandler
    Print #mFile, vbCrLf
  Exit Sub
ErrHandler:
  RaiseError ClassId, "WriteNewLineChar", Err.Number
End Sub

Public Sub WriteSeparatorLine(Optional bWriteNewLineCharAfter As Boolean = False)
  On Error GoTo ErrHandler
    If bWriteNewLineCharAfter Then
      Print #mFile, String(mSeparatorLineLength, mSeparatorLineChar) & vbCrLf
    Else
      Print #mFile, String(mSeparatorLineLength, mSeparatorLineChar)
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "WriteSeparatorLine", Err.Number
End Sub

Public Sub WriteText(sText As String, Optional bWriteNewLineCharAfter As Boolean = False)
  On Error GoTo ErrHandler
    If bWriteNewLineCharAfter Then
      Print #mFile, sText & vbCrLf
    Else
      Print #mFile, sText
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "WriteText", Err.Number
End Sub

Public Sub WriteTextLine(sText As String, Optional bWriteNewLineCharAfter As Boolean = False)
  On Error GoTo ErrHandler
    If bWriteNewLineCharAfter Then
      Print #mFile, sText & vbCrLf
    Else
      Print #mFile, sText
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "WriteTextLine", Err.Number
End Sub

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function CreateRandomFileName(sFilePrefix As String, nMaxNumberOfDocs As Long, _
                                      sExtension As String) As String
  Dim sTemp As String
  On Error GoTo ErrHandler
    Randomize
    sTemp = CStr(Int((nMaxNumberOfDocs * Rnd) + 1))
    CreateRandomFileName = sFilePrefix & sTemp & "." & sExtension
  Exit Function
ErrHandler:
  RaiseError ClassId, "CreateRandomFileName", Err.Number
End Function

Friend Function ExistsFile(FileName As String)
  On Error GoTo ErrHandler
    ExistsFile = (Len(Dir$(FileName))) > 0
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExistsFile", Err.Number
End Function



