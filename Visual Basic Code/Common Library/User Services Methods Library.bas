Attribute VB_Name = "MUSMethodsLibrary"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Biblioteca : MUSMethodsLibrary                               Parte   : User Services                     *
'* Fecha      : 31/Diciembre/2001                               Versión : 1.0                               *
'*                                                                                                          *
'* Descripción: Esta biblioteca proporciona los servicios del user services para aplicaciones Web.          *
'*              Los patrones de la interfaz HTML/DHTML se definen en archivos XML.                          *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId  As String = "MUSMethodsLibrary"

Private Const csXMLUSAttrName As String = "pattern"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function GetComboBox(cUSConstant As TEnumUSConstant, Optional vSelectedItem As Variant) As String
  Dim sComboPattern As String, sHTML As String
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    sHTML = GetXMLUSPattern(cUSConstant)
    If Not IsMissing(vSelectedItem) Then
      If IsNumeric(vSelectedItem) Then
        sHTML = Replace(sHTML, "<OPTION value=" & vSelectedItem & ">", _
                               "<OPTION SELECTED value=" & vSelectedItem & ">")
      Else
        sHTML = Replace(sHTML, "<OPTION value='" & vSelectedItem & "'>", _
                               "<OPTION SELECTED value='" & vSelectedItem & "'>")
      End If
    End If
    GetComboBox = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetComboBox", Err.Number
End Function

Public Function GetComboBoxWithRS(oRecordset As Recordset, sValueField As String, sDisplayField As String, _
                                  Optional vSelectedItem As Variant) As String
  Const cpComboOption As String = "<OPTION value=<@VALUE@>><@DISPLAY@></OPTION>"
  Dim sHTML As String, sTemp As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      If (.BOF And .EOF) Then
        Exit Function
      End If
      .MoveFirst
      Do While Not .EOF
        sTemp = Replace(cpComboOption, "<@VALUE@>", .Fields(sValueField))
        sHTML = sHTML & Replace(sTemp, "<@DISPLAY@>", .Fields(sDisplayField))
        .MoveNext
      Loop
    End With
    If Not IsMissing(vSelectedItem) Then
      If IsNumeric(vSelectedItem) Then
        sHTML = Replace(sHTML, "<OPTION value=" & vSelectedItem & ">", _
                               "<OPTION SELECTED value=" & vSelectedItem & ">")
      Else
        sHTML = Replace(sHTML, "<OPTION value='" & vSelectedItem & "'>", _
                               "<OPTION SELECTED value='" & vSelectedItem & "'>")
      End If
    End If
    GetComboBoxWithRS = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetComboBoxWithRS", Err.Number
End Function

Public Function GetXMLUSPattern(cUSConstant As TEnumUSConstant) As String
  Dim oXMLDOM As Object
  Dim sQueryItemName As String
  '**********************************************************************
  On Error GoTo ErrHandler
    sQueryItemName = GetUserServicesConstant(cUSConstant)
    Set oXMLDOM = CreateObject("ECEXMLDOMParser.CXMLDOM")
    GetXMLUSPattern = oXMLDOM.GetElementTextByName(GetConstant(cnXMLUSFile), _
                                                   "pattern", sQueryItemName, "source")
    Set oXMLDOM = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetXMLUSPattern", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function GetUserServicesConstant(cUSConstant As TEnumUSConstant) As String
  On Error GoTo ErrHandler
    GetUserServicesConstant = LoadResString(cUSConstant)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetUserServicesConstant", ErrUSConstantNotFound, cUSConstant
End Function

