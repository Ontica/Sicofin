Attribute VB_Name = "MDBXMLMethodsLib"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Biblioteca : MDBXMLMethodsLib                                Parte   : Data Services                     *
'* Fecha      : 24/Diciembre/2001                               Versión : 1.0                               *
'*                                                                                                          *
'* Descripción: Biblioteca con los servicios de consultas SQL en archivos XML.                              *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "MDBXMLMethodsLib"

Private Const csOracleIdentificator As String = "MSDAORA"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS PARA LA RECUPERACIÓN DE RECORDSETS                                                      *
'************************************************************************************************************

Public Function GetQueryString(cDataConstant As TEnumDataConstant) As String
  Dim sSQL As String
  '*************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetXMLQueryString(cDataConstant)
    GetQueryString = sSQL
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetQueryString", Err.Number
End Function

Public Function GetQueryStringWithPars(cDataConstant As TEnumDataConstant, vParsValues As Variant) As String
  Dim sSQL As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    If VarType(vParsValues) < vbArray Then
      sSQL = GetXMLQueryString(cDataConstant)
      sSQL = ParseParameters(sSQL, Array(vParsValues))
    Else
      sSQL = GetXMLQueryString(cDataConstant)
      sSQL = ParseParameters(sSQL, vParsValues)
    End If
    GetQueryStringWithPars = sSQL
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetQueryStringWithPars", Err.Number
End Function

Public Function GetQueryStringSQLClauses(cDataConstant As TEnumDataConstant, _
                                         sWhere As String, sOrderBy As String) As String
  Dim sSQL As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetXMLQueryStringWithSQLClauses(cDataConstant, sWhere, sOrderBy)
    GetQueryStringSQLClauses = sSQL
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetQueryStringSQLClauses", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function GetDataConstant(cDataConstant As TEnumDataConstant) As String
  On Error GoTo ErrHandler
    GetDataConstant = LoadResString(cDataConstant)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetDataConstant", ErrDataConstantNotFound, cDataConstant
End Function

Private Function GetXMLQueryString(cDataConstant As TEnumDataConstant) As String
  Dim oXMLDOM As Object
  Dim sXMLDOMFilePath As String, sQueryItemName As String
  '*****************************************************************************
  On Error GoTo ErrHandler
    sQueryItemName = GetDataConstant(cDataConstant)
    sXMLDOMFilePath = GetConstant(cnXMLQueryFilePath) & "\" & GetConstant(cnXMLQueryFileName)
    
    Set oXMLDOM = CreateObject("ECEXMLDOMParser.CXMLDOM")
    GetXMLQueryString = oXMLDOM.GetElementTextByName(sXMLDOMFilePath, "query", sQueryItemName, "source")
    Set oXMLDOM = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetXMLQueryString", Err.Number
End Function

Private Function GetXMLQueryStringWithSQLClauses(cDataConstant As TEnumDataConstant, _
                                                 sWhere As String, sOrderBy As String) As String
  Dim oXMLDOM As Object, oXMLDOMNode As Object
  Dim sXMLFilePath As String, sQueryItemName As String, sSQL As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    sQueryItemName = GetDataConstant(cDataConstant)
    sXMLFilePath = GetConstant(cnXMLQueryFilePath) & "\" & GetConstant(cnXMLQueryFileName)
    
    Set oXMLDOM = CreateObject("ECEXMLDOMParser.CXMLDOM")
    Dim sXPath As String
    sXPath = "query[@name = '" & sQueryItemName & "']"
    Set oXMLDOMNode = oXMLDOM.GetNode(sXMLFilePath, sXPath)
    sSQL = oXMLDOM.GetNodeElementText(oXMLDOMNode, "source")
    If Len(sWhere) = 0 Then
      sWhere = oXMLDOM.GetNodeElementText(oXMLDOMNode, "default_where")
    End If
    If Len(sOrderBy) = 0 Then
      sOrderBy = oXMLDOM.GetNodeElementText(oXMLDOMNode, "default_order")
    End If
    GetXMLQueryStringWithSQLClauses = ParseSQLClauses(sSQL, sWhere, sOrderBy)
    Set oXMLDOM = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetXMLQueryStringWithSQLClauses", Err.Number
End Function

Private Function ParseParameters(sSQL As String, vParsValues As Variant) As String
  Dim sTemp As String, i As Long
  '*******************************************************************************
  On Error GoTo ErrHandler
    sTemp = sSQL
    For i = LBound(vParsValues) To UBound(vParsValues)
      sTemp = Replace(sTemp, "%" & i + 1, vParsValues(i))
    Next i
    ParseParameters = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParseParameters", Err.Number
End Function

Private Function ParseSQLClauses(sSQL As String, sWhere As String, sOrderBy As String) As String
  Dim sTemp As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    sTemp = sSQL
    If Len(sWhere) <> 0 Then
      sTemp = Replace(sTemp, "<@WHERE@>", "WHERE " & sWhere)
    Else
      sTemp = Replace(sTemp, "<@WHERE@>", "")
    End If
    If Len(sOrderBy) <> 0 Then
      sTemp = Replace(sTemp, "<@ORDER_BY@>", "ORDER BY " & sOrderBy)
    Else
      sTemp = Replace(sTemp, "<@ORDER_BY@>", "")
    End If
    ParseSQLClauses = TrimAll(sTemp)
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParseSQLClauses", Err.Number
End Function

Private Function TrimAll(sString As Variant) As String
  Dim sTemp As String
  '***************************************************
  If Not IsNull(sString) Then
    sTemp = sString
  Else
    sTemp = ""
  End If
  sTemp = Trim$(sTemp)
  sTemp = Replace(sTemp, vbTab, " ")
  Do
    If InStr(sTemp, Space(2)) Then
      sTemp = Replace(sTemp, Space(2), " ")
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
