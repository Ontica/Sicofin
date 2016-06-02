Attribute VB_Name = "MDBMethodsLibrary"
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Biblioteca : MDBMethodsLibrary                               Parte   : Data Services                     *
'* Fecha      : 19/Diciembre/2001                               Versión : 2.0                               *
'*                                                                                                          *
'* Descripción: Esta biblioteca proporciona los servicios de acceso a datos a bases de datos ODBC u OLEDB   *
'*              a través del componente Microsoft ActiveX Data Objects (ADO) versión 2.5.                   *
'*              Las consultas de selección para bases de datos Oracle® se definen en archivos XML.          *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId  As String = "MDBMethodsLibrary"

Private Const csOracleIdentificator As String = "MSDAORA"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS PARA LA RECUPERACIÓN DE RECORDSETS                                                      *
'************************************************************************************************************

Public Function GetDataValue(sAppServer As String, cDataConstant As TEnumDataConstant, _
                             vParsValues As Variant) As Variant
  Dim oConnection As New Connection, oCommand As New Command
  Dim sConnString As String, sSQL As String, bUseOracle As Boolean
  '*************************************************************************************
  On Error GoTo ErrHandler
    With oCommand
      sConnString = GetConstant(sConstantName:=sAppServer)
      bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
      sSQL = Trim$(GetDataConstant(cDataConstant))
      .ActiveConnection = sConnString
      .CommandText = sSQL
      .CommandType = adCmdStoredProc
      If (VarType(vParsValues) < vbArray) Then
        GetParameters oCommand, cDataConstant, Array(vParsValues), bUseOracle
      Else
        GetParameters oCommand, cDataConstant, vParsValues, bUseOracle
      End If
      .Execute
      GetDataValue = .Parameters(.Parameters.Count - 1)
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetDataValue", Err.Number
End Function

Public Function GetRecordset(sAppServer As String, cDataConstant As TEnumDataConstant, _
                             Optional bUpdatable As Boolean = False) As Recordset
  Dim oRecordset As New Recordset, sConnString As String, bUseOracle As Boolean
  '*************************************************************************************
  On Error GoTo ErrHandler
    sConnString = GetConstant(sConstantName:=sAppServer)
    bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
    With oRecordset
      If (Not bUseOracle) And (Not bUpdatable) Then
        .CursorLocation = adUseClient
        .Open GetDataConstant(cDataConstant), sConnString, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      ElseIf (Not bUseOracle) And (bUpdatable) Then
        .CursorLocation = adUseServer
        .Open GetDataConstant(cDataConstant), sConnString, adOpenDynamic, adLockOptimistic
      ElseIf (bUseOracle) And (Not bUpdatable) Then
        .CursorLocation = adUseClient
        .Open GetXMLQueryString(cDataConstant), sConnString, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      ElseIf (bUseOracle) And (bUpdatable) Then
        .CursorLocation = adUseServer
        .Open GetXMLQueryString(cDataConstant), sConnString, adOpenDynamic, adLockOptimistic
      End If
    End With
    Set GetRecordset = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRecordset", Err.Number
End Function

Public Function GetRecordsetDef(sAppServer As String, cDataConstant As TEnumDataConstant, _
                                vParsValues As Variant) As Recordset
  Dim oCommand As New Command, oRecordset As New Recordset
  Dim sConnString As String, bUseOracle As Boolean, sSQL As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    sConnString = GetConstant(sConstantName:=sAppServer)
    bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
    With oCommand
      If (Not bUseOracle) And (VarType(vParsValues) < vbArray) Then
        sSQL = Trim$(GetDataConstant(cDataConstant))
        .ActiveConnection = sConnString
        .CommandText = sSQL
        .CommandType = adCmdStoredProc
        GetParameters oCommand, cDataConstant, Array(vParsValues), bUseOracle
      ElseIf (Not bUseOracle) And (VarType(vParsValues) >= vbArray) Then
        sSQL = Trim$(GetDataConstant(cDataConstant))
        .ActiveConnection = sConnString
        .CommandText = sSQL
        .CommandType = adCmdStoredProc
        GetParameters oCommand, cDataConstant, vParsValues, bUseOracle
      ElseIf (bUseOracle) And (VarType(vParsValues) < vbArray) Then
        sSQL = GetXMLQueryString(cDataConstant)
        sSQL = ParseParameters(sSQL, Array(vParsValues))
      ElseIf (bUseOracle) And (VarType(vParsValues) >= vbArray) Then
        sSQL = GetXMLQueryString(cDataConstant)
        sSQL = ParseParameters(sSQL, vParsValues)
      End If
    End With
    With oRecordset
      .CursorLocation = adUseClient
      If (Not bUseOracle) Then
        .Open oCommand, , adOpenDynamic, adLockBatchOptimistic
      ElseIf (bUseOracle) Then
        .Open sSQL, sConnString, adOpenDynamic, adLockBatchOptimistic
      End If
      Set .ActiveConnection = Nothing
    End With
    Set GetRecordsetDef = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRecordsetDef", Err.Number
End Function

Public Function GetRecordsetWithPars(sAppServer As String, cDataConstant As TEnumDataConstant, _
                                     vParsValues As Variant, _
                                     Optional bUpdatable As Boolean = False) As Recordset
  Dim oCommand As New Command, oRecordset As New Recordset
  Dim sConnString As String, bUseOracle As Boolean, sSQL As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    With oCommand
      sConnString = GetConstant(sConstantName:=sAppServer)
      bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
      If (Not bUseOracle) And (VarType(vParsValues) < vbArray) Then
        sSQL = Trim$(GetDataConstant(cDataConstant))
        .ActiveConnection = sConnString
        .CommandText = sSQL
        .CommandType = adCmdStoredProc
        GetParameters oCommand, cDataConstant, Array(vParsValues), bUseOracle
      ElseIf (Not bUseOracle) And (VarType(vParsValues) >= vbArray) Then
        sSQL = Trim$(GetDataConstant(cDataConstant))
        .ActiveConnection = sConnString
        .CommandText = sSQL
        .CommandType = adCmdStoredProc
        GetParameters oCommand, cDataConstant, vParsValues, bUseOracle
      ElseIf (bUseOracle) And (VarType(vParsValues) < vbArray) Then
        sSQL = GetXMLQueryString(cDataConstant)
        sSQL = ParseParameters(sSQL, Array(vParsValues))
      ElseIf (bUseOracle) And (VarType(vParsValues) >= vbArray) Then
        sSQL = GetXMLQueryString(cDataConstant)
        sSQL = ParseParameters(sSQL, vParsValues)
      End If
    End With
    With oRecordset
      If (Not bUseOracle) And (Not bUpdatable) Then
        .CursorLocation = adUseClient
        .Open oCommand, , adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      ElseIf (Not bUseOracle) And (bUpdatable) Then
        .CursorLocation = adUseServer
        .Open oCommand, , adOpenDynamic, adLockOptimistic
      ElseIf (bUseOracle) And (Not bUpdatable) Then
        .CursorLocation = adUseClient
        .Open sSQL, sConnString, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      ElseIf (bUseOracle) And (bUpdatable) Then
        .CursorLocation = adUseServer
        .Open sSQL, sConnString, adOpenDynamic, adLockOptimistic
      End If
    End With
    Set GetRecordsetWithPars = oRecordset
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRecordsetWithPars", Err.Number
End Function

Public Function GetRecordsetWithSQLClauses(sAppServer As String, cDataConstant As TEnumDataConstant, _
                                           sWhere As String, sOrderBy As String, _
                                           Optional bUpdatable As Boolean = False) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '***************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetXMLQueryStringWithSQLClauses(cDataConstant, sWhere, sOrderBy)
    With oRecordset
      If Not bUpdatable Then
        .CursorLocation = adUseClient
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      Else
        .CursorLocation = adUseServer
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      End If
    End With
    Set GetRecordsetWithSQLClauses = oRecordset
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRecordsetWithSQLClauses", Err.Number
End Function

Public Function GetRecordsetWithSQLString(sAppServer As String, sSQL As String, _
                                          Optional bUpdatable As Boolean = False) As Recordset
  Dim oRecordset As New Recordset
  '*******************************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      If Not bUpdatable Then
        .CursorLocation = adUseClient
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      Else
        .CursorLocation = adUseServer
        .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockOptimistic
      End If
    End With
    Set GetRecordsetWithSQLString = oRecordset
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRecordsetWithSQLString", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PÚBLICOS PARA LA MODIFICACIÓN DE DATOS                                                           *
'************************************************************************************************************

Public Sub AppendHistory(sAppServer As String, oRecordset As Recordset, _
                         Optional sKeyField As String = "", _
                         Optional sSequenceName As String = "")
  Dim oConnection As New Connection, sConnString As String, bUseOracle As Boolean
  '******************************************************************************
  On Error GoTo ErrHandler
    sConnString = GetConstant(sConstantName:=sAppServer)
    bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
    oConnection.Open sConnString
    With oRecordset
      If (Not bUseOracle) Then
        Set .ActiveConnection = oConnection
        .UpdateBatch
        Set .ActiveConnection = Nothing
      ElseIf (bUseOracle) Then
        .MoveFirst
        Do While Not .EOF
          If oRecordset(sKeyField) = 0 Then
            oRecordset(sKeyField) = NewRecordId(sAppServer, sSequenceName)
          End If
          .MoveNext
        Loop
        Set .ActiveConnection = oConnection
        .UpdateBatch
        Set .ActiveConnection = Nothing
      End If
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AppendHistory", Err.Number
End Sub

Public Function AppendRecordset(sAppServer As String, oRecordset As Recordset, _
                          Optional sKeyField As String = "", _
                          Optional sSequenceName As String = "") As Long
  Dim oConnection As New Connection, sConnString As String, bUseOracle As Boolean
  Dim nPos As Long, nNewId As Long
  '******************************************************************************
  On Error GoTo ErrHandler
    sConnString = GetConstant(sConstantName:=sAppServer)
    bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
    oConnection.Open sConnString
    nPos = oRecordset.AbsolutePosition
    With oRecordset
      If (Not bUseOracle) Or (Len(sSequenceName) = 0) Then
        Set .ActiveConnection = oConnection
        .UpdateBatch
        Set .ActiveConnection = Nothing
      ElseIf (bUseOracle) Or (Len(sSequenceName) <> 0) Then
        .MoveFirst
        Do While Not .EOF
          nNewId = NewRecordId(sAppServer, sSequenceName)
          oRecordset(sKeyField) = nNewId
          .MoveNext
        Loop
        Set .ActiveConnection = oConnection
        .UpdateBatch
        .AbsolutePosition = nPos
        Set .ActiveConnection = Nothing
      End If
    End With
    AppendRecordset = nNewId
  Exit Function
ErrHandler:
  RaiseError ClassId, "AppendRecordset", Err.Number
End Function

Public Function ExecuteCommand(sAppServer As String, cDataConstant As TEnumDataConstant, _
                               vParsValues As Variant, ParamArray vReturnValues() As Variant) As Long
  Dim oConnection As New Connection, oCommand As New Command
  Dim sConnString As String, sSQL As String
  Dim bUseOracle As Boolean, nResult As Long, nOutputPars As Long, i As Long, j As Long
  '**************************************************************************************************
  On Error GoTo ErrHandler
    With oCommand
      sConnString = GetConstant(sConstantName:=sAppServer)
      bUseOracle = (InStr(1, sConnString, csOracleIdentificator) <> 0)
      sSQL = Trim$(GetDataConstant(cDataConstant))
      .ActiveConnection = sConnString
      .CommandText = sSQL
      .CommandType = adCmdStoredProc
      If (VarType(vParsValues) < vbArray) Then
        GetParameters oCommand, cDataConstant, Array(vParsValues), bUseOracle
      ElseIf (VarType(vParsValues) >= vbArray) Then
        GetParameters oCommand, cDataConstant, vParsValues, bUseOracle
      End If
      .Execute nResult
      ExecuteCommand = nResult
      If Not IsMissing(vReturnValues) Then
        nOutputPars = (UBound(vReturnValues) - LBound(vReturnValues)) + 1
        j = LBound(vReturnValues)
        For i = (.Parameters.Count - nOutputPars) To (.Parameters.Count - 1)
          vReturnValues(j) = IIf(IsNull(.Parameters(i)), vReturnValues(j), .Parameters(i))
          j = j + 1
        Next i
      End If
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExecuteCommand", Err.Number
End Function

Public Function ExecuteCommandWithSQLString(sAppServer As String, sSQL As String) As Long
  Dim oCommand As New Command, nResult As Long
  '**************************************************************************************
  On Error GoTo ErrHandler
    With oCommand
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute nResult
      ExecuteCommandWithSQLString = nResult
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExecuteCommandWithSQLString", Err.Number
End Function

Public Function NewRecordId(sAppServer As String, sSequence As String) As Long
  Dim oRecordset As New Recordset
  '***************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      .Open "SELECT " & sSequence & ".NEXTVAL AS newRecordId FROM DUAL", _
            GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      NewRecordId = !NewRecordId
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "NewRecordId", Err.Number
End Function

Public Sub SaveRecordset(sAppServer As String, oRecordset As Recordset)
  Dim oConnection As New Connection
  '********************************************************************
  On Error GoTo ErrHandler
    oConnection.Open GetConstant(sConstantName:=sAppServer)
    With oRecordset
      Set .ActiveConnection = oConnection
      .UpdateBatch
      Set .ActiveConnection = Nothing
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveRecordset", Err.Number
End Sub

Public Function TrimAll(sString As Variant) As String
  Dim sTemp As String
  '**************************************************
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
  Dim sQueryItemName As String
  '*****************************************************************************
  On Error GoTo ErrHandler
    sQueryItemName = GetDataConstant(cDataConstant)
    Set oXMLDOM = CreateObject("ECEXMLDOMParser.CXMLDOM")
    GetXMLQueryString = oXMLDOM.GetElementTextByName(GetConstant(cnXMLQueryFile), _
                                                     "query", sQueryItemName, "source")
    Set oXMLDOM = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetXMLQueryString", Err.Number
End Function

Private Function GetXMLQueryStringWithSQLClauses(cDataConstant As TEnumDataConstant, _
                                                 sWhere As String, sOrderBy As String) As String
  Dim oXMLDOM As Object, oXMLDOMNode As Object
  Dim sQueryItemName As String, sSQL As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    sQueryItemName = GetDataConstant(cDataConstant)
    Set oXMLDOM = CreateObject("ECEXMLDOMParser.CXMLDOM")
    Dim sXPath As String
    sXPath = "query[@name = '" & sQueryItemName & "']"
    Set oXMLDOMNode = oXMLDOM.GetNode(GetConstant(cnXMLQueryFile), sXPath)
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
