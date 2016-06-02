VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStdAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Standard Account Manager (EFAStdActBS)          Parte   : Business services                 *
'* Clase      : CStdAccount                                     Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Operaciones sobre una cuenta del catálogo de cuentas estándar.                              *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CStandardAccount"

Private Const gcInifinteDate As Date = #12/31/2049#
Private Const cStdAccountConstantPatternPrefix = "StdAccntPtrn "

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function ChangeRole(sAppServer As String, nStdAccountId As Long, sNewRole As String, _
                           sSectors As String, sControlSectors As String, _
                           nUserId As Long, Optional dDate As Variant) As String
  Dim oRecordset As New Recordset, nParentId As Long, sOldRole As String, sTemp As String, sCurrenciesList As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    'If Not ValidateInsert(sAppServer, sStdAccountNumber, sStdAccountRole) Then
    '  Exit Function
    'End If
    If IsMissing(dDate) Then
      dDate = Date
    End If
    Set oRecordset = GetStandardAccountRS(sAppServer, nStdAccountId)
    sOldRole = oRecordset!rol_cuenta
    oRecordset!rol_cuenta = sNewRole
    Save sAppServer, oRecordset, CDate(dDate), nStdAccountId
    
    If (sNewRole <> "S") Then
      'AddCurrencies sAppServer, nStdAccountId, sCurrencies
      'AssignAreas sAppServer, nStdAccountId, sAreas
    End If
    DeleteSector sAppServer, nStdAccountId, 0
    If (sNewRole = "X") Then
      AddSectors sAppServer, nStdAccountId, sSectors, sControlSectors, CDate(dDate)
    End If
  
    If (sOldRole = "S" And sNewRole <> "S") Then
      sTemp = GetChildsStdAccounts(sAppServer, nStdAccountId, True)
      sCurrenciesList = GetChildsCurrencies(sAppServer, sTemp)
      AddCurrencies sAppServer, nStdAccountId, sCurrenciesList, CDate(dDate)
      If (sNewRole = "X") Then
        GetChildsSectors sAppServer, sTemp, sSectors, sControlSectors
        AddSectors sAppServer, nStdAccountId, sSectors, sControlSectors, CDate(dDate)
      End If
    End If
    DoConvertion sAppServer, nStdAccountId, sOldRole, nStdAccountId, sNewRole, nUserId, CDate(dDate)
    If (sOldRole = "S") And (sNewRole <> "S") Then
      DeleteChilds sAppServer, nStdAccountId
    End If
'    sTemp = CalculateParentNumber(sAppServer, nStdAccountTypeId, sStdAccountNumber)
'    Do
'      If Len(sTemp) <> 0 Then
'        nParentId = GetStdAccountId(sAppServer, nStdAccountTypeId, sTemp)
'        If (nParentId <> 0) Then
'          Set oRecordset = GetStandardAccountRS(sAppServer, nParentId)
'          If (oRecordset!rol_cuenta <> "S") Then
'            oRecordset!rol_cuenta = "S"
'            Save sAppServer, oRecordset, nParentId
'          End If
'        ElseIf (nParentId = 0) Then
'          Set oRecordset = GetStandardAccountRS(sAppServer)
'          With oRecordset
'            !id_tipo_cuentas_std = nStdAccountTypeId
'            !numero_cuenta_estandar = sTemp
'            !nombre_cuenta_estandar = sStdAccountName
'            !descripcion = sStdAccountDescription
'            !rol_cuenta = "S"
'            !id_tipo_cuenta = nStdAccountType
'            !naturaleza = sStdAccountNature
'          End With
'          Save sAppServer, oRecordset
'        End If
'        sTemp = CalculateParentNumber(sAppServer, nStdAccountTypeId, sTemp)
'      Else
'        Exit Do
'      End If
'    Loop
    ChangeRole = ""
  Exit Function
ErrHandler:
  RaiseError ClassId, "ChangeRole", Err.Number
End Function

Public Function Insert(sAppServer As String, nStdAccountTypeId As Long, sStdAccountNumber As String, _
                       sStdAccountName As String, sStdAccountDescription As String, _
                       sStdAccountRole As String, nStdAccountType As Long, sStdAccountNature As String, _
                       sCurrencies As String, sAreas As String, sSectors As String, sControlSectors As String, _
                       nUserId As Long, Optional dDate As Variant) As String
  Dim oRecordset As New Recordset, nStdAccountId As Long, nParentId As Long, sTemp As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    'If Not ValidateInsert(sAppServer, sStdAccountNumber, sStdAccountRole) Then
    '  Exit Function
    'End If
    If IsMissing(dDate) Then
      dDate = Date
    End If
    Set oRecordset = GetStandardAccountRS(sAppServer)
    With oRecordset
      !id_tipo_cuentas_std = nStdAccountTypeId
      !numero_cuenta_estandar = sStdAccountNumber
      !nombre_cuenta_estandar = sStdAccountName
      !descripcion = sStdAccountDescription
      !rol_cuenta = sStdAccountRole
      !id_tipo_cuenta = nStdAccountType
      !naturaleza = sStdAccountNature
    End With
    nStdAccountId = Save(sAppServer, oRecordset, CDate(dDate))
    If nStdAccountId <= 0 Then  'Problems
      Exit Function
    End If
    If (sStdAccountRole <> "S") Then
      AddCurrencies sAppServer, nStdAccountId, sCurrencies, CDate(dDate)
      AssignAreas sAppServer, nStdAccountId, sAreas, CDate(dDate)
    End If
    If (sStdAccountRole = "X") Then
      AddSectors sAppServer, nStdAccountId, sSectors, sControlSectors, CDate(dDate)
    End If
    If (sStdAccountRole <> "S") Then
      nParentId = GetParentId(sAppServer, nStdAccountId, CDate(dDate))
      If (nParentId <> 0) Then
        DoConvertion sAppServer, nStdAccountId, sStdAccountRole, nParentId, "S", nUserId, CDate(dDate)
      End If
    End If
    sTemp = CalculateParentNumber(sAppServer, nStdAccountTypeId, sStdAccountNumber)
    Do
      If Len(sTemp) <> 0 Then
        nParentId = GetStdAccountId(sAppServer, nStdAccountTypeId, sTemp)
        If (nParentId <> 0) Then
          Set oRecordset = GetStandardAccountRS(sAppServer, nParentId)
          If (oRecordset!rol_cuenta <> "S") Then
            oRecordset!rol_cuenta = "S"
            Save sAppServer, oRecordset, CDate(dDate), nParentId
          End If
        ElseIf (nParentId = 0) Then
          Set oRecordset = GetStandardAccountRS(sAppServer)
          With oRecordset
            !id_tipo_cuentas_std = nStdAccountTypeId
            !numero_cuenta_estandar = sTemp
            !nombre_cuenta_estandar = sStdAccountName
            !descripcion = sStdAccountDescription
            !rol_cuenta = "S"
            !id_tipo_cuenta = nStdAccountType
            !naturaleza = sStdAccountNature
          End With
          Save sAppServer, oRecordset, CDate(dDate)
        End If
        sTemp = CalculateParentNumber(sAppServer, nStdAccountTypeId, sTemp)
      Else
        Exit Do
      End If
    Loop
    Insert = ""
  Exit Function
ErrHandler:
  RaiseError ClassId, "Insert", Err.Number
End Function

Public Sub AddCurrencies(sAppServer As String, nStdAccountId As Long, sCurrenciesList As String, dDate As Date)
  Dim oRecordset As New Recordset, aCurrenciesIds As Variant, i As Long, sSQL As String
  '**************************************************************************************************
  On Error GoTo ErrHandler
    aCurrenciesIds = Split(sCurrenciesList, ",")
    sSQL = "SELECT * FROM COF_MAPEO_MONEDA WHERE id_cuenta_estandar = -1"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
      For i = LBound(aCurrenciesIds) To UBound(aCurrenciesIds)
        If Len(aCurrenciesIds(i)) <> 0 Then
          .AddNew
            !id_cuenta_estandar = nStdAccountId
            !id_moneda = CLng(aCurrenciesIds(i))
            !fecha_inicio = dDate
            !Fecha_Fin = #12/31/2049#
        End If
      Next i
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddCurrencies", Err.Number
End Sub

Public Sub AddSectors(sAppServer As String, nStdAccountId As Long, sSectorsList As String, _
                      sControlSectors As String, dDate As Date)
  Dim oRecordset As New Recordset, aSectorsIds As Variant, aControlSectors As Variant
  Dim i As Long, sSQL As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    aSectorsIds = Split(sSectorsList, ",")
    If Len(sControlSectors) <> 0 Then
      aControlSectors = Split(sControlSectors, ",")
    Else
      aControlSectors = Null
    End If
    With oRecordset
      sSQL = "SELECT * FROM COF_MAPEO_SECTOR WHERE id_cuenta_estandar = -1"
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
      For i = LBound(aSectorsIds) To UBound(aSectorsIds)
        If Len(aSectorsIds(i)) <> 0 Then
          .AddNew
            !id_cuenta_estandar = nStdAccountId
            !id_sector = CLng(aSectorsIds(i))
            If Not IsNull(aControlSectors) Then
              If ItemInArray(aControlSectors, aSectorsIds(i)) Then
                !rol_sector = "C"
              Else
                !rol_sector = "P"
              End If
            Else
              !rol_sector = "P"
            End If
            !fecha_inicio = dDate
            !Fecha_Fin = #12/31/2049#
        End If
      Next i
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AddSectors", Err.Number
End Sub

Public Sub AssignAreas(sAppServer As String, nStdAccountId As Long, sAreasList As String, dDate As Date)
  Dim oRecordset As New Recordset, aAreas As Variant, i As Long, sSQL As String
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    DeleteAreas sAppServer, nStdAccountId, CDate(dDate)
    If Len(sAreasList) = 0 Then
      Exit Sub
    End If
    aAreas = Split(sAreasList, ",")
    sSQL = "SELECT * FROM COF_MAPEO_AREA WHERE id_cuenta_estandar = -1"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
      For i = LBound(aAreas) To UBound(aAreas)
        If Len(aAreas(i)) <> 0 Then
          .AddNew
            !id_cuenta_estandar = nStdAccountId
            !patron_area = CStr(aAreas(i))
            !fecha_inicio = dDate
            !Fecha_Fin = #12/31/2049#
        End If
      Next i
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "AssignAreas", Err.Number
End Sub

Public Function ValidateAreas(sAppServer As String, sAreasList As String) As Boolean
ValidateAreas = True
End Function

Private Function ItemInArray(aArray As Variant, vItem As Variant) As Boolean
  Dim i As Long
  '*************************************************************************
  On Error GoTo ErrHandler
    ItemInArray = False
    For i = LBound(aArray) To UBound(aArray)
      If (CStr(aArray(i)) = CStr(vItem)) Then
        ItemInArray = True
        Exit Function
      End If
    Next i
  Exit Function
ErrHandler:
  RaiseError ClassId, "ItemInArray", Err.Number
End Function

Private Function DeleteCurrency(sAppServer As String, nStdAccountId As Long, nCurrencyId As Long, _
                                Optional dDate As Variant) As Boolean
  Dim oCommand As New Command, sSQL As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    DeleteCurrency = False
    If DoDeleteCurrency(sAppServer, nStdAccountId, nCurrencyId) Then
      sSQL = "DELETE FROM COF_MAPEO_MONEDA " & _
             "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND (id_moneda = " & nCurrencyId & ")"
      With oCommand
        .ActiveConnection = GetConstant(sConstantName:=sAppServer)
        .CommandText = sSQL
        .CommandType = adCmdText
        .Execute
      End With
      DeleteCurrency = True
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "DeleteCurrency", Err.Number
End Function

Private Function DeleteSector(sAppServer As String, nStdAccountId As Long, nSectorId As Long, _
                              Optional dDate As Variant) As Boolean
  Dim oCommand As New Command, sSQL As String
  '********************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If

    sSQL = "DELETE FROM COF_MAPEO_SECTOR " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    With oCommand
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute
    End With
    DeleteSector = True
  Exit Function
ErrHandler:
  RaiseError ClassId, "DeleteSector", Err.Number
End Function

Private Sub DeleteAreas(sAppServer As String, nStdAccountId As Long, Optional dDate As Date)
  Dim oCommand As New Command, sSQL As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    sSQL = "DELETE FROM COF_MAPEO_AREA " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    With oCommand
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CommandText = sSQL
      .CommandType = adCmdText
      .Execute
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DeleteAreas", Err.Number
End Sub

Public Function SaveBasicInfo(sAppServer As String, nStdAccountId As Long, sName As String, _
                              sDescription As String, nTypeId As Long, sNature As String) As Boolean
  Dim oRecordset As New Recordset, sSQL As String
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '**************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Cuenta_Estandar WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
      If Not (.BOF And .EOF) Then
        !nombre_cuenta_estandar = sName
        !descripcion = sDescription
        !id_tipo_cuenta = nTypeId
        !naturaleza = sNature
        .ActiveConnection = GetConstant(sConstantName:=sAppServer)
        .UpdateBatch
      End If
      .Close
      
      sSQL = "SELECT * FROM COF_Cuenta_Estandar_Hist WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
      Do While Not .EOF
        !nombre_cuenta_estandar = sName
        !descripcion = sDescription
        !id_tipo_cuenta = nTypeId
        !naturaleza = sNature
        .MoveNext
      Loop
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .UpdateBatch
      .Close
    End With
    SaveBasicInfo = True
  Exit Function
ErrHandler:
  RaiseError ClassId, "SaveBasicInfo", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Function DoDeleteCurrency(sAppServer As String, nStdAccountId As Long, _
                                  nCurrencyId As Long) As Boolean
DoDeleteCurrency = False
End Function

Private Function DoDeleteSector(sAppServer As String, nStdAccountId As Long, nSectorId As Long) As Boolean
DoDeleteSector = False
End Function

Public Function Save(sAppServer As String, oRecordset As Recordset, dDate As Date, Optional nStdAccountId As Long) As Long
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '***********************************************************************************************************************
  On Error GoTo ErrHandler
    If (nStdAccountId = 0) Then
      oRecordset!id_cuenta_estandar = 0
      Validate sAppServer, oRecordset
      nStdAccountId = AppendRecordset(sAppServer, oRecordset, "id_cuenta_estandar", "SEC_ID_CUENTA_ESTANDAR")
      SaveHistory sAppServer, oRecordset, dDate
    Else
      oRecordset!id_cuenta_estandar = nStdAccountId
      Validate sAppServer, oRecordset, False
      SaveRecordset sAppServer, oRecordset
      SaveHistory sAppServer, oRecordset, dDate
    End If
    Save = nStdAccountId
  Exit Function
ErrHandler:
  RaiseError ClassId, "Save", Err.Number
End Function

Private Sub SaveHistory(sAppServer As String, oRecordset As Recordset, dDate As Date)
  Dim oContext As ObjectContext, bObjectContextOK As Boolean, oHistoryRS As Recordset
  '**********************************************************************************
  On Error GoTo ErrHandler
    Set oHistoryRS = GetHistoryRS(sAppServer, oRecordset!id_cuenta_estandar)
    With oRecordset
      If Not (oHistoryRS.BOF And oHistoryRS.EOF) Then
        oHistoryRS!Fecha_Fin = dDate - 1
        'SaveHistory sAppServer, oRecordset
      End If
      oHistoryRS.AddNew
      oHistoryRS!id_cuenta_estandar = !id_cuenta_estandar
      oHistoryRS!id_tipo_cuentas_std = !id_tipo_cuentas_std
      oHistoryRS!numero_cuenta_estandar = !numero_cuenta_estandar
      oHistoryRS!nombre_cuenta_estandar = !nombre_cuenta_estandar
      oHistoryRS!descripcion = !descripcion
      oHistoryRS!id_tipo_cuenta = !id_tipo_cuenta
      oHistoryRS!rol_cuenta = !rol_cuenta
      oHistoryRS!naturaleza = !naturaleza
      oHistoryRS!fecha_inicio = dDate
      oHistoryRS!Fecha_Fin = gcInifinteDate
    End With
    AppendHistory sAppServer, oHistoryRS, "id_cuenta_estandar_hist", "SEC_ID_CUENTA_ESTANDAR_HIST"
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveHistory", Err.Number
End Sub

Private Function Exists(sAppServer As String, nStdAccountCatalogId As Long, _
                        sStdAccountNumber As String) As Boolean
  Exists = False
End Function

Private Function ExistsCatalog(sAppServer As String, nStdAccountCatalogId As Long) As Boolean
  ExistsCatalog = True
End Function

Private Function ExistsAccountType(sAppServer As String, nStdAccountTypeId As Long) As Boolean
  ExistsAccountType = True
End Function


Private Sub Validate(sAppServer As String, oRecordset As Recordset, Optional bForAppend As Boolean = True)
  Dim nResult As Long
  '*******************************************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      !nombre_cuenta_estandar = TrimAll(!nombre_cuenta_estandar)
      !descripcion = TrimAll(!descripcion)
      If Len(!nombre_cuenta_estandar) = 0 Then
        Err.Raise TEnumErrors.ErrNullAccountName
      End If
      If Not ExistsCatalog(sAppServer, !id_tipo_cuentas_std) Then
        Err.Raise TEnumErrors.ErrInvalidCatalogId
      End If
      If (!naturaleza <> "D") And (!naturaleza <> "A") Then
        Err.Raise TEnumErrors.ErrInvalidNature
      End If
      If (!rol_cuenta <> "S") And (!rol_cuenta <> "P") And (!rol_cuenta <> "C") And (!rol_cuenta <> "X") Then
        Err.Raise TEnumErrors.ErrInvalidRole
      End If
      If (Not ExistsAccountType(sAppServer, !id_tipo_cuenta)) Then
        Err.Raise TEnumErrors.ErrInvalidAccountType
      End If
      
      If (bForAppend) Then
        If Exists(sAppServer, !id_tipo_cuentas_std, !numero_cuenta_estandar) Then
          Err.Raise TEnumErrors.ErrAccountAlreadyExists
        End If
      Else
        'If Not Exists(sAppServer, !id_tipo_cuentas_std, !numero_cuenta_estandar) Then
        '  Err.Raise TEnumErrors.ErrAccountNotExists
        'End If
      End If
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Validate", Err.Number
End Sub

'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************
'*******************************************************************************************************

Public Function Role(sAppServer As String, nStdAccountId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, nResult As Long
  '************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      sSQL = "SELECT Rol_Cuenta FROM COF_CUENTA_ESTANDAR " & _
             "WHERE id_cuenta_estandar = " & nStdAccountId
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        Role = !rol_cuenta
      Else
        Role = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "Role", Err.Number
End Function

Public Function GetAccountRS(sAppServer As String, Optional nAccountId As Long = 0) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean, oRecordset As Recordset
  '*********************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set oRecordset = GetRecordsetDef(sAppServer, cnDefAccount, nAccountId)
    If nAccountId = 0 Then
      oRecordset.AddNew
    End If
    Set GetAccountRS = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetAccountRS", Err.Number
End Function

Public Function AssignToGralLedger(sAppServer As String, nGralLedgerId As Long, _
                                   nStdAccountId As Long) As Long
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oRecordset As Recordset, nNewAccountId As Long
  '******************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = GetAccountRS(sAppServer)
    With oRecordset
      !id_mayor = nGralLedgerId
      !id_cuenta_estandar = nStdAccountId
    End With
    nNewAccountId = AppendRecordset(sAppServer, oRecordset, "Id_cuenta", "SEC_ID_CUENTA")
    AssignToGralLedger = nNewAccountId
  Exit Function
ErrHandler:
  RaiseError ClassId, "AssignToGralLedger", Err.Number
End Function

Public Function Pattern(nStdAccountCategoryId As Long) As String
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim sPatternConstant As String
  '*************************************************************
  On Error GoTo ErrHandler
    sPatternConstant = cStdAccountConstantPatternPrefix & nStdAccountCategoryId
    Pattern = GetConstant(sConstantName:=sPatternConstant)
  Exit Function
ErrHandler:
  RaiseError ClassId, "Pattern", Err.Number
End Function

Public Sub SetPattern(nStdAccountCategoryId As Long, sNewPattern As String)
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim sPatternConstant As String
  '*************************************************************
  On Error GoTo ErrHandler
    sPatternConstant = cStdAccountConstantPatternPrefix & nStdAccountCategoryId
    WriteConstant sPatternConstant, sNewPattern
    DoEvents
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Pattern", Err.Number
End Sub

Public Function IsAccountNumberValid(nStdAccountCategoryId As Long, sStdAccountNumber As String) As Boolean
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim sPattern As String, sTemp As String, sTemp2 As String, i As Long, sChar As String
  '********************************************************************************************************
  On Error GoTo ErrHandler
    sPattern = Pattern(nStdAccountCategoryId)
    sTemp2 = Trim$(sStdAccountNumber)
    If Len(sTemp2) = 0 Then
      IsAccountNumberValid = False
      Exit Function
    End If
    For i = 1 To Len(sTemp2)
      sChar = Mid$(sTemp2, i, 1)
      If InStr(1, "0123456789", sChar) Then
        sTemp = sTemp & sChar
      ElseIf sChar <> "-" Then
        IsAccountNumberValid = False
        Exit Function
      End If
    Next i
    sTemp = Format(sTemp, "!" & sPattern)
    Do
      If InStr(sTemp, "--") Then
        sTemp = Replace(sTemp, "--", "")
      Else
        Exit Do
      End If
    Loop
    If (Right$(sTemp, 1) = "-") Then
      sTemp = Left$(sTemp, Len(sTemp) - 1)
    End If
    If (Mid$(sTemp, Len(sTemp) - 1, 1) = "-") Then
      sTemp = ""
    End If
    IsAccountNumberValid = (sTemp = sStdAccountNumber)
  Exit Function
ErrHandler:
  RaiseError ClassId, "IsAccountNumberValid", Err.Number
End Function

Public Function FormatAccountNumber(nStdAccountCategoryId As Long, sAccount As String) As String
  Dim sTemp As String, i As Long, sChar As String, sPattern As String
  '*********************************************************************************************
  On Error GoTo ErrHandler
    sAccount = Trim$(sAccount)
    sPattern = Pattern(nStdAccountCategoryId)
    If Len(sAccount) = 0 Then
      Exit Function
    End If
    For i = 1 To Len(sAccount)
      sChar = Mid$(sAccount, i, 1)
      If InStr(1, "0123456789", sChar) Then
        sTemp = sTemp & sChar
      ElseIf sChar <> "-" Then
        FormatAccountNumber = ""
        Exit Function
      End If
    Next i
    sTemp = Format(sTemp, "!" & sPattern)
    Do
      If InStr(sTemp, "--") Then
        sTemp = Replace(sTemp, "--", "")
      Else
        Exit Do
      End If
    Loop
    If Right$(sTemp, 1) = "-" Then
      sTemp = Left$(sTemp, Len(sTemp) - 1)
    End If
    FormatAccountNumber = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "FormatAccountNumber", Err.Number
End Function

Public Function Delete(sAppServer As String, nStdAccountId As Long, _
                       Optional dDate As Variant) As Boolean
  Dim oContext As ObjectContext, bObjectContextOK As Boolean, nResult As Long, sLogFile As String
  Dim nParentStdAccountId As Long
  '************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date - 1
    Else
      dDate = dDate - 1
    End If
    DeleteSector sAppServer, nStdAccountId, 0, dDate
    DeleteCurrency sAppServer, nStdAccountId, 0, dDate
    nResult = ExecuteCommand(sAppServer, cnDelStandardAccount, Array(nStdAccountId, dDate))
    If nResult = 0 Then
      RaiseError ClassId, "DeleteStandardAccount", TEnumErrors.ErrAccountNotExists, nStdAccountId
    ElseIf nResult < 0 Then
      RaiseError ClassId, "DeleteStandardAccount", TEnumErrors.ErrAccountInUse
    ElseIf nResult > 0 Then
      
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "Delete", Err.Number
End Function

Private Function DeleteChilds(sAppServer As String, nParentStdAccountId As Long, _
                              Optional dDate As Variant) As Boolean
  Dim aChilds As Variant, i As Long
  '********************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    aChilds = Split(GetChildsStdAccounts(sAppServer, nParentStdAccountId, False), ",")
    For i = LBound(aChilds) To UBound(aChilds)
      If IsNumeric(Trim$(aChilds(i))) Then
        If (CLng(aChilds(i)) <> nParentStdAccountId) Then
          Delete sAppServer, CLng(Trim$(aChilds(i))), dDate
        End If
      End If
    Next i
  Exit Function
ErrHandler:
  RaiseError ClassId, "DeleteChilds", Err.Number
End Function

Public Function GetSectorStandardAccounts(sAppServer As String, nSectorId) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set GetSectorStandardAccounts = GetRecordsetWithPars(sAppServer, cnQrySectorStandardAccounts, nSectorId, False)
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetSectorStandardAccounts", Err.Number
End Function

Public Function GetStandardAccount(sAppServer As String, nStdAccountId As Long) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '**********************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set GetStandardAccount = GetRecordsetWithPars(sAppServer, cnQryStandardAccount, nStdAccountId)
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetStandardAccount", Err.Number
End Function

Public Function GetStandardAccounts(sAppServer As String) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '*******************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set GetStandardAccounts = GetRecordset(sAppServer, cnQryStandardAccounts)
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetStandardAccounts", Err.Number
End Function

Public Function GetStandardAccountRS(sAppServer As String, Optional nStdAccountId As Long = 0) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean, oRecordset As Recordset
  '*************************************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set oRecordset = GetRecordsetDef(sAppServer, cnDefStandardAccount, nStdAccountId)
    If nStdAccountId = 0 Then
      oRecordset.AddNew
    End If
    Set GetStandardAccountRS = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetStandardAccountRS", Err.Number
End Function

Private Function GetHistoryRS(sAppServer As String, _
                                          Optional nStdAccountId As Long = 0) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  '***********************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    Set GetHistoryRS = GetRecordsetDef(sAppServer, cnDefStandardAccountHist, nStdAccountId)
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "GetHistoryRS", Err.Number
End Function

Private Sub ValidateStandardAccountNumber(sFormatStandardAccount As String, sNameStandardAccount As String)

End Sub

Public Function CalculateParentNumber(sAppServer As String, nStdAccountTypeId As Long, sStdAccountNumber As String) As String
  On Error GoTo ErrHandler
    If Len(sStdAccountNumber) > 4 Then
      CalculateParentNumber = Left$(sStdAccountNumber, Len(sStdAccountNumber) - 3)
    Else
      CalculateParentNumber = ""
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "CalculateParentNumber", Err.Number
End Function

Public Function GetParentId(sAppServer As String, nStdAccountId As Long, dDate As Date) As Long
  Dim oRecordset As New Recordset, sStdAccountNumber As String, nStdAccountTypeId As Long, nParentId As Long
  '**************************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = GetStdAccount(sAppServer, nStdAccountId, dDate)
    With oRecordset
      If Not (.BOF And .EOF) Then
        nStdAccountTypeId = !id_tipo_cuentas_std
        sStdAccountNumber = !numero_cuenta_estandar
      Else
        Err.Raise TEnumErrors.ErrAccountNotExists
      End If
      .Close
    End With
    Do
      If (Len(sStdAccountNumber) <> 0) And (nParentId = 0) Then
        sStdAccountNumber = CalculateParentNumber(sAppServer, nStdAccountTypeId, sStdAccountNumber)
        nParentId = GetStdAccountId(sAppServer, nStdAccountTypeId, sStdAccountNumber)
      Else
        Exit Do
      End If
    Loop
    GetParentId = nParentId
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetParentId", Err.Number
End Function

Public Function GetStdAccountId(sAppServer As String, nStdAccountTypeId As Long, _
                                sStdAccountNumber As String) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '*******************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar, fecha_fin FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (numero_cuenta_estandar = '" & sStdAccountNumber & "') AND " & _
           "(id_tipo_cuentas_std = " & nStdAccountTypeId & ") " & _
           "ORDER BY fecha_fin desc"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        If CDate(!Fecha_Fin) = CDate(GetConstant(cnLastSystemDate)) Then
          GetStdAccountId = !id_cuenta_estandar
        Else
          GetStdAccountId = -1 * !id_cuenta_estandar
        End If
      Else
        GetStdAccountId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountId", Err.Number
End Function

Public Function GetParent(sAppServer As String, nStdAccountTypeId As Long, _
                          sStdAccountNumber As String) As Recordset
  Dim oRecordset As New Recordset, sTemp As String, nParentId  As Long
  '*************************************************************************
  On Error GoTo ErrHandler
    nParentId = 0
    sTemp = sStdAccountNumber
    Do
      If (Len(sTemp) <> 0) And (nParentId = 0) Then
        sTemp = CalculateParentNumber(sAppServer, nStdAccountTypeId, sTemp)
        nParentId = GetStdAccountId(sAppServer, nStdAccountTypeId, sTemp)
      Else
        Exit Do
      End If
    Loop
    If (nParentId <> 0) Then
      Set GetParent = GetStdAccount(sAppServer, Abs(nParentId))
    Else
      Set GetParent = Nothing
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetParent", Err.Number
End Function

Public Function GetStdAccount(sAppServer As String, nStdAccountId As Long, _
                              Optional dDate As Variant) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    sSQL = "SELECT COF_Cuenta_Estandar_Hist.*, AO_Objects.Object_Name nombre_tipo_cuentas_std " & _
           "FROM COF_Cuenta_Estandar_Hist, AO_Objects " & _
           "WHERE (Id_Cuenta_Estandar = " & nStdAccountId & ") AND " & _
           "(COF_Cuenta_Estandar_Hist.id_tipo_cuentas_std = AO_Objects.object_id) AND " & _
           "(COF_Cuenta_Estandar_Hist.fecha_inicio <= '" & dDate & "') AND " & _
           "('" & dDate & "' <= COF_Cuenta_Estandar_Hist.fecha_fin)"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetStdAccount = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccount", Err.Number
End Function

Private Sub DoConvertion(sAppServer As String, nStdAccountId As Long, sActualRole As String, _
                         nParentId As Long, sNewRole As String, nUserId As Long, dDate As Date)
  Dim oBalances As New EFABalanceEngine.CEngine, oRecordset As New Recordset, aGralLedgers As Variant
  Dim oTransactionsRS As Recordset, oPostingsRS As Recordset, sPostingType As String, nCurrentCurrency As Long
  Dim sNumber As String, sNature As String, sParentNature As String, nTransactionId As Long, nSourceAccountId As Long
  Dim nGralLedgerId As Long, nCreditSum As Double, nDebitSum As Double, nTargetAccountId As Long, sParentRole As String
  Dim nSubsidiaryAccountId As Long, sParentNumberFrom As String, sParentNumberTo As String, nSectorId As Long
  '*****************************************************************************************************************
  On Error GoTo ErrHandler
    Set oRecordset = GetStdAccount(sAppServer, nStdAccountId, dDate)
    sNature = oRecordset!naturaleza
    sNumber = oRecordset!numero_cuenta_estandar
    oRecordset.Close
    Set oRecordset = Nothing
    
    If (nStdAccountId <> nParentId) Then
      Set oRecordset = GetStdAccount(sAppServer, nParentId, dDate)
      sParentNature = oRecordset!naturaleza
      sParentNumberFrom = oRecordset!numero_cuenta_estandar
      sParentNumberTo = oRecordset!numero_cuenta_estandar
      sParentRole = oRecordset!rol_cuenta
      If (sParentRole = "S") Then
        oRecordset.Close
        Exit Sub
      End If
      oRecordset.Close
      Set oRecordset = Nothing
      aGralLedgers = GetStdAccountGralLedgers(sAppServer, nParentId, False)
    ElseIf (nStdAccountId = nParentId) Then
      sParentNature = sNature
      sParentNumberFrom = sNumber
      If (sActualRole = "S" And sNewRole <> "S") Then
        sParentNumberTo = Left$(sParentNumberFrom & "-99-99-99-99-99-99", 22)
        aGralLedgers = GetStdAccountGralLedgers(sAppServer, nParentId, True)
      Else
        sParentNumberTo = sParentNumberFrom
        aGralLedgers = GetStdAccountGralLedgers(sAppServer, nParentId, False)
      End If
    End If
    
    If IsNull(aGralLedgers) Then   'No hay nada que convertir
      Exit Sub
    End If
    Set oTransactionsRS = GetTransactionsRS(sAppServer)
    Set oPostingsRS = GetPostingsRS(sAppServer)
    Set oRecordset = oBalances.BalancesRS(sAppServer, "####-##-##-##-##-##-##", aGralLedgers, False, True, True, _
                                          CDate(dDate), CDate(dDate), sParentNumberFrom, sParentNumberTo, nBalancesToDisplay:=6)
    With oRecordset
      Do While Not .EOF
        If nGralLedgerId <> !id_mayor Then
          nGralLedgerId = !id_mayor
          nTransactionId = PutTransaction(sAppServer, oTransactionsRS, nGralLedgerId, 338, _
                                          21, 25, "Conversión de cuentas", _
                                          CDate(dDate), Now(), nUserId, GetSupervisorId(sAppServer, nGralLedgerId))
        End If
        If (!saldo_actual <> 0) And (!rol_cuenta <> "S") Then
          If (!naturaleza = "D" And !saldo_actual > 0) Then
            sPostingType = "H"
            nCreditSum = nCreditSum + Abs(!saldo_actual)
          ElseIf (!naturaleza = "D") And (!saldo_actual < 0) Then
            sPostingType = "D"
            nDebitSum = nDebitSum + Abs(!saldo_actual)
          ElseIf (!naturaleza = "A") And (!saldo_actual > 0) Then
            sPostingType = "D"
            nDebitSum = nDebitSum + Abs(!saldo_actual)
          ElseIf (!naturaleza = "A") And (!saldo_actual < 0) Then
            sPostingType = "H"
            nCreditSum = nCreditSum + Abs(!saldo_actual)
          End If
          nSourceAccountId = GetAccountId(sAppServer, nGralLedgerId, nParentId)
          PutPosting sAppServer, oPostingsRS, nTransactionId, nSourceAccountId, !id_cuenta_auxiliar, _
                    !id_sector, 0, 0, "", "", "", Null, "", sPostingType, !id_moneda, _
                    Abs(!saldo_actual), Abs(!saldo_actual), True
          nTargetAccountId = GetAccountId(sAppServer, nGralLedgerId, nStdAccountId)
          If (sNewRole = sActualRole) Or (sNewRole = "S") Then
            nSubsidiaryAccountId = !id_cuenta_auxiliar
            If !id_sector <> 0 Then
              If SectorRole(sAppServer, nStdAccountId, !id_sector) = "P" Then
                nSubsidiaryAccountId = 0
              End If
            End If
            If (nCreditSum <> 0) Then
              PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, nSubsidiaryAccountId, _
                         !id_sector, 0, 0, "", "", "", Null, "", "D", !id_moneda, nCreditSum, nCreditSum, False
            End If
            If (nDebitSum <> 0) Then
              PutPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, nSubsidiaryAccountId, _
                         !id_sector, 0, 0, "", "", "", Null, "", "H", !id_moneda, nDebitSum, nDebitSum, False
            End If
          Else
            nSectorId = 0
            nSubsidiaryAccountId = 0
            If (sNewRole = "C" And sActualRole = "X") Or (sActualRole = "S" And sNewRole <> "P") Then
              nSubsidiaryAccountId = !id_cuenta_auxiliar
            End If
            If (!id_sector <> 0) And (sNewRole = "X") Then
              nSectorId = !id_sector
              If SectorRole(sAppServer, nStdAccountId, !id_sector) = "P" Then
                nSubsidiaryAccountId = 0
              End If
            End If
            If (nCreditSum <> 0) Then
              PutGroupedPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, nSubsidiaryAccountId, _
                         !id_moneda, nSectorId, "D", nCreditSum, nCreditSum, False
            End If
            If (nDebitSum <> 0) Then
              PutGroupedPosting sAppServer, oPostingsRS, nTransactionId, nTargetAccountId, nSubsidiaryAccountId, _
                         !id_moneda, nSectorId, "H", nDebitSum, nDebitSum, False
            End If
          End If
        End If
        nDebitSum = 0
        nCreditSum = 0
        .MoveNext
      Loop
      .Close
    End With
    SaveRecordset sAppServer, oTransactionsRS
    SaveRecordset sAppServer, oPostingsRS
  Exit Sub
ErrHandler:
  RaiseError ClassId, "DoConvertion", Err.Number
End Sub

Private Function GetSupervisorId(sAppServer As String, nGralLedgerId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************
  GetSupervisorId = 0
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetSupervisorId", Err.Number
End Function

Private Function GetStdAccountGralLedgers(sAppServer As String, nStdAccountId As Long, _
                                          Optional bGetChilds As Boolean = False) As Variant
  Dim oRecordset As New Recordset, sSQL As String, sChildsStdAccounts As String, aGralLedgers As Variant
  '*****************************************************************************************************
  On Error GoTo ErrHandler
    If Not bGetChilds Then
      sSQL = "SELECT DISTINCT id_mayor FROM COF_Cuenta " & _
             "WHERE (id_cuenta_estandar = " & nStdAccountId & ") " & _
             "ORDER BY id_mayor"      'eliminado <> 0
    Else
      sChildsStdAccounts = GetChildsStdAccounts(sAppServer, nStdAccountId)
      sSQL = "SELECT DISTINCT COF_Cuenta.id_mayor FROM COF_Cuenta " & _
             "WHERE (COF_Cuenta.id_cuenta_estandar IN (" & sChildsStdAccounts & ") )" & _
             "ORDER BY COF_Cuenta.id_mayor"      'eliminado <> 0
      
    End If
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If IsEmpty(aGralLedgers) Then
          ReDim aGralLedgers(0)
        Else
          ReDim Preserve aGralLedgers(UBound(aGralLedgers) + 1)
        End If
        aGralLedgers(UBound(aGralLedgers)) = !id_mayor
        .MoveNext
      Loop
      GetStdAccountGralLedgers = aGralLedgers
      Set .ActiveConnection = Nothing
    End With
  Exit Function
ErrHandler:
  GetStdAccountGralLedgers = Null
  RaiseError ClassId, "GetStdAccountGralLedgers", Err.Number
End Function

Public Function GetChildsStdAccounts(sAppServer As String, nStdAccountId As Long, _
                                     Optional bOnlyDetailAccounts As Boolean = True) As String
  Dim oRecordset As New Recordset, sSQL As String, sChildsStdAccounts As String
  Dim nStdAccountTypeId As Integer, sStdAccountNumber As String
  '************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_tipo_cuentas_std, numero_cuenta_estandar FROM COF_Cuenta_Estandar " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        nStdAccountTypeId = !id_tipo_cuentas_std
        sStdAccountNumber = !numero_cuenta_estandar
      Else
        Exit Function
      End If
      .Close
      sSQL = "SELECT id_cuenta_estandar FROM COF_Cuenta_Estandar " & _
             "WHERE (id_tipo_cuentas_std = " & nStdAccountTypeId & ") AND " & _
             "(numero_cuenta_estandar LIKE '" & sStdAccountNumber & "%') "
      If bOnlyDetailAccounts Then
        sSQL = sSQL & "AND (rol_cuenta <> 'S')"
      End If
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sChildsStdAccounts) <> 0 Then
          sChildsStdAccounts = sChildsStdAccounts & ", " & !id_cuenta_estandar
        Else
          sChildsStdAccounts = !id_cuenta_estandar
        End If
        .MoveNext
      Loop
      Set .ActiveConnection = Nothing
    End With
    GetChildsStdAccounts = sChildsStdAccounts
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetChildsStdAccounts", Err.Number
End Function

Private Sub PutPosting(sAppServer As String, oPostingsRS As Recordset, nTransactionId As Long, _
                       nAccountId As Long, nSubsidiaryAccountId As Long, nSectorId As Long, _
                       nReferencePostingId As Long, nResponsibilityAreaId As Long, sBudgetKey As Variant, _
                       sDisponibilityKey As Variant, sVerificationNumber As Variant, dPostingDate As Variant, _
                       sPostingConcept As Variant, sPostingType As String, nCurrencyId As Long, _
                       nAmount As Double, nBaseAmount As Double, bProtected As Boolean)
  On Error GoTo ErrHandler
    With oPostingsRS
      .AddNew
        !ID_MOVIMIENTO_TMP = GetObjectId(sAppServer, "SEC_ID_MOVIMIENTO_TMP")
        !ID_TRANSACCION = nTransactionId
        !id_cuenta = nAccountId
        !id_cuenta_auxiliar = nSubsidiaryAccountId
        !id_sector = nSectorId
        !id_movimiento_referencia = nReferencePostingId
        !id_area_responsabilidad = nResponsibilityAreaId
        !clave_presupuestal = sBudgetKey
        !clave_disponibilidad = sDisponibilityKey
        !numero_verificacion = sVerificationNumber
        !tipo_movimiento = sPostingType
        !fecha_movimiento = dPostingDate
        !concepto_movimiento = sPostingConcept
        !id_moneda = nCurrencyId
        !monto = nAmount
        !monto_moneda_base = nBaseAmount
        !PROTEGIDO = IIf(bProtected, 1, 0)
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutPosting", Err.Number
End Sub

Private Function PutTransaction(sAppServer As String, oTransactionRS As Recordset, _
                                nGralLedgerId As Long, nSourceId As Long, _
                                nTransactionTypeId As Long, nVoucherTypeId As Long, _
                                sTransactionConcept As String, dApplicationDate As Variant, _
                                dElaborationDate As Variant, _
                                nUserId As Long, nCheckerUserId As Long) As Long
  Dim nTransactionId As Long
  '******************************************************************************************
  On Error GoTo ErrHandler
    With oTransactionRS
      nTransactionId = GetObjectId(sAppServer, "SEC_ID_TRANSACCION")
      .AddNew
        !ID_TRANSACCION = nTransactionId
        !NUMERO_TRANSACCION = "No actualizada"
        !id_mayor = nGralLedgerId
        !ID_FUENTE = nSourceId
        !ID_TIPO_TRANSACCION = nTransactionTypeId
        !ID_TIPO_POLIZA = nVoucherTypeId
        !CONCEPTO_TRANSACCION = TrimAll(sTransactionConcept)
        !FECHA_AFECTACION = dApplicationDate
        !FECHA_REGISTRO = dElaborationDate
        !ID_ELABORADA_POR = nUserId
        !ID_AUTORIZADA_POR = nCheckerUserId
        !ESTA_ABIERTA = 1
    End With
    PutTransaction = nTransactionId
  Exit Function
ErrHandler:
  RaiseError ClassId, "PutTransaction", Err.Number
End Function

Private Function GetObjectId(sAppServer As String, sSequence As String) As Long
  Dim oRecordset As New Recordset
  '****************************************************************************
  On Error GoTo ErrHandler
    With oRecordset
      .Open "SELECT " & sSequence & ".NEXTVAL AS ID FROM DUAL", _
            GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      GetObjectId = !Id
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetObjectId", Err.Number
End Function

Private Function GetPostingsRS(sAppServer As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '**************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Movimiento_Tmp WHERE id_movimiento_tmp = -1"
    oRecordset.CursorLocation = adUseClient
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set GetPostingsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetPostingsRS", Err.Number
End Function

Private Function GetTransactionsRS(sAppServer As String) As Recordset
  Dim oRecordset As New Recordset, sSQL As String
  '******************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Transaccion WHERE (id_transaccion = -1)"
    oRecordset.CursorLocation = adUseClient
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenDynamic, adLockBatchOptimistic
    Set oRecordset.ActiveConnection = Nothing
    Set GetTransactionsRS = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetTransactionsRS", Err.Number
End Function

Private Sub SaveRecordset(sAppServer As String, oTransactionsRS As Recordset)
  Dim oCon As New Connection
  '**************************************************************************
  On Error GoTo ErrHandler
    oCon.Open GetConstant(sConstantName:=sAppServer)
    With oTransactionsRS
      Set .ActiveConnection = oCon
      .UpdateBatch
      Set .ActiveConnection = Nothing
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "SaveRecordset", Err.Number
End Sub

Private Function GetAccountId(sAppServer As String, nGralLedgerId As Long, nStdAccountId As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String, nAccountId As Long
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta FROM COF_Cuenta " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(id_mayor = " & nGralLedgerId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        nAccountId = !id_cuenta
      Else
        nAccountId = 0
      End If
      .Close
    End With
    If nAccountId = 0 Then
      nAccountId = AddStdAccountToGralLedger(sAppServer, nGralLedgerId, nStdAccountId)
    End If
    GetAccountId = nAccountId
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountId", Err.Number
End Function

Private Function AddStdAccountToGralLedger(sAppServer As String, nGralLedgerId As Long, _
                                           nStdAccountId As Long) As Long
  Dim oGralLedger As New AOGralLedger.CGralLedger, sRole As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    AddStdAccountToGralLedger = oGralLedger.AssignStdAccount(sAppServer, nGralLedgerId, nStdAccountId)
  Exit Function
ErrHandler:
  RaiseError ClassId, "AddStdAccountToGralLedger", Err.Number
End Function

Private Sub PutGroupedPosting(sAppServer As String, oPostingsRS As Recordset, nTransactionId As Long, _
                              nAccountId As Long, nSubsidiaryAccountId As Long, nCurrencyId As Long, _
                              nSectorId As Long, sPostingType As String, nAmount As Double, nBaseAmount As Double, _
                              bProtected As Boolean)
  On Error GoTo ErrHandler
    With oPostingsRS
      .Filter = "(id_transaccion = " & nTransactionId & ") AND (id_cuenta = " & nAccountId & ") AND " & _
                "(id_cuenta_auxiliar = " & nSubsidiaryAccountId & ") AND (id_moneda = " & nCurrencyId & ") AND " & _
                "(id_sector = " & nSectorId & ") AND (tipo_movimiento = '" & sPostingType & "')"
      If Not (.BOF And .EOF) Then
        !monto = (!monto + nAmount)
        !monto_moneda_base = (!monto_moneda_base + nBaseAmount)
      Else
        PutPosting sAppServer, oPostingsRS, nTransactionId, nAccountId, nSubsidiaryAccountId, _
                   nSectorId, 0, 0, "", "", "", Null, "", sPostingType, _
                   nCurrencyId, nAmount, nBaseAmount, IIf(bProtected, 1, 0)
      End If
      .Filter = adFilterNone
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "PutGroupedPosting", Err.Number
End Sub

Private Function SectorRole(sAppServer As String, nStdAccountId As Long, nSectorId As Long) As String
  Dim oRecordset As New Recordset, sSQL As String, nAccountId As Long
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT rol_sector FROM COF_Mapeo_Sector " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(id_sector = " & nSectorId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        SectorRole = !rol_sector
      Else
        SectorRole = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "SectorRole", Err.Number
End Function

Private Function GetChildsCurrencies(sAppServer As String, sChildsList As String) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT DISTINCT id_moneda FROM COF_Mapeo_Moneda " & _
           "WHERE (id_cuenta_estandar IN (" & sChildsList & ") )"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sTemp) <> 0 Then
          sTemp = sTemp & ", " & !id_moneda
        Else
          sTemp = !id_moneda
        End If
        .MoveNext
      Loop
      .Close
    End With
    GetChildsCurrencies = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetChildsCurrencies", Err.Number
End Function

Private Sub GetChildsSectors(sAppServer As String, sChildsList As String, sSectors As String, sControlSectors As String)
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String, sTemp2 As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT DISTINCT id_sector, rol_sector FROM COF_Mapeo_Sector " & _
           "WHERE (id_cuenta_estandar IN (" & sChildsList & ") )"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If Len(sTemp) <> 0 Then
          sTemp = sTemp & ", " & !id_sector
          If !rol_sector = "C" Then
            sTemp2 = sTemp2 & ", " & !id_sector
          End If
        Else
          sTemp = !id_sector
          If !rol_sector = "C" Then
            sTemp2 = !id_sector
          End If
        End If
        .MoveNext
      Loop
      .Close
    End With
    sSectors = sTemp
    sControlSectors = sTemp2
  Exit Sub
ErrHandler:
  RaiseError ClassId, "GetChildsSectors", Err.Number
End Sub
