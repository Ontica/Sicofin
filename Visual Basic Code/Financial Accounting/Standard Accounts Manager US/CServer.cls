VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Software Components                    Sistema : Financial Accounting              *
'* Componente : Standard Account Manager US (EFAStdActUS)       Parte   : User services                     *
'* Clase      : CManager                                        Patrón  : Stateless Class                   *
'* Fecha      : 28/Febrero/2002                                 Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Clase ráiz del componente "Standard Account Manager US".                                    *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CManager"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function CboStdAccountCatalogues(sAppServer As String, Optional nSelectedItemId As Long = 0) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT object_id, object_name " & _
           "FROM AO_Objects " & _
           "WHERE (entity_id = 1) AND (deleted = 0) " & _
           "ORDER BY presentation_order"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !object_id & ">" & !object_name & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    If nSelectedItemId <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItemId & ">", _
                             "<OPTION SELECTED value=" & nSelectedItemId & ">")
    Else
      sTemp = "<OPTION value=0>( Seleccionar catálogo estándar )</OPTION>" & vbCrLf & sTemp
    End If
    CboStdAccountCatalogues = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboStdAccountCatalogues", Err.Number
End Function

Public Function CboStdAccountNature(Optional sSelectedItem As String) As String
  Dim sTemp As String
  '****************************************************************************
  On Error GoTo ErrHandler
    sTemp = "<OPTION value='D'>Deudora</OPTION>" & vbCrLf & _
            "<OPTION value='A'>Acreedora</OPTION>"
    If (Len(sSelectedItem) <> 0) Then
      sTemp = Replace(sTemp, "<OPTION value='" & sSelectedItem & "'>", _
                             "<OPTION SELECTED value='" & sSelectedItem & "'>")
    End If
    CboStdAccountNature = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboStdAccountNature", Err.Number
End Function

Public Function CboStdAccountRoles(Optional sSelectedItem As String) As String
  Dim sTemp As String
  '***************************************************************************
  On Error GoTo ErrHandler
    sTemp = "<OPTION value='S'>Sumaria</OPTION>" & vbCrLf & _
            "<OPTION value='X'>Sectorizada (con o sin auxiliares)</OPTION>" & vbCrLf & _
            "<OPTION value='C'>De control (maneja auxiliares)</OPTION>" & vbCrLf & _
            "<OPTION value='P'>De detalle (no maneja auxiliares)</OPTION>"
    If (Len(sSelectedItem) <> 0) Then
      sTemp = Replace(sTemp, "<OPTION value='" & sSelectedItem & "'>", _
                             "<OPTION SELECTED value='" & sSelectedItem & "'>")
    End If
    CboStdAccountRoles = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboStdAccountRoles", Err.Number
End Function

Public Function FormatCurrenciesList(sAppServer As String, sCurrenciesList As String) As String
  FormatCurrenciesList = sCurrenciesList
End Function

Public Function FormatSectorsList(sAppServer As String, sSectorsList As String, sSectorsRoles As String) As String
  FormatSectorsList = sSectorsList & ":" & sSectorsRoles
End Function

Public Function FormatWildCharsList(sList As String) As String
  Dim sTemp As String
  '***********************************************************
  On Error Resume Next
    sTemp = Trim$(sList)
    sTemp = Replace(sTemp, "%", "*")
    sTemp = Replace(sTemp, "_", "?")
    sTemp = Replace(sTemp, ", ", ",")
    sTemp = Replace(sTemp, " ,", ",")
    sTemp = Replace(sTemp, "**", "*")
    sTemp = Replace(sTemp, ",,", ",")
    If InStr(1, sTemp, ",*,") Then sTemp = "*"
    If (Len(sTemp) > 1) And (Left$(sTemp, 2) = "*,") Then sTemp = "*"
    If (Len(sTemp) > 1) And (Right$(sTemp, 2) = ",*") Then sTemp = "*"
    FormatWildCharsList = sTemp
  Exit Function
End Function

Public Function CboStdAccountTypes(sAppServer As String, Optional nSelectedItemId As Long = 0) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '**********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT Object_Id, Object_Name FROM AO_Objects " & _
           "WHERE (Entity_Id = 2) ORDER BY Object_Name"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = sTemp & "<OPTION value=" & !object_id & ">" & !object_name & "</OPTION>" & vbCrLf
        .MoveNext
      Loop
      .Close
    End With
    If nSelectedItemId <> 0 Then
      sTemp = Replace(sTemp, "<OPTION value=" & nSelectedItemId & ">", _
                             "<OPTION SELECTED value=" & nSelectedItemId & ">")
    End If
    CboStdAccountTypes = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "CboStdAccountTypes", Err.Number
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

Public Function GetParentIdWithNumber(sAppServer As String, nStdAccountTypeId As Long, _
                                      sStdAccountNumber As String) As Long
  Dim oStdAccount As New EFAStdActBS.CStdAccount
  '*************************************************************************************
  On Error GoTo ErrHandler
    GetParentIdWithNumber = oStdAccount.GetParent(sAppServer, nStdAccountTypeId, _
                                                  sStdAccountNumber).Fields("id_cuenta_estandar")
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetParentIdWithNumber", Err.Number
End Function

Public Function GetStdAccountCategory(sAppServer As String, nCategoryId As Long) As Recordset
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '******************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM AO_Objects WHERE (Entity_Id = 1) ORDER BY Object_Name"
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetStdAccountCategory = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountCategory", Err.Number
End Function

Public Function GetStdAccountId(sAppServer As String, nStdAccountTypeId As Long, _
                                sStdAccountNumber As String) As Long
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '******************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_cuenta_estandar " & _
           "FROM COF_Cuenta_Estandar_Hist " & _
           "WHERE (numero_cuenta_estandar = '" & sStdAccountNumber & "') AND " & _
           "(id_tipo_cuentas_std = " & nStdAccountTypeId & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GetStdAccountId = !id_cuenta_estandar
      Else
        GetStdAccountId = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountId", Err.Number
End Function

Public Function StdAccountParentId(sAppServer As String, nStdAccountId As Long, _
                                   Optional dDate As Variant) As Long
  Dim oRecordset As New Recordset, sStdAccountNumber As String, nStdAccountTypeId As Long
  '**************************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
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
    sStdAccountNumber = StdAccountParentNumber(sAppServer, nStdAccountTypeId, sStdAccountNumber)
    If Len(sStdAccountNumber) <> 0 Then
      StdAccountParentId = GetStdAccountId(sAppServer, nStdAccountTypeId, sStdAccountNumber)
    Else
      StdAccountParentId = 0
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountParentId", Err.Number
End Function

Public Function CurrenciesList(sAppServer As String, Optional nStdAccountId As Long, _
                               Optional bCheckDisabled As Boolean = False, Optional dDate As Date) As String
  Const cCell = "<TD><INPUT type=checkbox name=chkCurrencies tagName='[<@CURRENCY_KEY@>] <@CURRENCY_NAME@>' value=<@CURRENCY_ID@> <@CHECKED@>>[<@CURRENCY_KEY@>] <@CURRENCY_NAME@></TD>"
  Dim oRecordset As New Recordset, oSelectedRS As New Recordset, sSQL As String
  Dim sTemp As String, sHTML As String, nCurrentCell As Long
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM AO_Currencies " & _
           "WHERE (deleted = 0) " & _
           "ORDER BY o_id_moneda"
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    sSQL = "SELECT COF_Mapeo_Moneda.id_moneda " & _
           "FROM COF_Mapeo_Moneda " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    oSelectedRS.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
    With oRecordset
      nCurrentCell = 0
      Do While Not .EOF
        nCurrentCell = nCurrentCell + 1
        sTemp = Replace(cCell, "<@CURRENCY_ID@>", !currency_id)
        sTemp = Replace(sTemp, "<@CURRENCY_KEY@>", !o_id_moneda)
        sTemp = Replace(sTemp, "<@CURRENCY_NAME@>", !currency_name)
        oSelectedRS.Filter = "id_moneda = " & !currency_id
        If Not bCheckDisabled Then
          If Not (oSelectedRS.BOF And oSelectedRS.EOF) Then
            sTemp = Replace(sTemp, "<@CHECKED@>", "checked")
          Else
            sTemp = Replace(sTemp, "<@CHECKED@>", "")
          End If
        Else
          If Not (oSelectedRS.BOF And oSelectedRS.EOF) Then
            sTemp = Replace(sTemp, "<@CHECKED@>", "checked disabled")
          Else
            sTemp = Replace(sTemp, "<@CHECKED@>", "")
          End If
        End If
        oSelectedRS.Filter = adFilterNone
        If (nCurrentCell Mod 2) = 1 Then
          sHTML = sHTML & "<TR>" & sTemp
        Else
          sHTML = sHTML & sTemp & "</TR>" & vbCrLf
        End If
        .MoveNext
      Loop
      .Close
    End With
    If (nCurrentCell Mod 2) = 1 Then
      sHTML = sHTML & "</TR>" & vbCrLf
    End If
    CurrenciesList = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "CurrenciesList", Err.Number
End Function

Public Function ResponsibilityAreasList(sAppServer As String, nStdAccountId As Long, _
                                        Optional dDate As Variant) As String
  Dim oRecordset As New Recordset, sSQL As String, sTemp As String
  '***********************************************************************************
  On Error GoTo ErrHandler
    If IsMissing(dDate) Then
      dDate = Date
    End If
    sSQL = "SELECT patron_area FROM COF_Mapeo_Area " & _
           "WHERE (id_cuenta_estandar = " & nStdAccountId & ") AND " & _
           "(fecha_inicio <= '" & dDate & "') AND ('" & dDate & "' <= fecha_fin) " & _
           "ORDER BY patron_area"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If (Len(sTemp) <> 0) Then
          sTemp = sTemp & "," & !patron_area
        Else
          sTemp = !patron_area
        End If
        .MoveNext
      Loop
    End With
    ResponsibilityAreasList = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "ResponsibilityAreasList", Err.Number
End Function

Public Function ResponsibilityAreasListWithArray(sAppServer As String, sAreasList As String, _
                                                 Optional bCheckBoxes As Boolean = False) As String
  Const cRow1 = "<TR><TD><INPUT type=checkbox name=chkAreas onclick='return chkAreas_onclick();' value=<@AREA_KEY@>></TD><TD>[<@AREA_KEY@>]</TD><TD><@AREA_NAME@></TD></TR>" & vbCrLf
  Const cRow2 = "<TR><TD>[<@AREA_KEY@>]</TD><TD><@AREA_NAME@></TD></TR>" & vbCrLf
  Dim oRecordset As New Recordset, oSelectedRS As New Recordset, sSQL As String
  Dim sTemp As String, sHTML As String, i As Long, sLike As String, aAreasKeys As Variant
  '*************************************************************************************************
  On Error GoTo ErrHandler
    sAreasList = Replace(sAreasList, "*", "%")
    sAreasList = Replace(sAreasList, "?", "_")
    aAreasKeys = Split(sAreasList, ",")
    For i = LBound(aAreasKeys) To UBound(aAreasKeys)
      If Len(Trim(aAreasKeys(i))) <> 0 Then
        If (Len(sLike) = 0) Then
          sLike = "( (participantKey LIKE '" & Trim(aAreasKeys(i)) & "') "
        Else
          sLike = sLike & "OR (participantKey LIKE '" & Trim(aAreasKeys(i)) & "') "
        End If
      End If
    Next i
    If Len(sLike) <> 0 Then
      sLike = sLike & ") "
    End If
    
    sSQL = "SELECT participantId, participantKey, participantName FROM MHParticipants " & _
           "WHERE (participantType = 'O') AND (status <> 'D') " & _
           IIf(Len(sLike) <> 0, " AND " & sLike, "") & _
           "ORDER BY participantKey, participantName"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = Replace(IIf(bCheckBoxes, cRow1, cRow2), "<@AREA_ID@>", !participantId)
        sTemp = Replace(sTemp, "<@AREA_KEY@>", !participantKey)
        sTemp = Replace(sTemp, "<@AREA_NAME@>", Left$(!participantName, 52))
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    ResponsibilityAreasListWithArray = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "ResponsibilityAreasListWithArray", Err.Number
End Function

Public Function RoleDescription(sRoleKey As String) As String
  On Error GoTo ErrHandler
    Select Case sRoleKey
      Case "S"
        RoleDescription = "Sumaria"
      Case "C"
        RoleDescription = "Control (maneja auxiliares)"
      Case "P"
        RoleDescription = "Detalle (no maneja auxiliares)"
      Case "X"
        RoleDescription = "Sectorizada (con o sin auxiliares)"
      Case Else
        RoleDescription = "¡Desconocido!"
    End Select
  Exit Function
ErrHandler:
  RaiseError ClassId, "RoleDescription", Err.Number
End Function

Public Function ChildsSectorsList(sAppServer As String, Optional nStdAccountId As Long, _
                                  Optional bCheckDisabled As Boolean = False, Optional dDate As Date) As String
  Const cRow = "<TR><TD><INPUT type=checkbox name=chkSectors tagName='[<@SECTOR_KEY@>] <@SECTOR_NAME@>' value=<@SECTOR_ID@> <@CHECKED@>>[<@SECTOR_KEY@>] <@SECTOR_NAME@></TD>" & vbCrLf & _
               "<TD><@ROL_SECTOR@></TD></TR>"
  Dim oRecordset As New Recordset, oSelectedRS As New Recordset, sSQL As String
  Dim oStdAccount As New EFAStdActBS.CStdAccount
  Dim sTemp As String, sHTML As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Sector WHERE (eliminado = 0) ORDER BY clave_sector "
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    
    sSQL = oStdAccount.GetChildsStdAccounts(sAppServer, nStdAccountId)
    sSQL = "SELECT * FROM COF_Mapeo_Sector WHERE (id_cuenta_estandar IN (" & sSQL & ") )"
    oSelectedRS.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
    With oRecordset
      Do While Not .EOF
        sTemp = Replace(cRow, "<@SECTOR_ID@>", !id_sector)
        sTemp = Replace(sTemp, "<@SECTOR_KEY@>", !clave_sector)
        sTemp = Replace(sTemp, "<@SECTOR_NAME@>", !nombre_sector)
        oSelectedRS.Filter = "id_sector = " & !id_sector
        If Not (oSelectedRS.BOF And oSelectedRS.EOF) Then
          sTemp = Replace(sTemp, "<@CHECKED@>", IIf(bCheckDisabled, "checked disabled", "checked"))
          sTemp = Replace(sTemp, "<@ROL_SECTOR@>", IIf(oSelectedRS!rol_sector = "P", "Detalle", "Control"))
        Else
          sTemp = Replace(sTemp, "<@CHECKED@>", "")
          sTemp = Replace(sTemp, "<@ROL_SECTOR@>", "")
        End If
        oSelectedRS.Filter = adFilterNone
        sTemp = Replace(sTemp, "<@CHECKED@>", "")
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    ChildsSectorsList = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "ChildsSectorsList", Err.Number
End Function

Public Function SectorsList(sAppServer As String, Optional nStdAccountId As Long, _
                            Optional bCheckDisabled As Boolean = False, Optional dDate As Date) As String
  Const cRow = "<TR><TD><INPUT type=checkbox name=chkSectors tagName='[<@SECTOR_KEY@>] <@SECTOR_NAME@>' value=<@SECTOR_ID@> <@CHECKED@>>[<@SECTOR_KEY@>] <@SECTOR_NAME@></TD>" & vbCrLf & _
               "<TD><@ROL_SECTOR@></TD></TR>"
  Dim oRecordset As New Recordset, oSelectedRS As New Recordset, sSQL As String
  Dim sTemp As String, sHTML As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Sector WHERE (eliminado = 0) ORDER BY clave_sector "
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    
    sSQL = "SELECT * FROM COF_Mapeo_Sector WHERE (id_cuenta_estandar = " & nStdAccountId & ")"
    oSelectedRS.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
    With oRecordset
      Do While Not .EOF
        sTemp = Replace(cRow, "<@SECTOR_ID@>", !id_sector)
        sTemp = Replace(sTemp, "<@SECTOR_KEY@>", !clave_sector)
        sTemp = Replace(sTemp, "<@SECTOR_NAME@>", !nombre_sector)
        oSelectedRS.Filter = "id_sector = " & !id_sector
        If Not (oSelectedRS.BOF And oSelectedRS.EOF) Then
          sTemp = Replace(sTemp, "<@CHECKED@>", IIf(bCheckDisabled, "checked disabled", "checked"))
          sTemp = Replace(sTemp, "<@ROL_SECTOR@>", IIf(oSelectedRS!rol_sector = "P", "Detalle", "Control"))
        Else
          sTemp = Replace(sTemp, "<@CHECKED@>", "")
          sTemp = Replace(sTemp, "<@ROL_SECTOR@>", "")
        End If
        oSelectedRS.Filter = adFilterNone
        sTemp = Replace(sTemp, "<@CHECKED@>", "")
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    SectorsList = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "SectorsList", Err.Number
End Function

Public Function SectorsListWithArray(sAppServer As String, sSectorsList As String) As String
  Const cRow = "<INPUT type=checkbox name=chkSectors tagName='[<@SECTOR_KEY@>] <@SECTOR_NAME@>' value=<@SECTOR_ID@> <@CHECKED@>>[<@SECTOR_KEY@>] <@SECTOR_NAME@><br>"
  Dim oRecordset As New Recordset, oSelectedRS As New Recordset, sSQL As String
  Dim sTemp As String, sHTML As String, aSectorsIds As Variant, i As Long
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT * FROM COF_Sector WHERE (eliminado = 0) ORDER BY clave_sector "
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    
    With oRecordset
      aSectorsIds = Split(sSectorsList, ",")
      For i = LBound(aSectorsIds) To UBound(aSectorsIds)
        If Len(Trim(aSectorsIds(i))) <> 0 Then
          .Filter = "id_sector = " & CLng(Trim(aSectorsIds(i)))
          If Not (.BOF And .EOF) Then
            sTemp = Replace(cRow, "<@SECTOR_ID@>", !id_sector)
            sTemp = Replace(sTemp, "<@SECTOR_KEY@>", !clave_sector)
            sTemp = Replace(sTemp, "<@SECTOR_NAME@>", !nombre_sector)
            sTemp = Replace(sTemp, "<@CHECKED@>", "")
            sHTML = sHTML & sTemp
          End If
          .Filter = adFilterNone
        End If
        'sTemp = Replace(sTemp, "<@CHECKED@>", "")
      Next i
    End With
    SectorsListWithArray = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "SectorsListWithArray", Err.Number
End Function

Public Function TblStdAccountCategories(sAppServer As String, Optional sOrderBy As String) As String
  Const cRow = "<TR><TD nowrap><A href='' onclick='return callEditor(2, <@ID@>);'><@NAME@></A></TD>" & vbCrLf & _
               "<TD><@DESCRIPTION@></b></TD>" & vbCrLf & _
               "<TD nowrap><@PATTERN@></b></TD></TR>" & vbCrLf
  Dim oRecordset As New Recordset
  Dim sSQL As String, sHTML As String, sTemp As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT AO_Objects.* " & _
           "FROM AO_Objects " & _
           "WHERE (AO_Objects.Entity_Id = 1) AND (Deleted = 0) " & _
           "ORDER BY "
    If Len(sOrderBy) = 0 Then
      sSQL = sSQL & "presentation_order"
    Else
      sSQL = sSQL & sOrderBy
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        sTemp = Replace(cRow, "<@ID@>", !object_id)
        sTemp = Replace(sTemp, "<@NAME@>", !object_name)
        sTemp = Replace(sTemp, "<@DESCRIPTION@>", IIf(IsNull(!Object_Description), "&nbsp;", !Object_Description))
        sTemp = Replace(sTemp, "<@PATTERN@>", !Object_Key)
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    TblStdAccountCategories = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "TblStdAccountCategories", Err.Number
End Function

Public Function TblStdAccounts(sAppServer As String, nStdAccountCategory As Long, dDate As Variant, _
                               Optional sFilter As String, Optional sOrderBy As String) As String
  Const cRow = "<TR><TD valign=top nowrap><A href='' onclick='return callEditor(2, <@ID@>);'><@NUMBER@></A></TD>" & vbCrLf & _
               "<TD valign=top width=60%><@NAME@></TD>" & vbCrLf & _
               "<TD valign=top nowrap><@ROLE@></TD>" & vbCrLf & _
               "<TD valign=top nowrap><@TYPE@></TD>" & vbCrLf & _
               "<TD valign=top nowrap><@NATURE@></TD>" & vbCrLf & _
               "<TD valign=top nowrap>" & _
               "<A href='' onclick='return callEditor(3, <@ID@>);'>Monedas</A> / " & _
               "<A href='' onclick='return callEditor(4, <@ID@>);'>Sectores</A> / " & _
               "<A href='' onclick='return callEditor(5, <@ID@>);'>Saldos</A> / " & _
               "<A href='' onclick='return callEditor(6, <@ID@>);'>Mayores</A> / " & _
               "<A href='' onclick='return callEditor(7, <@ID@>);'>Historia</A>" & _
               "</TD></TR>" & vbCrLf
  Dim oRecordset As New Recordset
  Dim sSQL As String, sHTML As String, sTemp As String, sRole As String, sCurrentSerie As String
  '*****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT COF_Cuenta_Estandar.*, AO_Objects.Object_Name " & _
           "FROM COF_Cuenta_Estandar, AO_Objects " & _
           "WHERE (COF_Cuenta_Estandar.Id_Tipo_Cuentas_Std = " & nStdAccountCategory & ") AND " & _
           "(COF_Cuenta_Estandar.Id_Tipo_Cuenta = AO_Objects.Object_Id) " & _
           IIf(Len(sFilter) <> 0, "AND (" & sFilter & ") ", "") & _
           "ORDER BY "
    If Len(sOrderBy) = 0 Then
      sSQL = sSQL & "Numero_Cuenta_Estandar"
    Else
      sSQL = sSQL & sOrderBy
    End If
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      Do While Not .EOF
        If (sCurrentSerie <> Left$(!numero_cuenta_estandar, 4)) Then     'OOJJOO
          sCurrentSerie = Left$(!numero_cuenta_estandar, 4)              'OOJJOO
          sHTML = sHTML & "<TR bgcolor=khaki><TD colspan=5><b>" & sCurrentSerie & "</b></TD>" & _
                          "<TD align=right>" & _
                          "<A href='' onclick='return callEditor(1, " & nStdAccountCategory & ");'><b>Agregar cuenta</b></A></TD></TR>"
        End If
        sRole = !rol_cuenta
        sTemp = Replace(cRow, "<@ID@>", !id_cuenta_estandar)
        sTemp = Replace(sTemp, "<@NUMBER@>", IIf(sRole <> "S", "<b>" & !numero_cuenta_estandar & "</b>", !numero_cuenta_estandar))
        sTemp = Replace(sTemp, "<@NAME@>", !Nombre_Cuenta_Estandar)
        sTemp = Replace(sTemp, "<@ROLE@>", IIf(sRole = "S", "Sumaria", IIf(sRole = "C", "Control", IIf(sRole = "P", "Detalle", "Sectorizada"))))
        sTemp = Replace(sTemp, "<@TYPE@>", !object_name)
        sTemp = Replace(sTemp, "<@NATURE@>", IIf(!Naturaleza = "D", "Deudora", "Acreedora"))
        sHTML = sHTML & sTemp
        .MoveNext
      Loop
      .Close
    End With
    TblStdAccounts = sHTML
  Exit Function
ErrHandler:
  RaiseError ClassId, "TblStdAccounts", Err.Number
End Function

Public Function NatureDescription(sNature As String) As String
  On Error GoTo ErrHandler
    Select Case sNature
      Case "D"
        NatureDescription = "Deudora"
      Case "A"
        NatureDescription = "Acreedora"
      Case Else
        NatureDescription = ""
    End Select
  Exit Function
ErrHandler:
  RaiseError ClassId, "NatureDescription", Err.Number
End Function

Public Function NotAddedAccountsList(sAppServer As String, nStdAccountTypeId As Long, sStdAccountNumber, _
                                     Optional ByRef nCount As Long) As String

End Function

Public Function TypeDescription(sAppServer As String, nStdAccountType As Long) As String
  Dim oRecordset As New Recordset, sSQL As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT object_name FROM AO_Objects " & _
           "WHERE (object_id = " & nStdAccountType & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
      If Not (.BOF And .EOF) Then
        TypeDescription = !object_name
      Else
        TypeDescription = ""
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "TypeDescription", Err.Number
End Function

Private Function StdAccountParentNumber(sAppServer, nStdAccountTypeId, sStdAccountNumber) As String
  On Error GoTo ErrHandler
    If Len(sStdAccountNumber) > 4 Then
      StdAccountParentNumber = Left$(sStdAccountNumber, Len(sStdAccountNumber) - 3)
    Else
      StdAccountParentNumber = ""
    End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountParentNumber", Err.Number
End Function
