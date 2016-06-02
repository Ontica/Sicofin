VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const ClassId As String = "CInterface"

Public Function Balance(sAppServer As String, nStdAccountTypeId As Long, _
                        dFromDate As Date, dToDate As Date, Optional bIncludeGeneralLedger As Boolean = True, _
                        Optional sFromAccount As String, Optional sToAccount As String, _
                        Optional sAccounts As String, _
                        Optional bIntoTable As Boolean = False) As Recordset
Dim oBalance             As AOGLBalances.CBalances
Dim oRSTotal             As Recordset
Dim oRs                  As Recordset
Dim dInidate             As Date
Dim dEndDate             As Date
Dim dCheck               As Date
Dim nMonth               As Long
Dim nYear                As Long
Dim nDay                 As Long
Dim bIni                 As Boolean
Dim vGl                  As Variant
Dim vGLToSend            As Variant
Dim vStdAccount          As Variant
Dim nStdAcc              As Long
Dim nGeneralLedger       As Long
Dim sGL                  As String

'***************************************************************************************************
On Error GoTo ErrHandler
  nMonth = Month(dFromDate)
  nYear = Year(dFromDate)
  bIni = True
  If nStdAccountTypeId <> 0 Then
    vStdAccount = nStdAccountTypeId
    vStdAccount = Array(vStdAccount)
  Else
    vStdAccount = GetStdAccountTypes(sAppServer)
  End If
  
  For nStdAcc = LBound(vStdAccount) To UBound(vStdAccount)
    vGl = GetGLs(sAppServer, CLng(vStdAccount(nStdAcc)))
    For nGeneralLedger = LBound(vGl) To UBound(vGl)
      If bIncludeGeneralLedger Then
        vGLToSend = vGl(nGeneralLedger)
        sGL = GetObjName(sAppServer, CLng(vGLToSend), "COF_MAYOR", "NUMERO_MAYOR", "ID_MAYOR")
      Else
        vGLToSend = vGl
          sGL = "CONSOLIDADO"
      End If

        dCheck = Day(dToDate) & "/" & nMonth & "/" & nYear
        
        Do While dToDate >= dCheck
          If nMonth = 1 Then
             dInidate = "02/01/" & nYear
             dEndDate = "31/01/" & nYear
           ElseIf nMonth = 2 Then
             dInidate = "01/02/" & nYear
             dEndDate = "29/02/" & nYear
           ElseIf nMonth = 3 Then
             dInidate = "01/03/" & nYear
             dEndDate = "31/03/" & nYear
           ElseIf nMonth = 4 Then
             dInidate = "01/04/" & nYear
             dEndDate = "30/04/" & nYear
           ElseIf nMonth = 5 Then
             dInidate = "01/05/" & nYear
             dEndDate = "31/05/" & nYear
           ElseIf nMonth = 6 Then
             dInidate = "01/06/" & nYear
             dEndDate = "30/06/" & nYear
           ElseIf nMonth = 7 Then
             dInidate = "01/07/" & nYear
             dEndDate = "31/07/" & nYear
           ElseIf nMonth = 8 Then
             dInidate = "01/08/" & nYear
             dEndDate = "31/08/" & nYear
           ElseIf nMonth = 9 Then
             dInidate = "01/09/" & nYear
             dEndDate = "30/09/" & nYear
           ElseIf nMonth = 10 Then
             dInidate = "01/10/" & nYear
             dEndDate = "31/10/" & nYear
           ElseIf nMonth = 11 Then
             dInidate = "01/11/" & nYear
             dEndDate = "30/11/" & nYear
           ElseIf nMonth = 12 Then
             dInidate = "01/12/" & nYear
             dEndDate = "31/12/" & nYear
           End If
           If (Month(dToDate) = nMonth) And (Year(dToDate) = nYear) Then
             dEndDate = dToDate
           End If
           Set oBalance = New AOGLBalances.CBalances
           If bIni Then
            Set oRSTotal = oBalance.BalanceISigro(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGLToSend, dInidate, dEndDate, _
                                            sFromAccount, sToAccount, 0, 0, 0, 0, 0, Date, 4, False)
            With oRSTotal
              If Not (.BOF And .EOF) Then
                Do While Not .EOF
                  !scon_numero_mayor = sGL
                  !scon_cuenta = Replace(!scon_cuenta, "-", "-")
                  .Update
                  .MoveNext
                Loop
              End If
              .MoveFirst
            End With
            bIni = False
            nMonth = nMonth + 1
            If nMonth > 12 Then
              nMonth = 1
              nYear = nYear + 1
            End If
           Else
             Set oRs = oBalance.BalanceISigro(sAppServer, "&&&&-&&-&&-&&-&&-&&-&&", vGLToSend, dInidate, dEndDate, _
                                              sFromAccount, sToAccount, sAccounts, bWithAux:=True)
             With oRs
              If Not (.BOF And .EOF) Then
                Do While Not .EOF
                  oRSTotal.AddNew
                  oRSTotal!scon_anio = !scon_anio
                  oRSTotal!scon_mes = !scon_mes
                  oRSTotal!scon_area = !scon_area
                  oRSTotal!scon_moneda = !scon_moneda
                  oRSTotal!scon_numero_mayor = sGL
                  oRSTotal!scon_cuenta = Replace(!scon_cuenta, "-", "-")
                  oRSTotal!scon_sector = !scon_sector
                  oRSTotal!scon_auxiliar = !scon_auxiliar
                  oRSTotal!scon_fecha_ultimo_movimiento = !scon_fecha_ultimo_movimiento
                  oRSTotal!scon_saldo = !scon_saldo
                  oRSTotal!scon_moneda_origen = !scon_moneda_origen
                  oRSTotal!scon_naturaleza_cuenta = !scon_naturaleza_cuenta
                  oRSTotal!scon_saldo_promedio = !scon_saldo_promedio
                  oRSTotal!scon_monto_debito = !scon_monto_debito
                  oRSTotal!scon_monto_credito = !scon_monto_credito
                  oRSTotal!scon_saldo_anterior = !scon_saldo_anterior
                  oRSTotal!scon_empresa = !scon_empresa
                  oRSTotal.Update
                  .MoveNext
                Loop
              End If
             End With
             nMonth = nMonth + 1
              If nMonth > 12 Then
                nMonth = 1
                nYear = nYear + 1
              End If
           End If
          If nMonth = 1 Or nMonth = 3 Or nMonth = 5 Or nMonth = 7 Or nMonth = 8 Or nMonth = 10 Or nMonth = 12 Then
            nDay = 31
          ElseIf nMonth = 2 Then
            nDay = 28
          ElseIf nMonth = 4 Or nMonth = 6 Or nMonth = 9 Or nMonth = 11 Then
            nDay = 30
          End If
          dCheck = nDay & "/" & nMonth & "/" & nYear
        Loop
      If bIncludeGeneralLedger Then
        nMonth = Month(dFromDate)
        nYear = Year(dFromDate)
      Else
        Exit For
      End If
    Next nGeneralLedger
    nMonth = Month(dFromDate)
    nYear = Year(dFromDate)
    bIni = True
  Next nStdAcc
  oRSTotal.MoveFirst
  
  If bIntoTable Then
    InsertIntoTable sAppServer, oRSTotal
    Set Balance = Nothing
  Else
    Set Balance = oRSTotal
  End If
  Exit Function
ErrHandler:
  RaiseError ClassId, "Balance", Err.Number
End Function


Private Sub InsertIntoTable(sAppServer As String, oRSIn As Recordset)
  Dim oCommand  As New Command
  Dim oConn As New Connection
  Dim oRSOut As New Recordset
  Dim sSQL As String
  '**********************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sSQL = " DELETE FROM SCON_SALDOS "
    
    Set oCommand.ActiveConnection = oConn
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute

    oConn.Close
    
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sSQL = " SELECT * FROM SCON_SALDOS WHERE SCON_ANIO = 0 "
    
    oRSOut.Open sSQL, oConn, adOpenDynamic, adLockOptimistic
    
    With oRSIn
      If Not (.BOF And .EOF) Then
        oConn.BeginTrans
        Do While Not .EOF
          oRSOut.AddNew
          oRSOut!scon_anio = !scon_anio
          oRSOut!scon_mes = !scon_mes
          oRSOut!scon_area = !scon_area
          oRSOut!scon_moneda = !scon_moneda
          oRSOut!scon_numero_mayor = !scon_numero_mayor
          oRSOut!scon_cuenta = !scon_cuenta
          oRSOut!scon_sector = !scon_sector
          oRSOut!scon_auxiliar = !scon_auxiliar
          oRSOut!scon_fecha_ultimo_movimiento = !scon_fecha_ultimo_movimiento
          oRSOut!scon_saldo = !scon_saldo
          oRSOut!scon_moneda_origen = !scon_moneda_origen
          oRSOut!scon_naturaleza_cuenta = !scon_naturaleza_cuenta
          oRSOut!scon_saldo_promedio = !scon_saldo_promedio
          oRSOut!scon_monto_debito = !scon_monto_debito
          oRSOut!scon_monto_credito = !scon_monto_credito
          oRSOut!scon_saldo_anterior = !scon_saldo_anterior
          oRSOut!scon_empresa = !scon_empresa
          oRSOut.Update
          .MoveNext
        Loop
        oConn.CommitTrans
      End If
    End With
    oRSOut.Close
    oRSIn.Close
  Exit Sub
ErrHandler:
  oConn.RollbackTrans
  RaiseError ClassId, "InsertIntoTable", Err.Number
End Sub


Public Function ExchangeRates(sAppServer As String, dFromDateSigro As Date, dToDateSigro As Date) As Recordset
  Dim oConn As New Connection, oRecordset As New Recordset, sSQL As String
  Dim sWhereToCurrency  As String
  Dim dToDate, dFromDate As String
  '**********************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    dFromDate = Format(dFromDateSigro, "DD/MMM/YYYY")
    dToDate = Format(dToDateSigro, "DD/MMM/YYYY")
    
    sSQL = " SELECT TO_CURRENCY_ID CURRENCY_ID, EXCHANGE_RATE, OBJECT_NAME " & _
           " FROM AO_EXCHANGE_RATES, AO_OBJECTS " & _
           " WHERE (AO_EXCHANGE_RATES.EXCHANGE_RATE_TYPE_ID = AO_OBJECTS.OBJECT_ID) AND " & _
           "       (FROM_DATE >= '" & dFromDate & "') AND " & _
           "       (TO_DATE <= '" & dToDate & "') AND " & _
           "       (FROM_CURRENCY_ID = 1) "
  
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set ExchangeRates = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "ExchangeRates", Err.Number
End Function

Public Function StandardAccounts(sAppServer As String, nStdAccountTypeId As Long, dDateSigro As Date, _
                         Optional sFromAccount As String, Optional sToAccount As String) As Recordset
  Dim oRs         As New ADODB.Recordset
  Dim oCon        As New ADODB.Connection
  Dim sSQL        As String
  Dim sSQL_A      As String
  Dim sSQL_B      As String
  Dim sWhere      As String
  Dim dDate       As String
  '****************************************************
  On Error GoTo ErrHandler
    oCon.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oCon.Open
    dDate = Format$(dDateSigro, "DD/MMM/YYYY")

    If (sFromAccount = "") And (sToAccount = "") Then
      sWhere = " "
    ElseIf (sFromAccount <> "") And (sToAccount <> "") Then
      sWhere = " WHERE (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sFromAccount & "' AND '" & sToAccount & "') "
    ElseIf (sFromAccount <> "") And (sToAccount = "") Then
      sWhere = " WHERE (NUMERO_CUENTA_ESTANDAR <= '" & sFromAccount & "') "
    End If

   
    sSQL_A = " SELECT COF_CUENTA_ESTANDAR_HIST.NUMERO_CUENTA_ESTANDAR, " & _
             "        COF_CUENTA_ESTANDAR_HIST.NOMBRE_CUENTA_ESTANDAR, ('00') Clave_Sector, " & _
             "        COF_CUENTA_ESTANDAR_HIST.NATURALEZA, " & _
             "        COF_CUENTA_ESTANDAR_HIST.DESCRIPCION, " & _
             "        COF_CUENTA_ESTANDAR_HIST.ROL_CUENTA, " & _
             "        OBJECT_NAME TIPO_CUENTA " & _
             " FROM COF_CUENTA_ESTANDAR_HIST, AO_OBJECTS " & _
             " WHERE (COF_CUENTA_ESTANDAR_HIST.ID_TIPO_CUENTA = AO_OBJECTS.OBJECT_ID(+)) AND " & _
             "       (COF_CUENTA_ESTANDAR_HIST.ROL_CUENTA IN ('S','P','C')) AND " & _
             "       (COF_CUENTA_ESTANDAR_HIST.ID_TIPO_CUENTAS_STD = " & nStdAccountTypeId & ") AND " & _
             "       (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dDate & "') AND " & _
             "       (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dDate & "') "

    sSQL_B = " SELECT COF_CUENTA_ESTANDAR_HIST.NUMERO_CUENTA_ESTANDAR, " & _
             "        COF_CUENTA_ESTANDAR_HIST.NOMBRE_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
             "        COF_CUENTA_ESTANDAR_HIST.NATURALEZA, " & _
             "        COF_CUENTA_ESTANDAR_HIST.DESCRIPCION, " & _
             "        COF_MAPEO_SECTOR.ROL_SECTOR ROL_CUENTA, " & _
             "        OBJECT_NAME TIPO_CUENTA " & _
             " FROM COF_CUENTA_ESTANDAR_HIST, AO_OBJECTS, COF_MAPEO_SECTOR, COF_SECTOR " & _
             " WHERE (COF_CUENTA_ESTANDAR_HIST.ID_TIPO_CUENTA = AO_OBJECTS.OBJECT_ID(+)) AND " & _
             "       (COF_CUENTA_ESTANDAR_HIST.ROL_CUENTA = 'X') AND " & _
             "       (COF_CUENTA_ESTANDAR_HIST.ID_TIPO_CUENTAS_STD = " & nStdAccountTypeId & " ) AND " & _
             "       (COF_CUENTA_ESTANDAR_HIST.ID_CUENTA_ESTANDAR = COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR) AND " & _
             "       (COF_MAPEO_SECTOR.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
             "       (TRUNC(TO_DATE(COF_CUENTA_ESTANDAR_HIST.FECHA_INICIO)) <= '" & dDate & "') AND " & _
             "       (TRUNC(TO_DATE(COF_CUENTA_ESTANDAR_HIST.FECHA_FIN)) >= '" & dDate & "') AND " & _
             "       (TRUNC(TO_DATE(COF_MAPEO_SECTOR.FECHA_INICIO)) <= '" & dDate & "') AND " & _
             "       (TRUNC(TO_DATE(COF_MAPEO_SECTOR.FECHA_FIN)) >= '" & dDate & "') "

    sSQL = " SELECT * " & _
           " FROM ( " & _
             sSQL_A & _
           " UNION " & _
             sSQL_B & _
           " ) A " & _
           sWhere & _
           " ORDER BY NUMERO_CUENTA_ESTANDAR "
           
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oCon, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set StandardAccounts = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "StandardAccounts", Err.Number
End Function

Public Function SubsidiaryAccounts(sAppServer As String) As Recordset
  Dim oRs         As New ADODB.Recordset
  Dim oCon        As New ADODB.Connection
  Dim sSQL        As String
  '******************************************************************
  On Error GoTo ErrHandler
    oCon.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oCon.Open
    
    sSQL = " SELECT COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR, " & _
           "        COF_CUENTA_AUXILIAR.NOMBRE_CUENTA_AUXILIAR, " & _
           "        COF_CUENTA_AUXILIAR.DESCRIPCION, " & _
           "        OBJECT_NAME TIPO_CUENTA_AUXILIAR " & _
           " FROM COF_CUENTA_AUXILIAR, COF_MAYOR_AUXILIAR, AO_OBJECTS " & _
           " WHERE (COF_CUENTA_AUXILIAR.ID_MAYOR_AUXILIAR = COF_MAYOR_AUXILIAR.ID_MAYOR_AUXILIAR) AND " & _
           "       (COF_MAYOR_AUXILIAR.ID_TIPO_MAYOR_AUXILIAR = AO_OBJECTS.OBJECT_ID) " & _
           " ORDER BY NUMERO_CUENTA_AUXILIAR "
           
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oCon, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set SubsidiaryAccounts = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "SubsidiaryAccounts", Err.Number
End Function

Private Function GetGLs(sAppServer As String, nStdAccountType As Long) As Variant
  Dim oConn   As New Connection
  Dim oRs     As New Recordset
  Dim sSQL    As String
  Dim vGl
  Dim i       As Long
  '**********************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

    sSQL = "SELECT ID_MAYOR FROM COF_MAYOR WHERE ID_TIPO_CUENTAS_STD = " & nStdAccountType

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
  
    ReDim vGl(0)
    With oRs
    Do While Not .EOF
      ReDim Preserve vGl(i)
      vGl(i) = !id_mayor
      i = i + 1
      .MoveNext
    Loop
  End With
  
  GetGLs = vGl
  'Array(1, 2, 3)
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetGLs", Err.Number
End Function

Private Function GetStdAccountTypes(sAppServer As String) As Variant
  Dim oConn   As New Connection
  Dim oRs     As New Recordset
  Dim sSQL    As String
  '*****************************************************************
On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

  sSQL = "SELECT DISTINCT(ID_TIPO_CUENTAS_STD) FROM COF_CUENTA_ESTANDAR_HIST "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
  
  GetStdAccountTypes = Array(1, 2)
Exit Function
ErrHandler:
  RaiseError ClassId, "GetStdAccountTypes", Err.Number
End Function

Private Function GetObjName(sAppServer As String, nId As Long, sTable As String, _
                            sFieldName As String, sFieldCondition As String, _
                            Optional sFieldName2 As String = "") As String
  Dim oConn    As New ADODB.Connection
  Dim oRs      As New ADODB.Recordset
   
  Dim sObjName As String
  Dim sSQL     As String
  '************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConn.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConn.Open
    
    If (Len(sFieldName2) = 0) Then
      sObjName = sFieldName
    ElseIf (Len(sFieldName2) <> 0) Then
      sObjName = sFieldName & " , " & sFieldName2
    End If
    
    sSQL = "SELECT " & sObjName & " FROM " & sTable & " WHERE " & sFieldCondition & " = " & nId
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      If (.BOF And .EOF) Then
        sObjName = " "
      Else
        If (Len(sFieldName2) = 0) Then
          sObjName = .Fields(0)
        ElseIf (Len(sFieldName2) <> 0) Then
          sObjName = ZeroFill(.Fields(0), 4) & "  " & .Fields(1)
        End If
      End If
    End With
    GetObjName = sObjName
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetObjName", Err.Number
End Function

Private Function ZeroFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
    ZeroFill = Mid("00000000000000000000000", 1, nLen - Len(sStr)) & sStr
  Exit Function
ErrorHandler:
  RaiseError ClassId, "ZeroFill", Err.Number
End Function
