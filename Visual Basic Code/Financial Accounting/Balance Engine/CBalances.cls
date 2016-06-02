VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Empiria® ***********************************************************************************************
'*                                                                                                          *
'* Solución   : Empiria® Components                             Sistema : Financial Accounting              *
'* Componente : Balance Engine (EFABalanceEngine)               Parte   : Business services                 *
'* Clase      : CBalances                                       Patrón  : Stateless Class                   *
'* Fecha      : 31/Diciembre/2001                               Versión : 1.0       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Proporciona las balanzas de comprobación del sistema de contabilidad finanaciera.           *
'*              Las balanzas son regresadas como Recordset de ADO 2.5 o como archivos XML.                  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 1999-2002. **
Option Explicit
Private Const ClassId As String = "CBalances"

Private Const cBalancesTableName As String = "TMP_Balanzas"
Private Const cAverageBalancesTableName As String = "TMP_Balanzas_Prom"

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function AverageBalancesRS(sAppServer As String, sAccountPattern As String, vGralLedgers As Variant, _
                           bConsolidateGralLedgers As Boolean, bReturnOnlyLastLevels As Boolean, _
                           bReturnSubsAccounts As Boolean, dFromDate As Date, dToDate As Date, _
                           Optional sFromAccount As String = "", Optional sToAccount As String = "", _
                           Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                           Optional nCurrencyType As Long = 0, Optional nAccountType As Long = 0, _
                           Optional sAccountNature As String = "", Optional nExcRateType As Long = 0, _
                           Optional nExcRateCurrency As Long = 0, Optional dExcRateDate As Date, _
                           Optional nBalancesToDisplay As Long = 4) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oConnection As New Connection, oCommand As New Command, oRecordset As New Recordset
  Dim sGralLedgersFilter As String, sTransactionsFilter As String, sStdAccountsFilter As String
  Dim sBalancesFilter As String, sSourceFilter As String, nBalancesStdAccountType As Long
  Dim sSQL As String, sStdAccts As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    
    If (IsArray(vGralLedgers)) Then
      nBalancesStdAccountType = GralLegerStdAccountType(sAppServer, CLng(vGralLedgers(0)))
      sGralLedgersFilter = "id_mayor IN (" & Join(vGralLedgers, ",") & ")"
    Else
      nBalancesStdAccountType = GralLegerStdAccountType(sAppServer, CLng(vGralLedgers))
      sGralLedgersFilter = "id_mayor = " & vGralLedgers
    End If
    If (nBalancesToDisplay = 5) Then
      bReturnOnlyLastLevels = True
    End If
    sStdAccountsFilter = StdAccountsFilterSQLStr(nBalancesStdAccountType, sFromAccount, sToAccount, _
                                                 nAccountType, sAccountNature)
    sSourceFilter = SourceFilterSQLStr(nCurrencyType)
    sTransactionsFilter = TransactionsFilterSQLStr(nTransactionTypeId, nVoucherTypeId)
    sBalancesFilter = BalancesFilterSQLStr(nBalancesToDisplay)
    If Len(sBalancesFilter) <> 0 Then
      sBalancesFilter = " AND " & sBalancesFilter
    End If
    
    sSQL = LastLevelsBalancesSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts, dFromDate, dToDate, _
                                    sGralLedgersFilter, sStdAccountsFilter, _
                                    sTransactionsFilter, sSourceFilter)
    If (nExcRateCurrency <> 0) And (nExcRateType <> 0) Then
      sSQL = ValorizateLastLevelBalances(sSQL, nExcRateType, nExcRateCurrency, dExcRateDate)
    End If
    TruncateTable sAppServer, cBalancesTableName
    oConnection.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConnection.Open
    Set oCommand.ActiveConnection = oConnection
    oCommand.CommandText = "INSERT INTO " & cBalancesTableName & " " & sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute

    If bReturnOnlyLastLevels Then
      sSQL = GetLastLevelsFromTempTableSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts)
    ElseIf Not bReturnOnlyLastLevels Then
      sSQL = GetAllLevelsFromTempTableSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts, sAccountPattern)
    End If

    sStdAccts = StdAccountHistCompleteSQLStr(dFromDate, dToDate, sStdAccountsFilter)

    sSQL = "SELECT A.id_mayor, A.id_moneda, o_id_moneda clave_moneda, currency_name nombre_moneda, naturaleza, " & _
           "id_tipo_cuenta, rol_cuenta, A.numero_cuenta_estandar, nombre_cuenta_estandar nombre_cuenta, " & _
           "A.id_sector, NVL(clave_sector, '00') clave_sector, " & _
           "A.id_cuenta_auxiliar, A.numero_cuenta_auxiliar, A.nombre_cuenta_auxiliar, " & _
           "saldo_inicial, cargos, abonos, saldo_actual, " & _
           "(DECODE(naturaleza, 'D', (cargos - abonos), (abonos - cargos))) movimientos, " & _
           "saldo_nv saldo_no_valorizado " & _
           "FROM ( " & sSQL & " ) A, ( " & sStdAccts & " ) B, COF_Sector C, AO_Currencies D " & _
           "WHERE (A.numero_cuenta_estandar = B.numero_cuenta_estandar) AND " & _
                 "(A.id_sector = C.id_sector(+)) AND (A.id_moneda = D.currency_id) " & sBalancesFilter & " " & _
           "ORDER BY A.id_mayor, o_id_moneda, naturaleza DESC, A.numero_cuenta_estandar, clave_sector "
          
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConnection, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set AverageBalancesRS = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "AverageBalancesRS", Err.Number
End Function

Public Function BalancesRS(sAppServer As String, sAccountPattern As String, vGralLedgers As Variant, _
                           bConsolidateGralLedgers As Boolean, bReturnOnlyLastLevels As Boolean, _
                           bReturnSubsAccounts As Boolean, dFromDate As Date, dToDate As Date, _
                           Optional sFromAccount As String = "", Optional sToAccount As String = "", _
                           Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                           Optional nCurrencyType As Long = 0, Optional nAccountType As Long = 0, _
                           Optional sAccountNature As String = "", Optional nExcRateType As Long = 0, _
                           Optional nExcRateCurrency As Long = 0, Optional dExcRateDate As Date, _
                           Optional nBalancesToDisplay As Long = 4) As Recordset
  Dim oContext As ObjectContext, bObjectContextOK As Boolean
  Dim oConnection As New Connection, oCommand As New Command, oRecordset As New Recordset
  Dim sGralLedgersFilter As String, sTransactionsFilter As String, sStdAccountsFilter As String
  Dim sBalancesFilter As String, sSourceFilter As String, nBalancesStdAccountType As Long
  Dim sSQL As String, sStdAccts As String
  '*********************************************************************************************************
  On Error GoTo ErrHandler
    Set oContext = GetObjectContext
    bObjectContextOK = Not (oContext Is Nothing)
    
    If (IsArray(vGralLedgers)) Then
      nBalancesStdAccountType = GralLegerStdAccountType(sAppServer, CLng(vGralLedgers(0)))
      sGralLedgersFilter = "id_mayor IN (" & Join(vGralLedgers, ",") & ")"
    Else
      nBalancesStdAccountType = GralLegerStdAccountType(sAppServer, CLng(vGralLedgers))
      sGralLedgersFilter = "id_mayor = " & vGralLedgers
    End If
    If (nBalancesToDisplay = 5) Then
      bReturnOnlyLastLevels = True
    End If
    sStdAccountsFilter = StdAccountsFilterSQLStr(nBalancesStdAccountType, sFromAccount, sToAccount, _
                                                 nAccountType, sAccountNature)
    sSourceFilter = SourceFilterSQLStr(nCurrencyType)
    sTransactionsFilter = TransactionsFilterSQLStr(nTransactionTypeId, nVoucherTypeId)
    sBalancesFilter = BalancesFilterSQLStr(nBalancesToDisplay)
    If Len(sBalancesFilter) <> 0 Then
      sBalancesFilter = " AND " & sBalancesFilter
    End If
    
    sSQL = LastLevelsBalancesSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts, dFromDate, dToDate, _
                                    sGralLedgersFilter, sStdAccountsFilter, _
                                    sTransactionsFilter, sSourceFilter)
    If (nExcRateCurrency <> 0) And (nExcRateType <> 0) Then
      sSQL = ValorizateLastLevelBalances(sSQL, nExcRateType, nExcRateCurrency, dExcRateDate)
    End If
    TruncateTable sAppServer, cBalancesTableName
    oConnection.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConnection.Open
    Set oCommand.ActiveConnection = oConnection
    oCommand.CommandText = "INSERT INTO " & cBalancesTableName & " " & sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute

    If bReturnOnlyLastLevels Then
      sSQL = GetLastLevelsFromTempTableSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts)
    ElseIf Not bReturnOnlyLastLevels Then
      sSQL = GetAllLevelsFromTempTableSQLStr(bConsolidateGralLedgers, bReturnSubsAccounts, sAccountPattern)
    End If

    sStdAccts = StdAccountHistCompleteSQLStr(dFromDate, dToDate, sStdAccountsFilter)

    sSQL = "SELECT A.id_mayor, A.id_moneda, o_id_moneda clave_moneda, currency_name nombre_moneda, naturaleza, " & _
           "id_tipo_cuenta, rol_cuenta, A.numero_cuenta_estandar, nombre_cuenta_estandar nombre_cuenta, " & _
           "A.id_sector, NVL(clave_sector, '00') clave_sector, " & _
           "A.id_cuenta_auxiliar, A.numero_cuenta_auxiliar, A.nombre_cuenta_auxiliar, " & _
           "saldo_inicial, cargos, abonos, saldo_actual, " & _
           "(DECODE(naturaleza, 'D', (cargos - abonos), (abonos - cargos))) movimientos, " & _
           "saldo_nv saldo_no_valorizado " & _
           "FROM ( " & sSQL & " ) A, ( " & sStdAccts & " ) B, COF_Sector C, AO_Currencies D " & _
           "WHERE (A.numero_cuenta_estandar = B.numero_cuenta_estandar) AND " & _
                 "(A.id_sector = C.id_sector(+)) AND (A.id_moneda = D.currency_id) " & sBalancesFilter & " " & _
           "ORDER BY A.id_mayor, o_id_moneda, naturaleza DESC, A.numero_cuenta_estandar, clave_sector "
          
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConnection, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set BalancesRS = oRecordset
    If bObjectContextOK Then oContext.SetComplete
  Exit Function
ErrHandler:
  If bObjectContextOK Then oContext.SetAbort
  RaiseError ClassId, "BalancesRS", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS CONSTRUCTORES DE CADENAS SQL PARA LA TABLA TEMPORAL DE SALDOS A ULTIMO NIVEL            *
'************************************************************************************************************

'CODIGO EN EL PROCEDIMIENTO LLAMADOR
'  Dim oBalanceAvg As Recordset
'  With oRecordset
'    .CursorLocation = adUseClient
'    .Open sSQL, oConnection, adOpenDynamic, adLockBatchOptimistic
'    Set .ActiveConnection = Nothing
'  End With
'
'  Set oBalanceAvg = BalanceAvgTotalRS(sAppServer, sAccountPattern, vGralLedgers, dFromDate, dToDate, _
'                                      sFromAccount, sToAccount, nTransactionTypeId, nVoucherTypeId, _
'                                      nCurrencyType, nExcRateType, nExcRateCurrency, dExcRateDate, _
'                                      nRowsToDisplay, bReturnSubsAccounts)
'  With oRecordset
'    If Not (.BOF And .EOF) Then
'      .MoveFirst
'      oBalanceAvg.MoveFirst
'      Do While Not .EOF
'        !saldo_promedio = oBalanceAvg!saldo_promedio
'        .Update
'        .MoveNext
'        oBalanceAvg.MoveNext
'      Loop
'    End If
'    .MoveFirst
'  End With
      
Private Function BeforePeriodBalancesSQLStr(bConsolidateGralLedgers As Boolean, bGetSubsAccounts As Boolean, _
                                           dFromDate As Date, dToDate As Date, sGralLedgersFilter As String, _
                                           sStdAccountsFilter As String) As String
  Dim sSQL As String
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryBeforePeriodBalances)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "A.id_mayor, "))
    sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", IIf(bGetSubsAccounts, ", A.id_cuenta_auxiliar", ""))
    sSQL = Replace(sSQL, "<@INITIAL_BALANCES@>", InitialBalancesSQLStr(bConsolidateGralLedgers, _
                                                                       bGetSubsAccounts, _
                                                                       dToDate, sGralLedgersFilter))
    sSQL = Replace(sSQL, "<@INITIAL_POSTINGS@>", BeforePeriodPostingsSQLStr(bConsolidateGralLedgers, _
                                                                       bGetSubsAccounts, _
                                                                       dFromDate, sGralLedgersFilter))
    sSQL = Replace(sSQL, "<@STD_ACCOUNTS@>", StdAccountsHistSQLStr(dFromDate, dToDate, sStdAccountsFilter))
    BeforePeriodBalancesSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "BeforePeriodBalancesSQLStr", Err.Number
End Function
    
Private Function BeforePeriodPostingsSQLStr(bConsolidateGralLedgers As Boolean, bGetSubsAccounts As Boolean, _
                                            dFromDate As Date, sGralLedgersFilter As String) As String
  Dim sSQL As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryBeforePeriodPostings)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "id_mayor, "))
    sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", IIf(bGetSubsAccounts, ", id_cuenta_auxiliar", ""))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_FILTER@>", sGralLedgersFilter)
    sSQL = Replace(sSQL, "<@FROM_DATE@>", dFromDate - 1)
    BeforePeriodPostingsSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "BeforePeriodPostingsSQLStr", Err.Number
End Function
               
Private Function InitialBalancesSQLStr(bConsolidateGralLedgers As Boolean, bGetSubsAccounts As Boolean, _
                                       dToDate As Date, sGralLedgersFilter As String) As String
  Dim sSQL As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryInitialBalances)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "id_mayor, "))
    sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", IIf(bGetSubsAccounts, ", id_cuenta_auxiliar", ""))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_FILTER@>", sGralLedgersFilter)
    sSQL = Replace(sSQL, "<@TO_DATE@>", dToDate)
    InitialBalancesSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "InitialBalancesSQLStr", Err.Number
End Function

Private Function InPeriodBalancesSQLStr(bConsolidateGralLedgers As Boolean, bGetSubsAccounts As Boolean, _
                                        dFromDate As Date, dToDate As Date, sGralLedgersFilter As String, _
                                        sTransactionsFilter As String) As String
  Dim sSQL As String
  '******************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryInPeriodBalances)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "id_mayor, "))
    sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", IIf(bGetSubsAccounts, ", id_cuenta_auxiliar", ""))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_FILTER@>", sGralLedgersFilter)
    sSQL = Replace(sSQL, "<@TRANSACTIONS_FILTER@>", IIf(Len(sTransactionsFilter) <> 0, _
                                                      " AND " & sTransactionsFilter, ""))
    sSQL = Replace(sSQL, "<@FROM_DATE@>", dFromDate)
    sSQL = Replace(sSQL, "<@TO_DATE@>", dToDate)
    InPeriodBalancesSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "InPeriodBalancesSQLStr", Err.Number
End Function

Private Function LastLevelsBalancesSQLStr(bConsolidateGralLedgers As Boolean, bGetSubsAccounts As Boolean, _
                                         dFromDate As Date, dToDate As Date, sGralLedgersFilter As String, _
                                         sStdAccountsFilter As String, sTransactionsFilter As String, _
                                         sSourceFilter As String) As String
  Dim sSQL As String
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryLastLevelsBalances)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS@>", IIf(bConsolidateGralLedgers, "(0) id_mayor", "A.id_mayor"))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "A.id_mayor, "))
    If bGetSubsAccounts Then
      sSQL = Replace(sSQL, "<@SUBS_ACCTS@>", "A.id_cuenta_auxiliar, " & _
                                             "MAX(D.numero_cuenta_auxiliar) numero_cuenta_auxiliar, " & _
                                             "MAX(D.nombre_cuenta_auxiliar) nombre_cuenta_auxiliar")
    Else
      sSQL = Replace(sSQL, "<@SUBS_ACCTS@>", "(0) id_cuenta_auxiliar, " & _
                                             "('') numero_cuenta_auxiliar, ('') nombre_cuenta_auxiliar")
    End If
    sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", IIf(bGetSubsAccounts, ", A.id_cuenta_auxiliar", ""))
    sSQL = Replace(sSQL, "<@SUBS_ACCTS_TABLE@>", IIf(bGetSubsAccounts, ", COF_Cuenta_Auxiliar D", ""))
    sSQL = Replace(sSQL, "<@SUBS_ACCTS_JOIN@> ", _
                              IIf(bGetSubsAccounts, "AND (A.id_cuenta_auxiliar = D.id_cuenta_auxiliar(+)) ", ""))
    
    sSQL = Replace(sSQL, "<@BEFORE_PERIOD_BAL@>", _
                         BeforePeriodBalancesSQLStr(bConsolidateGralLedgers, bGetSubsAccounts, _
                                                    dFromDate, dToDate, sGralLedgersFilter, _
                                                    sStdAccountsFilter))
    sSQL = Replace(sSQL, "<@IN_PERIOD_BAL@>", _
                         InPeriodBalancesSQLStr(bConsolidateGralLedgers, bGetSubsAccounts, _
                                                dFromDate, dToDate, sGralLedgersFilter, sTransactionsFilter))
    sSQL = Replace(sSQL, "<@STD_ACCOUNTS@>", StdAccountsHistSQLStr(dFromDate, dToDate, sStdAccountsFilter))
    sSQL = Replace(sSQL, "<@SOURCE_FILTER@>", IIf(Len(sSourceFilter) <> 0, " AND " & sSourceFilter, ""))
    LastLevelsBalancesSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "LastLevelsBalancesSQLStr", Err.Number
End Function
   
Private Function StdAccountsHistSQLStr(dFromDate As Date, dToDate As Date, _
                                       sStdAccountsFilter As String) As String
  Dim sSQL As String
  '***************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryStdAccountsHist)
    sSQL = Replace(sSQL, "<@STD_ACCOUNTS_FILTER@>", sStdAccountsFilter)
    sSQL = Replace(sSQL, "<@FROM_DATE@>", dFromDate)
    sSQL = Replace(sSQL, "<@TO_DATE@>", dToDate)
    StdAccountsHistSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountsHistSQLStr", Err.Number
End Function

Private Function ValorizateLastLevelBalances(sBalancesSQL As String, nExcRateType As Long, _
                                             nExcRateCurrency As Long, dExcRateDate As Date) As String
  Dim sPattern As String, sSQL As String
  '***********************************************************************************************************
  On Error GoTo ErrHandler
    sPattern = GetQueryString(cnQryValorizateToSourceCurrency)
    sSQL = Replace(sPattern, "<@BALANCES_TABLE@>", sBalancesSQL)
    sSQL = Replace(sSQL, "<@EXC_RATE_TYPE@>", nExcRateType)
    sSQL = Replace(sSQL, "<@EXC_RATE_DATE@>", dExcRateDate)
    sSQL = Replace(sSQL, "<@FROM_CURRENCY@>", cnSourceCurrency)
    If (nExcRateCurrency <> cnSourceCurrency) Then
      sPattern = GetQueryString(cnQryValorizateToTargetCurrency)
      sSQL = Replace(sPattern, "<@VALORIZATED_TABLE@>", sSQL)
      sSQL = Replace(sSQL, "<@EXC_RATE_TYPE@>", nExcRateType)
      sSQL = Replace(sSQL, "<@EXC_RATE_DATE@>", dExcRateDate)
      sSQL = Replace(sSQL, "<@TO_CURRENCY@>", cnSourceCurrency)
    End If
    ValorizateLastLevelBalances = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "ValorizateLastLevelBalances", Err.Number
End Function
'************************************************************************************************************
'* MÉTODOS PRIVADOS CONSTRUCTORES DE CADENAS SQL PARA LA CONSTRUCCION DE LA ESTRUCTURA DE REGRESO           *
'************************************************************************************************************

Private Function GetAllLevelsFromTempTableSQLStr(bConsolidateGralLedgers As Boolean, _
                                                 bGetSubsAccounts As Boolean, _
                                                 sAccountPattern As String) As String
  Dim sPattern, sSQL As String, nPos As Variant, i As Long
  '***********************************************************************************
  On Error GoTo ErrHandler
    nPos = GetPositions(sAccountPattern)
    sSQL = GetQueryString(cnQryAllLevelsFromTempTable2) & " "
    sPattern = GetQueryString(cnQryAllLevelsFromTempTable3) & " "
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & Replace(sPattern, "<@POS@>", i)
    Next i
    sSQL = sSQL & GetLastLevelsFromTempTableSQLStr(bConsolidateGralLedgers, bGetSubsAccounts)
    sPattern = GetQueryString(cnQryAllLevelsFromTempTable1)
    sSQL = Replace(sPattern, "<@SUBQUERY@>", sSQL)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS@>", IIf(bConsolidateGralLedgers, "(0) id_mayor", "id_mayor"))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "id_mayor, "))
    sSQL = Replace(sSQL, "<@TEMP_TABLE@>", cBalancesTableName)
    GetAllLevelsFromTempTableSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAllLevelsFromTempTableSQLStr", Err.Number
End Function

Private Function GetLastLevelsFromTempTableSQLStr(bConsolidateGralLedgers As Boolean, _
                                                  bGetSubsAccounts As Boolean) As String
  Dim sSQL As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryLastLevelsFromTempTable)
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS@>", IIf(bConsolidateGralLedgers, "(0) id_mayor", "id_mayor"))
    sSQL = Replace(sSQL, "<@GRAL_LEDGERS_CONS@>", IIf(bConsolidateGralLedgers, "", "id_mayor, "))
    If bGetSubsAccounts Then
      sSQL = Replace(sSQL, "<@SUBS_ACCTS@>", "id_cuenta_auxiliar, " & _
                                             "numero_cuenta_auxiliar, nombre_cuenta_auxiliar")
      sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", ", id_cuenta_auxiliar, numero_cuenta_auxiliar, " & _
                                                 "nombre_cuenta_auxiliar")
    Else
      sSQL = Replace(sSQL, "<@SUBS_ACCTS@>", "(0) id_cuenta_auxiliar, " & _
                                             "('') numero_cuenta_auxiliar, ('') nombre_cuenta_auxiliar")
      sSQL = Replace(sSQL, "<@GET_SUBS_ACCTS@>", "")
    End If
    sSQL = Replace(sSQL, "<@TEMP_TABLE@>", cBalancesTableName)
    GetLastLevelsFromTempTableSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetLastLevelsFromTempTableSQLStr", Err.Number
End Function

Private Function StdAccountHistCompleteSQLStr(dFromDate As Date, dToDate As Date, _
                                              sStdAccountsFilter As String) As String
  Dim sSQL As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    sSQL = GetQueryString(cnQryStdAccountHistComplete)
    sSQL = Replace(sSQL, "<@STD_ACCOUNTS_FILTER@>", sStdAccountsFilter)
    sSQL = Replace(sSQL, "<@FROM_DATE@>", dFromDate)
    sSQL = Replace(sSQL, "<@TO_DATE@>", dToDate)
    StdAccountHistCompleteSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountHistCompleteSQLStr", Err.Number
End Function
                                                             
'************************************************************************************************************
'* MÉTODOS PRIVADOS CONSTRUCTORES DE FILTROS SQL                                                            *
'************************************************************************************************************

Private Function BalancesFilterSQLStr(nRowsToDisplay As Long) As String
  Dim sSQL As String
  '********************************************************************
  On Error GoTo ErrHandler
    If (nRowsToDisplay = 1) Then
      sSQL = "(saldo_inicial <> 0)"
    ElseIf (nRowsToDisplay = 2) Then
      sSQL = "((cargos <> 0) OR (abonos <> 0))"
    ElseIf (nRowsToDisplay = 3) Then
      sSQL = "((saldo_inicial <> 0) OR (cargos <> 0) OR (abonos <> 0))"
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sSQL = "(saldo_actual < 0)"
    End If
    BalancesFilterSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalancesFilterSQLStr", Err.Number
End Function

Private Function SourceFilterSQLStr(nCurrencyType As Long) As String
  Dim sSQL As String
  '*****************************************************************
  On Error GoTo ErrHandler
    If (nCurrencyType > 0) Then
      sSQL = "(id_moneda = " & nCurrencyType & ") "
    ElseIf (nCurrencyType < 0) Then
      sSQL = "(id_moneda <> " & Abs(nCurrencyType) & ") "
    End If
    SourceFilterSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "SourceFilterSQLStr", Err.Number
End Function

Private Function StdAccountsFilterSQLStr(nStdAccountType As Long, _
                                         sFromAccount As String, sToAccount As String, _
                                         nAccountType As Long, sAccountNature As String) As String
  Dim sSQL As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    If (Len(sFromAccount) <> 0) And (Len(sToAccount) <> 0) Then
      sSQL = "((numero_cuenta_estandar >= '" & sFromAccount & "') AND " & _
             "(numero_cuenta_estandar <= '" & sToAccount & "'))"
    ElseIf (Len(sFromAccount) <> 0) And (Len(sToAccount) = 0) Then
      sSQL = "(numero_cuenta_estandar >= '" & sFromAccount & "')"
    ElseIf (Len(sFromAccount) = 0) And (Len(sToAccount) <> 0) Then
      sSQL = "(numero_cuenta_estandar <= '" & sToAccount & "')"
    End If
    
    If (nAccountType <> 0) Then
      If (Len(sSQL) <> 0) Then
        sSQL = sSQL & " AND "
      End If
      If (nAccountType > 0) Then
        sSQL = sSQL & "(id_tipo_cuenta = " & nAccountType & ")"
      ElseIf (nAccountType < 0) Then
        sSQL = sSQL & "(id_tipo_cuenta <> " & Abs(nAccountType) & ")"
      End If
    End If
    
    If (Len(sAccountNature) <> 0) Then
      If (Len(sSQL) <> 0) Then
        sSQL = sSQL & " AND "
      End If
      If (nAccountType > 0) Then
        sSQL = sSQL & " (naturaleza = '" & sAccountNature & "')"
      End If
    End If
    
    If Len(sSQL) <> 0 Then
      sSQL = "(id_tipo_cuentas_std = " & nStdAccountType & ") AND " & sSQL
    Else
      sSQL = "(id_tipo_cuentas_std = " & nStdAccountType & ")"
    End If
    StdAccountsFilterSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "StdAccountsFilterSQLStr", Err.Number
End Function
       
Private Function TransactionsFilterSQLStr(nTransactionTypeId As Long, nVoucherTypeId As Long) As String
  Dim sSQL As String
  '****************************************************************************************************
  On Error GoTo ErrHandler
    If (nTransactionTypeId > 0) Then
      sSQL = "(id_tipo_transaccion = " & nTransactionTypeId & ")"
    ElseIf (nTransactionTypeId < 0) Then
      sSQL = "(id_tipo_transaccion <> " & Abs(nTransactionTypeId) & ")"
    End If
    If (nVoucherTypeId <> 0) Then
      If (Len(sSQL) <> 0) Then
        sSQL = sSQL & " AND "
      End If
      If (nVoucherTypeId > 0) Then
        sSQL = sSQL & "(id_tipo_poliza = " & nVoucherTypeId & ")"
      ElseIf (nVoucherTypeId < 0) Then
        sSQL = sSQL & "(id_tipo_poliza <> " & Abs(nVoucherTypeId) & ")"
      End If
    End If
    TransactionsFilterSQLStr = sSQL
  Exit Function
ErrHandler:
  RaiseError ClassId, "TransactionsFilterSQLStr", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS MANEJADORES DE INFORMACIÓN DIVERSA QUE DEBE ELIMINARSE DE ESTA SECCIÓN EN EL FUTURO     *
'************************************************************************************************************
                                     
Private Sub Commit(sAppServer As String)
  Dim oCommand As New Command
  '*************************************
  On Error GoTo ErrHandler
    With oCommand
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CommandText = "COMMIT"
      .CommandType = adCmdText
      .Execute
    End With
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Commit", Err.Number
End Sub

Private Function GetPositions(ByVal sPattern As String) As Variant
  Dim sPositions As String, nPos As Long
  '***************************************************************
  On Error GoTo ErrHandler
    nPos = 0
    Do
      nPos = InStr(nPos + 1, sPattern, "-")
      If (nPos = 0) Then
        Exit Do
      ElseIf (nPos <> 0) Then
        If (Len(sPositions) = 0) Then
          sPositions = (nPos - 1)
        ElseIf (Len(sPositions) <> 0) Then
          sPositions = sPositions & "," & (nPos - 1)
        End If
      End If
    Loop
    If (Len(sPositions) = 0) Then
      sPositions = Len(sPattern)
    ElseIf (Len(sPositions) <> 0) Then
      sPositions = sPositions & "," & Len(sPattern)
    End If
    GetPositions = Split(sPositions, ",")
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetPositions", Err.Number
End Function

Private Function GralLegerStdAccountType(sAppServer As String, nGralLedger As Long) As Long
  Dim oRecordset As New Recordset, sSQL As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    sSQL = "SELECT id_tipo_cuentas_std " & _
           "FROM COF_Mayor " & _
           "WHERE (id_mayor = " & nGralLedger & ")"
    With oRecordset
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      If Not (.BOF And .EOF) Then
        GralLegerStdAccountType = !id_tipo_cuentas_std
      Else
        GralLegerStdAccountType = 0
      End If
      .Close
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GralLegerStdAccountType", Err.Number
End Function

Private Sub TruncateTable(sAppServer As String, sTableName As String)
  Dim oCommand As New Command
  '******************************************************************
  On Error GoTo ErrHandler
    With oCommand
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CommandText = "DELETE FROM " & sTableName
      .CommandType = adCmdText
      .Execute
    End With
    Commit sAppServer
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Truncate", Err.Number
End Sub

'*****************************************************************************************************
'**************                 SALDOS PROMEDIO                                       ****************
'*****************************************************************************************************

Private Function BalanceAvgTotalRS(sAppServer As String, sPattern As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                         Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                         Optional nCurrencyType As Long = 0, _
                         Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                         Optional dExcDate As Date, _
                         Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False) As ADODB.Recordset
  Dim oRs           As ADODB.Recordset
  Dim oRSTot        As ADODB.Recordset
  
  Dim dFecha        As Date
  Dim sAccount      As String
  Dim sSector       As String
  Dim nSaldoMN      As Double
  Dim nSaldoMNA     As Double
  Dim sAccName      As String
  
  Dim sNature       As String
  Dim sDateTime     As String
  Dim nCurrencyId   As Long
  
  Dim DebugAccount  As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    
    Set oRs = BalanceAvgRS(sAppServer, sPattern, vGL, dIniDate, dEndDate, sIniAccount, _
                           sEndAccount, nTransactionTypeId, nVoucherTypeId, nCurrencyType, _
                           nExcRateType, nExcRateCurrency, dExcDate, nOptionToDisplay, bWithAux)
    
    Set oRSTot = StructBalanceProm(sAppServer, sPattern, vGL, dIniDate, dEndDate, sIniAccount, _
                                   sEndAccount, nTransactionTypeId, nVoucherTypeId, nCurrencyType, _
                                   nExcRateType, nExcRateCurrency, dExcDate, nOptionToDisplay, bWithAux)
    DebugAccount = "0000"
    With oRs
      dFecha = !fecha_rev
      sAccount = !NUMERO_CUENTA_ESTANDAR
      sAccName = !nombre_cuenta
      sSector = !clave_sector
      nSaldoMN = 0
      nSaldoMNA = 0
      Do While Not (.EOF)
        
        If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
          If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPattern)) Then
            nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
                                
            If bWithAux Then
              With oRSTot
                .AddNew
                oRs.MovePrevious
                !ID_MONEDA = oRs!ID_MONEDA
                !o_id_moneda = oRs!o_id_moneda
                !NATURALEZA = oRs!NATURALEZA
                !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
                !clave_sector = oRs!clave_sector
                !nombre_cuenta = oRs!nombre_cuenta
                !ID_CUENTA_AUXILIAR = oRs!ID_CUENTA_AUXILIAR
                !numero_cuenta_auxiliar = oRs!numero_cuenta_auxiliar
                !id_tipo_cuenta = oRs!id_tipo_cuenta
                !rol_cuenta = oRs!rol_cuenta
                !from_movs = oRs!from_movs
                !saldo_promedio = nSaldoMNA
              End With
              oRs.MoveNext
            Else
              With oRSTot
                .AddNew
                oRs.MovePrevious
                !ID_MONEDA = oRs!ID_MONEDA
                !o_id_moneda = oRs!o_id_moneda
                !NATURALEZA = oRs!NATURALEZA
                !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
                !clave_sector = oRs!clave_sector
                !nombre_cuenta = oRs!nombre_cuenta
                !id_tipo_cuenta = oRs!id_tipo_cuenta
                !rol_cuenta = oRs!rol_cuenta
                !from_movs = oRs!from_movs
                !saldo_promedio = nSaldoMNA
              End With
              oRs.MoveNext
            End If
            
          ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPattern)) Then
            'Siguiente
          End If
          nSaldoMN = 0
          nSaldoMNA = 0
          sAccount = !NUMERO_CUENTA_ESTANDAR
          sAccName = !nombre_cuenta
          sSector = !clave_sector
          dFecha = !fecha_rev
        ElseIf (sAccount = !NUMERO_CUENTA_ESTANDAR) Then
          If (sSector <> !clave_sector) Then
            nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
            
            If bWithAux Then
              With oRSTot
                .AddNew
                oRs.MovePrevious
                !ID_MONEDA = oRs!ID_MONEDA
                !o_id_moneda = oRs!o_id_moneda
                !NATURALEZA = oRs!NATURALEZA
                !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
                !clave_sector = oRs!clave_sector
                !nombre_cuenta = oRs!nombre_cuenta
                !ID_CUENTA_AUXILIAR = oRs!ID_CUENTA_AUXILIAR
                !numero_cuenta_auxiliar = oRs!numero_cuenta_auxiliar
                !id_tipo_cuenta = oRs!id_tipo_cuenta
                !rol_cuenta = oRs!rol_cuenta
                !from_movs = oRs!from_movs
                !saldo_promedio = nSaldoMNA
              End With
              oRs.MoveNext
            Else
              With oRSTot
                .AddNew
                oRs.MovePrevious
                !ID_MONEDA = oRs!ID_MONEDA
                !o_id_moneda = oRs!o_id_moneda
                !NATURALEZA = oRs!NATURALEZA
                !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
                !clave_sector = oRs!clave_sector
                !nombre_cuenta = oRs!nombre_cuenta
                !id_tipo_cuenta = oRs!id_tipo_cuenta
                !rol_cuenta = oRs!rol_cuenta
                !from_movs = oRs!from_movs
                !saldo_promedio = nSaldoMNA
              End With
              oRs.MoveNext
            End If
            
            nSaldoMN = 0
            nSaldoMNA = 0
            sAccount = !NUMERO_CUENTA_ESTANDAR
            sAccName = !nombre_cuenta
            sSector = !clave_sector
            dFecha = !fecha_rev
          End If
        End If
        
        If (dFecha = #1/1/1990#) Then
           nSaldoMN = !SALDO_INICIAL
           nSaldoMNA = 0
           dFecha = dIniDate
        ElseIf (dFecha <> #1/1/1990#) Then
          nSaldoMNA = nSaldoMNA + (nSaldoMN * (!fecha_rev - dFecha))
          If (!NATURALEZA = "D") Then
            nSaldoMN = nSaldoMN + !Cargos - !ABONOS
          ElseIf (!NATURALEZA = "A") Then
            nSaldoMN = nSaldoMN - !Cargos + !ABONOS
          End If
          dFecha = !fecha_rev
        End If
        .MoveNext
      Loop
      nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
      
      If bWithAux Then
        With oRSTot
          .AddNew
          oRs.MovePrevious
          !ID_MONEDA = oRs!ID_MONEDA
          !o_id_moneda = oRs!o_id_moneda
          !NATURALEZA = oRs!NATURALEZA
          !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
          !clave_sector = oRs!clave_sector
          !nombre_cuenta = oRs!nombre_cuenta
          !ID_CUENTA_AUXILIAR = oRs!ID_CUENTA_AUXILIAR
          !numero_cuenta_auxiliar = oRs!numero_cuenta_auxiliar
          !id_tipo_cuenta = oRs!id_tipo_cuenta
          !rol_cuenta = oRs!rol_cuenta
          !from_movs = oRs!from_movs
          !saldo_promedio = nSaldoMNA
          .Update
        End With
        oRs.MoveNext
      Else
        With oRSTot
          .AddNew
          oRs.MovePrevious
          !ID_MONEDA = oRs!ID_MONEDA
          !o_id_moneda = oRs!o_id_moneda
          !NATURALEZA = oRs!NATURALEZA
          !NUMERO_CUENTA_ESTANDAR = oRs!NUMERO_CUENTA_ESTANDAR
          !clave_sector = oRs!clave_sector
          !nombre_cuenta = oRs!nombre_cuenta
          !id_tipo_cuenta = oRs!id_tipo_cuenta
          !rol_cuenta = oRs!rol_cuenta
          !from_movs = oRs!from_movs
          !saldo_promedio = nSaldoMNA
          .Update
        End With
        oRs.MoveNext
      End If
    End With
    oRSTot.MoveFirst
    Set BalanceAvgTotalRS = oRSTot
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceAvgTotalRS", Err.Number
End Function

Private Function BalanceAvgRS(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 0, _
                        Optional nExcRateCurrency As Long = 0, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As ADODB.Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sOptionToDisplay  As String
  
  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  
  '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      nGl = CLng(vGL)
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    nTypeStdAcc = GralLegerStdAccountType(sAppServer, nGl)
    
    If (nTransactionTypeId = 0) Then
      sWhereTran = ""
    ElseIf (nTransactionTypeId > 0) Then
      sWhereTran = " AND (ID_TIPO_TRANSACCION = " & nTransactionTypeId & ") "
    ElseIf (nTransactionTypeId < 0) Then
      sWhereTran = " AND (ID_TIPO_TRANSACCION <> " & Abs(nTransactionTypeId) & ") "
    End If
    
    If (nVoucherTypeId = 0) Then
      sWhereTran = sWhereTran & " "
    ElseIf (nVoucherTypeId > 0) Then
      sWhereTran = sWhereTran & " AND (ID_TIPO_POLIZA = " & nVoucherTypeId & ") "
    ElseIf (nVoucherTypeId < 0) Then
      sWhereTran = sWhereTran & " AND (ID_TIPO_POLIZA <> " & Abs(nVoucherTypeId) & ") "
    End If
        
    If (sIniAccount = "") And (sEndAccount = "") Then
      sWhere = " "
    ElseIf (sIniAccount <> "") And (sEndAccount <> "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sIniAccount & "' AND '" & sEndAccount & "') "
    ElseIf (sIniAccount <> "") And (sEndAccount = "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR <= '" & sIniAccount & "') "
    End If
    
    If (nCurrencyType = 0) Then
      sWhere = sWhere & " "
    ElseIf (nCurrencyType > 0) Then
      sWhere = sWhere & " AND (ID_MONEDA = " & nCurrencyType & ") "
    ElseIf (nCurrencyType < 0) Then
      sWhere = sWhere & " AND (ID_MONEDA <> " & Abs(nCurrencyType) & ") "
    End If
    
    If (nRowsToDisplay = 1) Then
      sOptionToDisplay = " WHERE (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " WHERE ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " WHERE ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " WHERE (SALDO_ACTUAL < 0) "
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    
    
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
            
    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
   
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        to_date('01/01/1990') FECHA_REV, max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
              " FROM ( " & _
              sSQL1 & _
              " UNION " & _
              sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, " & _
              "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA "
                         
    
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        FECHA_AFECTACION FECHA_REV, to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, FECHA_AFECTACION "
                       
    TruncateTable sAppServer, "tmp_SALDOS_PROM"
    
    sSQL = " INSERT INTO tmp_SALDOS_PROM " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.FECHA_REV, A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        FECHA_REV, sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                " FROM ( " & _
                  sSQL_SI & _
                " UNION " & _
                  sSQL2 & _
                    " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                        sWhere & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV " & _
              "  ) A " & _
              sOptionToDisplay
            
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = " tmp_SALDOS_PROM "
    
      sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, FECHA_REV, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4), FECHA_REV " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                    "           FECHA_REV, SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR, FECHA_REV " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV "
    '       " ORDER BY ID_MONEDA DESC, CUENTA DESC, SECTOR DESC "
    
    
    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    If nExcRateCurrency = 0 Then
      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR <> 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
'sOptionToDisplay

      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
'sOptionToDisplay
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         FECHA_REV, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                  "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnSourceCurrency & ") = " & cnSourceCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
               
'sOptionToDisplay &
      
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, " & _
               "         NUMERO_CUENTA_AUXILIAR, FECHA_REV, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
      
          If (cnSourceCurrency = nExcRateCurrency) Then
            sSQL = sSQL_A & _
                   " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
          End If
      Else
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         FECHA_REV, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnSourceCurrency & ") = " & cnSourceCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
'sOptionToDisplay &
      
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
                " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
        
        If (cnSourceCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
        End If
      End If
    End If
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set BalanceAvgRS = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceAvgRS", Err.Number
End Function

Private Function StructBalanceProm(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 0, _
                        Optional nExcRateCurrency As Long = 0, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sOptionToDisplay  As String
  
  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  
  '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConn.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConn.Open
    
    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      nGl = CLng(vGL)
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    'nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
    If (nTransactionTypeId = 0) Then
      sWhereTran = ""
    ElseIf (nTransactionTypeId > 0) Then
      sWhereTran = " AND (ID_TIPO_TRANSACCION = " & nTransactionTypeId & ") "
    ElseIf (nTransactionTypeId < 0) Then
      sWhereTran = " AND (ID_TIPO_TRANSACCION <> " & Abs(nTransactionTypeId) & ") "
    End If
    
    If (nVoucherTypeId = 0) Then
      sWhereTran = sWhereTran & " "
    ElseIf (nVoucherTypeId > 0) Then
      sWhereTran = sWhereTran & " AND (ID_TIPO_POLIZA = " & nVoucherTypeId & ") "
    ElseIf (nVoucherTypeId < 0) Then
      sWhereTran = sWhereTran & " AND (ID_TIPO_POLIZA <> " & Abs(nVoucherTypeId) & ") "
    End If
        
    If (sIniAccount = "") And (sEndAccount = "") Then
      sWhere = " "
    ElseIf (sIniAccount <> "") And (sEndAccount <> "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sIniAccount & "' AND '" & sEndAccount & "') "
    ElseIf (sIniAccount <> "") And (sEndAccount = "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR <= '" & sIniAccount & "') "
    End If
    
    If (nCurrencyType = 0) Then
      sWhere = sWhere & " "
    ElseIf (nCurrencyType > 0) Then
      sWhere = sWhere & " AND (ID_MONEDA = " & nCurrencyType & ") "
    ElseIf (nCurrencyType < 0) Then
      sWhere = sWhere & " AND (ID_MONEDA <> " & Abs(nCurrencyType) & ") "
    End If
    
    If (nRowsToDisplay = 1) Then
      sOptionToDisplay = " WHERE (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " WHERE ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " WHERE ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " WHERE (SALDO_ACTUAL < 0) "
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
            
    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
   
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        to_date('01/01/1990') FECHA_REV, max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
              " FROM ( " & _
              sSQL1 & _
              " UNION " & _
              sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, " & _
              "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA "
                         
    
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        FECHA_AFECTACION FECHA_REV, to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, FECHA_AFECTACION "
                       
    TruncateTable sAppServer, "tmp_SALDOS_PROM"
    
    sSQL = " INSERT INTO tmp_SALDOS_PROM " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.FECHA_REV, A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        FECHA_REV, sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                " FROM ( " & _
                  sSQL_SI & _
                " UNION " & _
                  sSQL2 & _
                    " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                "       (A.ID_MONEDA = 0) " & _
                        sWhere & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV " & _
              "  ) A " & _
              sOptionToDisplay
            
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = " tmp_SALDOS_PROM "
    
      sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, FECHA_REV, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4), FECHA_REV " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                    "           FECHA_REV, SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR, FECHA_REV " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV "
    '       " ORDER BY ID_MONEDA DESC, CUENTA DESC, SECTOR DESC "
    
               
    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    If nExcRateCurrency = 0 Then
      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        (0) SALDO_PROMEDIO , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) AND " & _
               "       (A.ID_MONEDA = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
        
        'sOptionToDisplay
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) AND " & _
               "       (A.ID_MONEDA = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
      
      'sOptionToDisplay
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                  "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND (A.ID_MONEDA = 0) " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
      'sOptionToDisplay &
      Else
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND (A.ID_MONEDA = 0) " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
                " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
      'sOptionToDisplay &
      End If
     End If
           ' " ORDER BY O_ID_MONEDA, CUENTA_STD.NATURALEZA DESC , CUENTA_STD.NUMERO_CUENTA_ESTANDAR,CLAVE_SECTOR "
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    Set StructBalanceProm = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "StructBalanceProm", Err.Number
End Function

