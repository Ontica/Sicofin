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
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : Fixed Reports DS (SCFFixedReportsDS)            Parte   : Data services                     *
'* Clase      : CBalances                                       Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Generador de balanzas. Los métodos de esta clase estaban en AOGLBalances desarrollado por   *
'*              Guillermo Rodríguez y que son empleados en los reportes contables fijos.                    *
'*              Reemplazar esta clase e incluir sus métodos en EFABalanceEngine.                            *
'*              Además es posible que los mismos métodos estén duplicados en la clase CTempBalances del     *
'*              componente EFABalanceReporter, la cual también deberá removerse de dicho componente.        *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CBalances"

'************************************************************************************************************
'* BALANZAS VIEJAS EMPLEADAS EN REPORTES (Dev: Guillermo Rodríguez)  / Reemplazar por EFABalanceEngine      *
'************************************************************************************************************

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
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
                       
    Truncate sAppServer, "tmp_SALDOS_PROM"
    
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
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
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
      
          If (cnFromCurrency = nExcRateCurrency) Then
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
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
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
        
        If (cnFromCurrency = nExcRateCurrency) Then
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

Private Function BalanceComparative(sAppServer As String, sPatron As String, vGL As Variant, _
                                    dIniDate As Date, dCuteDate As Date, _
                                    dIniDate2 As Date, dCuteDate2 As Date, _
                                    Optional sIniAccount As String = "", _
                                    Optional sEndAccount As String = "", _
                                    Optional nTransactionTypeId As Long = 0, _
                                    Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, _
                                    Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                                    Optional bWithAux As Boolean = False, _
                                    Optional bWithProm As Boolean = False) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sSQL_B              As String
  Dim sSQL_BA             As String
  Dim sSQL_SI_B           As String
  Dim sSQL1_B             As String
  Dim sSQL2_B             As String
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sOptionToDisplay  As String
  
  Dim sFromTable        As String
  Dim sFromTable_B       As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist      As String
  Dim sSQLStdAccHist_B    As String
  
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
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
            
    
    sSQL1_B = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
              "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
              " FROM COF_SALDO_INICIAL " & _
              " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate2 & "') " & _
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
            
    sSQL2_B = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
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
              "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate2 - 1) & "')  and " & _
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
   
   sSQLStdAccHist_B = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, " & _
                      "        NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                      " FROM COF_CUENTA_ESTANDAR_HIST " & _
                      " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate2 & "')) AND " & _
                      "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate2 & "'))) " & _
                      " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
                     
                     
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
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
                         
    sSQL_SI_B = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
                "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
                " FROM ( " & _
                  sSQL1_B & _
                " UNION " & _
                  sSQL2_B & _
                " ) A  , COF_CUENTA B, (" & sSQLStdAccHist_B & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA "
    
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
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sSQL2_B = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
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
             "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate2 & "')  and " & _
             "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate2 & "')  and " & _
             "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
             "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
             " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
             "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
                       

    Truncate sAppServer, "tmp_SALDOS"
    
    sSQL = " INSERT INTO tmp_SALDOS " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
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
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
            
    Truncate sAppServer, "tmp_SALDOS_B"
        
    sSQL_B = " INSERT INTO tmp_SALDOS_B " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                " FROM ( " & _
                  sSQL_SI_B & _
                " UNION " & _
                  sSQL2_B & _
                    " ) A  , COF_CUENTA B, (" & sSQLStdAccHist_B & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                        sWhere & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
              
              
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL_B
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = "tmp_SALDOS"
    sFromTable_B = "tmp_SALDOS_B"
    
    'proceso primer periodo
      sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                    "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    '       " ORDER BY ID_MONEDA DESC, CUENTA DESC, SECTOR DESC "
    
    'proceso segundo periodo
    
      sSQL_B = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
               "           (0) ID_CUENTA_AUXILIAR, " & _
               "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
               "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
               "    From  " & sFromTable_B & "  A " & _
               "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
               "          (nvl(ID_SECTOR,0) <> 0) " & _
               "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
               " Union "
    
    i = 0
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL_B = sSQL_B & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                        "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                        "     From  " & sFromTable_B & " A " & _
                        "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                        "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                        "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                        " Union "
    Next i
    
    sSQL_B = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
           " ( " & sSQL_B & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
           "       From  " & sFromTable_B & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    '       " ORDER BY ID_MONEDA DESC, CUENTA DESC, SECTOR DESC "
    
    
    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    sSQL2_B = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate2 & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate2 & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    If nExcRateCurrency = 0 Then
      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
               "        (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2 , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) "
               
               '& _
                       sOptionToDisplay
               
        sSQL_B = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
               "        (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL , " & _
               "        SALDO_INICIAL SALDO_INICIAL_2, CARGOS CARGOS_2, ABONOS ABONOS_2, SALDO_ACTUAL SALDO_ACTUAL_2 , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) "
               
               '& _
                       sOptionToDisplay
               
        sSQL = " select A.ID_MONEDA, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        A.CLAVE_SECTOR, " & _
               "        A.NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM ( " & _
                 sSQL & _
               " union " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NATURALEZA, NOMBRE_CUENTA_AUXILIAR, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
               "        (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) "
               
               '& sOptionToDisplay
               
        sSQL_B = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
                 "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                 "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                 "        (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
                 "        SALDO_INICIAL SALDO_INICIAL_2, CARGOS CARGOS_2, ABONOS ABONOS_2, SALDO_ACTUAL SALDO_ACTUAL_2, " & _
                 "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                 " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A , COF_SECTOR, " & _
                 "       AO_CURRENCIES " & _
                 " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                 "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
                 "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
                 "       (A.ID_CUENTA_AUXILIAR = 0) "
                 
                 '& sOptionToDisplay
                 
                 '& _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
                       
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        A.CLAVE_SECTOR, " & _
               "        A.NOMBRE_CUENTA, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM ( " & _
                 sSQL & _
               " UNION " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, NATURALEZA, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
               
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
               
        'sOptionToDisplay &
        
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A
        End If
     sSQL_BA = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
'sOptionToDisplay &
               
      sSQL_B = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_BA & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_B = sSQL_BA
        End If
        
        sSQL = " SELECT  id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                 sSQL & _
               " UNION " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, NATURALEZA, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
               
      Else
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                'sOptionToDisplay &
                
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A
        End If
        
     sSQL_BA = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
                'sOptionToDisplay &
                
        sSQL_B = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_BA & " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_B = sSQL_BA
        End If
                
        sSQL = " SELECT  id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
                "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
                "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
                "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                " FROM (" & _
                  sSQL & _
                " UNION " & _
                  sSQL_B & _
                " ) A " & _
                " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
                "          CLAVE_SECTOR, NATURALEZA, " & _
                "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
                " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
               
      End If
    End If
    
    If bWithProm Then
      With oRs
       .CursorLocation = adUseClient
       .Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
       Set .ActiveConnection = Nothing
      End With
    Else
      With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      End With
    End If
    Truncate sAppServer, "tmp_SALDOS"
    Truncate sAppServer, "tmp_SALDOS_B"
    Set BalanceComparative = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceComparative", Err.Number
End Function

Public Function BalanceCompByAccountRS(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        sAccount1 As String, sAccount2 As String, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 0, _
                        Optional nExcRateCurrency As Long = 0, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False, _
                        Optional bWithProm As Boolean = False) As ADODB.Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_AA           As String
  Dim sSQL_B            As String
  Dim sSQL_BB           As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  Dim sSLQ_TRUNC        As String
  
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
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
              "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
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
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
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
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
            
      Truncate sAppServer, "tmp_SALDOS"
      
      sSQL = " INSERT INTO tmp_SALDOS " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
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
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
            
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = "tmp_SALDOS"

      sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                    "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    
               
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
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
       sSQL_AA = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL,  NVL(EXCHANGE_RATE,1) EXCHANGE_RATE_1, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, (0) EXCHANGE_RATE_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (NUMERO_CUENTA_ESTANDAR in ('" & sAccount1 & "')) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, EXCHANGE_RATE "
      
      sSQL_A = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, " & _
               "         NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (NVL(EXCHANGE_RATE_1,1)/NVL(EXCHANGE_RATE,1)) EXCHANGE_RATE_1,  " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, (0) EXCHANGE_RATE_2,  " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_AA & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, EXCHANGE_RATE_1, EXCHANGE_RATE "
        
      If (cnFromCurrency = nExcRateCurrency) Then
        sSQL_A = sSQL_AA
      End If
      
      sSQL_BB = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, (0) EXCHANGE_RATE_1," & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_2, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2,  NVL(EXCHANGE_RATE,1) EXCHANGE_RATE_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (NUMERO_CUENTA_ESTANDAR in ('" & sAccount2 & "')) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, EXCHANGE_RATE "
               
      sSQL_B = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, " & _
               "         NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, (0) EXCHANGE_RATE_1," & _
               "         SUM(A.SALDO_INICIAL_2/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS_2/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS_2/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL_2/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         (NVL(EXCHANGE_RATE_2,1)/NVL(EXCHANGE_RATE,1)) EXCHANGE_RATE_2,  " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_BB & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS, EXCHANGE_RATE_2, EXCHANGE_RATE "
      
      If (cnFromCurrency = nExcRateCurrency) Then
        sSQL_B = sSQL_BB
      End If
           
        sSQL = " SELECT  id_moneda, " & _
               "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "         SUM(SALDO_ACTUAL) SALDO_ACTUAL, MAX(EXCHANGE_RATE_1) EXCHANGE_RATE, " & _
               "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, " & _
               "         SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, MAX(EXCHANGE_RATE_2) EXCHANGE_RATE_2 " & _
               " FROM (" & _
                 sSQL_A & _
               " UNION " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, " & _
               "          ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR " & _
               " ORDER BY NUMERO_CUENTA_AUXILIAR "
         
      Else
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
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
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS " & _
                " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      End If
     End If
    
    If bWithProm Then
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
        Set .ActiveConnection = Nothing
      End With
    Else
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    End If
    Set BalanceCompByAccountRS = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceCompByAccountRS", Err.Number
End Function

Public Function BalanceCompLastLevels(sAppServer As String, vGL As Variant, _
                                    dIniDate As Date, dCuteDate As Date, _
                                    dIniDate2 As Date, dCuteDate2 As Date, _
                                    Optional sIniAccount As String = "", _
                                    Optional sEndAccount As String = "", _
                                    Optional nTransactionTypeId As Long = 0, _
                                    Optional nVoucherTypeId As Long = 0, _
                                    Optional nCurrencyType As Long = 0, _
                                    Optional nExcRateType As Long = 0, _
                                    Optional nExcRateCurrency As Long = 0, _
                                    Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                                    Optional bWithAux As Boolean = False, _
                                    Optional bWithProm As Boolean = False) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sSQL_B              As String
  Dim sSQL_BA             As String
  Dim sSQL_SI_B           As String
  Dim sSQL1_B             As String
  Dim sSQL2_B             As String
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sOptionToDisplay  As String
  
  Dim sFromTable        As String
  Dim sFromTable_B       As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist      As String
  Dim sSQLStdAccHist_B    As String
  
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
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
            
    
    sSQL1_B = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
              "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
              " FROM COF_SALDO_INICIAL " & _
              " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate2 & "') " & _
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
            
    sSQL2_B = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
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
              "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate2 - 1) & "')  and " & _
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
   
   sSQLStdAccHist_B = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, " & _
                      "        NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                      " FROM COF_CUENTA_ESTANDAR_HIST " & _
                      " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate2 & "')) AND " & _
                      "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate2 & "'))) " & _
                      " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
                     
                     
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
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
                         
    sSQL_SI_B = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
                "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
                " FROM ( " & _
                  sSQL1_B & _
                " UNION " & _
                  sSQL2_B & _
                " ) A  , COF_CUENTA B, (" & sSQLStdAccHist_B & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA "
    
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
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sSQL2_B = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
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
             "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate2 & "')  and " & _
             "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate2 & "')  and " & _
             "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
             "       (COF_TRANSACCION." & sGl & " ) " & sWhereTran & _
             " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
             "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
                       

    Truncate sAppServer, "tmp_SALDOS"
    
    sSQL = " INSERT INTO tmp_SALDOS " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
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
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
            
    Truncate sAppServer, "tmp_SALDOS_B"
        
    sSQL_B = " INSERT INTO tmp_SALDOS_B " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                " FROM ( " & _
                  sSQL_SI_B & _
                " UNION " & _
                  sSQL2_B & _
                    " ) A  , COF_CUENTA B, (" & sSQLStdAccHist_B & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                        sWhere & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
              
              
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL_B
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = "tmp_SALDOS"
    sFromTable_B = "tmp_SALDOS_B"
    
    'proceso primer periodo
    
    'proceso segundo periodo
    
    
    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    sSQL2_B = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate2 & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate2 & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate2 & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            
    If nExcRateCurrency = 0 Then
      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
               "        (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2 , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , " & sFromTable & " A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) "
               
               '& _
                       sOptionToDisplay
               
        sSQL_B = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
               "        (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL , " & _
               "        SALDO_INICIAL SALDO_INICIAL_2, CARGOS CARGOS_2, ABONOS ABONOS_2, SALDO_ACTUAL SALDO_ACTUAL_2 , " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2_B & " ) CUENTA_STD , " & sFromTable_B & " A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) "
               
               '& _
                       sOptionToDisplay
               
        sSQL = " select A.ID_MONEDA, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        A.CLAVE_SECTOR, " & _
               "        A.NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM ( " & _
                 sSQL & _
               " union " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NATURALEZA, NOMBRE_CUENTA_AUXILIAR, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
               
               
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
               "        (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , " & sFromTable & " A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) "
               
               '& sOptionToDisplay
               
        sSQL_B = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
                 "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                 "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                 "        (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
                 "        SALDO_INICIAL SALDO_INICIAL_2, CARGOS CARGOS_2, ABONOS ABONOS_2, SALDO_ACTUAL SALDO_ACTUAL_2, " & _
                 "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                 " FROM ( " & sSQL2_B & " ) CUENTA_STD , " & sFromTable_B & " A , COF_SECTOR, " & _
                 "       AO_CURRENCIES " & _
                 " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                 "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
                 "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
                 "       (A.ID_CUENTA_AUXILIAR = 0) "
                 
                 '& sOptionToDisplay
                 
                 '& _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
                       
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        A.CLAVE_SECTOR, " & _
               "        A.NOMBRE_CUENTA, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM ( " & _
                 sSQL & _
               " UNION " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, NATURALEZA, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
               
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , " & sFromTable & " A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
               
        'sOptionToDisplay &
        
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A
        End If
     sSQL_BA = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOMBRE_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , " & sFromTable_B & " A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "
'sOptionToDisplay &
               
      sSQL_B = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_BA & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.NOMBRE_CUENTA_AUXILIAR, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_B = sSQL_BA
        End If
        
        sSQL = " SELECT  id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
               "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                 sSQL & _
               " UNION " & _
                 sSQL_B & _
               " ) A " & _
               " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
               "          CLAVE_SECTOR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, NATURALEZA, " & _
               "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
               
      Else
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , " & sFromTable & " A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                'sOptionToDisplay &
                
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A
        End If
        
     sSQL_BA = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , (" & sFromTable_B & " A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
                'sOptionToDisplay &
                
        sSQL_B = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_BA & " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_B = sSQL_BA
        End If
                
        sSQL = " SELECT  id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
                "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
                "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, (0) SALDO_PROMEDIO_2, " & _
                "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                " FROM (" & _
                  sSQL & _
                " UNION " & _
                  sSQL_B & _
                " ) A " & _
                " GROUP BY ID_MONEDA, O_ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA, " & _
                "          CLAVE_SECTOR, NATURALEZA, " & _
                "          ID_TIPO_CUENTA, ROL_CUENTA, FROM_MOVS " & _
                " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
               
      End If
    End If
    
    If bWithProm Then
      With oRs
       .CursorLocation = adUseClient
       .Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
       Set .ActiveConnection = Nothing
      End With
    Else
      With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      End With
    End If
    Set BalanceCompLastLevels = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceCompLastLevels", Err.Number
End Function

Private Function BalanceExtended(sAppServer As String, sPatron As String, vGL As Variant, _
                                dIniDate As Date, dCuteDate As Date, _
                                Optional sIniAccount As String = "", _
                                Optional sEndAccount As String = "", _
                                Optional nTransactionTypeId As Long = 0, _
                                Optional nVoucherTypeId As Long = 0, _
                                Optional nCurrencyType As Long = 0, _
                                Optional nExcRateType As Long = 0, _
                                Optional nExcRateCurrency As Long = 0, _
                                Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                                Optional bWithAux As Boolean = False, _
                                Optional bWithProm As Boolean = False) As Recordset
  Dim oConnection As New Connection, oCommand As New Command, oRecordset As New Recordset
  Dim sGl As String, sSQL As String, sSQL_A As String, sSQL1 As String, sSQL2 As String, sSQL_SI As String
  Dim sWhere As String, sWhereTran As String, sSelect As String, sInitSelect As String
  Dim sOptionToDisplay  As String, sFromTable As String, nPos As Variant, i As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  '***************************************************************
  On Error GoTo ErrHandler
    oConnection.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConnection.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConnection.Open

    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      nGl = CLng(vGL)
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
    
    If (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 1) Then
      sOptionToDisplay = " WHERE (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " WHERE ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " WHERE ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " WHERE (SALDO_ACTUAL < 0) "
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos, " & _
            "        SUM(SALDO_INICIAL) SALDO_INICIAL_MONEDA_BASE, (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE,   " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
               
    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
                     
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, " & _
              "        (0) Cargos, (0) Abonos,  " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
              "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_INICIAL_MONEDA_BASE, " & _
              "        (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE  " & _
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
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
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
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

     Truncate sAppServer, "TMP_Saldos_Ext"
     
     sSQL = " INSERT INTO TMP_Saldos_Ext " & _
            " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
            "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL, " & _
            "        A.SALDO_INICIAL_MONEDA_BASE, A.CARGOS_MONEDA_BASE, A.ABONOS_MONEDA_BASE, A.SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
            "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
            "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL, " & _
            "        sum(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, sum(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
            "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            sSQL_SI & _
            " UNION " & _
            sSQL2 & _
            " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
            sWhere & _
            " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
            "  ) A " & _
            sOptionToDisplay
            
    Set oCommand.ActiveConnection = oConnection
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
    
    sFromTable = " TMP_Saldos_Ext "
    
      sSQL = "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
             "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
             "     From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
                    "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                    "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, " & _
           "        SUM(SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, SUM(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
           " FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "              SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
           "              SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE  " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    
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
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
                       'sOptionToDisplay
      Else
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) and " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
       'sOptionToDisplay &
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS "
               
               'sOptionToDisplay &
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(A.SALDO_INICIAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) CARGOS_MONEDA_BASE, " & _
               "         SUM(A.ABONOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & "   ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
        
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
        End If
      Else
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (ID_CUENTA_AUXILIAR = 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS "
               'sOptionToDisplay &
      
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         SUM(A.SALDO_INICIAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) CARGOS_MONEDA_BASE, " & _
               "         SUM(A.ABONOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_MONEDA_BASE, (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
                       
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
        End If
      
      End If
    End If
    
    If bWithProm Then
      With oRecordset
        .CursorLocation = adUseClient
        .Open sSQL, oConnection, adOpenDynamic, adLockBatchOptimistic
        Set .ActiveConnection = Nothing
      End With
    Else
      With oRecordset
        .CursorLocation = adUseClient
        .Open sSQL, oConnection, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    End If
    
    Set BalanceExtended = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceExtended", Err.Number
End Function

Private Function BalanceExtendedProm(sAppServer As String, sPatron As String, vGL As Variant, _
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
  Dim oConnection As New Connection, oCommand As New Command, oRecordset As New Recordset
  Dim sGl As String, sSQL As String, sSQL_A As String, sSQL1 As String, sSQL2 As String, sSQL_SI As String
  Dim sWhere As String, sWhereTran As String, sSelect As String, sInitSelect As String
  Dim sOptionToDisplay  As String, sFromTable As String, nPos As Variant, i As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  '***************************************************************
  On Error GoTo ErrHandler
    oConnection.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConnection.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConnection.Open

    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      nGl = CLng(vGL)
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
    If (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 1) Then
      sOptionToDisplay = " WHERE (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " WHERE ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " WHERE ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " WHERE (SALDO_ACTUAL < 0) "
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos, " & _
            "        SUM(SALDO_INICIAL) SALDO_INICIAL_MONEDA_BASE, (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE,   " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
               
    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
                     
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        to_date('01/01/1990') FECHA_REV, max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, " & _
              "        (0) Cargos, (0) Abonos,  " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
              "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_INICIAL_MONEDA_BASE, " & _
              "        (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE  " & _
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
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
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

     Truncate sAppServer, "TMP_Saldos_Ext_PROM"
     
     sSQL = " INSERT INTO TMP_Saldos_Ext_PROM " & _
            " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
            "        A.FECHA_REV, A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL, " & _
            "        A.SALDO_INICIAL_MONEDA_BASE, A.CARGOS_MONEDA_BASE, A.ABONOS_MONEDA_BASE, A.SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
            "        FECHA_REV, sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
            "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL, " & _
            "        sum(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, sum(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
            "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            sSQL_SI & _
            " UNION " & _
            sSQL2 & _
            " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
            sWhere & _
            " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV " & _
            "  ) A " & _
            sOptionToDisplay
            
    Set oCommand.ActiveConnection = oConnection
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
    
    sFromTable = " TMP_Saldos_Ext_PROM "
      sSQL = "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
             "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
             "     From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
                    "           FECHA_REV, SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                    "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR, FECHA_REV " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "        FECHA_REV, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, " & _
           "        SUM(SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, SUM(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
           " FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "              FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
           "              SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE  " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV "
    
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
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "        SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
               'sOptionToDisplay
      Else
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "        SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) and " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
              'sOptionToDisplay

      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         FECHA_REV, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
               'sOptionToDisplay
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         FECHA_REV, SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) CARGOS_MONEDA_BASE, " & _
               "         SUM(A.ABONOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
        
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
        End If
      Else
      sSQL_A = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         FECHA_REV, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (ID_CUENTA_AUXILIAR = 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
              'sOptionToDisplay
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         FECHA_REV, SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) CARGOS_MONEDA_BASE, " & _
               "         SUM(A.ABONOS_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
      
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_A & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
        End If
      End If
    End If
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConnection, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set BalanceExtendedProm = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceExtendedProm", Err.Number
End Function

Private Function BalanceExtendedPromTotal(sAppServer As String, sPattern As String, vGL As Variant, _
                         dIniDate As Date, dEndDate As Date, _
                         Optional sIniAccount As String = "", Optional sEndAccount As String = "", _
                         Optional nTransactionTypeId As Long = 0, Optional nVoucherTypeId As Long = 0, _
                         Optional nCurrencyType As Long = 0, _
                         Optional nExcRateType As Long = 0, Optional nExcRateCurrency As Long = 0, _
                         Optional dExcDate As Date, _
                         Optional nOptionToDisplay As Long = 4, Optional bWithAux As Boolean = False) As Recordset
  Dim oRs           As Recordset
  Dim oRSTot        As Recordset
  
  Dim dFecha        As Date
  Dim sAccount      As String
  Dim sSector       As String
  Dim nSaldoMN      As Double
  Dim nSaldoMB      As Double
  Dim nSaldoMNA     As Double
  Dim nSaldoMBA     As Double
  Dim sAccName      As String
  
  Dim sNature       As String
  Dim sDateTime     As String
  Dim nCurrencyId   As Long
  
  Dim DebugAccount  As String
  '**************************************************************************************
  On Error GoTo ErrHandler
    
    Set oRs = BalanceExtendedProm(sAppServer, sPattern, vGL, dIniDate, dEndDate, sIniAccount, _
                          sEndAccount, nTransactionTypeId, nVoucherTypeId, nCurrencyType, _
                          nExcRateType, nExcRateCurrency, dExcDate, nOptionToDisplay, bWithAux)
    
    Set oRSTot = StructBalanceExtendedProm(sAppServer, sPattern, vGL, dIniDate, dEndDate, sIniAccount, _
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
      nSaldoMB = 0
      nSaldoMBA = 0
      Do While Not (.EOF)
        If (sAccount <> !NUMERO_CUENTA_ESTANDAR) Then
          If (Len(!NUMERO_CUENTA_ESTANDAR) <= Len(sPattern)) Then
            nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
            nSaldoMBA = (nSaldoMBA + (nSaldoMB * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
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
                !saldo_promedio_moneda_base = nSaldoMBA
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
                !saldo_promedio_moneda_base = nSaldoMBA
              End With
              oRs.MoveNext
            End If
          ElseIf (Len(!NUMERO_CUENTA_ESTANDAR) > Len(sPattern)) Then
            'Siguiente
          End If
          nSaldoMN = 0
          nSaldoMNA = 0
          nSaldoMB = 0
          nSaldoMBA = 0
          sAccount = !NUMERO_CUENTA_ESTANDAR
          sAccName = !nombre_cuenta
          sSector = !clave_sector
          dFecha = !fecha_rev
        ElseIf (sAccount = !NUMERO_CUENTA_ESTANDAR) Then
          If (sSector <> !clave_sector) Then
            nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
            nSaldoMBA = (nSaldoMBA + (nSaldoMB * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
                                
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
                !saldo_promedio_moneda_base = nSaldoMBA
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
                !saldo_promedio_moneda_base = nSaldoMBA
              End With
              oRs.MoveNext
            End If
            
            nSaldoMN = 0
            nSaldoMNA = 0
            nSaldoMB = 0
            nSaldoMBA = 0
            sAccount = !NUMERO_CUENTA_ESTANDAR
            sAccName = !nombre_cuenta
            sSector = !clave_sector
            dFecha = !fecha_rev
          End If
        End If
        
        If (dFecha = #1/1/1990#) Then
          nSaldoMN = !SALDO_INICIAL
          nSaldoMNA = 0
          nSaldoMB = !saldo_inicial_moneda_base
          nSaldoMBA = 0
          dFecha = dIniDate
        ElseIf (dFecha <> #1/1/1990#) Then
          nSaldoMNA = nSaldoMNA + (nSaldoMN * (!fecha_rev - dFecha))
          nSaldoMBA = nSaldoMBA + (nSaldoMB * (!fecha_rev - dFecha))
          If (!NATURALEZA = "D") Then
            nSaldoMN = nSaldoMN + !Cargos - !ABONOS
            nSaldoMB = nSaldoMB + !Cargos_moneda_base - !Abonos_moneda_base
          ElseIf (!NATURALEZA = "A") Then
            nSaldoMN = nSaldoMN - !Cargos + !ABONOS
            nSaldoMB = nSaldoMB - !Cargos_moneda_base + !Abonos_moneda_base
          End If
          dFecha = !fecha_rev
        End If
        .MoveNext
      Loop
      nSaldoMNA = (nSaldoMNA + (nSaldoMN * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
      nSaldoMBA = (nSaldoMBA + (nSaldoMB * ((dEndDate - dFecha) + 1))) / (dEndDate - dIniDate + 1)
      
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
          !saldo_promedio_moneda_base = nSaldoMBA
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
          !saldo_promedio_moneda_base = nSaldoMBA
          .Update
        End With
        oRs.MoveNext
      End If
    End With
    oRSTot.MoveFirst
    Set BalanceExtendedPromTotal = oRSTot
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceExtendedPromTotal", Err.Number
End Function

Public Function BalanceGeneral(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 0, _
                        Optional nExcRateCurrency As Long = 0, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False, _
                        Optional bIsExtended As Boolean = False, _
                        Optional bWithProm As Boolean = False, _
                        Optional bOnlyLastLevels As Boolean = False) As ADODB.Recordset
 ' Dim oBalance        As Recordset
 ' Dim oBalanceAvg     As Recordset
  Dim oBalance        As Object
  Dim oBalanceAvg     As Object

  '**************************************************************************************
  Set oBalance = CreateObject("ADODB.Recordset")
  Set oBalanceAvg = CreateObject("ADODB.Recordset")

  If Not bIsExtended And Not bWithProm Then
    
    Set oBalance = BalanceRS(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                             sIniAccount, sEndAccount, _
                             nTransactionTypeId, nVoucherTypeId, _
                             nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                             bWithAux, False, bOnlyLastLevels)
                           
  ElseIf Not bIsExtended And bWithProm Then
    Set oBalance = BalanceRS(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                           sIniAccount, sEndAccount, _
                           nTransactionTypeId, nVoucherTypeId, _
                           nCurrencyType, nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                           bWithAux, True)
                           
    Set oBalanceAvg = BalanceAvgTotalRS(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                                        sIniAccount, sEndAccount, _
                                        nTransactionTypeId, nVoucherTypeId, _
                                        nCurrencyType, _
                                        nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                                        bWithAux)
    With oBalance
      If Not (.BOF And .EOF) Then
        oBalance.MoveFirst
        oBalanceAvg.MoveFirst
        Do While Not .EOF
          !saldo_promedio = oBalanceAvg!saldo_promedio
          .Update
          .MoveNext
          oBalanceAvg.MoveNext
        Loop
      End If
      .MoveFirst
    End With
    
  ElseIf bIsExtended And Not bWithProm Then
    Set oBalance = BalanceExtended(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                           sIniAccount, sEndAccount, _
                           nTransactionTypeId, nVoucherTypeId, _
                           nCurrencyType, _
                           nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                           bWithAux)
                           
  ElseIf bIsExtended And bWithProm Then
    Set oBalance = BalanceExtended(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                           sIniAccount, sEndAccount, _
                           nTransactionTypeId, nVoucherTypeId, _
                           nCurrencyType, _
                           nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                           bWithAux, True)
                           
    Set oBalanceAvg = BalanceExtendedPromTotal(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                                        sIniAccount, sEndAccount, _
                                        nTransactionTypeId, nVoucherTypeId, _
                                        nCurrencyType, _
                                        nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                                        bWithAux)
    With oBalance
      If Not (.BOF And .EOF) Then
        oBalance.MoveFirst
        oBalanceAvg.MoveFirst
        Do While Not .EOF
          !saldo_promedio = oBalanceAvg!saldo_promedio
          !saldo_promedio_moneda_base = oBalanceAvg!saldo_promedio_moneda_base
          .Update
          .MoveNext
          oBalanceAvg.MoveNext
        Loop
      End If
      .MoveFirst
    End With
  End If
    Set BalanceGeneral = oBalance
    'oBalance.Close
    Set oBalance = Nothing
    'Set oBalance.ActiveConnection = Nothing
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceGeneral", Err.Number
End Function

Private Function BalanceRS(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 0, _
                        Optional nExcRateCurrency As Long = 0, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False, _
                        Optional bWithProm As Boolean = False, _
                        Optional bOnlyLastLevels As Boolean = False) As ADODB.Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_B            As String
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
  Dim vtimer            As Variant
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
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
      sOptionToDisplay = " AND (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " AND ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " AND ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " AND (SALDO_ACTUAL < 0) "
      bOnlyLastLevels = True
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
              "        max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
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
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
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
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
            
''      With oRs
''        .CursorLocation = adUseClient
''        vtimer = Timer
''        .Open sSQL2, oConn, adOpenStatic, adLockReadOnly
''        MsgBox Timer - vtimer
''        MsgBox .RecordCount
''        Set .ActiveConnection = Nothing
''      End With
''      oRs.Close
            
      Truncate sAppServer, "tmp_SALDOS"
      
      sSQL = " INSERT INTO tmp_SALDOS " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
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
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A "
              
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
'    vtimer = Timer
    oComm.Execute
'    MsgBox Timer - vtimer
    
      sFromTable = " tmp_SALDOS "

    If Not bOnlyLastLevels Then
      sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) AND " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
  
      nPos = GetPositions(sPatron)
      For i = LBound(nPos) To UBound(nPos)
        sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                      "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                      "     From  " & sFromTable & " A " & _
                      "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                      "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0)) " & _
                      "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                      " Union "
      Next i
      
      sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
             "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
             "        SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
             " ( " & sSQL & _
             "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
             "              SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
             "       From  " & sFromTable & " A " & _
             " ) " & _
             " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    Else
    
      If bWithAux Then
        sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "        SUM(SALDO_ACTUAL) SALDO_ACTUAL " & _
               " FROM  " & sFromTable & " A " & _
               " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
      Else
        sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "        SUM(SALDO_ACTUAL) SALDO_ACTUAL " & _
               " FROM  " & sFromTable & " A " & _
               " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR "
      End If
    
    End If
               
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
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                       sOptionToDisplay & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
                       sOptionToDisplay & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then

        sSQL_B = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
                  "       (A.ID_CUENTA_AUXILIAR <> 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID, " & cnFromCurrency & " ) = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & sOptionToDisplay & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "


        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_B & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_B & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
        End If
      Else
        
        sSQL_B = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID, " & cnFromCurrency & " ) = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & sOptionToDisplay & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "


        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM  (" & sSQL_B & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "

        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_B & _
                 " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
        End If
                
      End If
     End If
    
    If bWithProm Then
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenDynamic, adLockBatchOptimistic
        Set .ActiveConnection = Nothing
      End With
    Else
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    End If
    Set BalanceRS = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "BalanceRS", Err.Number
End Function

Public Function Comparative(sAppServer As String, sPatron As String, vGL As Variant, _
                            dIniDate As Date, dCuteDate As Date, _
                            dIniDate2 As Date, dCuteDate2 As Date, _
                            Optional sIniAccount As String = "", _
                            Optional sEndAccount As String = "", _
                            Optional nTransactionTypeId As Long = 0, _
                            Optional nVoucherTypeId As Long = 0, _
                            Optional nCurrencyType As Long = 0, _
                            Optional nExcRateType As Long = 0, _
                            Optional nExcRateCurrency As Long = 0, _
                            Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                            Optional bWithAux As Boolean = False, _
                            Optional bWithProm As Boolean = False) As Recordset
  Dim oBalanceComp        As Recordset
  Dim oBalanceCompProm    As Recordset
  '**************************************************************************************
  If Not bWithProm Then
    Set oBalanceComp = BalanceComparative(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                            dIniDate2, dCuteDate2, sIniAccount, sEndAccount, _
                           nTransactionTypeId, nVoucherTypeId, _
                           nCurrencyType, _
                           nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                           bWithAux)
                           
  ElseIf bWithProm Then
    Set oBalanceComp = BalanceComparative(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                           dIniDate2, dCuteDate2, sIniAccount, sEndAccount, _
                           nTransactionTypeId, nVoucherTypeId, _
                           nCurrencyType, _
                           nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                           bWithAux, True)
                           
    ' promedios para el primer periodo
    Set oBalanceCompProm = BalanceAvgTotalRS(sAppServer, sPatron, vGL, dIniDate, dCuteDate, _
                                        sIniAccount, sEndAccount, _
                                        nTransactionTypeId, nVoucherTypeId, _
                                        nCurrencyType, _
                                        nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                                        bWithAux)
    
    With oBalanceComp
      If Not (.BOF And .EOF) Then
        oBalanceComp.MoveFirst
        oBalanceCompProm.MoveFirst
        Do While Not .EOF
          If bWithAux Then
            If (!ID_MONEDA = oBalanceCompProm!ID_MONEDA) And (!NUMERO_CUENTA_ESTANDAR = oBalanceCompProm!NUMERO_CUENTA_ESTANDAR) And (!clave_sector = oBalanceCompProm!clave_sector) And (!ID_CUENTA_AUXILIAR = oBalanceCompProm!ID_CUENTA_AUXILIAR) Then
              !saldo_promedio = oBalanceCompProm!saldo_promedio
              .Update
              oBalanceCompProm.MoveNext
            End If
          Else
            If (!ID_MONEDA = oBalanceCompProm!ID_MONEDA) And (!NUMERO_CUENTA_ESTANDAR = oBalanceCompProm!NUMERO_CUENTA_ESTANDAR) And (!clave_sector = oBalanceCompProm!clave_sector) Then
              !saldo_promedio = oBalanceCompProm!saldo_promedio
              .Update
              oBalanceCompProm.MoveNext
            End If
          End If
          .MoveNext
        Loop
      End If
    End With
    
    ' promedios para el segundo periodo
    Set oBalanceCompProm = BalanceAvgTotalRS(sAppServer, sPatron, vGL, dIniDate2, dCuteDate2, _
                                        sIniAccount, sEndAccount, _
                                        nTransactionTypeId, nVoucherTypeId, _
                                        nCurrencyType, _
                                        nExcRateType, nExcRateCurrency, dExcDate, nRowsToDisplay, _
                                        bWithAux)
    
    With oBalanceComp
      If Not (.BOF And .EOF) Then
        oBalanceComp.MoveFirst
        oBalanceCompProm.MoveFirst
        Do While Not .EOF
          If bWithAux Then
            If (!ID_MONEDA = oBalanceCompProm!ID_MONEDA) And (!NUMERO_CUENTA_ESTANDAR = oBalanceCompProm!NUMERO_CUENTA_ESTANDAR) And (!clave_sector = oBalanceCompProm!clave_sector) And (!ID_CUENTA_AUXILIAR = oBalanceCompProm!ID_CUENTA_AUXILIAR) Then
              !saldo_promedio_2 = oBalanceCompProm!saldo_promedio
              .Update
              oBalanceCompProm.MoveNext
            End If
          Else
            If (!ID_MONEDA = oBalanceCompProm!ID_MONEDA) And (!NUMERO_CUENTA_ESTANDAR = oBalanceCompProm!NUMERO_CUENTA_ESTANDAR) And (!clave_sector = oBalanceCompProm!clave_sector) Then
              !saldo_promedio_2 = oBalanceCompProm!saldo_promedio
              .Update
              oBalanceCompProm.MoveNext
            End If
          End If
          .MoveNext
        Loop
      End If
      .MoveFirst
    End With
  End If
    Set Comparative = oBalanceComp
  Exit Function
ErrHandler:
  RaiseError ClassId, "Comparative", Err.Number
End Function

Private Function StructBalanceExtendedProm(sAppServer As String, sPatron As String, vGL As Variant, _
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
  Dim oConnection As New Connection, oCommand As New Command, oRecordset As New Recordset
  Dim sGl As String, sSQL As String, sSQL1 As String, sSQL2 As String, sSQL_SI As String
  Dim sWhere As String, sWhereTran As String, sSelect As String, sInitSelect As String
  Dim sOptionToDisplay  As String, sFromTable As String, nPos As Variant, i As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  '***************************************************************
  On Error GoTo ErrHandler
    oConnection.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConnection.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConnection.Open

    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      nGl = CLng(vGL)
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)
    
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
    If (nRowsToDisplay = 4) Then
      sOptionToDisplay = ""
    ElseIf (nRowsToDisplay = 1) Then
      sOptionToDisplay = " WHERE (SALDO_INICIAL <> 0) "
    ElseIf (nRowsToDisplay = 2) Then
      sOptionToDisplay = " WHERE ((CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 3) Then
      sOptionToDisplay = " WHERE ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0)) "
    ElseIf (nRowsToDisplay = 5) Then ' SOBREGIRADOS
      sOptionToDisplay = " WHERE (SALDO_ACTUAL < 0) "
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos, " & _
            "        SUM(SALDO_INICIAL) SALDO_INICIAL_MONEDA_BASE, (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE,   " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
               
    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "
                     
    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.ID_CUENTA, A.ID_SECTOR , A.ID_CUENTA_AUXILIAR ,   " & _
              "        to_date('01/01/1990') FECHA_REV, max(A.FECHA_SALDO_INICIAL) FECHA_SALDO_INICIAL, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, " & _
              "        (0) Cargos, (0) Abonos,  " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
              "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_INICIAL_MONEDA_BASE, " & _
              "        (0) Cargos_MONEDA_BASE, (0) Abonos_MONEDA_BASE  " & _
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
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos, " & _
            "        (0) SALDO_INICIAL_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto_Moneda_Base,0))) as Cargos_MONEDA_BASE, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto_Moneda_Base,0))) as Abonos_MONEDA_BASE " & _
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

     Truncate sAppServer, "TMP_Saldos_Ext_PROM"
     
     sSQL = " INSERT INTO TMP_Saldos_Ext_PROM " & _
            " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
            "        A.FECHA_REV, A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL, " & _
            "        A.SALDO_INICIAL_MONEDA_BASE, A.CARGOS_MONEDA_BASE, A.ABONOS_MONEDA_BASE, A.SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
            "        FECHA_REV, sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
            "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL, " & _
            "        sum(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, sum(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, " & _
            "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL_MONEDA_BASE)+sum(A.CARGOS_MONEDA_BASE)-sum(A.ABONOS_MONEDA_BASE)), " & _
            "                               (sum(A.SALDO_INICIAL_MONEDA_BASE)-sum(A.CARGOS_MONEDA_BASE)+sum(A.ABONOS_MONEDA_BASE)) )) SALDO_ACTUAL_MONEDA_BASE " & _
            " FROM ( " & _
            sSQL_SI & _
            " UNION " & _
            sSQL2 & _
            " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
            "       (A.ID_MONEDA = 0 ) " & _
            sWhere & _
            " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV " & _
            "  ) A " & _
            sOptionToDisplay
            
    Set oCommand.ActiveConnection = oConnection
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
    
    sFromTable = " TMP_Saldos_Ext_PROM "
    
    sSQL = "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
           "           FECHA_REV, SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
           "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
           "     From  " & sFromTable & "  A " & _
           "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
           "          (nvl(ID_SECTOR,0) <> 0) " & _
           "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4), FECHA_REV " & _
           " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR,(0) ID_CUENTA_AUXILIAR, " & _
                    "           FECHA_REV, SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                    "           SUM(A.SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, Sum(A.CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE,  Sum(A.ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(A.SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
                    "     From  " & sFromTable & " A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0))" & _
                    "    GROUP BY id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR, FECHA_REV " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "        FECHA_REV, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, " & _
           "        SUM(SALDO_INICIAL_MONEDA_BASE) SALDO_INICIAL_MONEDA_BASE, SUM(CARGOS_MONEDA_BASE) CARGOS_MONEDA_BASE, SUM(ABONOS_MONEDA_BASE) ABONOS_MONEDA_BASE, SUM(SALDO_ACTUAL_MONEDA_BASE) SALDO_ACTUAL_MONEDA_BASE " & _
           " FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "              FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, " & _
           "              SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE  " & _
           "       From  " & sFromTable & " A " & _
           " ) " & _
           " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_REV "
    
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
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        FECHA_REV, (0) SALDO_PROMEDIO, " & _
               "        (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) AND (A.ID_MONEDA = 0 ) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
               'sOptionToDisplay
      Else
        sSQL = " SELECT A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "        FECHA_REV, (0) SALDO_PROMEDIO, " & _
               "        (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) and " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) AND (A.ID_MONEDA = 0 ) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
               'sOptionToDisplay
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         FECHA_REV, (0) SALDO_PROMEDIO, " & _
               "         (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND (A.ID_MONEDA = 0) " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_REV "
               'sOptionToDisplay
      Else
        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         FECHA_REV, (0) SALDO_PROMEDIO, " & _
               "         (0) SALDO_PROMEDIO_MONEDA_BASE, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
               "   SELECT A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "          DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) as NOMBRE_CUENTA, " & _
               "          FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
               "          SALDO_INICIAL_MONEDA_BASE, CARGOS_MONEDA_BASE, ABONOS_MONEDA_BASE, SALDO_ACTUAL_MONEDA_BASE , " & _
               "          NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               "   FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
               "   WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "         (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "         (ID_CUENTA_AUXILIAR = 0) " & _
               "   ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
               " WHERE (C.CURRENCY_ID = " & nExcRateCurrency & ") and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND (A.ID_MONEDA = 0) " & _
               " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, FECHA_REV "
               
               'sOptionToDisplay
      End If
    End If
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConnection, adOpenDynamic, adLockBatchOptimistic
      Set .ActiveConnection = Nothing
    End With
    Set StructBalanceExtendedProm = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "StructBalanceExtendedProm", Err.Number
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
                       
    Truncate sAppServer, "tmp_SALDOS_PROM"
    
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

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Sub Commit(sAppServer As String)
  Dim oConn             As New Connection
  Dim oComm             As New Command
  '***************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = "COMMIT"
    oComm.CommandType = adCmdText
    oComm.Execute
    
    oConn.Close
    
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Commit", Err.Number
End Sub

Private Function GetPositions(ByVal sPatron As String) As Variant
  Dim sPositions   As String
  Dim nPos         As Long
  '*******************************************************
  On Error GoTo ErrHandler
    nPos = 0
    Do
      nPos = InStr(nPos + 1, sPatron, "-")
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
      sPositions = Len(sPatron)
    ElseIf (Len(sPositions) <> 0) Then
      sPositions = sPositions & "," & Len(sPatron)
    End If
    GetPositions = Split(sPositions, ",")
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetPositions", Err.Number
End Function

Private Sub Truncate(sAppServer As String, sTable As String)
  Dim oConn             As New Connection
  Dim oComm             As New Command
  Dim sSQL_TRUNC        As String
  '**************************************************************
  On Error GoTo ErrHandler
  
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    'oConn.ConnectionString = "DSN=COFINI;UID=INTERNAL;PWD=ORACLE"
    oConn.Open
    
    sSQL_TRUNC = "DELETE FROM " & sTable
    'sSQL_TRUNC = "DROP table " & sTable
    
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL_TRUNC
    oComm.CommandType = adCmdText
    oComm.Execute
    
    oConn.Close
    Commit sAppServer
    'CreateTable sAppServer, sTable
    
  Exit Sub
ErrHandler:
  RaiseError ClassId, "Truncate", Err.Number
End Sub

Private Function TypeStdAccount(sAppServer As String, nGl As Long) As Long
  Dim oConn             As New ADODB.Connection
  Dim oRs               As New ADODB.Recordset
  Dim sSQLTypeStdAcc    As String
  '**********************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    sSQLTypeStdAcc = " SELECT ID_TIPO_CUENTAS_STD FROM COF_MAYOR WHERE (ID_MAYOR = " & nGl & ")"
    
    With oRs
      '.CursorLocation = adUseClient
      .Open sSQLTypeStdAcc, oConn, adOpenStatic, adLockReadOnly
      TypeStdAccount = !id_tipo_cuentas_std
      .Close
      Set .ActiveConnection = Nothing
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "TypeStdAccount", Err.Number
End Function



