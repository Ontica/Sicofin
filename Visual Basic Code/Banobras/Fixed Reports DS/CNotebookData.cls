VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNotebookData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : Fixed Reports DS (SCFFixedReportsDS)            Parte   : Data services                     *
'* Clase      : CNotebookData                                   Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Genera las fuentes de datos para los reportes que constituyen el cuaderno de delegaciones.  *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CNotebookData"

Private Const StdAccHist As String = _
  " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR " & _
  " FROM COF_CUENTA_ESTANDAR_HIST " & _
  " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '<@dIniDate@>') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '<@dEndDate@>')) AND " & _
  "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '<@dIniDate@>') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '<@dEndDate@>'))) " & _
  " GROUP BY ID_CUENTA_ESTANDAR "

Private sStdAccHist As String

'************************************************************************************************************
'* MÉTODOS PÚBLICOS                                                                                         *
'************************************************************************************************************

Public Function GetObjName(sAppServer As String, nId As Long, sTable As String, _
                            sFieldName As String, sFieldCondition As String, _
                            Optional sFieldName2 As String = "") As String
  Dim oConn    As New ADODB.Connection
  Dim oRs      As New ADODB.Recordset
   
  Dim sObjName As String
  Dim sSQL     As String
  '************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
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

Public Function GetAccountName(sAppServer As String, nAcc As String, _
                              nGl As Long) As String
  Dim oConn    As New ADODB.Connection
  Dim oRs      As New ADODB.Recordset
   
  Dim sObjName As String
  Dim sSQL     As String
  '************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sSQL = " SELECT nvl(NOMBRE_CUENTA_ESTANDAR,'SIN NOMBRE') NOMBRE" & _
           " FROM COF_CUENTA_ESTANDAR, COF_CUENTA " & _
           " WHERE (COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR = COF_CUENTA.ID_CUENTA_ESTANDAR) and " & _
           "       (NUMERO_CUENTA_ESTANDAR  = '" & nAcc & "') and " & _
           "       (ID_MAYOR = " & nGl & " ) "
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      If (.BOF And .EOF) Then
        sObjName = " "
      Else
        sObjName = .Fields(0)
      End If
    End With
    GetAccountName = sObjName
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountName", Err.Number
End Function

Public Function GeneraRSNames(sAppServer As String, vGL As Variant)
  Dim oRecordset      As New ADODB.Recordset
  Dim oConn           As New ADODB.Connection
    
  Dim sSqlComp        As String
  Dim sGL             As String
  Dim sGLCondition    As String
'******************************************************************
 On Error GoTo ErrHandler
  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
  oConn.Open
   
    If (IsArray(vGL)) Then
      sGL = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGL & ") "
    Else
      sGL = vGL
      sGLCondition = "id_mayor = " & sGL
    End If
   
  sSqlComp = " SELECT COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR NUM_CUENTA, " & _
                    " COF_CUENTA_ESTANDAR.NOMBRE_CUENTA_ESTANDAR NOM_CUENTA,  " & _
                    " COF_CUENTA.ID_MAYOR " & _
             " From COF_CUENTA_ESTANDAR, COF_CUENTA " & _
             " WHERE (COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR = COF_CUENTA.ID_CUENTA_ESTANDAR(+)) AND " & _
                   " (COF_CUENTA." & sGLCondition & ") " & _
             " ORDER BY ID_MAYOR, NUMERO_CUENTA_ESTANDAR "
    
  With oRecordset
    .CursorLocation = adUseClient
    .Open sSqlComp, oConn, adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
  End With
   Set GeneraRSNames = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRSNames", Err.Number
End Function

Public Function GeneraAuxiliares(sAppServer As String, _
                                 dCuteDate As Date, _
                                 sAccounts As String, nGl As Long) As ADODB.Recordset
  Dim oConn                 As New ADODB.Connection
  Dim oRs                   As New ADODB.Recordset

  Dim sSQL                  As String
  Dim ssql1                 As String
  Dim ssql2                 As String
  Dim sSQLStdAccHist        As String
'  Dim sSql3                As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

'            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND " &

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',MONTO,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',MONTO,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
                    " NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

'    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'            "        NOMBRE_CUENTA_ESTANDAR, " & _
'            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
'            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
'            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
'            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'            "       and (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
'            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
'                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
'             " UNION " & _
'             sSql3 &

      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, " & _
             "        nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, " & _
             "        A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A  , COF_SECTOR , COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY  A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
                       " NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraAuxiliares = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraAuxiliares", Err.Number
End Function

Public Function GeneraResumen(sAppServer As String, _
                              dCuteDate As Date, _
                              sAccounts As String, nGl As Long) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oRs                  As New ADODB.Recordset

  Dim sSQL                 As String
  Dim ssql1                As String
  Dim ssql2                As String
'  Dim sSql3                As String
  Dim sSQLStdAccHist       As String
  '********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          NATURALEZA "

'            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',MONTO,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',MONTO,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "          NATURALEZA "

'    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'                     " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
'            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
'            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
'            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'            "       and (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
'            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
'                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'                    " COF_CUENTA_ESTANDAR.NATURALEZA "
'

'             " UNION " & _
'             sSql3 &

      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, " & _
             "        nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "        NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A, COF_SECTOR, COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY  A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
                       " NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 "
             
             '& _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
    
    
    sSQL = " SELECT CLAVE_SECTOR, sum(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
           " FROM ( " & sSQL & " ) A " & _
           " GROUP BY CLAVE_SECTOR " & _
           " ORDER BY CLAVE_SECTOR "
           
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraResumen = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraResumen", Err.Number
End Function

Public Function GetRS_R7006(sAppServer As String, dCuteDate As Date, dExcDate As Date, _
                            vAccounts As Variant, nGl As Long, _
                            Optional nExcRateType As Long = 46, _
                            Optional nExcRateCurrency As Long = 1) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oRs                  As New ADODB.Recordset

  Dim sGL                  As String
  Dim sSQL                 As String
  Dim ssql1                As String
  Dim ssql2                As String
'  Dim sSql3                As String

  Dim sCondition           As String
  
  Dim i                    As Long
  
  Dim sAccounts            As String
  Dim sWhereAcc            As String
  Dim sSQLStdAccHist       As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vAccounts)) Then
      sAccounts = Join(vAccounts, "','")
      sWhereAcc = " AND ( SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) in ('" & sAccounts & "')) "
    Else
      sAccounts = vAccounts
      sWhereAcc = " AND ( SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) = '" & sAccounts & "') "
    End If
    
    sGL = "id_mayor = " & nGl
    
                       
                       '" (AO_EXCHANGE_RATES.FROM_CURRENCY_ID = 1) and " & _
                       "       (EXCHANGE_RATE_TYPE = " & nExcType & ") and " & _
                       " ('" & dExcDate & "' >= nvl(FROM_DATE,'" & dExcDate & "') ) and " & _
                       " ('" & dExcDate & "' <= nvl(TO_DATE,'" & dExcDate & "') ) "

    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
      
      ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, " & _
                     " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
                     " COF_CUENTA_ESTANDAR.NATURALEZA, " & _
                     " sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
              " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
              " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
              "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
              "       (COF_SALDO_INICIAL." & sGL & ") " & _
                    sWhereAcc & _
              " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4), " & _
                    " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
                    " COF_CUENTA_ESTANDAR.NATURALEZA "

      ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, " & _
                     " COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
                     " COF_CUENTA_ESTANDAR.NATURALEZA, (0) SALDO_INICIAL, " & _
                     " (sum(decode(TIPO_MOVIMIENTO,'D',MONTO,0))) Cargos, " & _
                     " (sum(decode(TIPO_MOVIMIENTO,'H',MONTO,0))) Abonos " & _
              " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
              " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
                    " (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
                    " (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
                    " (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
                    " (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
                    " (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
                    " (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
                    " (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
                    " (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
                    " (COF_TRANSACCION. " & sGL & ") " & _
                    sWhereAcc & _
              " GROUP BY COF_MOVIMIENTO.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4), " & _
                    " COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, COF_CUENTA_ESTANDAR.NATURALEZA "

      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, " & _
             "       NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
             "       (decode(A.NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
             "                                (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)))*nvl(EXCHANGE_RATE,1)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A , COF_CUENTA_AUXILIAR , AO_EXCHANGE_RATES " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) and " & _
             "       (A.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+) ) and " & _
             "       (EXCHANGE_RATE_TYPE_ID(+) = " & nExcRateType & ") and " & _
             "       (FROM_CURRENCY_ID(+) = " & nExcRateCurrency & ") and " & _
             "       (FROM_DATE(+) <= '" & dExcDate & "') AND " & _
             "       (TO_DATE(+) >= '" & dExcDate & "') " & _
             " GROUP BY A.NUMERO_CUENTA_ESTANDAR, NUMERO_CUENTA_AUXILIAR, A.NATURALEZA, EXCHANGE_RATE, " & _
             "          NOMBRE_CUENTA_AUXILIAR " & _
             " ORDER BY NUMERO_CUENTA_AUXILIAR "
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
      
    Set GetRS_R7006 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRS_R7006", Err.Number
End Function

Public Function GetRS_7010(sAppServer As String, nGl As Long, _
                           dCuteDate As Date, _
                           sAccounts As String) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oCommand             As ADODB.Command
  Dim oRs                  As New ADODB.Recordset

  Dim sSQL                 As String
  Dim ssql1                As String
  Dim ssql2                As String
  Dim sSql3                As String
  Dim sSqlAux              As String
  Dim sSqlComp             As String
  Dim sSQLStdAccHist       As String
  '****************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
                     
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (FECHA_SALDO_INICIAL <= '" & (dCuteDate) & "') AND " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "          COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

'''''    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
'''''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
'''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
'''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'''''            "       and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
'''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
'''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, " & _
             "        A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, TO_DATE('') FECHA_REGISTRO, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A, COF_SECTOR, COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY  A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "           NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
    
    ' Para obtener los campos de los auxiliares
               
''''    sSqlAux = CreaSQLCamposAux()
''''
''''    sSqlComp = " SELECT A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOMBRE_SECTOR, A.CLAVE_SECTOR, " & _
''''               "        A.ID_CUENTA_AUXILIAR, nvl(A.NUMERO_CUENTA_AUXILIAR,'Sin Aux') numero_cuenta_auxiliar, " & _
''''               "        nvl(A.NOMBRE_CUENTA_AUXILIAR,'Sin Aux') nombre_cuenta_auxiliar, A.NATURALEZA, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
''''               "        NVL(B.FECHA_REGISTRO, TO_DATE('')) FECHA_REGISTRO " & _
''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
''''               " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOMBRE_SECTOR, A.CLAVE_SECTOR, " & _
''''               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, " & _
''''               "          B.FECHA_REGISTRO " & _
''''               " ORDER BY b.fecha_registro, a.NUMERO_CUENTA_ESTANDAR, a.numero_cuenta_auxiliar "
               

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetRS_7010 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRS_7010", Err.Number
End Function
                                  
Public Function GeneraRSAnalitico(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional bOnlyLastLevels As Boolean = False) As ADODB.Recordset

  'Dim oConn             As Object
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  'Dim oRs               As Object
  Dim oRs               As New ADODB.Recordset

  Dim sGL               As String

  Dim sSQL              As String

  Dim sSQL_B            As String

  Dim sSQL_SI           As String
  Dim ssql1             As String
  Dim ssql2             As String

  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sSelect           As String
  Dim sInitSelect       As String
  Dim sMaxSelect        As String
  Dim sOptionToDisplay  As String

  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long

  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long

  Dim sSQLStdAccHist    As String
  Dim sGlCons           As String
  Dim sGroupByGL        As String
  'Dim dinitimer  As Variant
  '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
      sGL = Join(vGL, ",")
      sGL = "id_mayor in (" & sGL & ") "
    Else
      nGl = CLng(vGL)
      sGL = vGL
      sGL = "id_mayor = " & sGL
    End If

    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)

    If (sIniAccount = "") And (sEndAccount = "") Then
      sWhere = " "
    ElseIf (sIniAccount <> "") And (sEndAccount <> "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sIniAccount & "' AND '" & sEndAccount & "') "
    ElseIf (sIniAccount <> "") And (sEndAccount = "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR <= '" & sIniAccount & "') "
    End If

    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", dIniDate)
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "

    ssql1 = " SELECT ID_MONEDA, COF_SALDO_INICIAL.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "        ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL,  COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (COF_SALDO_INICIAL." & sGL & ") AND " & _
            "       (FECHA_SALDO_INICIAL <= '" & (dCuteDate) & "') AND " & _
            "       (COF_SALDO_INICIAL.ID_CUENTA = B.ID_CUENTA) AND " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & sWhere & _
            " group by ID_MONEDA, COF_SALDO_INICIAL.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "          ID_SECTOR, ID_CUENTA_AUXILIAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION." & sGL & " ) AND " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = B.ID_CUENTA) AND " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & sWhere & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sSQL_SI = " SELECT A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, A.ID_SECTOR,  A.ID_CUENTA_AUXILIAR, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
              " FROM ( " & _
                ssql1 & _
              " UNION " & _
                ssql2 & _
              " ) A " & _
              " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, " & _
              "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = B.ID_CUENTA) AND " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION." & sGL & " ) " & sWhereTran & sWhere & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

      Truncate sAppServer, "tmp_SALDOS"

      sSQL = " INSERT INTO tmp_SALDOS " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                " FROM ( " & _
                  sSQL_SI & _
                " UNION " & _
                  ssql2 & _
                    " ) A " & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA "

'                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _

              '& _
              sOptionToDisplay    ' CHECAR OPCION DE MOVIMIENTOS<>0 O SALDOS<>0 '' se cambio a los querys originales

    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute

      sFromTable = " tmp_SALDOS "

'sSQL = ""

''    sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
''           "           (0) ID_CUENTA_AUXILIAR, " & _
''           "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
''           "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
''           "    From  " & sFromTable & "  A " & _
''           "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
''           "          not ((SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4) = NUMERO_CUENTA_ESTANDAR) AND (nvl(ID_SECTOR,0)=0)) " & _
''           "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
''           " Union "

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

        sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "        SUM(SALDO_ACTUAL) SALDO_ACTUAL " & _
               " FROM  " & sFromTable & " A " & _
               " GROUP BY ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR "

    End If

    ssql2 = " SELECT NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, " & _
            "        ID_TIPO_CUENTA, ROL_CUENTA  " & _
            " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
            " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
            "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) AND " & _
            "       (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  "

''''''''''    ssql2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
''''''''''            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
''''''''''            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
''''''''''            " FROM COF_CUENTA_ESTANDAR_HIST " & _
''''''''''            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
''''''''''            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
''''''''''            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
''''''''''            " GROUP BY NUMERO_CUENTA_ESTANDAR "


        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, CURRENCY_NAME NOMBRE_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        NOMBRE_CUENTA_ESTANDAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & ssql2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & _
                       sOptionToDisplay & _
               "       and (length(A.numero_cuenta_estandar)<= length('" & sPatron & "')) " & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRSAnalitico = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRSAnalitico", Err.Number
End Function

Public Function CreaSQLCamposAux(Optional bWithDir As Boolean = False) As String
  Dim ssql1     As String
  Dim sSql1A    As String
  Dim ssql2     As String
  Dim sSql3     As String
  Dim sSql4     As String
  Dim sSql5     As String
  Dim sSql6     As String
  Dim sSql7     As String
  Dim sSqlPart1 As String
  Dim sSqlCom   As String
  '*****************************************
  On Error GoTo ErrHandler

  If (bWithDir) Then
    ssql1 = " SELECT A.ID_CUENTA_AUXILIAR, MAX(A.DOMICILIO) DOMICILIO, MAX(A.RFC) RFC, MAX(A.FECHA_REGISTRO) FECHA_REGISTRO, " & _
              " MAX(A.FECHA_OFICIO) FECHA_OFICIO, MAX(A.NUMERO_OFICIO) NUMERO_OFICIO,  MAX(A.DESCRIPCION) DESCRIPCION, " & _
              " MAX(A.COLONIA) COLONIA, MAX(A.ESTADO) ESTADO "
  ElseIf Not (bWithDir) Then
    ssql1 = " SELECT A.ID_CUENTA_AUXILIAR, MAX(A.DOMICILIO) DOMICILIO, MAX(A.RFC) RFC, MAX(A.FECHA_REGISTRO) FECHA_REGISTRO, " & _
              " MAX(A.FECHA_OFICIO) FECHA_OFICIO, MAX(A.NUMERO_OFICIO) NUMERO_OFICIO,  MAX(A.DESCRIPCION) DESCRIPCION, " & _
              " MAX(A.CURP) CURP, MAX(A.CLAVE_RETENCION) CLAVE_RETENCION, MAX(A.IVA) IVA, MAX(A.ISR) ISR, " & _
              " MAX(A.ESTA_GRAVADO) ESTA_GRAVADO "
  End If
              
   sSql1A = " FROM ( " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING DOMICILIO, (' ') RFC, " & _
                   " TO_DATE('') FECHA_REGISTRO, TO_DATE('') FECHA_OFICIO, (' ') NUMERO_OFICIO, " & _
                   " (' ') DESCRIPCION, (' ') CURP, (' ') CLAVE_RETENCION, (0) IVA, " & _
                   " (0) ISR, (0) ESTA_GRAVADO, (' ') COLONIA, (' ') ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='DOMICILIO') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING  AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='RFC') " & _
            " Union "
            
    ssql2 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " COF_ATTR_CUENTA_AUXILIAR.VAL_DATE  AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='FECHA_REGISTRO') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, COF_ATTR_CUENTA_AUXILIAR.VAL_DATE AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='FECHA_OFICIO') " & _
            " Union "
            
    sSql3 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='NUMERO_OFICIO') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='DESCRIPCION') " & _
            " Union "
            
    sSql4 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='CURP') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='CLAVE_RETENCION') " & _
            " Union "
    sSql5 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, COF_ATTR_CUENTA_AUXILIAR.VAL_FLOAT AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='IVA') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " COF_ATTR_CUENTA_AUXILIAR.VAL_FLOAT AS ISR, (0) AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='ISR') " & _
            " Union "
            
    sSql6 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, COF_ATTR_CUENTA_AUXILIAR.VAL_BOOLEAN AS ESTA_GRAVADO, (' ') AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='ESTA_GRAVADO') " & _
            " Union " & _
            " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) AS ESTA_GRAVADO, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS COLONIA, (' ') AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='COLONIA') " & _
            " Union "

        
   sSql7 = " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, (' ') AS DOMICILIO, (' ') AS RFC, " & _
                   " TO_DATE('') AS FECHA_REGISTRO, TO_DATE('') AS FECHA_OFICIO, (' ') AS NUMERO_OFICIO, " & _
                   " (' ') AS DESCRIPCION, (' ') AS CURP, (' ') AS CLAVE_RETENCION, (0) AS IVA, " & _
                   " (0) AS ISR, (0) ESTA_GRAVADO, (' ') AS COLONIA, COF_ATTR_CUENTA_AUXILIAR.VAL_STRING AS ESTADO " & _
            " From AO_ATTR_DEF, AO_ENTITY_ATTR_DEF, COF_CUENTA_AUXILIAR, COF_ATTR_CUENTA_AUXILIAR " & _
            " WHERE  (AO_ATTR_DEF.ATTR_DEF_ID = AO_ENTITY_ATTR_DEF.ATTR_DEF_ID) AND " & _
                   " (COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR = COF_ATTR_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                   " (AO_ENTITY_ATTR_DEF.ENTITY_ATTR_DEF_ID = COF_ATTR_CUENTA_AUXILIAR.ENTITY_ATTR_DEF_ID) AND " & _
                   " (AO_ATTR_DEF.ATTR_DEF_NAME='ESTADO') " & _
        " ) A " & _
        " GROUP BY A.ID_CUENTA_AUXILIAR "
  
  sSqlPart1 = ssql1 & sSql1A & ssql2 & sSql3 & sSql4
  
  sSqlCom = sSqlPart1 & sSql5 & sSql6 & sSql7
  CreaSQLCamposAux = sSqlCom
Exit Function
ErrHandler:
  RaiseError ClassId, "CreaSQLCamposAux", Err.Number
End Function

Public Function GetRS_R7002(sAppServer As String, nGl As Long, _
                            dCuteDate As Date, sAccount As String, dExcDate As Date, _
                            Optional nExcRateType As Long = 46, _
                            Optional nExcRateCurrency As Long = 1) As ADODB.Recordset
                            
  Dim oConn                As New ADODB.Connection
  Dim oCommand             As ADODB.Command
  Dim oRs                  As New ADODB.Recordset

  Dim sGL                  As String
  Dim sSQL                 As String
  Dim ssql1                As String
  Dim ssql2                As String
  Dim sSql3                As String
  Dim sSQLA                As String

  Dim sCondition           As String
  Dim sGroupby             As String
  Dim sOrderby             As String
  
  Dim i                    As Long
  
  Dim sAccounts            As String
  Dim sAccCond             As String
  Dim sSQLStdAccHist       As String
  '***********************************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
        
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

'Debug.Print sSQLStdAccHist

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
        
  ssql1 = " SELECT   A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, " & _
         "           A.ID_MAYOR, A.NATURALEZA, NVL(COF_SECTOR.CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
         "           (DECODE(A.NATURALEZA,'D',SUM(A.SALDO_INICIAL)+SUM(A.CARGOS)-SUM(A.ABONOS), " & _
         "                                'A',SUM(A.SALDO_INICIAL)-SUM(A.CARGOS)+SUM(A.ABONOS)) * NVL(EXCHANGE_RATE,1)) IMPORTE " & _
         " FROM ( " & _
         "      SELECT NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
         "             Sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) AS CARGOS, (0) AS ABONOS,  " & _
         "             COF_SALDO_INICIAL.ID_MAYOR, COF_CUENTA_ESTANDAR.NATURALEZA,  COF_SALDO_INICIAL.ID_MONEDA, " & _
         "             COF_SALDO_INICIAL.ID_SECTOR " & _
         "      FROM  (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_CUENTA, COF_SALDO_INICIAL " & _
         "      WHERE (COF_SALDO_INICIAL.ID_CUENTA = COF_CUENTA.ID_CUENTA) AND " & _
         "            (COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR = COF_CUENTA.ID_CUENTA_ESTANDAR) AND " & _
         "            (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & " ) AND " & _
         "            (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR LIKE '" & sAccount & "%')  AND " & _
         "            (COF_SALDO_INICIAL.FECHA_SALDO_INICIAL <=' " & dCuteDate & "') " & _
         "     GROUP BY NUMERO_CUENTA_ESTANDAR,COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_MONEDA , " & _
         "              COF_SALDO_INICIAL.ID_MAYOR, COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR " & _
     " Union "
 ssql2 = "      SELECT NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
         "            (0) SALDO_INICIAL, " & _
         "            SUM(DECODE(TIPO_MOVIMIENTO, 'D' ,MONTO,0 )) CARGOS , " & _
         "            SUM(DECODE(TIPO_MOVIMIENTO, 'H' ,MONTO,0 )) ABONOS ,   " & _
         "            COF_TRANSACCION.ID_MAYOR, COF_CUENTA_ESTANDAR.NATURALEZA, COF_MOVIMIENTO.ID_MONEDA, " & _
         "            COF_MOVIMIENTO.ID_SECTOR " & _
         "      FROM (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_CUENTA, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
         "      WHERE (COF_MOVIMIENTO.ID_CUENTA = COF_CUENTA.ID_CUENTA) AND " & _
         "            (COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR = COF_CUENTA.ID_CUENTA_ESTANDAR) AND " & _
         "            (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) AND " & _
         "            (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) AND " & _
         "            (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) AND " & _
         "            (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) AND " & _
         "            (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) AND " & _
         "            (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "') AND " & _
         "            (FECHA_AFECTACION >= NVL(fecha_saldo_inicial,'01/01/1990')) AND " & _
         "            (COF_TRANSACCION.ID_MAYOR = " & nGl & " ) AND " & _
         "            (NUMERO_CUENTA_ESTANDAR LIKE '" & sAccount & "%') " & _
         "      GROUP BY NUMERO_CUENTA_ESTANDAR,COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_MONEDA, TIPO_MOVIMIENTO, " & _
         "               COF_TRANSACCION.ID_MAYOR, COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR " & _
         " ) A, COF_SECTOR, AO_EXCHANGE_RATES "
         
 sSql3 = " WHERE (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
         "       (A.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND " & _
         "       (AO_EXCHANGE_RATES.EXCHANGE_RATE_TYPE_ID(+) = " & nExcRateType & ") and " & _
         "       (AO_EXCHANGE_RATES.FROM_CURRENCY_ID(+) = " & nExcRateCurrency & ") and " & _
         "       (AO_EXCHANGE_RATES.FROM_DATE(+) <= '" & dExcDate & "') AND " & _
         "       (AO_EXCHANGE_RATES.TO_DATE(+) >= '" & dExcDate & "') " & _
         " GROUP BY  A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR,  " & _
         "           A.ID_MAYOR, A.NATURALEZA, A.NOMBRE_CUENTA_ESTANDAR, EXCHANGE_RATE " & _
         " HAVING ((DECODE(A.NATURALEZA,'D', SUM(A.SALDO_INICIAL) + SUM(A.CARGOS) - SUM(A.ABONOS), " & _
         "                             'A', SUM(A.SALDO_INICIAL) - SUM(A.CARGOS) + SUM(A.ABONOS)) * NVL(EXCHANGE_RATE,1)) <> 0) "

    sSQL = ssql1 & ssql2 & sSql3
    
    
    
    sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, " & _
           "         A.ID_MAYOR, A.NATURALEZA, CLAVE_SECTOR, " & _
           "         SUM(A.IMPORTE) IMPORTE " & _
           " FROM ( " & sSQL & " ) A " & _
           " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, " & _
           "          A.ID_MAYOR, A.NATURALEZA, CLAVE_SECTOR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
      
    Set GetRS_R7002 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRS_R7002", Err.Number
End Function

Public Function GeneraRecordsetC7016(sAppServer As String, vGL As Variant, sSubAcc As String, _
                                     dIniDate As Date, dEndDate As Date, _
                                     Optional dExcDate As Date, _
                                     Optional nExcRateType As Long = 46, _
                                     Optional nExcRateCurrency As Long = 1)
  Dim oConn           As New ADODB.Connection
  Dim oRecordset      As New ADODB.Recordset
  
  Dim sSQL            As String
  Dim ssql1           As String
  Dim ssql2           As String
  Dim sSqlAux         As String
  Dim sSql3103        As String
  Dim sSql1801        As String
  Dim sGL             As String
  Dim sSQLStdAccHist  As String
  '************************************************************************
  On Error GoTo ErrHandler
  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
  oConn.Open
  
    If (IsArray(vGL)) Then
      sGL = Join(vGL, ",")
      sGL = "id_mayor in (" & sGL & ") "
    Else
      sGL = vGL
      sGL = "id_mayor = " & sGL
    End If

    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", dIniDate)
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dEndDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "

    ssql1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGL & ") and (FECHA_SALDO_INICIAL <= '" & dEndDate & "') and " & _
            "       (NVL(ID_CUENTA_AUXILIAR,0) <> 0) " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "
            
    
    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dEndDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGL & " ) and " & _
            "       (NVL(COF_MOVIMIENTO.ID_CUENTA_AUXILIAR,0) <> 0) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
    
    sSql3103 = " SELECT A.ID_CUENTA_AUXILIAR, (0) ST_A, " & _
              "         sum(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,7), " & _
              "                    '3103-" & sSubAcc & "', (A.SALDO_INICIAL-A.CARGOS+A.ABONOS),0)) ST_B " & _
              " FROM ( " & _
                ssql1 & _
              " UNION " & _
                ssql2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY A.ID_CUENTA_AUXILIAR "

    sSql1801 = " SELECT A.ID_CUENTA_AUXILIAR, " & _
              "        sum(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,7), " & _
              "                   '1801-" & sSubAcc & "', (A.SALDO_INICIAL+A.CARGOS-A.ABONOS),0)) ST_A , " & _
              "        (0) ST_B " & _
              " FROM ( " & _
              ssql1 & _
              " UNION " & _
              ssql2 & _
              " ) A, COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY A.ID_CUENTA_AUXILIAR "
    
'''''    sSqlAux = CreaSQLCamposAux()
'''''
'''''    sSQL = " SELECT NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
'''''              "     A.ID_CUENTA_AUXILIAR,   " & _
'''''              "     sum(nvl(A.ST_A,0)) ST_A , sum(nvl(A.ST_B,0)) ST_B , " & _
'''''              "     NVL(E.FECHA_REGISTRO, TO_DATE('')) FECHA_REGISTRO   " & _
'''''              " FROM ( " & _
'''''                sSql1801 & _
'''''              " UNION " & _
'''''                sSql3103 & _
'''''              " ) A, COF_CUENTA_AUXILIAR D, ( " & sSqlAux & " ) E " & _
'''''              " WHERE (A.ID_CUENTA_AUXILIAR = D.ID_CUENTA_AUXILIAR) and " & _
'''''              "       (A.ID_CUENTA_AUXILIAR = E.ID_CUENTA_AUXILIAR(+)) " & _
'''''              " GROUP BY A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
'''''              "          NOMBRE_CUENTA_AUXILIAR, FECHA_REGISTRO " & _
'''''              " ORDER BY FECHA_REGISTRO, NUMERO_CUENTA_AUXILIAR "
              
    sSQL = " SELECT NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, " & _
              "     A.ID_CUENTA_AUXILIAR,   " & _
              "     sum(nvl(A.ST_A,0)) ST_A , sum(nvl(A.ST_B,0)) ST_B , " & _
              "     TO_DATE('') FECHA_REGISTRO   " & _
              " FROM ( " & _
                sSql1801 & _
              " UNION " & _
                sSql3103 & _
              " ) A, COF_CUENTA_AUXILIAR D " & _
              " WHERE (A.ID_CUENTA_AUXILIAR = D.ID_CUENTA_AUXILIAR) " & _
              " GROUP BY A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
              "          NOMBRE_CUENTA_AUXILIAR " & _
              " ORDER BY NUMERO_CUENTA_AUXILIAR "
              
'    sSqlComp = " SELECT A.NUM_AUX, A.NOM_AUX, A.AUX, SUM(A.ST_A) AS ST_A, SUM (A.ST_B) AS ST_B, NVL(B.FECHA_REGISTRO, TO_DATE('')) FECHA_REGISTRO " & _
               " FROM  ( " & sSqlComp & " ) A, ( " & sSqlAux & " ) B " & _
               " WHERE (A.AUX = B.ID_CUENTA_AUXILIAR(+)) " & _
               " GROUP BY A.NUM_AUX, A.NOM_AUX, A.AUX, B.FECHA_REGISTRO " & _
               " ORDER BY B.FECHA_REGISTRO, A.NUM_AUX "
               
               
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordsetC7016 = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordsetC7016", Err.Number
End Function

Public Function GetObjNameStr(sAppServer As String, sNumber As String, sTable As String, _
                            sFieldName As String, sFieldCondition As String, _
                            Optional sFieldName2 As String = "") As String
  Dim oConn    As New ADODB.Connection
  Dim oRs      As New ADODB.Recordset
   
  Dim sObjName As String
  Dim sSQL     As String
  '************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (Len(sFieldName2) = 0) Then
      sObjName = sFieldName
    ElseIf (Len(sFieldName2) <> 0) Then
      sObjName = sFieldName & " , " & sFieldName2
    End If
    
    sSQL = "SELECT " & sObjName & " FROM " & sTable & " WHERE " & sFieldCondition & " = '" & sNumber & "'"
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
    GetObjNameStr = sObjName
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetObjNameStr", Err.Number
End Function

Public Function GetRS_7014(sAppServer As String, nGl As Long, _
                           dCuteDate As Date, _
                           sAccounts As String) As ADODB.Recordset
  Dim oConn                 As New ADODB.Connection
  Dim oCommand              As ADODB.Command
  Dim oRs                   As New ADODB.Recordset

  Dim sSQL                  As String
  Dim ssql1                 As String
  Dim ssql2                 As String
  Dim sSql3                 As String
  Dim sSqlAux               As String
  Dim sSqlComp              As String
  Dim sSQLStdAccHist        As String
  '****************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') AND " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
                    " COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "       COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

'''''    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
'''''            "        NOMBRE_CUENTA_ESTANDAR, COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, " & _
'''''            "        (0) SALDO_INICIAL,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
'''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
'''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'''''            "       and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
'''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
'''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, " & _
             "        A.ID_CUENTA_AUXILIAR, nvl(COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
             "        nvl(COF_CUENTA_AUXILIAR.NOMBRE_CUENTA_AUXILIAR, ' ') NOMBRE_CUENTA_AUXILIAR, " & _
             "        nvl(COF_CUENTA_AUXILIAR.DESCRIPCION, ' ') DESCRIPCION, " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
               ssql1 & _
             " UNION " & _
               ssql2 & _
             " ) A, COF_SECTOR, COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "          NOMBRE_CUENTA_AUXILIAR, COF_CUENTA_AUXILIAR.DESCRIPCION, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
    
    ' Para obtener los campos de los auxiliares
               
''''''    sSqlAux = CreaSQLCamposAux()
''''''
''''''    sSqlComp = " SELECT A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
''''''                        " SUM(A.SA) AS SA, A.FECHA_ULT, " & _
''''''                        " A.NOM_SECTOR, A.NUM_SECTOR , A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
''''''                        "  nvl(B.DESCRIPCION,' ') DESCRIPCION  " & _
''''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
''''''               " GROUP BY A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
''''''                       " A.FECHA_ULT, A.NOM_SECTOR, A.NUM_SECTOR, A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
''''''                       " B.DESCRIPCION "
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetRS_7014 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRS_7014", Err.Number
End Function

Public Function GeneraRS_7021(sAppServer As String, _
                              dCuteDate As Date, _
                              vGL As Variant) As Recordset
  Dim oConn               As New ADODB.Connection
  Dim oRs                 As New ADODB.Recordset

  Dim sSQL                As String
  Dim ssql1               As String
  Dim ssql2               As String
  Dim sSql3               As String
  
  Dim sGL                 As String
  Dim sAccounts           As String
  Dim sAccounts1          As String
  Dim sAccounts2          As String
  Dim sSQLStdAccHist      As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      sGL = Join(vGL, ",")
      sGL = "id_mayor in (" & sGL & ") "
    Else
      sGL = vGL
      sGL = "id_mayor = " & sGL
    End If
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    sAccounts1 = "'1350-00-00-00-04-01-01','1350-00-00-00-04-01-02','1350-00-00-00-04-01-03','1350-00-00-00-04-01-04'," & _
                 "'1350-00-00-00-04-02-01','1350-00-00-00-04-03-01','1350-00-00-00-04-04-01','1350-00-00-00-06-02-01'," & _
                 "'1350-00-00-00-04-02-04','1350-00-00-00-07-02-01','1350-00-00-00-08-02-01'," & _
                 "'1350-00-00-00-04-02-30','1350-00-00-00-04-03-30','1350-00-00-00-04-04-30','1350-00-00-00-06-02-30'," & _
                 "'1350-00-00-00-07-02-30','1350-00-00-00-08-02-30'," & _
                 "'1319-52-00-00-04-01-02','1319-52-00-00-04-02-00','1319-52-00-00-04-03-00','1319-52-00-00-04-04-00'," & _
                 "'1319-52-00-00-06-02-00','1319-52-00-00-07-02-00','1319-52-00-00-08-02-00'," & _
                 "'1319-30-52-00-04-01-00','1319-30-52-00-04-02-00','1319-30-52-00-04-03-00','1319-30-52-00-06-01-00'," & _
                 "'1319-30-52-00-07-01-00','1319-30-52-00-08-01-00'," & _
                 "'1319-30-52-00-04-04-00','1319-30-52-00-04-05-00','1319-30-52-00-04-06-00','1319-30-52-00-06-02-00'," & _
                 "'1319-30-52-00-07-02-00','1319-30-52-00-08-02-00'," & _
                 "'1314-01-50-00-04-02-00','1314-01-50-00-04-03-00','1314-01-50-00-04-04-00','1314-01-50-00-06-00-00'," & _
                 "'1314-01-50-00-07-00-00','1314-01-50-00-08-00-00'," & _
                 "'1314-02-50-00-04-02-00','1314-02-50-00-04-03-00','1314-02-50-00-04-04-00','1314-02-50-00-06-00-00'," & _
                 "'1314-02-50-00-07-00-00','1314-02-50-00-08-00-00'," & _
                 "'1320-50-00-00-04-02-02','1320-50-00-00-04-03-02','1320-50-00-00-04-04-02','1320-50-00-00-06-02-00'," & _
                 "'1320-50-00-00-07-02-00','1320-50-00-00-08-02-00'," & _
                 "'1320-50-00-00-04-02-03','1320-50-00-00-04-03-03','1320-50-00-00-04-04-03','1320-50-00-00-06-03-00'," & _
                 "'1320-50-00-00-07-03-00','1320-50-00-00-08-03-00'," & _
                 "'6378-02-02-00-04-19-00','6378-02-02-00-04-20-00','6378-02-02-00-04-21-00','6378-02-02-00-04-25-00'," & _
                 "'6378-02-02-00-04-26-00','6378-02-02-00-04-27-00'," & _
                 "'6378-02-02-00-04-13-00','6378-02-02-00-04-14-00','6378-02-02-00-04-15-00','6378-02-02-00-04-31-00'," & _
                 "'6378-02-02-00-04-32-00','6378-02-02-00-04-33-00'," & _
                 "'6378-02-02-00-04-22-00','6378-02-02-00-04-23-00','6378-02-02-00-04-24-00','6378-02-02-00-04-28-00'," & _
                 "'6378-02-02-00-04-29-00','6378-02-02-00-04-30-00',"
                
    sAccounts2 = "'6378-02-02-00-04-16-00','6378-02-02-00-04-17-00','6378-02-02-00-04-18-00','6378-02-02-00-04-34-00'," & _
                 "'6378-02-02-00-04-35-00','6378-02-02-00-04-36-00'," & _
                 "'6378-02-02-00-04-01-00','6378-02-02-00-04-11-00','6378-02-02-00-04-12-00','6378-02-02-00-04-06-00'," & _
                 "'6378-02-02-00-04-07-00','6378-02-02-00-04-08-00'"
                 
    sAccounts = sAccounts1 & sAccounts2

    ssql1 = " SELECT COF_CUENTA.ID_CUENTA,COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) as Cargos, (0) as Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) and " & _
            "       (COF_SALDO_INICIAL." & sGL & ") " & _
            " GROUP BY COF_CUENTA.ID_CUENTA,COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR "
            
    ssql2 = " SELECT COF_CUENTA.ID_CUENTA,COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) AND " & _
            "       (COF_TRANSACCION." & sGL & ") " & _
            " GROUP BY COF_CUENTA.ID_CUENTA, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
    
    sSQL = " SELECT A.ID_CUENTA,ID_CUENTA_AUXILIAR, " & _
           "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
           "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)))) SALDO " & _
           " FROM ( " & _
             ssql1 & _
           " UNION " & _
             ssql2 & _
           " ) A, COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
           " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
           "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
           " GROUP BY  A.ID_CUENTA, ID_CUENTA_AUXILIAR, NATURALEZA " & _
           " HAVING (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
           "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS))) <> 0 ) "

    ssql1 = "select NUMERO_CUENTA_AUXILIAR AUX, NOMBRE_CUENTA_AUXILIAR ACREDITADO, " & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1350-00-00-00-04-01-01',SALDO,'1350-00-00-00-04-01-02',SALDO,'1350-00-00-00-04-01-03',SALDO," & _
                            "'1350-00-00-00-04-01-04',SALDO,'1350-00-00-00-04-02-01',SALDO,'1350-00-00-00-04-03-01',SALDO," & _
                            "'1350-00-00-00-04-04-01',SALDO,'1350-00-00-00-06-02-01',SALDO,'1350-00-00-00-04-02-04',SALDO," & _
                            "'1350-00-00-00-07-02-01',SALDO,'1350-00-00-00-08-02-01',SALDO)),0) CAPITAL_VIG," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1350-00-00-00-04-02-30',SALDO,'1350-00-00-00-04-03-30',SALDO,'1350-00-00-00-04-04-30',SALDO," & _
                            "'1350-00-00-00-06-02-30',SALDO,'1350-00-00-00-07-02-30',SALDO,'1350-00-00-00-08-02-30',SALDO)),0) AMORT_X_TRA," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1319-52-00-00-04-01-02',SALDO,'1319-52-00-00-04-02-00',SALDO,'1319-52-00-00-04-03-00',SALDO," & _
                            "'1319-52-00-00-04-04-00',SALDO,'1319-52-00-00-06-02-00',SALDO,'1319-52-00-00-07-02-00',SALDO," & _
                            "'1319-52-00-00-08-02-00',SALDO)),0) INT_VIGENTE," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1319-30-52-00-04-01-00',SALDO,'1319-30-52-00-04-02-00',SALDO,'1319-30-52-00-04-03-00',SALDO," & _
                            "'1319-30-52-00-06-01-00',SALDO,'1319-30-52-00-07-01-00',SALDO,'1319-30-52-00-08-01-00',SALDO)),0) INTS_NOR_X_TRASP," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1319-30-52-00-04-04-00',SALDO,'1319-30-52-00-04-05-00',SALDO,'1319-30-52-00-04-06-00',SALDO," & _
                            "'1319-30-52-00-06-02-00',SALDO,'1319-30-52-00-07-02-00',SALDO,'1319-30-52-00-08-02-00',SALDO)),0) INTS_COR_X_TRASP," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1314-01-50-00-04-02-00',SALDO,'1314-01-50-00-04-03-00',SALDO,'1314-01-50-00-04-04-00',SALDO," & _
                            "'1314-01-50-00-06-00-00',SALDO,'1314-01-50-00-07-00-00',SALDO,'1314-01-50-00-08-00-00',SALDO)),0) CAP_VENC_EXIGIBLE," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1314-02-50-00-04-02-00',SALDO,'1314-02-50-00-04-03-00',SALDO,'1314-02-50-00-04-04-00',SALDO," & _
                            "'1314-02-50-00-06-00-00',SALDO,'1314-02-50-00-07-00-00',SALDO,'1314-02-50-00-08-00-00',SALDO)),0) CAP_VENC_NO_EXIGIBLE,"
                            
    ssql2 = "        NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1320-50-00-00-04-02-02',SALDO,'1320-50-00-00-04-03-02',SALDO,'1320-50-00-00-04-04-02',SALDO," & _
                            "'1320-50-00-00-06-02-00',SALDO,'1320-50-00-00-07-02-00',SALDO,'1320-50-00-00-08-02-00',SALDO)),0) INTERESES_VENCIDS," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'1320-50-00-00-04-02-03',SALDO,'1320-50-00-00-04-03-03',SALDO,'1320-50-00-00-04-04-03',SALDO," & _
                            "'1320-50-00-00-06-03-00',SALDO,'1320-50-00-00-07-03-00',SALDO,'1320-50-00-00-08-03-00',SALDO)),0) INTERESES_CORRIDS," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'6378-02-02-00-04-19-00',SALDO,'6378-02-02-00-04-20-00',SALDO,'6378-02-02-00-04-21-00',SALDO," & _
                            "'6378-02-02-00-04-25-00',SALDO,'6378-02-02-00-04-26-00',SALDO,'6378-02-02-00-04-27-00',SALDO)),0) INTER_NORM_EJ_CSO," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'6378-02-02-00-04-13-00',SALDO,'6378-02-02-00-04-14-00',SALDO,'6378-02-02-00-04-15-00',SALDO," & _
                            "'6378-02-02-00-04-31-00',SALDO,'6378-02-02-00-04-32-00',SALDO,'6378-02-02-00-04-33-00',SALDO)),0) INTER_NORM_EJ_ANT," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'6378-02-02-00-04-22-00',SALDO,'6378-02-02-00-04-23-00',SALDO,'6378-02-02-00-04-24-00',SALDO," & _
                            "'6378-02-02-00-04-28-00',SALDO,'6378-02-02-00-04-29-00',SALDO,'6378-02-02-00-04-30-00',SALDO)),0) INTER_CORR_EJ_CSO," & _
                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR, " & _
                            "'6378-02-02-00-04-16-00',SALDO,'6378-02-02-00-04-17-00',SALDO,'6378-02-02-00-04-18-00',SALDO," & _
                            "'6378-02-02-00-04-34-00',SALDO,'6378-02-02-00-04-35-00',SALDO,'6378-02-02-00-04-36-00',SALDO)),0) INTER_CORR_EJ_ANT,"
                            
    sSql3 = "        NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
                            "'6378-02-02-00-04-01-00',SALDO,'6378-02-02-00-04-11-00',SALDO,'6378-02-02-00-04-12-00',SALDO," & _
                            "'6378-02-02-00-04-06-00',SALDO,'6378-02-02-00-04-07-00',SALDO,'6378-02-02-00-04-08-00',SALDO)),0) INTER_MORATOR " & _
            " FROM ( " & sSQL & " ) A, cof_cuenta B, (" & sSQLStdAccHist & ") C, COF_CUENTA_AUXILIAR D " & _
            " WHERE (A.id_cuenta = B.id_cuenta) and " & _
            "       (B.id_cuenta_estandar = C.id_cuenta_estandar) and " & _
            "       (A.id_cuenta_auxiliar = D.id_cuenta_auxiliar) " & _
            " GROUP BY NUMERO_CUENTA_AUXILIAR , NOMBRE_CUENTA_AUXILIAR " & _
            " ORDER BY NUMERO_CUENTA_AUXILIAR "

    sSQL = ssql1 & ssql2 & sSql3

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRS_7021 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraRS_7021", Err.Number
End Function

Public Function GetRS_7024(sAppServer As String, _
                           dCuteDate As Date, _
                           sAccounts As String, nGl As Long) As Recordset
  Dim oConn                 As New ADODB.Connection
  Dim oCommand              As ADODB.Command
  Dim oRs                   As New ADODB.Recordset

  Dim sSQL                  As String
  Dim ssql1                 As String
  Dim ssql2                 As String
  Dim sSql3                 As String
  Dim sSqlAux               As String
  Dim sSqlComp              As String
  Dim sSQLStdAccHist        As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, " & _
            "        (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "          COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

''    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
''            "       and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, " & _
             "        A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        (' ') RFC, (' ') DOMICILIO, (' ') COLONIA, (' ') ESTADO, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
               ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A, COF_SECTOR, COF_CUENTA_AUXILIAR" & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "          NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
    
    ' Para obtener los campos de los auxiliares
               
''''    sSqlAux = CreaSQLCamposAux(True)
''''
''''    sSqlComp = " SELECT A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''               "        A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, " & _
''''               "        SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
''''               "        NVL(B.RFC,' ') RFC, NVL(B.DOMICILIO,' ') DOMICILIO, NVL(B.COLONIA,' ') COLONIA, NVL(B.ESTADO,' ') ESTADO " & _
''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0)" & _
''''               " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, B.RFC, B.DOMICILIO, B.COLONIA, B.ESTADO " & _
''''               " ORDER BY A.NUMERO_CUENTA_ESTANDAR, A.NUMERO_CUENTA_AUXILIAR "
               
               
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetRS_7024 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRS_7024", Err.Number
End Function

Public Function GetRS_7028(sAppServer As String, nGl As Long, _
                           dCuteDate As Date, _
                           sAccounts As String) As Recordset
  Dim oConn                 As New ADODB.Connection
  Dim oCommand              As ADODB.Command
  Dim oRs                   As New ADODB.Recordset

  Dim sSQL                  As String
  Dim ssql1                 As String
  Dim ssql2                 As String
  Dim sSql3                 As String
  Dim sSqlAux               As String
  Dim sSqlComp              As String
  Dim sSQLStdAccHist        As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
    
    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, " & _
            "        (0) SALDO_INICIAL,  " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

''''    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
''''            "       and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
      sSQL = " SELECT A.NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOMBRE_SECTOR, " & _
             "        nvl(CLAVE_SECTOR, '00') CLAVE_SECTOR, " & _
             "        A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        TO_DATE('') FECHA_OFICIO, (' ') NUMERO_OFICIO, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A, COF_SECTOR, COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
             "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
             " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "          NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR "
    
    ' Para obtener los campos de los auxiliares
               
''''    sSqlAux = CreaSQLCamposAux()
''''
''''    sSqlComp = " SELECT A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOMBRE_SECTOR, " & _
''''               "        A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, " & _
''''               "        A.NATURALEZA, " & _
''''               "        SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
''''               "        NVL(B.FECHA_OFICIO, TO_DATE('')) FECHA_OFICIO, NVL(B.NUMERO_OFICIO,' ') NUMERO_OFICIO  " & _
''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
''''               " GROUP BY A.NUMERO_CUENTA_ESTANDAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOMBRE_SECTOR, A.CLAVE_SECTOR, " & _
''''               "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, " & _
''''               "          B.FECHA_OFICIO, B.NUMERO_OFICIO " & _
''''               " ORDER BY B.FECHA_OFICIO, A.NUMERO_CUENTA_ESTANDAR, A.NUMERO_CUENTA_AUXILIAR "
               

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetRS_7028 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetRS_7028", Err.Number
End Function

Public Function GetRS_7000(sAppServer As String, nGl As Long, _
                           dCuteDate As Date, _
                           sAccounts As String) As ADODB.Recordset
  Dim oConn                 As New ADODB.Connection
  Dim oCommand              As ADODB.Command
  Dim oRs                   As New ADODB.Recordset

  Dim sSQL                  As String
  Dim ssql1                 As String
  Dim ssql2                 As String
  Dim sSql3                 As String
  Dim sSqlAux               As String
  Dim sSqlComp              As String
  Dim sSQLStdAccHist        As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", "01/01/1990")
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)

    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "

    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "          NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
            "        NOMBRE_CUENTA_ESTANDAR, COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) SALDO_INICIAL,  " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, " & _
            "      COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) AND " & _
            "       (COF_TRANSACCION.ID_MAYOR = " & nGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
                    " NATURALEZA, NOMBRE_CUENTA_ESTANDAR "

'''    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'''            "        NOMBRE_CUENTA_ESTANDAR, " & _
'''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  Sum(COF_MOVIMIENTO.MONTO) AS Cargos, (0) as Abonos " & _
'''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
'''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='D') AND " & _
'''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'''            "       and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
'''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,  " & _
'''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
                   
      sSQL = " SELECT NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, " & _
             "        A.ID_CUENTA_AUXILIAR, nvl(COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
             "        nvl(COF_CUENTA_AUXILIAR.NOMBRE_CUENTA_AUXILIAR, ' ') NOMBRE_CUENTA_AUXILIAR, " & _
             "        nvl(COF_CUENTA_AUXILIAR.DESCRIPCION, ' ') DESCRIPCION, " & _
             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, " & _
             "        sum(A.SALDO_INICIAL) SALDO_INICIAL, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
             "        decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) SALDO_ACTUAL " & _
             " FROM ( " & _
             ssql1 & _
             " UNION " & _
             ssql2 & _
             " ) A , COF_CUENTA_AUXILIAR" & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) " & _
             " GROUP BY  NUMERO_CUENTA_ESTANDAR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
             "           NOMBRE_CUENTA_AUXILIAR, DESCRIPCION, NATURALEZA,  NOMBRE_CUENTA_ESTANDAR " & _
             " HAVING decode(A.NATURALEZA,'D',sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS), " & _
             "                                sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
             " ORDER BY NUMERO_CUENTA_ESTANDAR, NUMERO_CUENTA_AUXILIAR "
    
    ' Para obtener los campos de los auxiliares
               
''''''    sSqlAux = CreaSQLCamposAux()
''''''
''''''    sSqlComp = " SELECT A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
''''''                        " SUM(A.SA) AS SA, A.FECHA_ULT, " & _
''''''                        " A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
''''''                        "  nvl(B.DESCRIPCION,' ') DESCRIPCION  " & _
''''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
''''''               " GROUP BY A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
''''''                       " A.FECHA_ULT, A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
''''''                       " B.DESCRIPCION " & _
''''''               " ORDER BY A.Cuenta, NUMERO_CUENTA_AUXILIAR "
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetRS_7000 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GetRS_7000", Err.Number
End Function

'************************************************************************************************************
'* MÉTODOS PRIVADOS                                                                                         *
'************************************************************************************************************

Private Sub CreateGralMovTBL(oConn As ADODB.Connection, _
                            sGLCondition As String, _
                            sTableName As String)
  Dim oCommand             As New ADODB.Command
  Dim sSQL                 As String
  Dim sCondition           As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    If (Len(sGLCondition) = 0) Then
      GoTo ErrHandler
    ElseIf (Len(sGLCondition) <> 0) Then
    ' "INSERT INTO " & sTableName & " "
    '" Create Table " & sTableName & " AS "
     sSQL = "INSERT INTO " & sTableName & " " & _
           " SELECT COF_MOVIMIENTO.ID_CUENTA,COF_MOVIMIENTO.ID_SECTOR, " & _
           "        COF_MOVIMIENTO.ID_CUENTA_AUXILIAR,COF_MOVIMIENTO.TIPO_MOVIMIENTO, " & _
           "        COF_MOVIMIENTO.ID_MONEDA,COF_MOVIMIENTO.MONTO,COF_MOVIMIENTO.MONTO_MONEDA_BASE, " & _
           "        NVL(COF_SALDO_INICIAL.SALDO_INICIAL,0) AS SALDO_INICIAL, " & _
           "        NVL(COF_SALDO_INICIAL.FECHA_SALDO_INICIAL,'01/01/1990') AS FECHA_SALDO_INICIAL, " & _
           "        COF_TRANSACCION.ID_MAYOR,COF_TRANSACCION.FECHA_AFECTACION, NUMERO_TRANSACCION,(1) FROM_MOVS,CONCEPTO_TRANSACCION " & _
           " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
           " WHERE (((COF_MOVIMIENTO.ID_CUENTA=COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
           "         (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR=COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) AND " & _
           "         (COF_MOVIMIENTO.ID_SECTOR=COF_SALDO_INICIAL.ID_SECTOR(+)) AND " & _
           "         (COF_MOVIMIENTO.ID_MONEDA=COF_SALDO_INICIAL.ID_MONEDA(+))) and " & _
           "         (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION)) and " & _
           "    ( COF_TRANSACCION." & sGLCondition & ") "
             
      Set oCommand.ActiveConnection = oConn
      oCommand.CommandText = sSQL
      oCommand.CommandType = adCmdText
      oCommand.Execute
    End If
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CreateGralMovTbl", Err.Number
End Sub

Private Sub CreateMovTBL(oConn As ADODB.Connection, _
                            sGLCondition As String, sMovKind As String, _
                            dIniDate As Date, dEndDate As Date, _
                            sTableName As String, sMovName As String, _
                            Optional bWithAux As Boolean = False, _
                            Optional bGrouping As Boolean = True, _
                            Optional bWithFec As Boolean = False, _
                            Optional bWithGL As Boolean = False)
  Dim oCommand             As New ADODB.Command
  Dim sSQL                 As String
  Dim sCondition           As String
  Dim sAuxCondition        As String
  Dim sAux                 As String
  
  Dim sAuxTBLFields        As String
  Dim sGLTBLFields         As String
  
  Dim sGLField             As String
  Dim sGLJoin              As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    If (Len(sGLCondition) = 0) Then
      GoTo ErrHandler
    ElseIf (Len(sGLCondition) <> 0) Then
      sGLField = " "
      sGLJoin = " "
      sAuxTBLFields = " "
      sGLTBLFields = " "
      If (bWithAux) And (bGrouping) Then
        sAuxCondition = " ID_CUENTA_AUXILIAR as AUX, MAX(FECHA_AFECTACION) as FECHA_ULT, "
        sAux = ", ID_CUENTA_AUXILIAR "
        sAuxTBLFields = " AUX, FECHA_ULT, "
      ElseIf (bWithAux) And Not (bGrouping) Then
        sAuxCondition = " ID_CUENTA_AUXILIAR as AUX, FECHA_AFECTACION as FECHA_ULT,  " & _
                        " NUMERO_TRANSACCION,CONCEPTO_TRANSACCION,  "
        sAux = " ,ID_CUENTA_AUXILIAR , FECHA_AFECTACION ,  " & _
               " NUMERO_TRANSACCION,CONCEPTO_TRANSACCION "
        sAuxTBLFields = " AUX, FECHA_ULT, NUMERO_TRANSACCION, CONCEPTO_TRANSACCION,  "
      ElseIf Not (bWithAux) And (bWithFec) Then
        sAuxCondition = " FECHA_AFECTACION as FECHA_ULT, "
        sAux = " , FECHA_AFECTACION "
        sAuxTBLFields = " FECHA_ULT, "
      ElseIf Not (bWithAux) And Not (bWithFec) Then
        sAuxCondition = "  "
        sAux = " "
        If (bWithGL) Then
          sGLField = " COF_CUENTA.ID_MAYOR, "
          sGLJoin = " (tmp_MovsTBL.ID_MAYOR = COF_CUENTA.ID_MAYOR) and "
          sGLTBLFields = " ID_MAYOR, "
        End If
      End If
      ' AND (ROL_CUENTA <> 'S')
      sCondition = "((TIPO_MOVIMIENTO= '" & sMovKind & "' ) " & _
                   " AND (FECHA_AFECTACION >= '" & dIniDate & "')" & _
                   " And (FECHA_AFECTACION <= '" & dEndDate & "'))  " & _
                   " and ( fecha_afectacion >= fecha_saldo_inicial)"
                   
      ' "INSERT INTO " & stablename & " (ID_MONEDA, CUENTA, SECTOR, " & sauxTBLFields & sGLTBLFields  & _
                    " AMOUNT, " & sMovName & " ) "
      '" CREATE TABLE " & sTableName & " as "
      sSQL = "INSERT INTO " & sTableName & " (ID_MONEDA, CUENTA, SECTOR, " & sAuxTBLFields & sGLTBLFields & _
                    " AMOUNT, " & sMovName & " ,FROM_MOVS ) " & _
             " SELECT ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR AS CUENTA, " & _
                    " ID_SECTOR AS SECTOR, " & sAuxCondition & sGLField & _
                    " sum(MONTO) as AMOUNT, Sum(MONTO_MONEDA_BASE) AS " & sMovName & ", (1) FROM_MOVS  " & _
             " FROM tmp_MovsTBL, COF_CUENTA , COF_CUENTA_ESTANDAR " & _
             " WHERE ((tmp_MovsTBL.ID_CUENTA = COF_CUENTA.ID_CUENTA) and " & sGLJoin & _
                    " (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR)) and " & _
             "   (" & sCondition & " ) " & _
             " GROUP BY ID_MONEDA, " & sGLField & " COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, ID_SECTOR " & sAux
      
      Set oCommand.ActiveConnection = oConn
      oCommand.CommandText = sSQL
      oCommand.CommandType = adCmdText
      oCommand.Execute
    End If
  Exit Sub
ErrHandler:
   RaiseError ClassId, "CreateMovTBL", Err.Number
End Sub

Private Sub CreateTotTBL(oConn As ADODB.Connection, sPatron As String, sGLCondition As String, _
                            sTableName As String, sFromTable As String, sMovName As String, _
                            Optional bWithFec As Boolean = False, _
                            Optional bWithGL As Boolean = False)
  Dim oCommand             As New ADODB.Command
  Dim sSQL                 As String
  Dim nPos                 As Variant
  Dim i                    As Long
  Dim sFecCondition        As String
  Dim sGLField             As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    
    If (bWithFec) Then
      sFecCondition = " , FECHA_ULT"
    ElseIf Not (bWithFec) Then
      sFecCondition = " "
    End If
    If (bWithGL) Then
      sGLField = " , ID_MAYOR"
    ElseIf Not (bWithGL) Then
      sGLField = " "
    End If
' Apartir del 26/Oct/2000 ->
    sSQL = "    SELECT id_moneda, SUBSTR( " & sFromTable & ".CUENTA,1, 4 ) AS CUENTA, (0) AS SECTOR, Sum( " & sFromTable & "." & sMovName & ") As " & sMovName & ", " & _
           "           Sum( " & sFromTable & ".AMOUNT) As AMOUNT " & sFecCondition & sGLField & ", (0) FROM_MOVS " & _
           "     From  " & sFromTable & _
           "    Where (Length(CUENTA) >= 4) and " & _
           "          not ((SUBSTR( " & sFromTable & ".CUENTA,1, 4) = CUENTA) and (nvl(SECTOR,0)=0)) " & _
           "    GROUP BY id_moneda " & sGLField & " , SUBSTR(CUENTA,1, 4) " & sFecCondition & _
           " Union "
    
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & " ) AS CUENTA, SECTOR, Sum( " & sFromTable & "." & sMovName & ") As " & sMovName & ", " & _
                    "           Sum( " & sFromTable & ".AMOUNT) As AMOUNT " & sFecCondition & sGLField & ", (0) FROM_MOVS " & _
                    "     From  " & sFromTable & _
                    "    Where (Length(CUENTA) >=  " & nPos(i) & " ) and " & _
                    "          (SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & ") <> CUENTA) " & _
                    "    GROUP BY id_moneda " & sGLField & " , SUBSTR(CUENTA,1,  " & nPos(i) & "  ), SECTOR " & sFecCondition & _
                    " Union "
' <- Apartir del 26/Oct/2000
' Hasta el 26/Oct/2000
'      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & " ) AS CUENTA, (0) AS SECTOR, Sum( " & sFromTable & "." & sMovName & ") As " & sMovName & ", " & _
'                    "           Sum( " & sFromTable & ".AMOUNT) As AMOUNT " & sFecCondition & _
'                    "     From  " & sFromTable & _
'                    "    Where (Length(CUENTA) >=  " & nPos(i) & " ) " & _
'                    "    GROUP BY id_moneda, SUBSTR(CUENTA,1,  " & nPos(i) & "  ) " & sFecCondition & _
'                    " Union "
    Next i
    ' "INSERT INTO " & sTableName & " (ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & " ) "
    '" CREATE TABLE " & sTableName & " AS "
    sSQL = "INSERT INTO " & sTableName & " (ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & " ,FROM_MOVS) " & _
           " SELECT * FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & ",FROM_MOVS " & _
           "         From  " & sFromTable & _
           " )"
    
    Set oCommand.ActiveConnection = oConn
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CreateTotTbl", Err.Number
End Sub

Private Sub CreateTotTBL58(oConn As ADODB.Connection, sPatron As String, sGLCondition As String, _
                            sTableName As String, sFromTable As String, sMovName As String, _
                            Optional bWithFec As Boolean = False, _
                            Optional bWithGL As Boolean = False)
  Dim oCommand             As New ADODB.Command
  Dim sSQL                 As String
  Dim nPos                 As Variant
  Dim i                    As Long
  Dim sFecCondition        As String
  Dim sGLField             As String
  '****************************************************************************************
  On Error GoTo ErrHandler

    If (bWithFec) Then
      sFecCondition = " , FECHA_ULT"
    ElseIf Not (bWithFec) Then
      sFecCondition = " "
    End If
    If (bWithGL) Then
      sGLField = " , ID_MAYOR"
    ElseIf Not (bWithGL) Then
      sGLField = " "
    End If

    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      '
      '
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & " ) AS CUENTA,(0) SECTOR, Sum( " & sFromTable & "." & sMovName & ") As " & sMovName & ", " & _
                    "           Sum( " & sFromTable & ".AMOUNT) As AMOUNT " & sFecCondition & sGLField & ", (0) FROM_MOVS " & _
                    "     From  " & sFromTable & _
                    "    Where (Length(CUENTA) >=  " & nPos(i) & " ) and " & _
                    "          not ((SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & ") = CUENTA) and (NVL(SECTOR,0)=0)) " & _
                    "    GROUP BY id_moneda " & sGLField & " , SUBSTR(CUENTA,1,  " & nPos(i) & "  ) " & sFecCondition & _
                    " Union "
    Next i

    sSQL = "INSERT INTO " & sTableName & " (ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & " ,FROM_MOVS) " & _
           " SELECT * FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & ",FROM_MOVS " & _
           "         From  " & sFromTable & _
           " )"

    Set oCommand.ActiveConnection = oConn
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CreateTotTbl58", Err.Number
End Sub

Private Sub CreateTotTblAnalitico(oConn As ADODB.Connection, sPatron As String, sGLCondition As String, _
                            sTableName As String, sFromTable As String, sMovName As String, _
                            Optional bWithFec As Boolean = False, _
                            Optional bWithGL As Boolean = False)
  Dim oCommand             As New ADODB.Command
  Dim sSQL                 As String
  Dim nPos                 As Variant
  Dim i                    As Long
  Dim sFecCondition        As String
  Dim sGLField             As String
  '****************************************************************************************
  On Error GoTo ErrHandler

    If (bWithFec) Then
      sFecCondition = " , FECHA_ULT"
    ElseIf Not (bWithFec) Then
      sFecCondition = " "
    End If
    If (bWithGL) Then
      sGLField = " , ID_MAYOR"
    ElseIf Not (bWithGL) Then
      sGLField = " "
    End If

    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
    '                     "          not ((SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & ") = CUENTA) and (NVL(SECTOR,0)=0)) "
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR( " & sFromTable & ".CUENTA,1, " & nPos(i) & " ) AS CUENTA,(0) SECTOR, Sum( " & sFromTable & "." & sMovName & ") As " & sMovName & ", " & _
                    "           Sum( " & sFromTable & ".AMOUNT) As AMOUNT " & sFecCondition & sGLField & ", (0) FROM_MOVS " & _
                    "     From  " & sFromTable & _
                    "    Where (Length(CUENTA) >=  " & nPos(i) & " ) " & _
                    "    GROUP BY id_moneda " & sGLField & " , SUBSTR(CUENTA,1,  " & nPos(i) & "  ) " & sFecCondition
      If (UBound(nPos) <> i) Then
         sSQL = sSQL & " Union "
      End If
    Next i

    sSQL = "INSERT INTO " & sTableName & " (ID_MONEDA,CUENTA,SECTOR,AMOUNT, " & sMovName & sFecCondition & sGLField & " ,FROM_MOVS) " & _
           " SELECT * FROM " & _
           " ( " & sSQL & " ) "

    Set oCommand.ActiveConnection = oConn
    oCommand.CommandText = sSQL
    oCommand.CommandType = adCmdText
    oCommand.Execute
  Exit Sub
ErrHandler:
  RaiseError ClassId, "CreateTotTblAnalitico", Err.Number
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

Private Function ZeroFill(ByVal sStr As String, nLen As Long) As String
  On Error GoTo ErrorHandler
    ZeroFill = Mid("00000000000000000000000", 1, nLen - Len(sStr)) & sStr
  Exit Function
ErrorHandler:
  RaiseError ClassId, "ZeroFill", Err.Number
End Function

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

Private Sub Truncate(sAppServer As String, sTableName As String)
  Dim oCommand As New Command
  '*************************************************************
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
  RaiseError ClassId, "TruncateTable", Err.Number
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
