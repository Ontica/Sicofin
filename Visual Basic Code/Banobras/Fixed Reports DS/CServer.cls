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
'*** Sistema de contabilidad financiera (SICOFIN) ***********************************************************
'*                                                                                                          *
'* Solución   : Customer Components                             Sistema : Financial Accounting              *
'* Componente : Fixed Reports DS (SCFFixedReportsDS)            Parte   : Data services                     *
'* Clase      : CServer                                         Patrón  : Stateless Class                   *
'* Fecha      : 31/Enero/2002                                   Versión : 1.1       Versión patrón: 1.0     *
'*                                                                                                          *
'* Descripción: Genera las fuentes de datos para los reportes contables fijos excepto los del cuaderno.     *
'*                                                                                                          *
'****************************************************** Copyright © La Vía Ontica, S.C. México, 2001-2002. **
Option Explicit
Private Const ClassId As String = "CServer"

Private sStdAccHist As String
Private Const StdAccHist As String = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR " & _
                                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '<@dIniDate@>') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '<@dEndDate@>')) AND " & _
                                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '<@dIniDate@>') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '<@dEndDate@>'))) " & _
                                     " GROUP BY ID_CUENTA_ESTANDAR "


'************************************************************************************************************
'* OTROS MÉTODOS PÚBLICOS  (Queries para generar reportes no encontrados en AOGLBalances viejo)             *
'************************************************************************************************************

Public Function AccountType(sAppServer As String, nGl As Long) As Long
  Dim oConn             As New ADODB.Connection
  Dim oRs               As New ADODB.Recordset
  Dim sSQLAccType       As String
  '*******************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    sSQLAccType = " SELECT MIN(ID_GRUPO_MAYOR) id_grupo_mayor FROM COF_ELEMENTO_GRUPO_MAYOR WHERE (ID_MAYOR = " & nGl & ")"
    
    With oRs
      '.CursorLocation = adUseClient
      .Open sSQLAccType, oConn, adOpenStatic, adLockReadOnly
      AccountType = !id_grupo_mayor
      .Close
      Set .ActiveConnection = Nothing
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "AccountType", Err.Number
End Function

Public Function VouchersRS(sAppServer As String, vGL As Variant, dIniDate As Date, dCuteDate As Date, _
                           bByAfectationDate As Boolean, bUpdatedVoucher As Boolean) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  Dim sGl               As String
  Dim sSQL              As String
  Dim sWhere            As String
  Dim nPos              As Variant
  Dim i                 As Long
  '****************************************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGl = "id_mayor IN (" & sGl & ") "
    Else
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    If bByAfectationDate Then
      sWhere = " (TRUNC(TO_DATE(COF_TRANSACCION.FECHA_AFECTACION)) >= '" & dIniDate & "')  AND " & _
               " (TRUNC(TO_DATE(COF_TRANSACCION.FECHA_AFECTACION)) <= '" & dCuteDate & "')  AND "
    Else
      sWhere = " (TRUNC(TO_DATE(COF_TRANSACCION.FECHA_REGISTRO)) >= '" & dIniDate & "')  AND " & _
               " (TRUNC(TO_DATE(COF_TRANSACCION.FECHA_REGISTRO)) <= '" & dCuteDate & "')  AND "
               
    End If
    If bUpdatedVoucher Then
      sSQL = " SELECT COF_MAYOR.ID_MAYOR, COF_MAYOR.NUMERO_MAYOR, COF_TRANSACCION.NUMERO_TRANSACCION, " & _
             "        COF_TRANSACCION.FECHA_AFECTACION, COF_TRANSACCION.FECHA_REGISTRO, " & _
             "        PARTICIPANTNAME, COF_TRANSACCION.CONCEPTO_TRANSACCION, " & _
             "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
             "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
             " FROM COF_MOVIMIENTO, COF_TRANSACCION, MHPARTICIPANTS, COF_MAYOR  " & _
             " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
             "       (COF_TRANSACCION.ID_ELABORADA_POR = MHPARTICIPANTS.PARTICIPANTID(+)) AND " & _
             "       (COF_TRANSACCION.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND " & _
             sWhere & _
             "       (COF_TRANSACCION." & sGl & " ) AND (ESTA_ABIERTA = 0) " & _
             " GROUP BY COF_MAYOR.ID_MAYOR, COF_MAYOR.NUMERO_MAYOR, COF_TRANSACCION.ID_TRANSACCION, " & _
             "          COF_TRANSACCION.NUMERO_TRANSACCION, COF_TRANSACCION.FECHA_AFECTACION, " & _
             "          COF_TRANSACCION.FECHA_REGISTRO, PARTICIPANTNAME, " & _
             "          COF_TRANSACCION.CONCEPTO_TRANSACCION " & _
             " ORDER BY COF_MAYOR.NUMERO_MAYOR, NUMERO_TRANSACCION "
    Else
      sSQL = " SELECT COF_MAYOR.ID_MAYOR, COF_MAYOR.NUMERO_MAYOR, COF_TRANSACCION.NUMERO_TRANSACCION, " & _
             "        COF_TRANSACCION.FECHA_AFECTACION, COF_TRANSACCION.FECHA_REGISTRO, " & _
             "        PARTICIPANTNAME, COF_TRANSACCION.CONCEPTO_TRANSACCION, " & _
             "        (SUM(DECODE(TIPO_MOVIMIENTO,'D',Monto,0))) AS CARGOS, " & _
             "        (SUM(DECODE(TIPO_MOVIMIENTO,'H',Monto,0))) AS ABONOS " & _
             " FROM COF_MOVIMIENTO_TMP, COF_TRANSACCION, MHPARTICIPANTS, COF_MAYOR " & _
             " WHERE (COF_MOVIMIENTO_TMP.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
             "       (COF_TRANSACCION.ID_ELABORADA_POR = MHPARTICIPANTS.PARTICIPANTID(+)) AND " & _
             "       (COF_TRANSACCION.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND " & _
             sWhere & _
             "       (COF_TRANSACCION." & sGl & " ) AND (ESTA_ABIERTA = 1) " & _
             " GROUP BY COF_MAYOR.ID_MAYOR, COF_MAYOR.NUMERO_MAYOR, COF_TRANSACCION.ID_TRANSACCION, " & _
             "          COF_TRANSACCION.NUMERO_TRANSACCION, COF_TRANSACCION.FECHA_AFECTACION, " & _
             "          COF_TRANSACCION.FECHA_REGISTRO, PARTICIPANTNAME, " & _
             "          COF_TRANSACCION.CONCEPTO_TRANSACCION " & _
             " ORDER BY COF_MAYOR.NUMERO_MAYOR, NUMERO_TRANSACCION "
    End If
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set VouchersRS = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "VouchersRS", Err.Number
End Function


Public Function ParamWithDestiny(sAppServer As String, dDate As Date, _
                                    nRegla As Long, nLevel As Long, _
                                    Optional nEntityGroup As Long = 0, _
                                    Optional nSourceGroup As Long) As Recordset
  Dim oRecordset As New ADODB.Recordset, sSQL As String, sSQL2 As String
  Dim sSQLStdAccHist As String
  Dim sWhere As String
  '*****************************************************************************
  On Error GoTo ErrHandler
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dDate)
    
    
    If (nEntityGroup <> 0 And nSourceGroup <> 0) Then
      sWhere = " AND (COF_GRUPO_CUENTA.ID_ENTIDAD_AGRUPADOR = " & nEntityGroup & ") AND " & _
               " (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = " & nSourceGroup & ") "
    Else
      sWhere = ""
    End If
    
    sSQL = " SELECT COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN, " & _
           "        DECODE(COF_GRUPO_CUENTA.ID_ENTIDAD_AGRUPADOR,1, NVL(OBJECT_NAME, ' '), " & _
           "               9, NVL(NOMBRE_MAYOR||' ['||NUMERO_MAYOR||']',' '), 16,NVL(NOMBRE_GRUPO_MAYOR,' ')) NOMBRE_GRUPO, " & _
           "        COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR NUMERO_CUENTA_ORIGEN, " & _
           "        NVL(COF_SECTOR.CLAVE_SECTOR,'SIN SECTOR') CLAVE_SECTOR_ORIGEN, " & _
           "        NVL(COF_GRUPO_CUENTA.ID_MONEDA_ORIGEN,0) ID_MONEDA_ORIGEN, NVL(AO_CURRENCIES.ABBREV,'TODAS LAS MONEDAS') ABBREV_MONEDA_ORIGEN, " & _
           "        NVL(COF_CUENTA_ESTANDAR_2.NUMERO_CUENTA_ESTANDAR,' ') NUMERO_CUENTA_DESTINO, " & _
           "         NVL(COF_SECTOR_2.CLAVE_SECTOR,'SIN SECTOR') CLAVE_SECTOR_DESTINO," & _
           "         NVL(COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR_DESTINO, " & _
           "        NVL(COF_CUENTA_ESTANDAR_SOB.NUMERO_CUENTA_ESTANDAR,' ') NUMERO_CUENTA_SOB, " & _
           "         NVL(COF_SECTOR_SOB.CLAVE_SECTOR,'SIN SECTOR') CLAVE_SECTOR_SOB, " & _
           "         NVL(COF_CUENTA_AUXILIAR_SOB.NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR_SOB, " & _
           "         NVL(VERSION_REGLA,0) VERSION_REGLA " & _
           " FROM COF_REGLA_CONTABLE, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR_2, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR_SOB, " & _
           "      COF_SECTOR, COF_SECTOR COF_SECTOR_2, COF_SECTOR COF_SECTOR_SOB, " & _
           "      COF_CUENTA_AUXILIAR, COF_CUENTA_AUXILIAR COF_CUENTA_AUXILIAR_SOB, " & _
           "      COF_GRUPO_MAYOR, COF_MAYOR, COF_GRUPO_CUENTA, COF_GRUPO_CUENTA_BIS,  " & _
           "      AO_CURRENCIES, AO_OBJECTS "
           
  sSQL2 = " WHERE (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR >= COF_GRUPO_CUENTA.CUENTA_ORIGEN_INICIAL) and " & _
           "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR <= COF_GRUPO_CUENTA.CUENTA_ORIGEN_FINAL) and " & _
           "       (COF_REGLA_CONTABLE.ID_REGLA_CONTABLE = " & nRegla & " ) AND " & _
           "       (COF_GRUPO_CUENTA.ID_REGLA_CONTABLE = " & nRegla & " ) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.NIVEL <= " & nLevel & " ) AND " & _
           "       (COF_CUENTA_ESTANDAR.ID_TIPO_CUENTAS_STD = COF_REGLA_CONTABLE.ID_TIPO_CUENTAS_STD) AND " & _
           "       (COF_GRUPO_CUENTA.ID_GRUPO_CUENTA = COF_GRUPO_CUENTA_BIS.ID_GRUPO_CUENTA) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = AO_OBJECTS.OBJECT_ID(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = COF_MAYOR.ID_MAYOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = COF_GRUPO_MAYOR.ID_GRUPO_MAYOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_SECTOR_ORIGEN = COF_SECTOR.ID_SECTOR(+)) AND " & _
           "       (ABS(COF_GRUPO_CUENTA.ID_MONEDA_ORIGEN) = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_CUENTA_DESTINO = COF_CUENTA_ESTANDAR_2.ID_CUENTA_ESTANDAR(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_SECTOR_DESTINO = COF_SECTOR_2.ID_SECTOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_AUXILIAR_DESTINO = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_CUENTA_DESTINO_SOBREGIRO = COF_CUENTA_ESTANDAR_SOB.ID_CUENTA_ESTANDAR(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_SECTOR_DESTINO_SOBREGIRO = COF_SECTOR_SOB.ID_SECTOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.ID_AUXILIAR_DESTINO_SOBREGIRO = COF_CUENTA_AUXILIAR_SOB.ID_CUENTA_AUXILIAR(+)) AND " & _
           "       ((TRUNC(TO_DATE(COF_GRUPO_CUENTA_BIS.FECHA_INICIO)) <= '" & dDate & "') AND (TRUNC(TO_DATE(COF_GRUPO_CUENTA_BIS.FECHA_FIN)) >= '" & dDate & "')) " & _
                   sWhere & _
           " ORDER BY  COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_eSTANDAR, COF_CUENTA_ESTANDAR_2.NUMERO_CUENTA_ESTANDAR "
    sSQL = sSQL & sSQL2
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set ParamWithDestiny = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParamWithDestiny", Err.Number
End Function


Public Function ParamWithoutDestiny(sAppServer As String, dDate As Date, _
                                    nRegla As Long, nLevel As Long) As Recordset
  Dim oRecordset As New ADODB.Recordset, sSQL As String, sSQL2 As String
  Dim sSQLStdAccHist As String
  '*******************************************************************
  On Error GoTo ErrHandler
  
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dDate)
    
    sSQL = " SELECT COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN, " & _
           "        DECODE(COF_GRUPO_CUENTA.ID_ENTIDAD_AGRUPADOR,1, NVL(OBJECT_NAME, ' '), " & _
           "               9, NVL(NOMBRE_MAYOR,' '), 16,NVL(NOMBRE_GRUPO_MAYOR,' ')) NOMBRE_GRUPO, " & _
           "        COF_GRUPO_CUENTA.NUMERO_GRUPO_CUENTA, COF_GRUPO_CUENTA.NOMBRE_GRUPO_CUENTA, " & _
           "        COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR NUMERO_CUENTA_ORIGEN, " & _
           "        NVL(COF_SECTOR.CLAVE_SECTOR,'SIN SECTOR') CLAVE_SECTOR_ORIGEN, " & _
           "        NVL(COF_GRUPO_CUENTA.ID_MONEDA_ORIGEN,0) ID_MONEDA_ORIGEN, " & _
           "        NVL(AO_CURRENCIES.ABBREV,'SIN MONEDA') ABBREV_MONEDA_ORIGEN, " & _
           "        NVL(VERSION_REGLA,0) VERSION_REGLA " & _
           " FROM COF_REGLA_CONTABLE, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, " & _
           "      COF_SECTOR, COF_GRUPO_MAYOR, COF_MAYOR, COF_GRUPO_CUENTA, COF_GRUPO_CUENTA_BIS,  " & _
           "      AO_CURRENCIES, AO_OBJECTS "
           
  sSQL2 = " WHERE (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR >= COF_GRUPO_CUENTA.CUENTA_ORIGEN_INICIAL) and " & _
           "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR <= COF_GRUPO_CUENTA.CUENTA_ORIGEN_FINAL) and " & _
           "       (COF_REGLA_CONTABLE.ID_REGLA_CONTABLE = " & nRegla & " ) AND " & _
           "       (COF_GRUPO_CUENTA.ID_REGLA_CONTABLE = " & nRegla & " ) AND " & _
           "       (COF_GRUPO_CUENTA_BIS.NIVEL <= " & nLevel & " ) AND " & _
           "       (COF_CUENTA_ESTANDAR.ID_TIPO_CUENTAS_STD = COF_REGLA_CONTABLE.ID_TIPO_CUENTAS_STD) AND " & _
           "       (COF_GRUPO_CUENTA.ID_GRUPO_CUENTA = COF_GRUPO_CUENTA_BIS.ID_GRUPO_CUENTA) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = AO_OBJECTS.OBJECT_ID(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = COF_MAYOR.ID_MAYOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_AGRUPADOR_ORIGEN = COF_GRUPO_MAYOR.ID_GRUPO_MAYOR(+)) AND " & _
           "       (COF_GRUPO_CUENTA.ID_SECTOR_ORIGEN = COF_SECTOR.ID_SECTOR(+)) AND " & _
           "       (ABS(COF_GRUPO_CUENTA.ID_MONEDA_ORIGEN) = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
           "       ((TRUNC(TO_DATE(COF_GRUPO_CUENTA_BIS.FECHA_INICIO)) <= '" & dDate & "') AND (TRUNC(TO_DATE(COF_GRUPO_CUENTA_BIS.FECHA_FIN)) >= '" & dDate & "')) " & _
           " ORDER BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_eSTANDAR "
    sSQL = sSQL & sSQL2
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set ParamWithoutDestiny = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "ParamWithoutDestiny", Err.Number
End Function

Public Function SQLStandardAccountHist(dIniDate As Date, dCuteDate As Date, _
                                       Optional OptionToInclude As Long = 0) As String
  
  sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", dIniDate)
  sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)
  
  If OptionToInclude = 0 Then
    SQLStandardAccountHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                             "        NUMERO_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTAS_STD  " & _
                             " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                             " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                             "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "

  ElseIf OptionToInclude = 1 Then
    SQLStandardAccountHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, " & _
                             "        NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, NATURALEZA, " & _
                             "        ID_TIPO_CUENTAS_STD, ID_TIPO_CUENTA, ROL_CUENTA  " & _
                             " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                             " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                             "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) "
  
  End If
                     

    Exit Function
ErrHandler:
  RaiseError ClassId, "SQLStandardAccountHist", Err.Number
End Function

Public Function GeneraRecordset18(sAppServer As String, nGl As Variant, vGL As Variant, _
                                   dIniDate As Date, dCuteDate As Date, dExcDate As Date, _
                                   Optional nExcRateType As Long = 46, _
                                   Optional nExcRateCurrency As Long = 1) As Recordset
  Dim oConn           As New Connection
  Dim oRecordset      As New Recordset

  Dim sSQL            As String
  Dim sSQL_A          As String
  
  Dim sSQL1           As String
  Dim sSQL2           As String
  
  Dim sGl             As String
  Dim sGLCondition    As String
  
  Dim sSQLStdAccHist   As String
  Dim sSQLStdAccHist2  As String
  '****************************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
    Else
      sGl = vGL
    End If

    sGLCondition = "id_mayor in (" & sGl & "," & nGl & ") "

    sSQLStdAccHist = SQLStandardAccountHist(dIniDate, dCuteDate)
    sSQLStdAccHist2 = SQLStandardAccountHist(dIniDate, dCuteDate, 1)
                     
    sSQL1 = " SELECT id_moneda, substr(NUMERO_CUENTA_ESTANDAR,1,4) as Cuenta, COF_SALDO_INICIAL.ID_MAYOR, " & _
            "        decode(COF_SALDO_INICIAL.ID_MAYOR," & nGl & ",sum(COF_SALDO_INICIAL.SALDO_INICIAL),0) SI, " & _
            "        (0) as Cargos, (0) as Abonos, " & _
            "        decode(COF_SALDO_INICIAL.ID_MAYOR," & nGl & ",0,sum(COF_SALDO_INICIAL.SALDO_INICIAL)) SI_ALL, " & _
            "         (0) as Cargos_ALL, (0) as Abonos_ALL " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_SALDO_INICIAL." & sGLCondition & ") " & _
            " GROUP BY ID_MONEDA, COF_SALDO_INICIAL.ID_MAYOR,substr(NUMERO_CUENTA_ESTANDAR,1,4) "
    
  

    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, substr(NUMERO_CUENTA_ESTANDAR,1,4) as Cuenta , COF_TRANSACCION.ID_MAYOR, " & _
            "        (0) as SI, " & _
            "        decode(COF_TRANSACCION.ID_MAYOR," & nGl & ",(sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))),0) as Cargos, " & _
            "        decode(COF_TRANSACCION.ID_MAYOR," & nGl & ",(sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))),0) as Abonos, " & _
            "        (0) as SI_ALL, " & _
            "        decode(COF_TRANSACCION.ID_MAYOR," & nGl & ",0,(sum(decode(TIPO_MOVIMIENTO,'D',Monto,0)))) as Cargos_ALL , " & _
            "        decode(COF_TRANSACCION.ID_MAYOR," & nGl & ",0,(sum(decode(TIPO_MOVIMIENTO,'H',Monto,0)))) as Abonos_ALL  " & _
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
            "       (COF_TRANSACCION." & sGLCondition & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR,substr(NUMERO_CUENTA_ESTANDAR,1,4) "


sSQL = " SELECT ID_MONEDA, (NATURALEZA) NAT, A.Cuenta, NOMBRE_CUENTA_ESTANDAR NOMBRE, " & _
           "        (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
           "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) )) SA, " & _
           "        (decode(NATURALEZA,'D',(sum(A.SI_ALL)+sum(A.CARGOS_ALL)-sum(A.ABONOS_ALL)), " & _
           "                               (sum(A.SI_ALL)-sum(A.CARGOS_ALL)+sum(A.ABONOS_ALL)))) SA_OTROS " & _
           " FROM ( " & _
           sSQL1 & _
           " UNION " & _
           sSQL2 & _
           " ) A  , ( " & sSQLStdAccHist2 & " ) B " & _
           " WHERE (A.CUENTA = B.NUMERO_CUENTA_ESTANDAR(+)) AND " & _
           "       (LENGTH(B.NUMERO_CUENTA_ESTANDAR) = 4 ) AND " & _
           "       (ID_TIPO_CUENTAS_STD = 1) " & _
           " GROUP BY A.ID_MONEDA, B.NATURALEZA, A.CUENTA, NOMBRE_cuenta_estandar " & _
           " HAVING (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
           "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS))) <> 0 ) OR " & _
           "        (decode(NATURALEZA,'D',(sum(A.SI_ALL)+sum(A.CARGOS_ALL)-sum(A.ABONOS_ALL)), " & _
           "                               (sum(A.SI_ALL)-sum(A.CARGOS_ALL)+sum(A.ABONOS_ALL))) <> 0 ) "
  
    
    sSQL_A = " SELECT (1) id_moneda, NAT, A.Cuenta, NOMBRE, SUM(NVL(EXCHANGE_RATE,1)*A.SA) SA, " & _
             "                                              SUM(NVL(EXCHANGE_RATE,1)*A.SA_OTROS) SA_OTROS  " & _
             " FROM ( " & sSQL & " ) A , AO_EXCHANGE_RATES B " & _
             " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
             "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
             "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
             "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
             "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
             " GROUP BY NAT, CUENTA, NOMBRE "
             
    sSQL = " SELECT NAT, A.Cuenta, NOMBRE, SUM(A.SA/NVL(EXCHANGE_RATE,1)) SA, " & _
           "                               SUM(A.SA_OTROS/NVL(EXCHANGE_RATE,1)) SA_OTROS  " & _
           " FROM ( " & sSQL_A & " ) A , AO_EXCHANGE_RATES B " & _
           " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
           "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
           "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
           "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
           "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
           " GROUP BY NAT, CUENTA, NOMBRE " & _
           " ORDER BY NAT DESC, A.CUENTA "
           
    If (cnFromCurrency = nExcRateCurrency) Then
      sSQL = sSQL_A & _
             " ORDER BY NAT DESC, A.CUENTA "
    End If
    
  With oRecordset
    .CursorLocation = adUseClient
    .Open sSQL, oConn, adOpenForwardOnly, adLockReadOnly
    Set .ActiveConnection = Nothing
  End With
   Set GeneraRecordset18 = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset18", Err.Number
End Function

Public Function GeneraRecordset19(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As Recordset
  Dim oConn             As New Connection
  Dim oComm             As New Command
  Dim oRs               As New Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  'Dim sSelect           As String
  'Dim sInitSelect       As String
  'Dim sMaxSelect        As String
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
                       
      sSQL = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
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
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A " & _
              sOptionToDisplay
            
    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
            

      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, NUMERO_MAYOR, NOMBRE_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR, COF_MAYOR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
               "       (A.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) AND " & _
               "       (length(A.NUMERO_CUENTA_ESTANDAR) <= " & Len(sPatron) & " )" & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, NUMERO_MAYOR "
      Else
        
       sSQL_A = " SELECT A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, " & _
                "        sum(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, sum(ABONOS) ABONOS, " & _
                "        sum(SALDO_ACTUAL) SALDO_ACTUAL " & _
                " FROM ( " & sSQL & " ) A " & _
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR "
        
        
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, NUMERO_MAYOR, NOMBRE_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL_A & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_MAYOR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_MAYOR = COF_MAYOR.ID_MAYOR(+))  AND " & _
               "       (length(A.NUMERO_CUENTA_ESTANDAR) <= " & Len(sPatron) & " )" & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_MAYOR "
               'sOptionToDisplay &
               ' AND (A.ID_CUENTA_AUXILIAR = 0)
      End If
    
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    'Debug.Print sSQL
    Set GeneraRecordset19 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset19", Err.Number
End Function

Public Function GeneraRecordset51_a_54(sAppServer As String, vGL As Variant, _
                                       dEndDate As Date, dExcDate As Date, _
                                       Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1) As Recordset
  Dim oConn           As New ADODB.Connection
  Dim oRs             As New ADODB.Recordset
  
  Dim sGl             As String
  
  Dim sSQL            As String
  Dim sSql_SA         As String
  Dim sSql_SAA        As String
  Dim sSQL1           As String
  Dim sSQL2           As String
  
  Dim sSQLStdAccHist  As String
    '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_CUENTA, sum(SALDO_INICIAL) SI, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dEndDate & "') " & _
            " group by ID_MONEDA, ID_CUENTA "

    
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        (0) SI, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dEndDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA "
    
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dEndDate, 1)
    
   sSql_SAA = " SELECT (1) ID_MONEDA_VAL, A.ID_MONEDA, NATURALEZA, (NUMERO_CUENTA_ESTANDAR) CUENTA, (ID_TIPO_CUENTA) TIPO_CUENTA, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) ) * NVL(EXCHANGE_RATE,1) ) SA " & _
              " FROM ( " & _
              sSQL1 & _
              " UNION " & _
              sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C, AO_EXCHANGE_RATES D " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) and " & _
              "       (A.ID_MONEDA = D.TO_CURRENCY_ID(+)) and " & _
              "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND " & _
              "       (ID_TIPO_CUENTA in (3,4,5,6,7,8,10)) and (SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2)<>'50') " & _
              " GROUP BY ID_MONEDA, NATURALEZA, NUMERO_CUENTA_ESTANDAR, ID_TIPO_CUENTA, EXCHANGE_RATE "
    
    sSql_SA = " SELECT A.ID_MONEDA, NATURALEZA, CUENTA, TIPO_CUENTA, " & _
              "         SUM(A.SA/NVL(EXCHANGE_RATE,1)) SA " & _
              " FROM ( " & sSql_SAA & " ) A  , AO_EXCHANGE_RATES D " & _
              " WHERE (A.ID_MONEDA_VAL = D.FROM_CURRENCY_ID(+)) and " & _
              "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
              " GROUP BY ID_MONEDA, NATURALEZA, CUENTA, TIPO_CUENTA "
    
    If (cnFromCurrency = nExcRateCurrency) Then
      sSql_SA = sSql_SAA
    End If
          
     sSQL = " SELECT NATURALEZA, SUBSTR( A.CUENTA,1, 2 ) AS CUENTA, MAX(TIPO_CUENTA) TIPO_CUENTA, " & _
            "        (sum(decode(ID_MONEDA,1,SA,28,SA,0))) SALDO_MN, " & _
            "        (sum(decode(ID_MONEDA,1,0,28,0,SA)))     SALDO_OTRA, " & _
            "        SUM(A.SA) SALDO_TOTAL " & _
            " From  ( " & sSql_SA & " )  A " & _
            " GROUP BY NATURALEZA, SUBSTR(CUENTA,1, 2)  " & _
            " ORDER BY NATURALEZA DESC, SUBSTR(CUENTA,1, 2) "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset51_a_54 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset51_a_54", Err.Number
End Function

Public Function GeneraRecordset55_56(sPatron As String, sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, dExcDate As Date, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sSQL_MN           As String
  Dim sSQL_MNA          As String
  Dim sSQL_ME           As String
  Dim sSQL_MEA          As String
  
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
            "       (COF_TRANSACCION." & sGl & " ) " & _
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
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A "
            
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = " tmp_SALDOS "
    
    sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
           "           (0) ID_CUENTA_AUXILIAR, " & _
           "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
           "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
           "    From  " & sFromTable & "  A " & _
           "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
           "          not ((SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4) = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_SECTOR,0)=0))" & _
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
            
    sSQL_MNA = " SELECT  A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_MN, (0) SALDO_ACTUAL_ME, " & _
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
                 " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA in (1,28)) and " & _
               "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "

     sSQL_MN = " SELECT  A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
               "         SUM(A.SALDO_ACTUAL_MN/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_MN, (0) SALDO_ACTUAL_ME, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_MNA & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (" & cnFromCurrency & "  = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "

        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_MN = sSQL_MNA
        End If
        
       sSQL_MEA = " SELECT  A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                  "         (0) SALDO_ACTUAL_MN, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_ME, " & _
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
                     " ) A , AO_EXCHANGE_RATES B " & _
                   " WHERE (not(A.ID_MONEDA in (1,28))) and " & _
                   "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                   "          A.ROL_CUENTA, A.FROM_MOVS "
                  
        sSQL_ME = " SELECT  A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                  "         (0) SALDO_ACTUAL_MN, SUM(A.SALDO_ACTUAL_ME/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_ME, " & _
                  "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                  " FROM (" & sSQL_MEA & " ) A , AO_EXCHANGE_RATES B " & _
                   " WHERE (" & cnFromCurrency & " = B.FROM_CURRENCY_ID(+)) and " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                   "          A.ROL_CUENTA, A.FROM_MOVS "
                  
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL_ME = sSQL_MEA
        End If
        
          sSQL = " SELECT  A.NATURALEZA, " & _
                 "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                 "         SUM(SALDO_ACTUAL_MN) SALDO_ACTUAL_MN, SUM(SALDO_ACTUAL_ME) SALDO_ACTUAL_ME, " & _
                 "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                 " FROM (" & _
                   sSQL_MN & _
                 " UNION " & _
                   sSQL_ME & _
                 " ) A " & _
                 " WHERE ((SALDO_ACTUAL_MN <> 0) OR (SALDO_ACTUAL_ME <> 0)) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '1503%')) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '50%')) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '90%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '91%')) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '92%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '93%')) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '94%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '95%')) AND " & _
                 "       (NOT(NUMERO_CUENTA_ESTANDAR LIKE '96%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '97%')) " & _
                 " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                 "          A.ROL_CUENTA, A.FROM_MOVS " & _
                 " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
    
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRecordset55_56 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset55_56", Err.Number
End Function

Public Function GeneraParametros(sAppServer As String, nReportId As Long) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oRs               As New ADODB.Recordset
  
  Dim sSQL              As String
  
  '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sSQL = " SELECT * " & _
           " FROM COF_PARAMETROS_REPORTES " & _
           " WHERE (REPORT_ID = " & nReportId & ")"
            
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraParametros = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraParametros", Err.Number
End Function

Public Function GeneraRecordset58(sPatron As String, sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, dExcDate As Date, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 2, _
                        Optional sAccounts As String, Optional nReport As Long = 58) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sSQL_MN           As String
  Dim sSQL_MEA          As String
  Dim sSQL_ME           As String
  Dim sWhereAccount     As String
  
  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long
  
  Dim sSQLStdAccHist    As String
  Dim sWhereReport      As String
  
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
    
    If sAccounts <> "" Then
      sWhereAccount = " NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ") AND "
    End If
    
    If (nReport = 58) Then
      sWhereReport = " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '1503%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '50%')) AND " & _
                     " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '90%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '91%')) AND " & _
                     " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '92%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '93%')) AND " & _
                     " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '94%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '95%')) AND " & _
                     " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '96%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '97%')) AND " & _
                     " (NOT(NUMERO_CUENTA_ESTANDAR LIKE '98%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '99%')) AND "
    ElseIf ((nReport = 65) Or (nReport = 102)) Then
      sWhereReport = ""
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
            "       (COF_TRANSACCION." & sGl & " ) " & _
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
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV " & _
              "  ) A "
            
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = " tmp_SALDOS_PROM "
    
    sSQL = ""
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
                    "           (0) ID_CUENTA_AUXILIAR,  " & _
                    "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
                    "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, FECHA_REV  " & _
                    "    From  " & sFromTable & "  A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= " & nPos(i) & ") and " & _
                    "          not ((SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_SECTOR,0)=0) and (nvl(ID_CUENTA_AUXILIAR,0)=0)) " & _
                    "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & "), FECHA_REV " & _
                    " Union "
           
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, FECHA_REV " & _
           " FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, FECHA_REV " & _
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
               
    sSQL_MN = " SELECT A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_MN, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_MN, " & _
              "        SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_MN, " & _
              "        (0) SALDO_INICIAL_ME, (0) CARGOS_ME, (0) ABONOS_ME, " & _
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
                  " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA in (1,28)) and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID,1) = 1) and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "

    sSQL_MEA = " SELECT A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        (0) SALDO_INICIAL_MN, (0) CARGOS_MN, (0) ABONOS_MN, " & _
              "        SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_ME, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_ME, " & _
              "        SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_ME, " & _
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
                  " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (not (A.ID_MONEDA in (1,28))) and " & _
                "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
    
    sSQL_ME = " SELECT A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        (0) SALDO_INICIAL_MN, (0) CARGOS_MN, (0) ABONOS_MN, " & _
              "        SUM(A.SALDO_INICIAL_ME/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_ME, " & _
              "        SUM(A.CARGOS_ME/NVL(EXCHANGE_RATE,1)) CARGOS_ME, " & _
              "        SUM(A.ABONOS_ME/NVL(EXCHANGE_RATE,1)) ABONOS_ME, " & _
              "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
              " FROM (" & sSQL_MEA & ") A , AO_EXCHANGE_RATES B " & _
              " WHERE (" & cnFromCurrency & " = B.FROM_CURRENCY_ID(+)) and " & _
              "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
              " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
              "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV "
    
      If (cnFromCurrency = nExcRateCurrency) Then
        sSQL_ME = sSQL_MEA
      End If
    
      sSQL = " SELECT A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        sum(saldo_inicial_mn) SALDO_INICIAL_MN, sum(CARGOS_MN) CARGOS_MN, SUM(ABONOS_MN) ABONOS_MN, " & _
              "        SUM(SALDO_INICIAL_ME) SALDO_INICIAL_ME, SUM(CARGOS_ME) CARGOS_ME, SUM(ABONOS_ME) ABONOS_ME, " & _
              "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
              " FROM (" & _
                sSQL_MN & _
              " UNION " & _
                sSQL_ME & _
              " ) A " & _
              " WHERE  (length(NUMERO_CUENTA_ESTANDAR)) <= (length('" & sPatron & "')) and " & sWhereReport & _
              "        (not(NUMERO_CUENTA_ESTANDAR like '%00-00-00')) and " & _
              "        (not(NUMERO_CUENTA_ESTANDAR like '%00-00')) and " & _
              "        (CLAVE_SECTOR = '00') AND " & sWhereAccount & _
              "        ((SALDO_INICIAL_MN <> 0) OR (CARGOS_MN <> 0) OR (ABONOS_MN <> 0) OR " & _
              "        (SALDO_INICIAL_ME <> 0) OR (CARGOS_ME <> 0) OR (ABONOS_ME <> 0)) " & _
              " GROUP BY A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, FECHA_REV, A.CLAVE_SECTOR, A.ID_TIPO_CUENTA, A.ROL_CUENTA, " & _
              "          A.NOMBRE_CUENTA, A.FROM_MOVS  " & _
              " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, FECHA_REV "
   
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset58 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset58", Err.Number
End Function

Public Function GeneraRecordset59(sAppServer As String, vGL As Variant, dEndDate As Date, _
                                  dExcDate As Date, _
                                  Optional nExcRateType As Long = 46, Optional nExcRateCurrency As Long = 1)
  Dim oConn           As New ADODB.Connection
  Dim oRecordset      As New ADODB.Recordset
  
  Dim sSQL1           As String
  Dim sSQL2           As String
  Dim sSql_SA         As String
  Dim sSql_SAA        As String
  Dim sSQL            As String
  
  Dim sGl             As String
  Dim sSQLStdAccHist  As String
  '**********************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

    sGl = vGL
    sGl = "id_mayor = " & sGl

    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        sum(SALDO_INICIAL) SI, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dEndDate & "') " & _
            "       and (NVL(ID_CUENTA_AUXILIAR,0) <> 0) " & _
            " group by ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR "
            
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) SI, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dEndDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) and " & _
            "       (NVL(COF_MOVIMIENTO.ID_CUENTA_AUXILIAR,0) <> 0) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dEndDate, 1)
    
    sSql_SAA = " SELECT NUMERO_CUENTA_ESTANDAR CUENTA, " & _
              "        A.ID_MONEDA, CURRENCY_NAME NOM_MONEDA, NVL(CLAVE_SECTOR, '00') SECTOR , " & _
              "        NATURALEZA NAT, NUMERO_CUENTA_AUXILIAR NUM_AUX,   " & _
              "        NOMBRE_CUENTA_AUXILIAR NOM_AUX, NVL(EXCHANGE_RATE,1) EXC_RATE, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) )) ST, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)))*NVL(EXCHANGE_RATE,1)) S_Val " & _
              " FROM ( " & _
                sSQL1 & _
              " UNION " & _
                sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C, AO_CURRENCIES D, AO_EXCHANGE_RATES E, COF_SECTOR F, COF_CUENTA_AUXILIAR G " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) and " & _
              "       (a.ID_MONEDA = D.CURRENCY_ID) and (a.ID_MONEDA = E.TO_CURRENCY_ID(+)) and " & _
              "       (A.ID_SECTOR = F.ID_SECTOR(+)) and (A.ID_CUENTA_AUXILIAR = G.ID_CUENTA_AUXILIAR) and " & _
              "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') AND " & _
              "       (SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2) in ('13','61','63')) " & _
              " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, CURRENCY_NAME ,  " & _
              "        NUMERO_CUENTA_AUXILIAR , NOMBRE_CUENTA_AUXILIAR , NATURALEZA , EXCHANGE_RATE " & _
              " HAVING (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) ) <> 0) "
              
    sSql_SA = " SELECT CUENTA, A.ID_MONEDA, NOM_MONEDA, SECTOR, NAT, NUM_AUX, NOM_AUX, " & _
              "        (NVL(EXC_RATE,1)/NVL(EXCHANGE_RATE,1)) EXC_RATE, " & _
              "        SUM(A.ST) ST, SUM(S_Val/NVL(EXCHANGE_RATE,1)) S_VAL, " & _
              " FROM ( " & sSql_SAA & ") A, AO_EXCHANGE_RATES E " & _
              " WHERE (" & cnFromCurrency & " = E.FROM_CURRENCY_ID(+)) and " & _
              "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
              " GROUP BY A.ID_MONEDA, CUENTA, SECTOR, NOM_MONEDA, NUM_AUX, NOM_AUX, NAT, " & _
              "          EXC_RATE, EXCHANGE_RATE " & _
              " ORDER BY SUBSTR(CUENTA,1,4), SECTOR, SUBSTR(CUENTA,6), NUM_AUX "
              
      If (cnFromCurrency = nExcRateCurrency) Then
        sSql_SA = sSql_SAA & _
                  " ORDER BY SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4), CLAVE_SECTOR, SUBSTR(NUMERO_CUENTA_ESTANDAR,6), NUMERO_CUENTA_AUXILIAR  "
      End If


    With oRecordset
      .CursorLocation = adUseClient
      .Open sSql_SA, oConn, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset59 = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset59", Err.Number
End Function

Public Function GeneraRecordset60(sAppServer As String, dCuteDate As Date, _
                                  sAccounts As String) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oCommand             As ADODB.Command
  Dim oRs                  As New ADODB.Recordset

  Dim sGl                  As String
  Dim sSQL                 As String
  Dim sSQL1                As String
  Dim sSQL2                As String
  Dim sSql3                As String
  Dim sAccountCondition    As String
  Dim sSQLStdAccHist       As String
  Dim sSQLGL               As String
  '**********************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dCuteDate)

    sSQL1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
            "       COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, COF_CUENTA_ESTANDAR.NATURALEZA, " & _
            "       COF_SALDO_INICIAL.SALDO_INICIAL AS SI, (0) as Cargos, (0) as Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL  " & _
            " WHERE (COF_SALDO_INICIAL.ID_CUENTA = COF_CUENTA.ID_CUENTA) AND " & _
            "       (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR In ('" & sAccounts & "'))  " & _
            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
            "       COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, COF_CUENTA_ESTANDAR.NATURALEZA, " & _
            "       COF_SALDO_INICIAL.SALDO_INICIAL "
        
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        COF_CUENTA_ESTANDAR.NATURALEZA, (0) as SI, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
            " FROM COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, " & _
            "      COF_TRANSACCION, COF_SALDO_INICIAL, COF_ELEMENTO_GRUPO_MAYOR " & _
            " WHERE (COF_MOVIMIENTO.ID_CUENTA = COF_CUENTA.ID_CUENTA) and " & _
            "       (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  AND " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) AND " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR In ('" & sAccounts & "')) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NATURALEZA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sSQL = " SELECT A.ID_MONEDA, A.Cuenta, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , A.NATURALEZA, sum(A.SI) SI, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos " & _
           " FROM ( " & _
           sSQL1 & _
           " UNION " & _
           sSQL2 & _
           " ) A, COF_CUENTA_AUXILIAR " & _
           " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) " & _
           " GROUP BY A.ID_MONEDA, A.Cuenta, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA " & _
           " ORDER BY A.ID_MONEDA, A.Cuenta, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA "
        
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset60 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraRecordset60", Err.Number
End Function

Public Function GeneraRecordset61(sAppServer As String, vGL As Variant, _
                                  dEndDate As Date, _
                                  dExcDate As Date, _
                                  Optional nExcRateType As Long = 46, _
                                  Optional nExcRateCurrency As Long = 1) As Recordset
  Dim oRecordset      As New Recordset
  Dim oConn           As New Connection
  Dim sSQL1           As String
  Dim sSql1A          As String
  Dim sSql1B          As String
  Dim sSql1C          As String
  Dim sSQL2           As String
  Dim sSql2A          As String
  Dim sSql2B          As String
  Dim sSql2C          As String
  Dim sSql3           As String
  Dim sSql4           As String
  
  Dim sSqlComp        As String
  Dim sGl             As String
  Dim sCondition      As String
  Dim sGLCondition    As String
  Dim sSQLStdAccHist  As String
  '************************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
  
     sCondition = " AND (fecha_afectacion <= '" & dEndDate & "') "
  
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGLCondition = "id_mayor = " & sGl
    End If
    
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dEndDate)
    CreateGralMovTBL oConn, sGLCondition, "tmp_MovsTBL"
    
    sSQL1 = " SELECT A.CUENTA, A.ID_MONEDA, CURRENCY_NAME NOM_MONEDA, NVL(A.SECTOR,0) SECTOR, (0) AUX, A.NATURALEZA NAT, " & _
                 " ('0') NUM_AUX, ('0') NOM_AUX, (-1) EXC_RATE, A.ID_MAYOR, " & _
                 " DECODE(A.NATURALEZA, 'D', SUM(A.SI)+SUM(A.CARGOS)-SUM(A.ABONOS), " & _
                                      " 'A', SUM(A.SI)-SUM(A.CARGOS)+SUM(A.ABONOS)) ST " & _
          " FROM ( "

          sSql1A = " SELECT COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR CUENTA, COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_SECTOR SECTOR, " & _
                         " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR AUX, NATURALEZA, COF_SALDO_INICIAL.ID_MAYOR, " & _
                         " (Sum(COF_SALDO_INICIAL.SALDO_INICIAL)) SI, (0) Cargos, (0) Abonos " & _
                  " From COF_SALDO_INICIAL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (COF_SALDO_INICIAL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                        " (COF_SALDO_INICIAL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND (COF_SALDO_INICIAL." & sGLCondition & ") AND (FECHA_SALDO_INICIAL<='" & dEndDate & "')  " & _
                  "       AND (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_SECTOR, " & _
                           " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR , NATURALEZA, COF_SALDO_INICIAL.ID_MAYOR, EXCHANGE_RATE " & _
                  " Union "
                  
         sSql1B = " SELECT COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR CUENTA, ID_MONEDA, ID_SECTOR SECTOR, ID_CUENTA_AUXILIAR AUX, NATURALEZA, tmp_MovsTBL.ID_MAYOR, " & _
                          " (0) SI, " & _
                          " (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
                          " (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
                  " From tmp_MovsTBL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR " & _
                  " WHERE (tmp_MovsTBL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                        " (FECHA_AFECTACION>=NVL(fecha_saldo_inicial,'01/01/1990')) " & sCondition & _
                  " GROUP BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, ID_MONEDA, ID_SECTOR, ID_CUENTA_AUXILIAR, NATURALEZA, tmp_MovsTBL.ID_MAYOR " & _
          " ) A, AO_CURRENCIES " & _
          " WHERE (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) " & _
          " GROUP BY A.CUENTA, A.ID_MONEDA, CURRENCY_NAME, A.SECTOR, A.NATURALEZA, A.ID_MAYOR " & _
          " Union "

  sSQL2 = " SELECT A.CUENTA, A.ID_MONEDA, ('0') NOM_MONEDA, NVL(A.SECTOR,0) SECTOR, NVL(A.AUX,0) AUX, A.NATURALEZA NAT, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUM_AUX, " & _
                 " NVL(NOMBRE_CUENTA_AUXILIAR,' ') NOM_AUX, A.EXC_RATE, A.ID_MAYOR, " & _
                 " DECODE(A.NATURALEZA, 'D', SUM(A.SI)+SUM(A.CARGOS)-SUM(A.ABONOS), " & _
                                      " 'A', SUM(A.SI)-SUM(A.CARGOS)+SUM(A.ABONOS)) ST " & _
          " FROM ( "
         sSql2A = " SELECT COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR CUENTA, COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_SECTOR SECTOR, " & _
                         " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR AUX, NATURALEZA, NVL(EXCHANGE_RATE,1) EXC_RATE, COF_SALDO_INICIAL.ID_MAYOR, " & _
                         " (Sum(COF_SALDO_INICIAL.SALDO_INICIAL)) SI, (0) Cargos, (0) Abonos " & _
                  " From COF_SALDO_INICIAL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (COF_SALDO_INICIAL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                        " (COF_SALDO_INICIAL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND (COF_SALDO_INICIAL." & sGLCondition & ") AND " & _
                        " (FECHA_SALDO_INICIAL<='" & dEndDate & "') " & _
                  "       AND (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_SECTOR, " & _
                           " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR , NATURALEZA, EXCHANGE_RATE, COF_SALDO_INICIAL.ID_MAYOR " & _
                  " Union "

         sSql2B = " SELECT COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR CUENTA, ID_MONEDA, ID_SECTOR SECTOR, ID_CUENTA_AUXILIAR AUX, NATURALEZA, " & _
                          " NVL(EXCHANGE_RATE,1) EXC_RATE, tmp_MovsTBL.ID_MAYOR, " & _
                          " (0) SI, " & _
                          " (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
                          " (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
                  " From tmp_MovsTBL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (tmp_MovsTBL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                        " (tmp_MovsTBL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND " & _
                        " (FECHA_AFECTACION>=NVL(fecha_saldo_inicial,'01/01/1990')) " & sCondition & _
                  "       AND (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, ID_MONEDA, ID_SECTOR, ID_CUENTA_AUXILIAR, NATURALEZA, " & _
                        " EXCHANGE_RATE , tmp_MovsTBL.ID_MAYOR "
          
          sSql2C = " ) A, COF_CUENTA_AUXILIAR " & _
           " WHERE (A.AUX = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) " & _
           " GROUP BY A.CUENTA, A.ID_MONEDA, A.SECTOR, A.AUX, A.NATURALEZA, " & _
                    " NUMERO_CUENTA_AUXILIAR , NOMBRE_CUENTA_AUXILIAR, A.EXC_RATE, A.ID_MAYOR "
                    
  sSqlComp = sSQL1 & sSql1A & sSql1B & sSQL2 & sSql2A & sSql2B & sSql2C
  
  sSqlComp = " SELECT A.*, SUBSTR(A.CUENTA,1,4) CUENTA1,  NVL(SUBSTR(A.CUENTA,6,LENGTH(A.CUENTA)-5),' ') AS SUBCUENTA, " & _
                      " NVL(COF_SECTOR.CLAVE_SECTOR,'00') CVE_SECTOR, NVL(COF_SECTOR.NOMBRE_SECTOR,' ') NOM_SECTOR, " & _
                      " NVL(COF_MAYOR.NUMERO_MAYOR,' ') NUM_MAYOR, NVL(COF_MAYOR.NOMBRE_MAYOR,' ') NOM_MAYOR " & _
             " FROM ( " & sSqlComp & " ) A, COF_SECTOR, COF_MAYOR " & _
             " WHERE (A.SECTOR = COF_SECTOR.ID_SECTOR(+)) AND (A.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND (A.ST <> 0) " & _
             " ORDER BY NAT DESC, CUENTA1, ID_MONEDA, SUBCUENTA, SECTOR, AUX, EXC_RATE "
  
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSqlComp, oConn, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset61 = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset61", Err.Number
End Function

Public Function GeneraRecordset62(ByVal sAccs As String, sAppServer As String, vGL As Variant, _
                                   dIniDate As Date, dEndDate As Date) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oCommand             As ADODB.Command
  Dim oRs                  As New ADODB.Recordset

  Dim sSQL                 As String
  Dim sGl                  As String
  Dim sGLCondition         As String
  Dim sAccounts            As String
  Dim sAccCond             As String
  Dim vAccounts            As Variant
  Dim i                    As Long
  
  Dim sSQLStdAccHist       As String
  '**********************************
  On Error GoTo ErrHandler

    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGLCondition = "id_mayor = " & sGl
    End If
    vAccounts = Split(sAccs, ",")
    If (IsArray(vAccounts)) Then
      sAccounts = vAccounts(0)
      sAccCond = " ((NUMERO_CUENTA_ESTANDAR like '" & sAccounts & "%')  "
      For i = LBound(vAccounts) + 1 To UBound(vAccounts)
        sAccounts = vAccounts(i)
        sAccCond = sAccCond & " or (NUMERO_CUENTA_ESTANDAR like '" & sAccounts & "%') "
      Next i
      sAccCond = sAccCond & " ) and "
    ElseIf Not (IsArray(vAccounts)) Then
      sAccounts = vAccounts
      sAccCond = sAccCond & " (NUMERO_CUENTA_ESTANDAR like '" & sAccounts & "%') and "
    End If

    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    sSQLStdAccHist = SQLStandardAccountHist(dIniDate, dEndDate, 1)
    
    sSQL = " SELECT  COF_TRANSACCION.ID_MAYOR, COF_TRANSACCION.FECHA_AFECTACION as FEC_AFE, " & _
           "         COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR as CUENTA, NVL(COF_SECTOR.CLAVE_SECTOR,'00') as CLAVE_SECTOR, " & _
           "         COF_CUENTA_ESTANDAR.NOMBRE_CUENTA_ESTANDAR as NOMBRE_CUENTA, " & _
           "         NVL(COF_TRANSACCION.NUMERO_TRANSACCION,'0') as POLIZA, AO_CURRENCIES.O_ID_MONEDA as MONEDA, " & _
           "         NVL(COF_TRANSACCION.CONCEPTO_TRANSACCION,' ') as CONCEPTO, COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR as AUX, " & _
           "         COF_MOVIMIENTO.TIPO_MOVIMIENTO, MONTO_MONEDA_BASE as MONTO " & _
           " FROM    COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_TRANSACCION,COF_MOVIMIENTO,COF_CUENTA_AUXILIAR, " & _
           "         AO_CURRENCIES , COF_SECTOR " & _
           " WHERE  ((COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
           "         (COF_TRANSACCION.ID_TRANSACCION = COF_MOVIMIENTO.ID_TRANSACCION) AND " & _
           "         (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) AND " & _
           "         (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
           "         (COF_MOVIMIENTO.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID) AND " & _
           "         (COF_MOVIMIENTO.ID_SECTOR = COF_SECTOR.ID_SECTOR)) AND " & _
           "         (COF_TRANSACCION." & sGLCondition & ") AND " & _
           "         (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "') AND " & _
                      sAccCond & _
           "         (COF_TRANSACCION.FECHA_AFECTACION <= '" & dEndDate & "') " & _
           " ORDER BY   COF_TRANSACCION.FECHA_AFECTACION, AO_CURRENCIES.O_ID_MONEDA, " & _
           "         COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,COF_SECTOR.CLAVE_SECTOR, " & _
           "         COF_TRANSACCION.NUMERO_TRANSACCION, " & _
           "         COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR, COF_MOVIMIENTO.TIPO_MOVIMIENTO "
                                
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset62 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraRecordset62", Err.Number
End Function


Public Function GeneraRecordset9(sAppServer As String, vGL As Variant, _
                                 dIniDate As Date, dEndDate As Date) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oCommand             As ADODB.Command
  Dim oRs                  As New ADODB.Recordset

  Dim sSQL                 As String
  Dim sGl                  As String
  Dim sGLCondition         As String
  Dim sAccCond             As String
  Dim vAccounts            As Variant
  
  Dim sSQLStdAccHist       As String
  '**********************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGLCondition = "id_mayor = " & sGl
    End If
    
    sSQLStdAccHist = SQLStandardAccountHist(dIniDate, dEndDate, 1)
    
    sSQL = " SELECT  COF_TRANSACCION.ID_MAYOR, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR CUENTA, " & _
           "         NVL(COF_SECTOR.CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
           "         COF_CUENTA_ESTANDAR.NOMBRE_CUENTA_ESTANDAR NOMBRE_CUENTA, " & _
           "         NVL(COF_TRANSACCION.NUMERO_TRANSACCION,'0') POLIZA, AO_CURRENCIES.O_ID_MONEDA MONEDA, " & _
           "         AO_CURRENCIES.CURRENCY_NAME, " & _
           "         NVL(COF_TRANSACCION.CONCEPTO_TRANSACCION,' ') as CONCEPTO, " & _
           "         NVL(COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,'0') AUX, " & _
           "         COF_MOVIMIENTO.TIPO_MOVIMIENTO, MONTO " & _
           " FROM    COF_MOVIMIENTO, COF_TRANSACCION, COF_CUENTA_AUXILIAR, " & _
           "         AO_CURRENCIES, COF_SECTOR, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR  " & _
           " WHERE  (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) AND " & _
           "        (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
           "        (COF_MOVIMIENTO.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID) AND " & _
           "        (COF_MOVIMIENTO.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
           "        (COF_MOVIMIENTO.ID_CUENTA = COF_CUENTA.ID_CUENTA) AND " & _
           "        (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
           "        (COF_TRANSACCION." & sGLCondition & ") AND " & _
           "        (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "') AND " & _
           "        (COF_TRANSACCION.FECHA_AFECTACION <= '" & dEndDate & "') " & _
           " ORDER BY AO_CURRENCIES.O_ID_MONEDA, " & _
           "          COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR,COF_SECTOR.CLAVE_SECTOR, " & _
           "          COF_TRANSACCION.NUMERO_TRANSACCION, " & _
           "          COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR, COF_MOVIMIENTO.TIPO_MOVIMIENTO "
                                
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset9 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraRecordset9", Err.Number
End Function

Public Function GeneraRecordset64(sAppServer As String, vGL As Variant, _
                                  dIniDate As Date, dCuteDate As Date, Optional sIniAccount As String = "", _
                                  Optional sEndAccount As String = "") As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String
  
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2, sSql3      As String

  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long

  Dim sSQLStdAccHist    As String
  Dim sWhere            As String
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

    If (sIniAccount = "") And (sEndAccount = "") Then
      sWhere = " "
    ElseIf (sIniAccount <> "") And (sEndAccount <> "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sIniAccount & "' AND '" & sEndAccount & "') "
    ElseIf (sIniAccount <> "") And (sEndAccount = "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR <= '" & sIniAccount & "') "
    End If

    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)


    sSQL1 = " SELECT ID_MONEDA, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "

    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & Format((dIniDate - 1), "DD/MM/YYYY") & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sStdAccHist = Replace(StdAccHist, "<@dIniDate@>", dIniDate)
    sStdAccHist = Replace(sStdAccHist, "<@dEndDate@>", dCuteDate)
    
    sSQLStdAccHist = " SELECT A.ID_CUENTA_ESTANDAR_HIST, A.ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, " & _
                     " NOMBRE_CUENTA_ESTANDAR, NATURALEZA, ID_TIPO_CUENTAS_STD, ID_TIPO_CUENTA, ROL_CUENTA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST A, (" & sStdAccHist & ") B " & _
                     " WHERE (A.ID_CUENTA_ESTANDAR_HIST = B.ID_CUENTA_ESTANDAR_HIST) AND " & _
                     "       (A.ID_CUENTA_ESTANDAR = B.ID_CUENTA_ESTANDAR) AND " & _
                     "       (A.ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )"
    
    sSQL_SI = " SELECT id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR,  " & _
              "        to_date('01/01/1990') fecha_rev, ('') CONCEPTO_MOVIMIENTO, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
              " FROM ( " & _
                sSQL1 & _
              " UNION " & _
                sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4), NATURALEZA "
    
    sSql3 = " SELECT COF_MOVIMIENTO.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, " & _
            "        FECHA_AFECTACION FECHA_REV, CONCEPTO_MOVIMIENTO, (0) SALDO_INICIAL, " & _
            "        sum(decode(TIPO_MOVIMIENTO,'D',Monto,0)) Cargos, " & _
            "        sum(decode(TIPO_MOVIMIENTO,'H',Monto,0)) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI, " & _
            "      COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = B.ID_CUENTA) and " & _
            "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4), FECHA_AFECTACION, " & _
            "          CONCEPTO_MOVIMIENTO "
   
    sSQL = " SELECT id_moneda, o_id_moneda, currency_name nombre_moneda, A.numero_cuenta_estandar cuenta, " & _
           "        NVL(nombre_cuenta_estandar,' ') nombre_cuenta, concepto_movimiento, " & _
           "        NVL(fecha_rev,'01/01/1990') fecha_ult, " & _
           "        Saldo_Inicial SI, Cargos, Abonos, NATURALEZA " & _
           " FROM ( " & _
             sSQL_SI & _
           " UNION " & _
             sSql3 & _
           " ) A, (" & sSQLStdAccHist & ") B, ao_currencies " & _
           " WHERE (A.numero_cuenta_estandar = B.numero_cuenta_estandar) AND " & _
           "       (A.id_moneda = ao_currencies.currency_id(+)) " & _
           " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, FECHA_rev  "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset64 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset64", Err.Number
End Function

Public Function GeneraRecordset68(sAppServer As String, vGL As Variant, _
                                  dEndDate As Date) As Recordset
  Dim oConn           As New ADODB.Connection
  Dim oRs             As New ADODB.Recordset
  
  Dim sGl             As String
  
  Dim sSQL            As String
  Dim sSql_SA         As String
  Dim sSQL1           As String
  Dim sSQL2           As String
  
  Dim sSQLStdAccHist  As String
    '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    sSQL1 = " SELECT ID_MONEDA, ID_MAYOR, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dEndDate & "') " & _
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
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dEndDate) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
            
    
    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dEndDate, 1)
    
      sSQL = " SELECT NUMERO_CUENTA_ESTANDAR CUENTA, A.ID_MONEDA, CURRENCY_NAME NOM_MONEDA, NVL(CLAVE_SECTOR,' ') SECTOR, " & _
             "       DECODE(NVL(A.ID_SECTOR,0),0, NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE,   " & _
             "       NATURALEZA NAT, A.ID_MAYOR, NUMERO_MAYOR NUM_MAYOR, " & _
             "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
             "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) ST " & _
             " FROM ( " & _
               sSQL1 & _
             " UNION " & _
               sSQL2 & _
             " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C, AO_CURRENCIES D, COF_SECTOR E, COF_MAYOR F " & _
             " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
             "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
             "       (A.ID_MONEDA = D.CURRENCY_ID(+)) AND " & _
             "       (A.ID_SECTOR = E.ID_SECTOR(+)) AND " & _
             "       (A.ID_MAYOR = F.ID_MAYOR(+)) AND " & _
             "       (NUMERO_CUENTA_ESTANDAR LIKE '1103%') " & _
             " GROUP BY NUMERO_CUENTA_ESTANDAR, A.ID_MONEDA, CLAVE_SECTOR, NATURALEZA," & _
             "          A.ID_MAYOR, NUMERO_MAYOR, DECODE(NVL(A.ID_SECTOR,0),0, NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR), " & _
             "          CURRENCY_NAME " & _
             " ORDER BY ID_MONEDA, NUMERO_MAYOR, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
             
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset68 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset68", Err.Number
End Function


Public Function GeneraRecordset69(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        vAccountGroups As Variant, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_B            As String
  Dim sSQL_C            As String
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
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
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
            

      sSQL_A = " SELECT A.ID_MAYOR, sum(A.ABONOS) ABONOS, (0) CARGOS_2, (0) SALDO_ACTUAL_2, (0) SALDO_ACTUAL_3 " & _
                 " FROM ( " & sSQL2 & ") A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(0) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MAYOR"

            
      sSQL_B = " SELECT A.ID_MAYOR, (0) ABONOS, SUM(A.CARGOS) CARGOS_2, " & _
               "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                 "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL_2, (0) SALDO_ACTUAL_3 " & _
                 " FROM ( " & _
                   sSQL_SI & _
                 " UNION " & _
                   sSQL2 & _
                 " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(1) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MAYOR, NATURALEZA "

      sSQL_C = " SELECT A.ID_MAYOR, (0) ABONOS, (0) CARGOS_2, (0) SALDO_ACTUAL_2, " & _
                 "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                 "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL_3 " & _
                 " FROM ( " & _
                   sSQL_SI & _
                 " UNION " & _
                   sSQL2 & _
                 " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(2) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MAYOR, NATURALEZA "
    
        sSQL = " SELECT  NUMERO_MAYOR, NOMBRE_MAYOR, DECODE(COF_GRUPO_MAYOR.ID_GRUPO_MAYOR, 23, 2, 14, 2, 1) ORDER_ID_GRUPO_MAYOR, " & _
               "         SUM(ABONOS) ABONOS, SUM(CARGOS_2) CARGOS_2, " & _
               "         SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, SUM(SALDO_ACTUAL_3) SALDO_ACTUAL_3 " & _
               " FROM (" & _
                 sSQL_A & _
               " UNION " & _
                 sSQL_B & _
               " UNION " & _
                 sSQL_C & _
               " ) A, COF_MAYOR, COF_ELEMENTO_GRUPO_MAYOR, COF_GRUPO_MAYOR " & _
               " WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR) AND " & _
               "       (A.ID_MAYOR = COF_ELEMENTO_GRUPO_MAYOR.id_mayor) AND " & _
               "       (COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR = COF_GRUPO_MAYOR.ID_GRUPO_MAYOR) AND " & _
               "       (COF_GRUPO_MAYOR.CLIP_GRUPO_MAYOR = 2) " & _
               " GROUP BY NUMERO_MAYOR, NOMBRE_MAYOR, COF_GRUPO_MAYOR.ID_GRUPO_MAYOR " & _
               " ORDER BY ORDER_ID_GRUPO_MAYOR, NUMERO_MAYOR, NOMBRE_MAYOR "
         
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRecordset69 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset69", Err.Number
End Function

Public Function GeneraRecordset152(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        vAccountGroups As Variant, dExcDate As Date, _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1, _
                        Optional nRowsToDisplay As Long = 4) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_B            As String
  Dim sSQL_C            As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sWhere            As String
  Dim sWhereTran        As String
  Dim sOptionToDisplay  As String
  
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
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
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
            

      sSQL_A = " SELECT A.ID_MAYOR, A.ID_MONEDA, " & _
               "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
               "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL, " & _
               "        (0) SALDO_ACTUAL_2 " & _
               " FROM ( " & _
                 sSQL_SI & _
               " UNION " & _
                 sSQL2 & _
               " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
               " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
               "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
               "       ((NUMERO_CUENTA_ESTANDAR like (" & vAccountGroups(0) & ")) OR " & _
               "       (NUMERO_CUENTA_ESTANDAR like (" & vAccountGroups(1) & "))) " & _
                       sWhere & _
               " GROUP BY A.ID_MAYOR, A.ID_MONEDA, NATURALEZA "

      sSQL_B = " SELECT A.ID_MAYOR, DECODE(NUMERO_CUENTA_ESTANDAR, '5000-99', 2,28) ID_MONEDA, " & _
               "        (0) SALDO_ACTUAL, " & _
               "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
               "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL_2 " & _
               " FROM ( " & _
                 sSQL_SI & _
               " UNION " & _
                 sSQL2 & _
               " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
               " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
               "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
               "       ((NUMERO_CUENTA_ESTANDAR like (" & vAccountGroups(2) & ")) OR " & _
               "       (NUMERO_CUENTA_ESTANDAR like (" & vAccountGroups(3) & "))) " & _
               " GROUP BY A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "

'sWhere
        sSQL = " SELECT  DECODE(NVL(SUB_NUMERO_MAYOR,0),0,NUMERO_MAYOR||'    '||'0000',SUB_NUMERO_MAYOR||'    '||NUMERO_MAYOR) NUMERO_MAYOR_COMPLETO, " & _
               "         O_ID_MONEDA, SUM(SALDO_ACTUAL) SALDO_ACTUAL, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, " & _
               "         EXCHANGE_RATE " & _
               " FROM (" & _
                 sSQL_A & _
               " UNION " & _
                 sSQL_B & _
               " ) A, COF_MAYOR, AO_EXCHANGE_RATES, AO_CURRENCIES " & _
               " WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR) AND " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID) AND " & _
               "       (A.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) and " & _
               "       (EXCHANGE_RATE_TYPE_ID(+) = " & nExcRateType & ") and " & _
               "       (FROM_CURRENCY_ID(+) = " & nExcRateCurrency & ") and " & _
               "       (FROM_DATE(+) <= '" & dExcDate & "') AND " & _
               "       (TO_DATE(+) >= '" & dExcDate & "') " & _
               " GROUP BY DECODE(NVL(SUB_NUMERO_MAYOR,0),0,NUMERO_MAYOR||'    '||'0000',SUB_NUMERO_MAYOR||'    '||NUMERO_MAYOR), " & _
               "          O_ID_MONEDA, EXCHANGE_RATE " & _
               " ORDER BY NUMERO_MAYOR_COMPLETO "
         
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRecordset152 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset152", Err.Number
End Function

Public Function GeneraRecordset70(sAppServer As String, dCuteDate As Date) As Recordset

  Dim oConn                 As New ADODB.Connection
  Dim oCommand              As ADODB.Command
  Dim oRs                   As New ADODB.Recordset

  Dim sGl                   As String
  Dim sSQL                  As String
  Dim sSQL1                 As String
  Dim sSQL2                 As String
  Dim sSql3                 As String
  Dim sSQLStdAccHist        As String
  Dim sSQL_SI               As String
  Dim sSQL_CA               As String
  '**********************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open

    sSQLStdAccHist = SQLStandardAccountHist("01/01/1990", dCuteDate)
    
    sSQL_SI = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_MAYOR, COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR, " & _
              "        substr(NUMERO_CUENTA_ESTANDAR,1,4) Cuenta, NATURALEZA, " & _
              "        Sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) Cargos,(0) Abonos " & _
              " FROM COF_ELEMENTO_GRUPO_MAYOR, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
              " WHERE ((COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND  " & _
              "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) AND " & _
              "       (COF_CUENTA.ID_MAYOR = COF_ELEMENTO_GRUPO_MAYOR.ID_MAYOR)) AND " & _
              "       (COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR In (17,18,19,20,25,23)) and " & _
              "       substr(NUMERO_CUENTA_ESTANDAR,1,4) in ('1101', '1103', '1202', '1204', '1206') " & _
              " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_MAYOR, COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR, " & _
              "       substr(NUMERO_CUENTA_ESTANDAR,1,4), NATURALEZA "

    sSQL_CA = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA.ID_MAYOR, COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR, " & _
              "        substr(NUMERO_CUENTA_ESTANDAR,1,4) Cuenta, NATURALEZA, (0) SI, " & _
              "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
              "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
              " FROM  COF_ELEMENTO_GRUPO_MAYOR, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, " & _
              "       COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL " & _
              " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
              "       (COF_MOVIMIENTO.ID_MONEDA = COF_SALDO_INICIAL.ID_MONEDA(+)) and " & _
              "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
              "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
              "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
              "       (COF_MOVIMIENTO.ID_CUENTA = COF_CUENTA.ID_CUENTA) AND " & _
              "       (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
              "       (COF_CUENTA.ID_MAYOR = COF_ELEMENTO_GRUPO_MAYOR.ID_MAYOR) AND " & _
              "       (COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR In (17,18,19,20,25,23)) and  " & _
              "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "') AND " & _
              "        SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4) in ('1101', '1103', '1202', '1204', '1206') and " & _
              "       (FECHA_AFECTACION > nvl(FECHA_SALDO_INICIAL,'01/01/1990')) " & _
              "       GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA.ID_MAYOR, COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR, " & _
              "                SUBSTR(NUMERO_CUENTA_ESTANDAR,1,4), NATURALEZA "

    sSQL1 = " SELECT DECODE(ID_GRUPO_MAYOR, 23, 2, 1) OBJECT, A.ID_MONEDA, A.ID_MAYOR, A.ID_GRUPO_MAYOR, A.CUENTA, A.NATURALEZA, " & _
            "        sum(A.SI) SI, sum(A.Cargos) CARGOS, sum(A.Abonos) ABONOS " & _
            " FROM ( " & _
                sSQL_SI & _
            " Union " & _
                sSQL_CA & _
            " ) A  " & _
            " GROUP BY A.ID_MONEDA, A.ID_MAYOR, A.ID_GRUPO_MAYOR, A.CUENTA, A.NATURALEZA "

    sSQL = " SELECT S.ID_MONEDA, S.OBJECT, S.ID_MAYOR, NOMBRE_MAYOR, " & _
                  " DECODE(S.ID_GRUPO_MAYOR, 17, '1936',COF_MAYOR.NUMERO_MAYOR) NUMERO_MAYOR , S.CUENTA,  S.CARGOS, S.ABONOS, (SI+CARGOS-ABONOS) SA " & _
           " FROM ( " & _
           sSQL1 & _
           " ) S , COF_MAYOR " & _
           " WHERE S.ID_MAYOR = COF_MAYOR.ID_MAYOR " & _
           " ORDER BY  S.ID_MONEDA, S.OBJECT, NUMERO_MAYOR, NOMBRE_MAYOR, S.CUENTA "
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset70 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset70", Err.Number
End Function


Public Function GeneraRecordset71(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        vAccountGroups As Variant, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As Recordset
                                 
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_B            As String
  Dim sSQL_C            As String
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
            

      sSQL_A = " SELECT A.ID_MAYOR, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_ACTUAL_2, (0) SALDO_ACTUAL_3 " & _
               " FROM ( " & _
                 " SELECT A.ID_MONEDA, A.ID_MAYOR, " & _
                 "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                 "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                 " FROM ( " & _
                   sSQL_SI & _
                 " UNION " & _
                   sSQL2 & _
                 " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(0) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NATURALEZA " & _
               "  ) A, AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY ID_MAYOR "

            
      sSQL_B = " SELECT A.ID_MAYOR, (0) SALDO_ACTUAL, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_2, (0) SALDO_ACTUAL_3 " & _
               " FROM ( " & _
                 " SELECT A.ID_MONEDA, A.ID_MAYOR, " & _
                 "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                 "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                 " FROM ( " & _
                   sSQL_SI & _
                 " UNION " & _
                   sSQL2 & _
                 " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(1) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NATURALEZA " & _
               "  ) A, AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY ID_MAYOR "
               
      sSQL_C = " SELECT A.ID_MAYOR, (0) SALDO_ACTUAL, (0) SALDO_ACTUAL_2, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_3 " & _
               " FROM ( " & _
                 " SELECT A.ID_MONEDA, A.ID_MAYOR, " & _
                 "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
                 "                              (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_ACTUAL " & _
                 " FROM ( " & _
                   sSQL_SI & _
                 " UNION " & _
                   sSQL2 & _
                 " ) A  , COF_C1UENTA B, (" & sSQLStdAccHist & ") C " & _
                 " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                 "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) AND " & _
                 "       (NUMERO_CUENTA_ESTANDAR in (" & vAccountGroups(2) & ")) " & _
                         sWhere & _
                 " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NATURALEZA " & _
               "  ) A, AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY ID_MAYOR "
    
    
        
        sSQL = " SELECT  NUMERO_MAYOR, NOMBRE_MAYOR, min(ID_GRUPO_MAYOR) id_grupo_mayor, " & _
               "         SUM(SALDO_ACTUAL) SALDO_ACTUAL, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, " & _
               "         SUM(SALDO_ACTUAL_3) SALDO_ACTUAL_3 " & _
               " FROM (" & _
                 sSQL_A & _
               " UNION " & _
                 sSQL_B & _
               " UNION " & _
                 sSQL_C & _
               " ) A, COF_MAYOR, COF_ELEMENTO_GRUPO_MAYOR " & _
               " WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR) AND " & _
               "       (A.ID_MAYOR = COF_ELEMENTO_GRUPO_MAYOR.id_mayor) " & _
               " GROUP BY NUMERO_MAYOR, NOMBRE_MAYOR " & _
               " ORDER BY ID_GRUPO_MAYOR, NUMERO_MAYOR "
         
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRecordset71 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset71", Err.Number
End Function

Public Function GeneraRecordset73(sAppServer As String, sPatron As String, vGL As Variant, _
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
                        Optional bOnlyLastLevels As Boolean = False) As Recordset
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
            
      Truncate sAppServer, "tmp_SALDOS_MAYOR"
      
      sSQL = " INSERT INTO tmp_SALDOS_MAYOR " & _
              " SELECT A.ID_MAYOR, A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MAYOR, A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
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
                " GROUP BY A.ID_MAYOR, A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A "
              
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
      sFromTable = " tmp_SALDOS_MAYOR "

    If Not bOnlyLastLevels Then
      sSQL = "    SELECT id_mayor, id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1,4) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
             "           (0) ID_CUENTA_AUXILIAR, " & _
             "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
             "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
             "    From  " & sFromTable & "  A " & _
             "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) AND " & _
             "          (nvl(ID_SECTOR,0) <> 0) " & _
             "    GROUP BY id_mayor, id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, 4) " & _
             " Union "
  
      nPos = GetPositions(sPatron)
      For i = LBound(nPos) To UBound(nPos)
        sSQL = sSQL & "    SELECT id_mayor, id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
                      "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS,  Sum(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
                      "     From  " & sFromTable & " A " & _
                      "    Where (Length(NUMERO_CUENTA_ESTANDAR) >=  " & nPos(i) & " ) and " & _
                      "          not ((SUBSTR( A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_CUENTA_AUXILIAR,0)=0)) " & _
                      "    GROUP BY id_mayor, id_moneda , SUBSTR(NUMERO_CUENTA_ESTANDAR,1,  " & nPos(i) & "  ), ID_SECTOR " & _
                      " Union "
      Next i
      
      sSQL = " SELECT ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
             "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
             "        SUM(SALDO_ACTUAL) SALDO_ACTUAL FROM " & _
             " ( " & sSQL & _
             "       SELECT ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
             "              SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
             "       From  " & sFromTable & " A " & _
             " ) " & _
             " GROUP BY ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
    Else
    
      If bWithAux Then
        sSQL = " SELECT ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "        SUM(SALDO_ACTUAL) SALDO_ACTUAL " & _
               " FROM  " & sFromTable & " A " & _
               " GROUP BY ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR "
      Else
        sSQL = " SELECT ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, (0) ID_CUENTA_AUXILIAR, " & _
               "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
               "        SUM(SALDO_ACTUAL) SALDO_ACTUAL " & _
               " FROM  " & sFromTable & " A " & _
               " GROUP BY ID_MAYOR, ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR "
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

      Else
        sSQL = " select A.ID_MAYOR, DECODE(NVL(SUB_NUMERO_MAYOR,0),0,NUMERO_MAYOR||'    '||'0000',SUB_NUMERO_MAYOR||'    '||NUMERO_MAYOR) NOMBRE_MAYOR, " & _
               "        A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
               "        DECODE(NVL(A.ID_SECTOR,0),0,NOMBRE_CUENTA_ESTANDAR, NOMBRE_SECTOR) NOMBRE_CUENTA, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_MAYOR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) AND " & _
               "       (A.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) " & _
                       sOptionToDisplay & _
               " ORDER BY NOMBRE_MAYOR, O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      End If
    ElseIf nExcRateCurrency <> 0 Then
    
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
    Set GeneraRecordset73 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset73", Err.Number
End Function

Public Function GeneraRecordset101(sPatron As String, sAppServer As String, nGl As Long, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, dExcDate As Date, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1, _
                        Optional sAccounts As String) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset
  
  Dim sGl               As String
  
  Dim sSQL              As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
  Dim sSQL_MN, sSQL_ME  As String
  Dim sWhereAccount     As String
  
  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long
  
  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGlStdAcc         As Long
  
  Dim sSQLStdAccHist    As String
  
  '***************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    If (IsArray(vGL)) Then
      nGlStdAcc = CLng(vGL(0))
      sGl = Join(vGL, ",")
    Else
      nGlStdAcc = CLng(vGL)
      sGl = vGL
    End If
    
    sGl = "id_mayor in (" & sGl & "," & nGl & ") "
    
    nTypeStdAcc = TypeStdAccount(sAppServer, nGlStdAcc)
    
    If sAccounts <> "" Then
      sWhereAccount = " NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ") AND "
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
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_TRANSACCION.ID_MAYOR, COF_MOVIMIENTO.ID_CUENTA , " & _
            "          COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, FECHA_AFECTACION "
                       
    Truncate sAppServer, "tmp_SALDOS_PROM_B"

    sSQL = " INSERT INTO tmp_SALDOS_PROM_B " & _
              " SELECT A.ID_MONEDA, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.FECHA_REV, A.SALDO_INICIAL, A.CARGOS, A.ABONOS, " & _
              "        A.SALDO_INICIAL_ALL, A.CARGOS_ALL, A.ABONOS_ALL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, FECHA_REV, " & _
                "        decode(A.id_mayor, " & nGl & ", sum(A.SALDO_INICIAL),0) SALDO_INICIAL, " & _
                "        decode(A.id_mayor, " & nGl & ", sum(A.CARGOS),0) CARGOS, " & _
                "        decode(A.id_mayor, " & nGl & ", sum(A.ABONOS),0) ABONOS, " & _
                "        decode(A.id_mayor, " & nGl & ", 0,sum(A.SALDO_INICIAL)) SALDO_INICIAL_ALL, " & _
                "        decode(A.id_mayor, " & nGl & ", 0,sum(A.CARGOS)) CARGOS_ALL, " & _
                "        decode(A.id_mayor, " & nGl & ", 0,sum(A.ABONOS)) ABONOS_ALL " & _
                " FROM ( " & _
                  sSQL_SI & _
                " UNION " & _
                  sSQL2 & _
                    " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
                " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
                "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
                " GROUP BY A.ID_MONEDA, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA, FECHA_REV, A.ID_MAYOR " & _
              "  ) A "
            
    Set oComm.ActiveConnection = oConn
    oComm.CommandText = sSQL
    oComm.CommandType = adCmdText
    oComm.Execute
    
    sFromTable = " tmp_SALDOS_PROM_B "
    sSQL = ""
    nPos = GetPositions(sPatron)
    For i = LBound(nPos) To UBound(nPos)
      sSQL = sSQL & "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & " ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
                    "           (0) ID_CUENTA_AUXILIAR,  " & _
                    "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
                    "           SUM(A.SALDO_INICIAL_ALL) SALDO_INICIAL_ALL, Sum(A.CARGOS_ALL) CARGOS_ALL, Sum(A.ABONOS_ALL) ABONOS_ALL, FECHA_REV  " & _
                    "    From  " & sFromTable & "  A " & _
                    "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= " & nPos(i) & ") and " & _
                    "          not ((SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & ") = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_SECTOR,0)=0) and (nvl(ID_CUENTA_AUXILIAR,0)=0)) " & _
                    "    GROUP BY id_moneda, SUBSTR(NUMERO_CUENTA_ESTANDAR,1, " & nPos(i) & "), FECHA_REV " & _
                    " Union "
    Next i
    
    sSQL = " SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
           "        SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, " & _
           "        SUM(SALDO_INICIAL_ALL) SALDO_INICIAL_ALL, SUM(CARGOS_ALL) CARGOS_ALL, SUM(ABONOS_ALL) ABONOS_ALL, FECHA_REV " & _
           " FROM " & _
           " ( " & sSQL & _
           "       SELECT ID_MONEDA, NUMERO_CUENTA_ESTANDAR, ID_SECTOR, ID_CUENTA_AUXILIAR, SALDO_INICIAL, CARGOS, ABONOS, SALDO_INICIAL_ALL, CARGOS_ALL, ABONOS_ALL, FECHA_REV " & _
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
            
    
    sSQL_MN = " SELECT A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        SUM(1*A.SALDO_INICIAL) SALDO_INICIAL, SUM(1*A.CARGOS) CARGOS, " & _
              "        SUM(1*A.ABONOS) ABONOS, " & _
              "        SUM(1*A.SALDO_INICIAL_ALL) SALDO_INICIAL_ALL, SUM(1*A.CARGOS_ALL) CARGOS_ALL, " & _
              "        SUM(1*A.ABONOS_ALL) ABONOS_ALL, " & _
              "        NVL(EXCHANGE_RATE,1) EXCHANGE_RATE, " & _
              "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
              " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        NOMBRE_CUENTA_ESTANDAR NOMBRE_CUENTA, " & _
                  "        FECHA_REV, SALDO_INICIAL, CARGOS, ABONOS, SALDO_INICIAL_ALL, CARGOS_ALL, ABONOS_ALL," & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                  "       (ID_CUENTA_AUXILIAR = 0) " & _
                  " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS, FECHA_REV, EXCHANGE_RATE "

    
      sSQL = " SELECT A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, FECHA_REV, " & _
              "        saldo_inicial, CARGOS, ABONOS, " & _
              "        SALDO_INICIAL_ALL, CARGOS_ALL, ABONOS_ALL, " & _
              "        A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, A.EXCHANGE_RATE " & _
              " FROM (" & _
                sSQL_MN & _
              " ) A " & _
              " WHERE  (length(NUMERO_CUENTA_ESTANDAR)) <= (length('" & sPatron & "')) and " & _
              "        (not(NUMERO_CUENTA_ESTANDAR like '%00-00-00')) and " & _
              "        (not(NUMERO_CUENTA_ESTANDAR like '%00-00')) and " & _
              "        (not(NUMERO_CUENTA_ESTANDAR like '%00')) and " & _
              "        (CLAVE_SECTOR = '00') AND " & sWhereAccount & _
              "        ((SALDO_INICIAL <> 0) OR (CARGOS <> 0) OR (ABONOS <> 0) OR " & _
              "        (SALDO_INICIAL_ALL <> 0) OR (CARGOS_ALL <> 0) OR (ABONOS_ALL <> 0)) " & _
              " ORDER BY A.ID_MONEDA, NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, FECHA_REV "
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset101 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset101", Err.Number
End Function

Public Function GeneraRecordset106(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Cartera", dIniDate, dCuteDate)
    'Debug.Print sSqlAux
    

    sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Cartera') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset106 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset106", Err.Number
End Function

Public Function GeneraRecordset107(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Administrativos", dIniDate, dCuteDate)
    'Debug.Print sSqlAux
    

        sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Administrativos') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "
           
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset107 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset107", Err.Number
End Function

Public Function GeneraRecordset108(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Activos", dIniDate, dCuteDate)
    
    sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Activos') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset108 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset108", Err.Number
End Function

Public Function GeneraRecordset109(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Pendientes", dIniDate, dCuteDate)
    
    sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Pendientes') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset109 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset109", Err.Number
End Function

Public Function GeneraRecordset144(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Fideicomisos", dIniDate, dCuteDate)
        sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Fideicomisos') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset144 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset144", Err.Number
End Function

Public Function GeneraRecordset156(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date) As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String

  ' Historicos
  Dim nGl               As Long
  Dim sSqlAux           As String
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
    
    sSqlAux = SQLCreaCamposAux(sAppServer, "Proveedores", dIniDate, dCuteDate)
    sSQL = " SELECT * " & _
           " FROM (" & sSqlAux & ") auxiliares " & _
           " WHERE (auxiliares. " & sGl & ") and " & _
           "       (object_name = 'Proveedores') " & _
           " ORDER BY NOMBRE_CUENTA_AUXILIAR "

    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset156 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset156", Err.Number
End Function

Public Function GeneraRecordset110(sPatron As String, sAppServer As String, vGL As Variant, _
                                   dIniDate As Date, dCuteDate As Date, _
                                   Optional sFromAccount As String = "", Optional sToAccount As String = "", _
                                   Optional sFromSubsAccount As String, Optional sToSubsAccount As String = "") As Recordset
  Dim oConn             As New ADODB.Connection
  Dim oComm             As New ADODB.Command
  Dim oRs               As New ADODB.Recordset

  Dim sGl               As String

  Dim sSQL              As String
  Dim sSQL_A            As String
  Dim sSQL_B            As String
  
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2, sSql3      As String


  Dim sSQL_MN, sSQL_ME  As String

  Dim sFromTable        As String
  Dim nPos              As Variant
  Dim i                 As Long

  ' Historicos
  Dim nTypeStdAcc       As Long
  Dim nGl               As Long

  Dim sSQLStdAccHist    As String
  Dim sSQLStdAccHist2   As String
  
  Dim sWhere As String, sWhereAux As String
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
    
    sWhere = LikeSQLString("numero_cuenta_estandar", sFromAccount, sToAccount)
    If Len(sWhere) <> 0 Then
      sWhere = " AND " & sWhere
    End If
    
    sWhereAux = LikeSQLString("numero_cuenta_auxiliar", sFromSubsAccount, sToSubsAccount)
    If Len(sWhereAux) <> 0 Then
      sWhereAux = " AND " & sWhereAux
    End If
    nTypeStdAcc = TypeStdAccount(sAppServer, nGl)


    sSQL1 = " SELECT ID_MONEDA, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, " & _
            "        FECHA_SALDO_INICIAL, sum(SALDO_INICIAL) SALDO_INICIAL, (0) Cargos, (0) Abonos " & _
            " FROM COF_SALDO_INICIAL " & _
            " WHERE (COF_SALDO_INICIAL." & sGl & ") and (FECHA_SALDO_INICIAL <= '" & dCuteDate & "') " & _
            " group by ID_MONEDA, ID_CUENTA, ID_SECTOR, ID_CUENTA_AUXILIAR, FECHA_SALDO_INICIAL "

    sSQL2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA , " & _
            "        COF_MOVIMIENTO.ID_SECTOR , COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        to_date('01/01/1990') FECHA_SALDO_INICIAL , (0) SALDO_INICIAL, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & (dIniDate - 1) & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) " & _
            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA, " & _
            "          COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

    sSQLStdAccHist = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, ID_CUENTA_ESTANDAR, " & _
                     "        NUMERO_CUENTA_ESTANDAR, NATURALEZA  " & _
                     " FROM COF_CUENTA_ESTANDAR_HIST " & _
                     " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
                     "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) " & _
                     " GROUP BY ID_CUENTA_ESTANDAR, NUMERO_CUENTA_ESTANDAR, NATURALEZA "

    sSQL_SI = " SELECT (0) id_movimiento, A.ID_MONEDA, A.ID_CUENTA, " & _
              "        A.ID_SECTOR , A.ID_CUENTA_AUXILIAR,   " & _
              "        to_date('01/01/1990') fecha_rev, ('0000') NUMERO_TRANSACCION, ('0000') CONCEPTO_TRANSACCION, " & _
              "        (decode(NATURALEZA,'D',(sum(A.SALDO_INICIAL)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
              "                               (sum(A.SALDO_INICIAL)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO_INICIAL, (0) Cargos, (0) Abonos  " & _
              " FROM ( " & _
                sSQL1 & _
              " UNION " & _
                sSQL2 & _
              " ) A  , COF_CUENTA B, (" & sSQLStdAccHist & ") C " & _
              " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
              "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
              " GROUP BY A.ID_MONEDA, A.ID_CUENTA, " & _
              "          A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NATURALEZA "

    sSql3 = " SELECT COF_MOVIMIENTO.id_movimiento, COF_MOVIMIENTO.ID_MONEDA, COF_MOVIMIENTO.ID_CUENTA, " & _
            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        FECHA_AFECTACION FECHA_REV, NUMERO_TRANSACCION, CONCEPTO_TRANSACCION, (0) SALDO_INICIAL, " & _
            "        (decode(TIPO_MOVIMIENTO,'D',Monto,0)) Cargos, " & _
            "        (decode(TIPO_MOVIMIENTO,'H',Monto,0)) Abonos " & _
            " FROM COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL virtual_SI  " & _
            " WHERE (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_MONEDA = virtual_SI.ID_MONEDA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = virtual_SI.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = virtual_SI.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = virtual_SI.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= '" & dIniDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION > nvl(virtual_SI.FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_TRANSACCION." & sGl & " ) "
            
    sSQLStdAccHist2 = SQLStandardAccountHist(dIniDate, dCuteDate, 1)
    
    
    sSQL_A = " SELECT id_movimiento, id_moneda, naturaleza, numero_cuenta_estandar cuenta, NVL(clave_sector,'00') clave_Sector, " & _
             "        NVL(nombre_sector,' ') n_sector, NVL(nombre_cuenta_estandar,' ') nombre_cuenta, " & _
             "        NVL(numero_cuenta_auxiliar,'0') clave_aux, NVL(nombre_cuenta_auxiliar,' ') nombre_aux, " & _
             "        NVL(fecha_rev,'01/01/1990') fecha_ult, NVL(numero_transaccion,'0') poliza, NVL(concepto_transaccion,' ') concepto, " & _
             "        Saldo_Inicial SI, Cargos, Abonos, NVL(id_tipo_cuenta,0) tipo_cuenta, NVL(rol_cuenta,' ') ROL "
             
    sSQL_B = " FROM ( " & _
                sSQL_SI & _
             " UNION " & _
               sSql3 & _
             " ) A, cof_cuenta, (" & sSQLStdAccHist2 & ") cof_cuenta_estandar, cof_sector, cof_cuenta_auxiliar " & _
             " WHERE ((A.id_cuenta = cof_cuenta.id_cuenta) AND " & _
             "        (cof_cuenta.id_cuenta_estandar = cof_cuenta_estandar.id_cuenta_estandar) AND " & _
             "        (A.id_sector = cof_sector.id_sector(+)) AND " & _
             "        (A.id_cuenta_auxiliar = cof_cuenta_auxiliar.id_cuenta_auxiliar(+))) AND " & _
             "        (Length(numero_cuenta_estandar) <= Length('" & sPatron & "')) AND " & _
             "        (cof_cuenta. " & sGl & ") " & sWhere & sWhereAux & _
             " ORDER BY ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
             "          CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, FECHA_rev, NUMERO_TRANSACCION  "

    sSQL = sSQL_A & sSQL_B
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset110 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset110", Err.Number
End Function

Public Function GeneraRecordset113(sAppServer As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sFromAccount As String = "", _
                        Optional sToAccount As String = "", _
                        Optional sFromSubsidiaryAcc As String = "", _
                        Optional sToSubsidiaryAcc As String = "", _
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

  Dim sSQL_B            As String

  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
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
        
    If (Len(sFromAccount) <> 0) Or (Len(sToAccount) <> 0) Then
      sWhere = " AND " & LikeSQLString("numero_cuenta_estandar", sFromAccount, sToAccount)
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
              "  ) A "
              
              '& _
              sOptionToDisplay    ' CHECAR OPCION DE MOVIMIENTOS<>0 O SALDOS<>0
            
      Set oComm.ActiveConnection = oConn
      oComm.CommandText = sSQL
      oComm.CommandType = adCmdText
      oComm.Execute
      
      sFromTable = " tmp_SALDOS "

    sSQL2 = " SELECT MAX(ID_CUENTA_ESTANDAR_HIST) ID_CUENTA_ESTANDAR_HIST, NUMERO_CUENTA_ESTANDAR, " & _
            "        MAX(NOMBRE_CUENTA_ESTANDAR) NOMBRE_CUENTA_ESTANDAR, MAX(NATURALEZA) NATURALEZA,  " & _
            "        MAX(ID_TIPO_CUENTA) ID_TIPO_CUENTA, MAX(ROL_CUENTA) ROL_CUENTA " & _
            " FROM COF_CUENTA_ESTANDAR_HIST " & _
            " WHERE (((TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_INICIO)) <= '" & dCuteDate & "')) AND " & _
            "        ((TRUNC(TO_DATE(FECHA_FIN)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(FECHA_FIN)) >= '" & dCuteDate & "'))) AND " & _
            "         (ID_TIPO_CUENTAS_STD = " & nTypeStdAcc & " )  " & _
            " GROUP BY NUMERO_CUENTA_ESTANDAR "
                
    sWhere = ""
    If (Len(sFromSubsidiaryAcc) <> 0) Or (Len(sToSubsidiaryAcc) <> 0) Then
      sWhere = " AND " & LikeSQLString("numero_cuenta_auxiliar", sFromSubsidiaryAcc, sToSubsidiaryAcc)
    End If
    
    If nExcRateCurrency = 0 Then
      If bWithAux Then
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR, " & _
               "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD ,  " & sFromTable & "  A , COF_SECTOR, " & _
               "       AO_CURRENCIES, COF_CUENTA_AUXILIAR " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR <> 0) " & sWhere & sOptionToDisplay & _
               " ORDER BY O_ID_MONEDA, NUMERO_CUENTA_AUXILIAR, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      Else
        sSQL = " select A.ID_MONEDA, O_ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, " & _
               "        NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR, " & _
               "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , (0) SALDO_PROMEDIO, " & _
               "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
               " FROM ( " & sSQL2 & " ) CUENTA_STD , " & sFromTable & " A , COF_SECTOR, " & _
               "       AO_CURRENCIES " & _
               " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
               "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) and " & _
               "       (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND " & _
               "       (A.ID_CUENTA_AUXILIAR = 0) " & sWhere & sOptionToDisplay & _
               " ORDER BY O_ID_MONEDA, NATURALEZA DESC, A.NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
      End If
    ElseIf nExcRateCurrency <> 0 Then
      If bWithAux Then
        sSQL_B = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        NOMBRE_CUENTA_ESTANDAR, " & _
                  "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                  "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD ,  " & sFromTable & "  A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
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
               " GROUP BY C.O_ID_MONEDA, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS "

        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & sSQL_B & " ) A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_TIPO_CUENTA, " & _
               "          A.ROL_CUENTA, A.FROM_MOVS " & _
               " ORDER BY O_ID_MONEDA, NUMERO_CUENTA_AUXILIAR, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
               
        If (cnFromCurrency = nExcRateCurrency) Then
          sSQL = sSQL_B & _
                 " ORDER BY O_ID_MONEDA, NUMERO_CUENTA_AUXILIAR, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR "
        End If
      Else
        sSQL_B = " SELECT (1) id_moneda, C.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL, SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM (" & _
                  " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                  "        NOMBRE_CUENTA_ESTANDAR, " & _
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
                "       (NVL(B.FROM_CURRENCY_ID, " & cnFromCurrency & " ) = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & sOptionToDisplay & _
                " GROUP BY C.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_TIPO_CUENTA, " & _
                "          A.ROL_CUENTA, A.FROM_MOVS "

        sSQL = " SELECT (" & nExcRateCurrency & ") id_moneda, A.O_ID_MONEDA, A.NATURALEZA, " & _
               "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, (0) SALDO_PROMEDIO, " & _
               "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
               " FROM  (" & sSQL_B & ") A , AO_EXCHANGE_RATES B " & _
               " WHERE (A.ID_MONEDA = B.FROM_CURRENCY_ID(+)) and " & _
               "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
               "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
               "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
               "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
               " GROUP BY A.O_ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA_ESTANDAR, A.ID_TIPO_CUENTA, " & _
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
    Set GeneraRecordset113 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset113", Err.Number
End Function


Public Function GeneraRecordset115(sAppServer As String, sParticipantType As String, dDate As Date, _
                                   sSQLOrder As String, sStatus As String) As Recordset
  Dim oRs            As New Recordset
  Dim oCon           As New Connection
  Dim sSQL           As String
  Dim sSQLOrderTemp  As String
  Dim sStatusTemp    As String
 '****************************************************
  On Error GoTo ErrHandler
    oCon.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oCon.Open

    If sSQLOrder = "" Then
      sSQLOrderTemp = " ORDER BY PARTICIPANTKEY "
    Else
      sSQLOrderTemp = " ORDER BY " & sSQLOrder
    End If
  
    If sStatus = "" Then
      sStatusTemp = ""
    Else
      sStatusTemp = " (Status = '" & sStatus & "') AND "
    End If
  
    sSQL = " SELECT * " & _
           " FROM MHPARTICIPANTS " & _
           " WHERE " & sStatus & _
           "       (TRUNC(TO_DATE(FROMDATE)) <= '" & dDate & "') AND " & _
           "       (TRUNC(TO_DATE(TODATE)) >= '" & dDate & "') AND " & _
           "       (ParticipantType = '" & sParticipantType & "') " & _
             sSQLOrderTemp


    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oCon, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    
    Set GeneraRecordset115 = oRs
  Exit Function
ErrHandler:
    RaiseError ClassId, "GeneraRecordset115", Err.Number
End Function

Public Function GeneraRecordset116(sAppServer As String, nStdAccountTypeId As Long, dDate As Date, _
                          sFromAccount As String, sToAccount As String) As Recordset
  Dim oRs         As New ADODB.Recordset
  Dim oCon        As New ADODB.Connection
  Dim sSQL        As String
  Dim sSQL_A      As String
  Dim sSQL_B      As String
  Dim sWhere      As String
  '****************************************************
  On Error GoTo ErrHandler
    oCon.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oCon.Open
    
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
    Set GeneraRecordset116 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset116", Err.Number
End Function

Public Function GeneraRecordset117(sAppServer As String, nStdAccountTypeId As Long, _
                                  sSQLOrder As String) As Recordset
  Dim oRs               As New Recordset
  Dim oCon              As New Connection
  Dim sSQL              As String
  Dim sSQLOrderTemp     As String
 '****************************************************
 On Error GoTo ErrHandler
   oCon.ConnectionString = GetConstant(sConstantName:=sAppServer)
   oCon.Open
   
   If (sSQLOrder = "") Then
     sSQLOrderTemp = " ORDER BY NOMBRE_GRUPO_MAYOR, NUMERO_MAYOR "
   Else
     sSQLOrderTemp = " ORDER BY NUMERO_MAYOR "
   End If
   
   sSQL = " SELECT COF_MAYOR.NUMERO_MAYOR, COF_MAYOR.NOMBRE_MAYOR, " & _
          "        COF_GRUPO_MAYOR.NOMBRE_GRUPO_MAYOR " & _
          " FROM COF_GRUPO_MAYOR, COF_ELEMENTO_GRUPO_MAYOR, COF_MAYOR " & _
          " WHERE (COF_MAYOR.ID_MAYOR = COF_ELEMENTO_GRUPO_MAYOR.ID_MAYOR(+)) AND " & _
          "       (COF_ELEMENTO_GRUPO_MAYOR.ID_GRUPO_MAYOR = COF_GRUPO_MAYOR.ID_GRUPO_MAYOR(+)) AND " & _
          "       (COF_MAYOR.ID_TIPO_CUENTAS_STD = " & nStdAccountTypeId & ") AND " & _
          "       (COF_GRUPO_MAYOR.CLIP_GRUPO_MAYOR = 2) " & _
          sSQLOrderTemp
   
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oCon, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset117 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset117", Err.Number
End Function

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
      
      sCondition = "((TIPO_MOVIMIENTO= '" & sMovKind & "' ) " & _
                   " AND (FECHA_AFECTACION >= '" & dIniDate & "')" & _
                   " And (FECHA_AFECTACION <= '" & dEndDate & "'))  " & _
                   " and ( fecha_afectacion >= fecha_saldo_inicial)"
                   
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

Public Function GetExchangeRate(sAppServer As String, nToCurrency As Long, _
                                nExcRateType As Long, nExcRateCurrency As Long, _
                                dExcDate As Date) As Double
  Dim oConn    As New ADODB.Connection
  Dim oRs      As New ADODB.Recordset
  Dim sSQL     As String
  '************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    sSQL = " SELECT EXCHANGE_RATE FROM AO_EXCHANGE_RATES WHERE " & _
           " (FROM_CURRENCY_ID = " & nExcRateCurrency & ") AND " & _
           " (TO_CURRENCY_ID = " & nToCurrency & ") AND " & _
           " (EXCHANGE_RATE_TYPE_ID = " & nExcRateType & ") AND " & _
           " (FROM_DATE <= '" & dExcDate & "') AND " & _
           " (TO_DATE >= '" & dExcDate & "') "
           
    With oRs
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      If (.BOF And .EOF) Then
        GetExchangeRate = 1
      Else
        GetExchangeRate = !EXCHANGE_RATE
      End If
      .Close
      Set .ActiveConnection = Nothing
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetExchangeRate", Err.Number
End Function

Public Function GetObjName(sAppServer As String, nId As Long, sTable As String, _
                            sFieldName As String, sFieldCondition As String, _
                            Optional sFieldName2 As String = "") As String
  Dim oConn       As New ADODB.Connection
  Dim oRs         As New ADODB.Recordset
  Dim sObjName    As String
  Dim sSQL        As String
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
      '.CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      If (.BOF And .EOF) Then
        sObjName = " "
      Else
        If (Len(sFieldName2) = 0) Then
          sObjName = .Fields(0)
        ElseIf (Len(sFieldName2) <> 0) Then
          sObjName = ZeroFill(.Fields(0), 4) & "  " & .Fields(1)
        End If
      End If
      .Close
      Set .ActiveConnection = Nothing
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
      '.CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      If (.BOF And .EOF) Then
        sObjName = " "
      Else
        sObjName = .Fields(0)
      End If
      .Close
      Set .ActiveConnection = Nothing
    End With
    GetAccountName = sObjName
  Exit Function
ErrHandler:
  RaiseError ClassId, "GetAccountName", Err.Number
End Function

Public Function GetTransactionRS(sAppServer As String, nTransactionId As Long) As Recordset
  Dim oRecordset As New Recordset
  Dim sSQL As String
  '******************************************************************
  On Error GoTo ErrorHandler
    sSQL = "SELECT COF_Transaccion.*, nvl(COF_mayor.nombre_mayor,' ') nombre_mayor, " & _
           " nvl(COF_mayor.numero_mayor,' ') numero_mayor, " & _
           " COF_mayor.id_moneda_base, nvl(COF_fuente.nombre_fuente,' ') nombre_fuente" & _
           " FROM COF_Transaccion, COF_mayor, COF_fuente " & _
           " WHERE (COF_Transaccion.id_mayor = COF_mayor.id_mayor) AND " & _
           " (COF_Transaccion.id_fuente = COF_fuente.id_fuente) AND " & _
           " (id_transaccion = " & nTransactionId & ")"
    With oRecordset
      .ActiveConnection = GetConstant(sConstantName:=sAppServer)
      .CursorLocation = adUseClient
      .Open sSQL, , adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GetTransactionRS = oRecordset
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetTransactionRS", Err.Number
End Function

Public Function GetPostingsRS(sAppServer As String, nTransactionId As Long) As ADODB.Recordset
  Dim oRecordset As New Recordset
  Dim sSQL As String
  '*********************************************************
  On Error GoTo ErrorHandler
     sSQL = "SELECT COF_movimiento.id_movimiento, COF_movimiento.id_transaccion," & _
             " COF_movimiento.id_cuenta, COF_movimiento.id_cuenta_auxiliar, COF_movimiento.id_sector, " & _
             " COF_movimiento.id_area_responsabilidad, COF_movimiento.id_clave_presupuestal, COF_movimiento.tipo_movimiento, " & _
             " COF_movimiento.fecha_movimiento, nvl(COF_movimiento.concepto_movimiento,' ')as concepto_movimiento, " & _
             " COF_movimiento.id_moneda, COF_movimiento.monto, COF_movimiento.monto_moneda_base, COF_cuenta_estandar.numero_cuenta_estandar, " & _
             " COF_cuenta_estandar.nombre_cuenta_estandar, nvl(COF_cuenta_auxiliar.numero_cuenta_auxiliar,' ') as numero_cuenta_auxiliar, " & _
             " nvl(COF_cuenta_auxiliar.nombre_cuenta_auxiliar, ' ') nombre_cuenta_auxiliar, nvl(COF_sector.clave_sector,' ') clave_sector, " & _
             " nvl(COF_sector.nombre_sector,' ') nombre_sector, " & _
             " AO_Currencies.abbrev, nvl(COF_estr_presupuestal.clave_presupuestal,' ') clave_presupuestal, " & _
             " nvl(COF_area_responsabilidad.clave_area_responsabilidad, ' ') clave_area_responsabilidad, AO_Currencies.o_id_moneda " & _
           " FROM COF_movimiento, COF_cuenta, COF_cuenta_estandar, COF_cuenta_auxiliar, COF_sector, AO_Currencies, COF_estr_presupuestal, COF_area_responsabilidad " & _
           " WHERE (COF_movimiento.id_cuenta = COF_cuenta.id_cuenta) AND" & _
             " (COF_cuenta.id_cuenta_estandar = COF_cuenta_estandar.id_cuenta_estandar) AND " & _
             " (COF_movimiento.id_cuenta_auxiliar = COF_cuenta_auxiliar.id_cuenta_auxiliar (+)) AND " & _
             " (COF_movimiento.id_sector = COF_sector.id_sector (+) ) AND " & _
             " (COF_movimiento.id_moneda = AO_Currencies.currency_id) AND " & _
             " (COF_movimiento.id_area_responsabilidad = COF_area_responsabilidad.id_area_responsabilidad (+) ) AND " & _
             " (COF_movimiento.id_clave_presupuestal = COF_estr_presupuestal.id_clave_presupuestal (+) ) AND " & _
             " (COF_movimiento.id_transaccion = " & nTransactionId & ") " & _
           " ORDER BY COF_movimiento.id_movimiento"
    oRecordset.Open sSQL, GetConstant(sConstantName:=sAppServer), adOpenForwardOnly, adLockReadOnly
    oRecordset.ActiveConnection = Nothing
    Set GetPostingsRS = oRecordset
  Exit Function
ErrorHandler:
  RaiseError ClassId, "GetPostingsRS", Err.Number
End Function

Public Function GeneraRSNames(sAppServer As String, vGL As Variant)
  Dim oRecordset      As New ADODB.Recordset
  Dim oConn           As New ADODB.Connection
    
  Dim sSqlComp        As String
  Dim sGl             As String
  Dim sGLCondition    As String
  Dim nGl             As Long
  Dim nStdAccountTypeId As Long
'******************************************************************
 On Error GoTo ErrHandler
  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
  oConn.Open
    
    If (IsArray(vGL)) Then
      nGl = CLng(vGL(0))
    Else
      nGl = CLng(vGL)
    End If
    nStdAccountTypeId = TypeStdAccount(sAppServer, nGl)
    sSqlComp = " SELECT DISTINCT(NUMERO_CUENTA_ESTANDAR) NUM_CUENTA, " & _
               "        NOMBRE_CUENTA_ESTANDAR NOM_CUENTA " & _
               " From COF_CUENTA_ESTANDAR_HIST " & _
               " WHERE (ID_TIPO_CUENTAS_STD = " & nStdAccountTypeId & ") " & _
               " ORDER BY NUMERO_CUENTA_ESTANDAR "
      
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

Public Function GenerateRSTransactions(sAppServer As String, vGL As Variant, _
                                        dInitDate As Date, dEndDate As Date, _
                                        Optional bOpen As Boolean = False, _
                                        Optional nTransKind As Long = 0) As ADODB.Recordset
  Dim oConn                As New ADODB.Connection
  Dim oRs                  As New ADODB.Recordset

  Dim sGl                  As String
  Dim sSQL                 As String
  Dim sCondition           As String
  Dim sGLCondition         As String
  Dim sTransCondition      As String
  Dim sMovsTable           As String
  '*******************************************************************************************
  On Error GoTo ErrHandler
    sCondition = " AND (COF_TRANSACCION.fecha_afectacion >= '" & dInitDate & "') AND " & _
                 " (COF_TRANSACCION.fecha_afectacion <= '" & dEndDate & "')"
    If (nTransKind <> 0) Then
      sTransCondition = " And (ID_TIPO_TRANSACCION = " & nTransKind & ") "
    ElseIf (nTransKind = 0) Then
      sTransCondition = " "
    End If
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGLCondition = "id_mayor = " & sGl
    End If
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    If (bOpen) Then
      sSQL = " SELECT A.FECHA_AFECTACION, A.FECHA_REGISTRO, A.ID_MAYOR, NOMBRE_MAYOR, A.ID_TRANSACCION,  A.NUMERO_TRANSACCION, A.CONCEPTO_TRANSACCION, A.PARTICIPANT_NAME, A.ESTA_ABIERTA, " & _
             "        sum(A.CARGOS) as Cargos, sum(A.ABONOS) as Abonos " & _
             " FROM  ( " & _
             "       SELECT FECHA_AFECTACION, FECHA_REGISTRO, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' ') NUMERO_TRANSACCION, " & _
             "              COF_TRANSACCION.CONCEPTO_TRANSACCION, nvl(MH_PARTICIPANTS.PARTICIPANT_NAME,' ') PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, " & _
             "              (0) as CARGOS , sum(COF_MOVIMIENTO_TMP.MONTO_MONEDA_BASE) AS ABONOS " & _
             "       FROM   COF_TRANSACCION, COF_MOVIMIENTO_TMP, MH_PARTICIPANTS " & _
             "       WHERE  (COF_TRANSACCION.ID_TRANSACCION = COF_MOVIMIENTO_TMP.ID_TRANSACCION ) and " & _
             "              (COF_TRANSACCION.ID_ELABORADA_POR = MH_PARTICIPANTS.PARTICIPANT_ID(+)) and   " & _
             "              (COF_MOVIMIENTO_TMP.TIPO_MOVIMIENTO = 'H') And (COF_TRANSACCION." & sGLCondition & ") " & sCondition & sTransCondition & _
             "       GROUP BY FECHA_AFECTACION, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' '), " & _
             "               COF_TRANSACCION.CONCEPTO_TRANSACCION , PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, FECHA_REGISTRO " & _
             " Union " & _
             "       SELECT FECHA_AFECTACION, FECHA_REGISTRO, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' ') NUMERO_TRANSACCION, " & _
             "              COF_TRANSACCION.CONCEPTO_TRANSACCION, nvl(MH_PARTICIPANTS.PARTICIPANT_NAME,' ') PARTICIPANT_NAME, " & _
             "              COF_TRANSACCION.ESTA_ABIERTA, sum(COF_MOVIMIENTO_TMP.MONTO_MONEDA_BASE) AS CARGOS, " & _
             "              (0) as ABONOS " & _
             "       FROM   COF_TRANSACCION, COF_MOVIMIENTO_TMP, MH_PARTICIPANTS " & _
             "       WHERE  (COF_TRANSACCION.ID_TRANSACCION = COF_MOVIMIENTO_TMP.ID_TRANSACCION ) and " & _
             "              (COF_TRANSACCION.ID_ELABORADA_POR = MH_PARTICIPANTS.PARTICIPANT_ID(+)) and   " & _
             "              (COF_MOVIMIENTO_TMP.TIPO_MOVIMIENTO = 'D') And (COF_TRANSACCION." & sGLCondition & ") " & sCondition & sTransCondition & _
             "       GROUP BY FECHA_AFECTACION, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' '), " & _
             "               COF_TRANSACCION.CONCEPTO_TRANSACCION , PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, FECHA_REGISTRO " & _
             " ) A , COF_MAYOR  WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR) " & _
             " GROUP BY A.FECHA_AFECTACION, A.ID_MAYOR, A.ID_TRANSACCION, A.NUMERO_TRANSACCION, A.CONCEPTO_TRANSACCION, A.PARTICIPANT_NAME, A.ESTA_ABIERTA, A.FECHA_REGISTRO, NOMBRE_MAYOR "
    ElseIf Not (bOpen) Then
      sSQL = " SELECT A.FECHA_AFECTACION, A.FECHA_REGISTRO, A.ID_MAYOR, NOMBRE_MAYOR, A.ID_TRANSACCION,  A.NUMERO_TRANSACCION, A.CONCEPTO_TRANSACCION, A.PARTICIPANT_NAME, A.ESTA_ABIERTA, " & _
             "        sum(A.CARGOS) as Cargos, sum(A.ABONOS) as Abonos " & _
             " FROM  ( " & _
             "       SELECT FECHA_AFECTACION, FECHA_REGISTRO, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' ') NUMERO_TRANSACCION, " & _
             "              COF_TRANSACCION.CONCEPTO_TRANSACCION, nvl(MH_PARTICIPANTS.PARTICIPANT_NAME,' ') PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, " & _
             "              (0) as CARGOS , sum(COF_MOVIMIENTO.MONTO_MONEDA_BASE) AS ABONOS " & _
             "       FROM   COF_TRANSACCION, COF_MOVIMIENTO, MH_PARTICIPANTS " & _
             "       WHERE  (COF_TRANSACCION.ID_TRANSACCION = COF_MOVIMIENTO.ID_TRANSACCION ) and " & _
             "              (COF_TRANSACCION.ID_ELABORADA_POR = MH_PARTICIPANTS.PARTICIPANT_ID(+)) and   " & _
             "              (COF_MOVIMIENTO.TIPO_MOVIMIENTO = 'H') And (COF_TRANSACCION." & sGLCondition & ") " & sCondition & sTransCondition & _
             "       GROUP BY FECHA_AFECTACION, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' '), " & _
             "               COF_TRANSACCION.CONCEPTO_TRANSACCION , PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, FECHA_REGISTRO " & _
             " Union " & _
             "       SELECT FECHA_AFECTACION, FECHA_REGISTRO, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' ') NUMERO_TRANSACCION, " & _
             "              COF_TRANSACCION.CONCEPTO_TRANSACCION, nvl(MH_PARTICIPANTS.PARTICIPANT_NAME,' ') PARTICIPANT_NAME, " & _
             "              COF_TRANSACCION.ESTA_ABIERTA, sum(COF_MOVIMIENTO.MONTO_MONEDA_BASE) AS CARGOS, " & _
             "              (0) as ABONOS " & _
             "       FROM   COF_TRANSACCION, COF_MOVIMIENTO, MH_PARTICIPANTS " & _
             "       WHERE  (COF_TRANSACCION.ID_TRANSACCION = COF_MOVIMIENTO.ID_TRANSACCION ) and " & _
             "              (COF_TRANSACCION.ID_ELABORADA_POR = MH_PARTICIPANTS.PARTICIPANT_ID(+)) and   " & _
             "              (COF_MOVIMIENTO.TIPO_MOVIMIENTO = 'D') And (COF_TRANSACCION." & sGLCondition & ") " & sCondition & sTransCondition & _
             "       GROUP BY FECHA_AFECTACION, ID_MAYOR, COF_TRANSACCION.ID_TRANSACCION, NVL(COF_TRANSACCION.NUMERO_TRANSACCION,' '), " & _
             "               COF_TRANSACCION.CONCEPTO_TRANSACCION , PARTICIPANT_NAME, COF_TRANSACCION.ESTA_ABIERTA, FECHA_REGISTRO " & _
             " ) A, COF_MAYOR  WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR) " & _
             " GROUP BY A.FECHA_AFECTACION, A.ID_MAYOR, A.ID_TRANSACCION, A.NUMERO_TRANSACCION, A.CONCEPTO_TRANSACCION, A.PARTICIPANT_NAME, A.ESTA_ABIERTA, A.FECHA_REGISTRO, NOMBRE_MAYOR "
    End If
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GenerateRSTransactions = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordsetTransactions", Err.Number
End Function

Public Function SQLCreaCamposAux(sAppServer As String, sAttributesCategory As String, _
                                 Optional dIniDate As Date, Optional dEndDate As Date, Optional bWithDir As Boolean = False) As String
  Dim oConn     As New Connection
  Dim oRs       As New Recordset
  Dim sSQL1     As String
  Dim sSql1A    As String
  Dim sSqlCom   As String
  
  Dim sSQLIni       As String
  Dim sSQL          As String

  Dim sSQLCampos    As String
  Dim sSQLUnion     As String
  Dim nPos          As Long
  Dim nReg          As Long
  Dim nCount        As Long
  Dim sSQLSelec     As Long
  Dim sCampo        As String
  Dim sNombreCampo  As String
  
  '************************************************************************************
  
  On Error GoTo ErrHandler
  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
  oConn.Open
  
  sSQLIni = " SELECT NameSpace, DataType " & _
            " FROM AOENTITYATTRDEF, AO_OBJECTS " & _
            " WHERE (AOENTITYATTRDEF.ObjectId = AO_OBJECTS.Object_Id) AND " & _
            "       (AOENTITYATTRDEF.ENTITYID = 11) and (OBJECT_NAME = '" & sAttributesCategory & "')"
  
    With oRs
      .CursorLocation = adUseClient
      .Open sSQLIni, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
      
      If Not (.BOF And .EOF) Then
        nReg = .RecordCount
        Do While Not .EOF
          sSQLCampos = sSQLCampos & ", Max(" & !namespace & ") " & !namespace
          .MoveNext
        Loop

        .MoveFirst
        nPos = .AbsolutePosition
        For nCount = 1 To nReg
          Do While Not .EOF
             sNombreCampo = !namespace
            If (.AbsolutePosition = nPos) Then
              If sSQLUnion <> "" Then
                sSQLUnion = sSQLUnion & " UNION "
              End If
              sCampo = !namespace
              
              If (!DataType = 1) Or (!DataType = 2) Then
                sSQL = sSQL & ", ValString " & sNombreCampo
              ElseIf (!DataType = 3) Then
                sSQL = sSQL & ", ValDate " & sNombreCampo
              ElseIf (!DataType = 5) Or (!DataType = 6) Or (!DataType = 7) Or (!DataType = 8) Or (!DataType = 9) Then
                sSQL = sSQL & ", ValNumeric " & sNombreCampo
              End If
            Else
              If (!DataType = 1) Or (!DataType = 2) Then
                sSQL = sSQL & ", (' ') " & sNombreCampo
              ElseIf (!DataType = 3) Then
                sSQL = sSQL & ", to_date('') " & sNombreCampo
              ElseIf (!DataType = 5) Or (!DataType = 6) Or (!DataType = 7) Or (!DataType = 8) Or (!DataType = 9) Then
                sSQL = sSQL & ", (0) " & sNombreCampo
              End If
            End If
            .MoveNext
            If .EOF Then
              sSQLUnion = sSQLUnion & _
                          " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR " & sSQL & _
                          " From AOEntityAttrDef, COF_CUENTA_AUXILIAR_ATTRS, COF_CUENTA_AUXILIAR " & _
                          " WHERE (AOEntityAttrDef.EntityAttrDefId = COF_CUENTA_AUXILIAR_ATTRS.EntityAttrDefId) AND " & _
                          "       (COF_CUENTA_AUXILIAR_ATTRS.Id_Cuenta_Auxiliar = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
                          "       (AOEntityAttrDef.Namespace = '" & sCampo & "') AND " & _
                          "       (((TRUNC(TO_DATE(FromDate)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FromDate)) <= '" & dEndDate & "')) AND " & _
                          "        ((TRUNC(TO_DATE(ToDate)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(ToDate)) >= '" & dEndDate & "'))) "
  
              sSQL = ""
            End If
          Loop
          .MoveFirst
          nPos = nPos + 1
        Next nCount
      End If
    End With
    
    If Len(sSQLUnion) <> 0 Then
      sSQL1 = " SELECT A.ID_CUENTA_AUXILIAR " & sSQLCampos & _
              " FROM ( " & sSQLUnion & " ) A " & _
              " GROUP BY A.ID_CUENTA_AUXILIAR "
    
    sSqlCom = " SELECT A.*, B.NUMERO_CUENTA_AUXILIAR, B.NOMBRE_CUENTA_AUXILIAR, B.DESCRIPCION, " & _
              "        D.OBJECT_NAME, C.ID_MAYOR " & _
              " FROM ( " & sSQL1 & " ) A, COF_CUENTA_AUXILIAR B, COF_MAYOR_AUXILIAR C, AO_OBJECTS D " & _
              " WHERE (B.ID_CUENTA_AUXILIAR = A.ID_CUENTA_AUXILIAR(+)) AND " & _
              "       (B.ID_MAYOR_AUXILIAR = C.ID_MAYOR_AUXILIAR) AND " & _
              "       (C.ID_TIPO_MAYOR_AUXILIAR = D.OBJECT_ID) "
    Else
      sSQL1 = ""
      sSqlCom = " SELECT B.ID_CUENTA_AUXILIAR, B.NUMERO_CUENTA_AUXILIAR, B.NOMBRE_CUENTA_AUXILIAR, B.DESCRIPCION, " & _
          "        D.OBJECT_NAME, C.ID_MAYOR " & _
          " FROM  COF_CUENTA_AUXILIAR B, COF_MAYOR_AUXILIAR C, AO_OBJECTS D " & _
          " WHERE (B.ID_MAYOR_AUXILIAR = C.ID_MAYOR_AUXILIAR) AND " & _
          "       (C.ID_TIPO_MAYOR_AUXILIAR = D.OBJECT_ID) "

    End If
  
      
  'sSqlCom = ssql1 & sSql1A
  SQLCreaCamposAux = sSqlCom
Exit Function
ErrHandler:
  RaiseError ClassId, "CreaSQLCamposAux", Err.Number
End Function


'''''Public Function SQLCreaCamposAux(sAppServer As String, sAttributesCategory As String, _
'''''                                Optional dIniDate As Date, Optional dEndDate As Date, Optional bWithDir As Boolean = False) As String
'''''  Dim oConn     As New Connection
'''''  Dim oRs       As New Recordset
'''''  Dim ssql1     As String
'''''  Dim sSql1A    As String
'''''  Dim sSqlCom   As String
'''''
'''''  Dim sSQLIni       As String
'''''  Dim sSQL          As String
'''''
'''''  Dim sSQLCampos    As String
'''''  Dim sSQLUnion     As String
'''''  Dim nPos          As Long
'''''  Dim nReg          As Long
'''''  Dim nCount        As Long
'''''  Dim sSQLSelec     As Long
'''''  Dim sCampo        As String
'''''  Dim sNombreCampo  As String
'''''
'''''  '************************************************************************************
'''''
'''''  On Error GoTo ErrHandler
'''''  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
'''''  oConn.Open
'''''
'''''  sSQLIni = " SELECT NameSpace, DataType " & _
'''''            " FROM AOENTITYATTRDEF, AO_OBJECTS " & _
'''''            " WHERE (AOENTITYATTRDEF.ObjectId = AO_OBJECTS.Object_Id) AND " & _
'''''            "       (AOENTITYATTRDEF.ENTITYID = 11) and (OBJECT_NAME = '" & sAttributesCategory & "')"
'''''
'''''    With oRs
'''''      .CursorLocation = adUseClient
'''''      .Open sSQLIni, oConn, adOpenStatic, adLockReadOnly
'''''      Set .ActiveConnection = Nothing
'''''
'''''      If Not (.BOF And .EOF) Then
'''''        nReg = .RecordCount
'''''        Do While Not .EOF
'''''
'''''''''          sSQLCampos = sSQLCampos & ", Max(" & Mid(!namespace, InStr(!namespace, ".") + 1, Len(!namespace) - InStr(!namespace, ".")) & ") " & _
'''''''''                                    Mid(!namespace, InStr(!namespace, ".") + 1, Len(!namespace) - InStr(!namespace, "."))
'''''
'''''          sSQLCampos = sSQLCampos & ", Max(" & !namespace & ") " & !namespace
'''''
'''''          .MoveNext
'''''        Loop
'''''      End If
'''''      .MoveFirst
'''''      nPos = .AbsolutePosition
'''''      For nCount = 1 To nReg
'''''        Do While Not .EOF
'''''           'sNombreCampo = Mid(!namespace, InStr(!namespace, ".") + 1, Len(!namespace) - InStr(!namespace, "."))
'''''           sNombreCampo = !namespace
'''''          If (.AbsolutePosition = nPos) Then
'''''            If sSQLUnion <> "" Then
'''''              sSQLUnion = sSQLUnion & " UNION "
'''''            End If
'''''            sCampo = !namespace
'''''
'''''            If (!DataType = 1) Or (!DataType = 2) Then
'''''              sSQL = sSQL & ", ValString " & sNombreCampo
'''''            ElseIf (!DataType = 3) Then
'''''              sSQL = sSQL & ", ValDate " & sNombreCampo
'''''            ElseIf (!DataType = 5) Or (!DataType = 6) Or (!DataType = 7) Or (!DataType = 8) Or (!DataType = 9) Then
'''''              sSQL = sSQL & ", ValNumeric " & sNombreCampo
'''''            End If
'''''          Else
'''''            If (!DataType = 1) Or (!DataType = 2) Then
'''''              sSQL = sSQL & ", (' ') " & sNombreCampo
'''''            ElseIf (!DataType = 3) Then
'''''              sSQL = sSQL & ", to_date('') " & sNombreCampo
'''''            ElseIf (!DataType = 5) Or (!DataType = 6) Or (!DataType = 7) Or (!DataType = 8) Or (!DataType = 9) Then
'''''              sSQL = sSQL & ", (0) " & sNombreCampo
'''''            End If
'''''          End If
'''''          .MoveNext
'''''          If .EOF Then
'''''            sSQLUnion = sSQLUnion & _
'''''                        " SELECT COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR, " & _
'''''                        "        COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR, COF_CUENTA_AUXILIAR.NOMBRE_CUENTA_AUXILIAR, COF_CUENTA_AUXILIAR.DESCRIPCION, " & _
'''''                        "        AO_OBJECTS.OBJECT_NAME, COF_MAYOR_AUXILIAR.ID_MAYOR " & sSQL & _
'''''                        " From COF_MAYOR_AUXILIAR, AO_OBJECTS, AOEntityAttrDef, COF_CUENTA_AUXILIAR_ATTRS, COF_CUENTA_AUXILIAR " & _
'''''                        " WHERE (COF_MAYOR_AUXILIAR.ID_TIPO_MAYOR_AUXILIAR = AO_OBJECTS.OBJECT_ID) AND " & _
'''''                        "       (AOEntityAttrDef.EntityAttrDefId = COF_CUENTA_AUXILIAR_ATTRS.EntityAttrDefId) AND " & _
'''''                        "       (COF_CUENTA_AUXILIAR_ATTRS.Id_Cuenta_Auxiliar = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND " & _
'''''                        "       (COF_MAYOR_AUXILIAR.ID_MAYOR_AUXILIAR = COF_CUENTA_AUXILIAR.ID_MAYOR_AUXILIAR) AND " & _
'''''                        "       (AOEntityAttrDef.Namespace = '" & sCampo & "') AND " & _
'''''                        "       (((TRUNC(TO_DATE(FromDate)) <= '" & dIniDate & "') OR (TRUNC(TO_DATE(FromDate)) <= '" & dEndDate & "')) AND " & _
'''''                       "        ((TRUNC(TO_DATE(ToDate)) >= '" & dIniDate & "') OR (TRUNC(TO_DATE(ToDate)) >= '" & dEndDate & "'))) "
'''''
'''''            sSQL = ""
'''''          End If
'''''        Loop
'''''        .MoveFirst
'''''        nPos = nPos + 1
'''''      Next nCount
'''''    End With
'''''
'''''''    SELECT COF_CUENTA_AUXILIAR_ATTRS.ValString AS Domicilio,
'''''''           COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,
'''''''           COF_MAYOR_AUXILIAR.ID_MAYOR
'''''''    From COF_MAYOR_AUXILIAR, AO_OBJECTS, AOEntityAttrDef, COF_CUENTA_AUXILIAR_ATTRS, COF_CUENTA_AUXILIAR
'''''''    WHERE (COF_MAYOR_AUXILIAR.ID_TIPO_MAYOR_AUXILIAR = AO_OBJECTS.OBJECT_ID) AND
'''''''          (AOEntityAttrDef.EntityAttrDefId = COF_CUENTA_AUXILIAR_ATTRS.EntityAttrDefId) AND
'''''''          (COF_CUENTA_AUXILIAR_ATTRS.Id_Cuenta_Auxiliar = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR) AND
'''''''          (COF_MAYOR_AUXILIAR.ID_MAYOR_AUXILIAR = COF_CUENTA_AUXILIAR.ID_MAYOR_AUXILIAR) AND
'''''''          (COF_MAYOR_AUXILIAR.ID_MAYOR = 9) AND
'''''''          (AO_OBJECTS.OBJECT_NAME = 'Administrativos') AND
'''''''          (AOEntityAttrDef.Namespace = 'Domicilio')
'''''
'''''    ssql1 = " SELECT A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.DESCRIPCION, " & _
'''''            "        A.OBJECT_NAME, A.ID_MAYOR " & sSQLCampos
'''''
'''''    sSql1A = " FROM ( " & sSQLUnion & _
'''''             " ) A " & _
'''''             " GROUP BY A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.DESCRIPCION, A.OBJECT_NAME, A.ID_MAYOR "
'''''
'''''  sSqlCom = ssql1 & sSql1A
'''''  SQLCreaCamposAux = sSqlCom
'''''Exit Function
'''''ErrHandler:
'''''  RaiseError ClassId, "CreaSQLCamposAux", Err.Number
'''''End Function

'***************************************************************
'  Utils
'****************************************************************

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

'**************************************************************************************
'
'**************************************************************************************

Public Function GeneraRecordset74(sAppServer As String, sPatron As String, vGL As Variant, _
                                   dIniDate As Date, dCuteDate As Date, _
                                   dIniDate2 As Date, dCuteDate2 As Date, _
                                   Optional sIniAccount As String = "", _
                                   Optional sEndAccount As String = "", _
                                   Optional nExcRateType As Long = 0, _
                                    Optional nExcRateCurrency As Long = 0, _
                                   Optional dExcDate As Date) As Recordset
                                 
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
  Dim sSelect           As String
  Dim sInitSelect       As String
  Dim sMaxSelect        As String
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
  Dim sAAC              As String
  Dim sAAC_B            As String
  
  Dim sAAN              As String
  Dim sAAN_B            As String
  
  Dim sCAC              As String
  Dim sCAC_B            As String
  
  Dim sCAN              As String
  Dim sCAN_B            As String
  
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
    
    If (sIniAccount = "") And (sEndAccount = "") Then
      sWhere = " "
    ElseIf (sIniAccount <> "") And (sEndAccount <> "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR BETWEEN '" & sIniAccount & "' AND '" & sEndAccount & "') "
    ElseIf (sIniAccount <> "") And (sEndAccount = "") Then
      sWhere = " AND (NUMERO_CUENTA_ESTANDAR <= '" & sIniAccount & "') "
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
                       

    sSQL = "  SELECT A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
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
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A "
            
    sSQL_B = " SELECT A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, " & _
              "        A.SALDO_INICIAL, A.CARGOS, A.ABONOS, A.SALDO_ACTUAL " & _
              " FROM ( " & _
                " SELECT A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR,   " & _
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
                " GROUP BY A.ID_MONEDA, A.ID_MAYOR, NUMERO_CUENTA_ESTANDAR, " & _
                "          A.ID_SECTOR , A.ID_CUENTA_AUXILIAR, NATURALEZA " & _
              "  ) A "
              
    
    
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
            

'activo primer periodo
      sAAN = " SELECT  A.ID_MAYOR, " & _
             "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL*(-1)),NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL)) SALDO_INICIAL, " & _
             "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.CARGOS*(-1)),NVL(EXCHANGE_RATE,1)*A.CARGOS)) CARGOS, " & _
             "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.ABONOS*(-1)),NVL(EXCHANGE_RATE,1)*A.ABONOS)) ABONOS, " & _
             "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL*(-1)),NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL)) SALDO_ACTUAL, " & _
             "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
             "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
             "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
             " FROM (" & _
                  " select A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) AND " & _
                  "       ((ID_TIPO_CUENTA = 3) OR ((ID_TIPO_CUENTA = 10) AND (CUENTA_STD.NUMERO_CUENTA_ESTANDAR LIKE '31%'))) " & _
                  " ) A , AO_EXCHANGE_RATES B " & _
              " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
              "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dCuteDate & "') <= '" & dCuteDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dCuteDate & "') >= '" & dCuteDate & "') " & _
              " GROUP BY A.ID_MAYOR "
                
      sAAN_B = " SELECT  A.ID_MAYOR, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
               "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
               " FROM (" & sAAN & ") A , AO_EXCHANGE_RATES B " & _
                " WHERE (1 = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dCuteDate & "') <= '" & dCuteDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dCuteDate & "') >= '" & dCuteDate & "') " & _
                " GROUP BY A.ID_MAYOR "
                
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sAAN_B = sAAN
        End If
        
'Activo segundo periodo

     sAAC = " SELECT A.ID_MAYOR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL*(-1)),NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL)) SALDO_INICIAL_2, " & _
               "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.CARGOS*(-1)),NVL(EXCHANGE_RATE,1)*A.CARGOS)) CARGOS_2, " & _
               "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.ABONOS*(-1)),NVL(EXCHANGE_RATE,1)*A.ABONOS)) ABONOS_2, " & _
               "         SUM(DECODE(SUBSTR(NUMERO_CUENTA_ESTANDAR,1,2),'31',(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL*(-1)),NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL)) SALDO_ACTUAL_2, " & _
               "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
               "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
               " FROM (" & _
                  " select A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) AND " & _
                  "       ((ID_TIPO_CUENTA = 3) OR ((ID_TIPO_CUENTA = 10) AND (CUENTA_STD.NUMERO_CUENTA_ESTANDAR LIKE '31%'))) " & _
                  " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.ID_MAYOR "
                
      sAAC_B = " SELECT A.ID_MAYOR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_2, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_2, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_2, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_2, " & _
               "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
               "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
               " FROM (" & sAAC & " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (1 = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.ID_MAYOR "
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sAAC_B = sAAC
        End If
                
'*****cuenta PRIMER PERIODO
      sCAN = " SELECT  A.ID_MAYOR, " & _
             "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
             "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
             "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_3, " & _
             "         SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_3, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_3, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_3, " & _
             "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
             " FROM (" & _
                  " select A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
                  " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (CUENTA_STD.NUMERO_CUENTA_ESTANDAR = '6320-95') " & _
                  " ) A , AO_EXCHANGE_RATES B " & _
              " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
              "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
              "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
              "       (NVL(FROM_DATE,'" & dCuteDate & "') <= '" & dCuteDate & "') AND " & _
              "       (NVL(TO_DATE,'" & dCuteDate & "') >= '" & dCuteDate & "') " & _
              " GROUP BY A.ID_MAYOR "
                
        sCAN_B = " SELECT  A.ID_MAYOR, " & _
                 "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
                 "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
                 "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_3, " & _
                 "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_3, " & _
                 "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_3, " & _
                 "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_3, " & _
                 "         (0) SALDO_INICIAL_4, (0) CARGOS_4, (0) ABONOS_4, (0) SALDO_ACTUAL_4 " & _
                 " FROM (" & sCAN & ") A , AO_EXCHANGE_RATES B " & _
                 " WHERE (1 = B.FROM_CURRENCY_ID(+)) and " & _
                 "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                 "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                 "       (NVL(FROM_DATE,'" & dCuteDate & "') <= '" & dCuteDate & "') AND " & _
                 "       (NVL(TO_DATE,'" & dCuteDate & "') >= '" & dCuteDate & "') " & _
                 " GROUP BY A.ID_MAYOR "
                
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sCAN_B = sCAN
        End If
        
  ' CUENTA SEGUNDO PERIODO
  
     sCAC = " SELECT A.ID_MAYOR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_INICIAL) SALDO_INICIAL_4, " & _
               "         SUM(NVL(EXCHANGE_RATE,1)*A.CARGOS) CARGOS_4, SUM(NVL(EXCHANGE_RATE,1)*A.ABONOS) ABONOS_4, SUM(NVL(EXCHANGE_RATE,1)*A.SALDO_ACTUAL) SALDO_ACTUAL_4 " & _
               " FROM (" & _
                  " select A.ID_MONEDA, A.ID_MAYOR, A.NUMERO_CUENTA_ESTANDAR, " & _
                  "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL " & _
                  " FROM ( " & sSQL2_B & " ) CUENTA_STD , ( " & sSQL_B & " ) A " & _
                  " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                  "       (CUENTA_STD.NUMERO_CUENTA_ESTANDAR = '6320-95') " & _
                  " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.ID_MAYOR "
                
      sCAC_B = " SELECT A.ID_MAYOR, " & _
               "         (0) SALDO_INICIAL, (0) CARGOS, (0) ABONOS, (0) SALDO_ACTUAL, " & _
               "         (0) SALDO_INICIAL_2, (0) CARGOS_2, (0) ABONOS_2, (0) SALDO_ACTUAL_2, " & _
               "         (0) SALDO_INICIAL_3, (0) CARGOS_3, (0) ABONOS_3, (0) SALDO_ACTUAL_3, " & _
               "         SUM(A.SALDO_INICIAL/NVL(EXCHANGE_RATE,1)) SALDO_INICIAL_4, " & _
               "         SUM(A.CARGOS/NVL(EXCHANGE_RATE,1)) CARGOS_4, " & _
               "         SUM(A.ABONOS/NVL(EXCHANGE_RATE,1)) ABONOS_4, " & _
               "         SUM(A.SALDO_ACTUAL/NVL(EXCHANGE_RATE,1)) SALDO_ACTUAL_4 " & _
               " FROM (" & sCAC & " ) A , AO_EXCHANGE_RATES B " & _
                " WHERE (1 = B.FROM_CURRENCY_ID(+)) and " & _
                "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                " GROUP BY A.ID_MAYOR "
                
        If (cnFromCurrency = nExcRateCurrency) Then
          sCAC_B = sCAC
        End If


        sSQL = " SELECT  A.id_mayor, NUMERO_MAYOR, NOMBRE_MAYOR, DECODE(D.ID_GRUPO_MAYOR,23,2,17,1,18,1,19,1,20,1,25,1,3) ORDEN, " & _
                "         SUM(SALDO_INICIAL) SALDO_INICIAL, SUM(CARGOS) CARGOS, SUM(ABONOS) ABONOS, SUM(SALDO_ACTUAL) SALDO_ACTUAL, " & _
                "         SUM(SALDO_INICIAL_2) SALDO_INICIAL_2, SUM(CARGOS_2) CARGOS_2, SUM(ABONOS_2) ABONOS_2, SUM(SALDO_ACTUAL_2) SALDO_ACTUAL_2, " & _
                "         SUM(SALDO_INICIAL_3) SALDO_INICIAL_3, SUM(CARGOS_3) CARGOS_3, SUM(ABONOS_3) ABONOS_3, SUM(SALDO_ACTUAL_3) SALDO_ACTUAL_3, " & _
                "         SUM(SALDO_INICIAL_4) SALDO_INICIAL_4, SUM(CARGOS_4) CARGOS_4, SUM(ABONOS_4) ABONOS_4, SUM(SALDO_ACTUAL_4) SALDO_ACTUAL_4 " & _
                " FROM (" & _
                  sAAN_B & _
                " UNION " & _
                  sAAC_B & _
                " UNION " & _
                  sCAN_B & _
                " UNION " & _
                  sCAC_B & _
                " ) A, COF_MAYOR B, COF_ELEMENTO_GRUPO_MAYOR C, COF_GRUPO_MAYOR D " & _
                " WHERE (A.ID_MAYOR = B.ID_MAYOR(+)) AND " & _
                "       (A.ID_MAYOR = C.ID_MAYOR(+)) AND " & _
                "       (C.ID_GRUPO_MAYOR = D.ID_GRUPO_MAYOR(+)) AND " & _
                "       (D.CLIP_GRUPO_MAYOR = 2) " & _
                " GROUP BY A.ID_MAYOR, NUMERO_MAYOR, NOMBRE_MAYOR, D.ID_GRUPO_MAYOR "
                
        sSQL = " SELECT * " & _
               " FROM (" & sSQL & ") A " & _
               " ORDER BY A.ORDEN, NUMERO_MAYOR, NOMBRE_MAYOR "
               
    
    With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset74 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset74", Err.Number
End Function

Public Function GeneraRecordset96(sAppServer As String, vGL As Variant, dIniDate As Date, dEndDate As Date, _
                                  dExcDate As Date, Optional nExcRateType As Long = 46, _
                                  Optional nExcRateCurrency As Long = 1) As Recordset
  Dim oRecordset      As New Recordset
  Dim oConn           As New Connection
  Dim sSQL1           As String
  Dim sSql1A          As String
  Dim sSql1B          As String
  Dim sSql1C          As String
  Dim sSQL2           As String
  Dim sSql2A          As String
  Dim sSql2B          As String
  Dim sSql2C          As String
  Dim sSql3           As String
  Dim sSql4           As String
  
  Dim sSqlComp        As String
  Dim sGl             As String
  Dim sCondition      As String
  Dim sGLCondition    As String
  Dim sSQLStdAccHist  As String
'**********************************************
 On Error GoTo ErrHandler
  oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
  oConn.Open
  
     sCondition = " AND (fecha_afectacion >= '" & dIniDate & "') AND " & _
                  " (fecha_afectacion <= '" & dEndDate & "') "
  
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGLCondition = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGLCondition = "id_mayor = " & sGl
    End If
    
    CreateGralMovTBL oConn, sGLCondition, "tmp_MovsTBL"
    sSQLStdAccHist = SQLStandardAccountHist(dIniDate, dEndDate, 1)
    
    sSQL2 = " SELECT A.CUENTA, A.NOM_CUENTA, SUBSTR(A.CUENTA,1,4) CUENTA1, NVL(SUBSTR(A.CUENTA,6,LENGTH(A.CUENTA)-5),' ') SUBCUENTA, A.ID_MONEDA, AO_CURRENCIES.CURRENCY_NAME NOM_MONEDA, A.SECTOR, A.AUX, A.NATURALEZA NAT, NUMERO_CUENTA_AUXILIAR NUM_AUX, " & _
                 " NOMBRE_CUENTA_AUXILIAR NOM_AUX, A.EXC_RATE, A.ID_MAYOR, " & _
                 " DECODE(A.NATURALEZA, 'D', SUM(A.SI)+SUM(A.CARGOS)-SUM(A.ABONOS), " & _
                                      " 'A', SUM(A.SI)-SUM(A.CARGOS)+SUM(A.ABONOS)) ST " & _
          " FROM ( "
          
         sSql2A = " SELECT NUMERO_CUENTA_ESTANDAR CUENTA, NOMBRE_CUENTA_ESTANDAR NOM_CUENTA, " & _
                  "        COF_SALDO_INICIAL.ID_MONEDA, COF_SALDO_INICIAL.ID_SECTOR SECTOR, " & _
                         " COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR AUX, NATURALEZA, NVL(EXCHANGE_RATE,1) EXC_RATE, " & _
                         " COF_SALDO_INICIAL.ID_MAYOR, " & _
                         " DECODE(Sum(COF_SALDO_INICIAL.SALDO_INICIAL),0,0,Sum(COF_SALDO_INICIAL.SALDO_INICIAL)) SI, (0) Cargos, (0) Abonos " & _
                  " From COF_SALDO_INICIAL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (COF_SALDO_INICIAL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND " & _
                  "       (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                  "       (COF_SALDO_INICIAL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND " & _
                  "       (COF_SALDO_INICIAL." & sGLCondition & ") AND " & _
                  "       (FECHA_SALDO_INICIAL<='" & dEndDate & "') AND " & _
                  "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                  "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, COF_SALDO_INICIAL.ID_MONEDA, " & _
                  "          COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR , " & _
                  "          NATURALEZA, EXCHANGE_RATE, COF_SALDO_INICIAL.ID_MAYOR " & _
                  " Union "

         sSql2B = " (SELECT NUMERO_CUENTA_ESTANDAR CUENTA, NOMBRE_CUENTA_ESTANDAR NOM_CUENTA, ID_MONEDA, " & _
                  "         ID_SECTOR SECTOR, ID_CUENTA_AUXILIAR AUX, NATURALEZA, " & _
                  "         NVL(EXCHANGE_RATE,1) EXC_RATE, tmp_MovsTBL.ID_MAYOR, " & _
                  "         (0) SI, Sum(MONTO) Cargos, (0) Abonos " & _
                  " From tmp_MovsTBL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (tmp_MovsTBL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND " & _
                  "       (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                  "       (tmp_MovsTBL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND " & _
                  "       (TIPO_MOVIMIENTO='D') AND (ROL_CUENTA<>'S') AND " & _
                  "       (FECHA_AFECTACION>=NVL(fecha_saldo_inicial,'01/01/1990')) " & sCondition & " AND " & _
                  "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                  "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY NUMERO_CUENTA_ESTANDAR, NOMBRE_CUENTA_ESTANDAR, ID_MONEDA, ID_SECTOR, " & _
                  "          ID_CUENTA_AUXILIAR, NATURALEZA, EXCHANGE_RATE , tmp_MovsTBL.ID_MAYOR) " & _
                  " Union "

         sSql2C = " (SELECT NUMERO_CUENTA_ESTANDAR CUENTA, NOMBRE_CUENTA_ESTANDAR NOM_CUENTA, ID_MONEDA, " & _
                  "         ID_SECTOR SECTOR, ID_CUENTA_AUXILIAR AUX, NATURALEZA, " & _
                          " NVL(EXCHANGE_RATE,1) EXC_RATE, tmp_MovsTBL.ID_MAYOR, " & _
                          " (0) SI, (0) Cargos, Sum(MONTO) Abonos " & _
                  " From tmp_MovsTBL, COF_CUENTA, (" & sSQLStdAccHist & ") COF_CUENTA_ESTANDAR, AO_EXCHANGE_RATES " & _
                  " WHERE (tmp_MovsTBL.ID_CUENTA=COF_CUENTA.ID_CUENTA) AND " & _
                  "       (COF_CUENTA.ID_CUENTA_ESTANDAR=COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) AND " & _
                  "       (tmp_MovsTBL.ID_MONEDA = AO_EXCHANGE_RATES.TO_CURRENCY_ID(+)) AND " & _
                  "       (TIPO_MOVIMIENTO='H') And (ROL_CUENTA<>'S') AND " & _
                  "       (FECHA_AFECTACION>=NVL(fecha_saldo_inicial,'01/01/1990')) " & sCondition & " AND " & _
                  "       (NVL(EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(FROM_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                  "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  " GROUP BY COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, COF_CUENTA_ESTANDAR.NOMBRE_CUENTA_ESTANDAR, ID_MONEDA, ID_SECTOR, ID_CUENTA_AUXILIAR, NATURALEZA, " & _
                  "          EXCHANGE_RATE , tmp_MovsTBL.ID_MAYOR) " & _
           " ) A, COF_CUENTA_AUXILIAR, AO_CURRENCIES " & _
           " WHERE (A.AUX = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_MONEDA = AO_CURRENCIES.CURRENCY_ID(+)) AND ((A.CUENTA = '1103-04') OR (A.CUENTA = '1103-05')) " & _
           " GROUP BY A.CUENTA, A.NOM_CUENTA, A.ID_MONEDA, AO_CURRENCIES.CURRENCY_NAME, A.SECTOR, A.AUX, A.NATURALEZA, " & _
                    " NUMERO_CUENTA_AUXILIAR , NOMBRE_CUENTA_AUXILIAR, A.EXC_RATE, A.ID_MAYOR " & _
           " ORDER BY CUENTA1, A.SECTOR, SUBCUENTA, A.ID_MONEDA, A.AUX "
                       
    sSqlComp = sSQL2 & sSql2A & sSql2B & sSql2C
    sSqlComp = " SELECT A.*, NVL(COF_MAYOR.NUMERO_MAYOR,' ') NUM_MAYOR, NVL(COF_MAYOR.NOMBRE_MAYOR,' ') NOM_MAYOR, NVL(COF_SECTOR.CLAVE_SECTOR,'00') CVE_SECTOR  " & _
               " FROM ( " & sSqlComp & " ) A, COF_MAYOR, COF_SECTOR " & _
               " WHERE (A.ID_MAYOR = COF_MAYOR.ID_MAYOR(+)) AND (A.SECTOR = COF_SECTOR.ID_SECTOR(+)) AND (A.ST <> 0) "
  
    With oRecordset
      .CursorLocation = adUseClient
      .Open sSqlComp, oConn, adOpenForwardOnly, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRecordset96 = oRecordset
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset96", Err.Number
End Function

Public Function GeneraRecordset99(sAppServer As String, sPatron As String, vGL As Variant, _
                        dIniDate As Date, dCuteDate As Date, _
                        Optional sIniAccount As String = "", _
                        Optional sEndAccount As String = "", _
                        Optional nTransactionTypeId As Long = 0, _
                        Optional nVoucherTypeId As Long = 0, _
                        Optional nCurrencyType As Long = 0, _
                        Optional nExcRateType As Long = 46, _
                        Optional nExcRateCurrency As Long = 1, _
                        Optional dExcDate As Date, Optional nRowsToDisplay As Long = 4, _
                        Optional bWithAux As Boolean = False) As Recordset
                                   
  Dim oConn             As New Connection
  Dim oComm             As New Command
  Dim oRs               As New Recordset
  
  Dim sGl               As String
  
  Dim sSQL_A            As String
  Dim sSQL              As String
  Dim sSQL_SI           As String
  Dim sSQL1             As String
  Dim sSQL2             As String
  
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
                       
      Truncate sAppServer, "tmp_Saldos"
      
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
    
    sFromTable = " tmp_SALDOS "

    sSQL = "    SELECT id_moneda, SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4 ) NUMERO_CUENTA_ESTANDAR, (0) ID_SECTOR, " & _
           "           (0) ID_CUENTA_AUXILIAR, " & _
           "           SUM(A.SALDO_INICIAL) SALDO_INICIAL, Sum(A.CARGOS) CARGOS, Sum(A.ABONOS) ABONOS, " & _
           "           SUM(A.SALDO_ACTUAL) SALDO_ACTUAL " & _
           "    From  " & sFromTable & "  A " & _
           "    Where (Length(NUMERO_CUENTA_ESTANDAR) >= 4) and " & _
           "          not ((SUBSTR(A.NUMERO_CUENTA_ESTANDAR,1, 4) = NUMERO_CUENTA_ESTANDAR) and (nvl(ID_SECTOR,0)=0))" & _
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
            
      If bWithAux Then
        If (cnFromCurrency <> nExcRateCurrency) Then
          sSQL_A = " SELECT  id_moneda, NVL(C.CURRENCY_NAME,' ') CURRENCY_NAME, C.O_ID_MONEDA, A.NATURALEZA, " & _
                   "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
                   "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, " & _
                   "         SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, NVL(B.EXCHANGE_RATE,1) EXCHANGE_RATE_1, " & _
                   "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                   " FROM (" & _
                      " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                      "        NVL(NOMBRE_CUENTA_ESTANDAR, ' ') NOMBRE_CUENTA, " & _
                      "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                      "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                      "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                      " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                      " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                      "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                      "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                      " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                   " WHERE (A.ID_MONEDA = C.CURRENCY_ID) and (A.ID_MONEDA <> 1) AND (A.ID_MONEDA <> 28) AND " & _
                   "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (length(NUMERO_cuenta_ESTANDAR) <= " & Len(sPatron) & ") and (A.clave_sector = '00') and (SALDO_ACTUAL <> 0) AND " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, EXCHANGE_RATE, A.NOMBRE_CUENTA, CURRENCY_NAME, A.ID_TIPO_CUENTA, " & _
                   "          A.ROL_CUENTA, A.FROM_MOVS, C.O_ID_MONEDA "
                   
            sSQL = " SELECT  id_moneda, CURRENCY_NAME, A.O_ID_MONEDA, A.NATURALEZA, " & _
                   "         A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                   "         ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
                   "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, " & _
                   "         SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                   "         (0) SALDO_PROMEDIO, (NVL(EXCHANGE_RATE_1,1)/NVL(EXHANGE_RATE,1)) EXCHANGE_RATE, " & _
                   "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                   " FROM (" & sSQL_A & " ) A , AO_EXCHANGE_RATES B " & _
                   " WHERE (" & cnFromCurrency & " = B.FROM_CURRENCY_ID(+)) and " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   "       AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5011%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5051%')) " & _
                   " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
                   "          A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, EXCHANGE_RATE_1, EXCHANGE_RATE,  " & _
                   "          A.NOMBRE_CUENTA, CURRENCY_NAME, A.ID_TIPO_CUENTA, " & _
                   "          A.ROL_CUENTA, A.FROM_MOVS, A.O_ID_MONEDA, " & _
                   " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, " & _
                   "          NUMERO_CUENTA_AUXILIAR, O_ID_MONEDA "
                 
        Else
            sSQL = " SELECT  id_moneda, NVL(C.CURRENCY_NAME,' ') CURRENCY_NAME, C.O_ID_MONEDA, A.NATURALEZA, " & _
                   "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
                   "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                   "         (0) SALDO_PROMEDIO, NVL(B.EXCHANGE_RATE,1) EXCHANGE_RATE, A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                   " FROM (" & _
                      " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                      "        NVL(NOMBRE_CUENTA_ESTANDAR, ' ') NOMBRE_CUENTA, " & _
                      "        A.ID_CUENTA_AUXILIAR, NVL(NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, " & _
                      "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                      "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                      " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR, COF_CUENTA_AUXILIAR " & _
                      " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                      "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                      "       (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0) " & _
                      " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                   " WHERE (A.ID_MONEDA = C.CURRENCY_ID) and (A.ID_MONEDA <> 1) AND (A.ID_MONEDA <> 28) AND " & _
                   "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (length(NUMERO_cuenta_ESTANDAR) <= " & Len(sPatron) & ") and (A.clave_sector = '00') and (SALDO_ACTUAL <> 0) AND " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   "       AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5011%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5051%')) " & _
                   " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, EXCHANGE_RATE, A.NOMBRE_CUENTA, CURRENCY_NAME, A.ID_TIPO_CUENTA, " & _
                   "          A.ROL_CUENTA, A.FROM_MOVS, C.O_ID_MONEDA " & _
                   " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, NUMERO_CUENTA_AUXILIAR, O_ID_MONEDA "
                   'sOptionToDisplay
        End If
                     
      Else
        If (cnFromCurrency <> nExcRateCurrency) Then
        
          sSQL_A = " SELECT  id_moneda, NVL(C.CURRENCY_NAME,' ') CURRENCY_NAME, C.O_ID_MONEDA, A.NATURALEZA, " & _
                   "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                   "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, " & _
                   "         SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, NVL(B.EXCHANGE_RATE,1) EXCHANGE_RATE_1, " & _
                   "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                   " FROM (" & _
                      " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                      "        NVL(NOMBRE_CUENTA_ESTANDAR, ' ') NOMBRE_CUENTA, " & _
                      "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                      "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                      " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                      " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                      "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                      "       (ID_CUENTA_AUXILIAR = 0) " & _
                      " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                    " WHERE (A.ID_MONEDA = C.CURRENCY_ID) and (A.ID_MONEDA <> 1) AND (A.ID_MONEDA <> 28) AND " & _
                    "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (length(NUMERO_cuenta_ESTANDAR) <= " & Len(sPatron) & ") and (A.clave_sector = '00') and (SALDO_ACTUAL <> 0) AND " & _
                    "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                    "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                    "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                    "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                    " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, EXCHANGE_RATE, A.NOMBRE_CUENTA, CURRENCY_NAME, A.ID_TIPO_CUENTA, " & _
                    "          A.ROL_CUENTA, A.FROM_MOVS, C.O_ID_MONEDA "
          
            sSQL = " SELECT  id_moneda, CURRENCY_NAME, A.O_ID_MONEDA, A.NATURALEZA, " & _
                   "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                   "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, " & _
                   "         SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, " & _
                   "         (0) SALDO_PROMEDIO, (NVL(EXCHANGE_RATE_1,1)/NVL(EXCHANGE_RATE,1)) EXCHANGE_RATE, " & _
                   "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                   " FROM (" & sSQL_A & ") A , AO_EXCHANGE_RATES B " & _
                   " WHERE (" & cnFromCurrency & " = B.FROM_CURRENCY_ID(+)) AND " & _
                   "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                   "       (NVL(B.TO_CURRENCY_ID," & nExcRateCurrency & ") = " & nExcRateCurrency & ") and " & _
                   "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                   "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                   "       AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5011%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5051%')) " & _
                   " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, " & _
                   "          EXCHANGE_RATE_1, EXCHANGE_RATE, A.NOMBRE_CUENTA, CURRENCY_NAME, " & _
                   "          A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS, A.O_ID_MONEDA " & _
                   " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, O_ID_MONEDA "
      
        Else
          sSQL = " SELECT  id_moneda, NVL(C.CURRENCY_NAME,' ') CURRENCY_NAME, C.O_ID_MONEDA, A.NATURALEZA, " & _
                 "         A.NUMERO_CUENTA_ESTANDAR , A.CLAVE_SECTOR, A.NOMBRE_CUENTA, " & _
                 "         SUM(A.SALDO_INICIAL) SALDO_INICIAL, SUM(A.CARGOS) CARGOS, " & _
                 "         SUM(A.ABONOS) ABONOS, SUM(A.SALDO_ACTUAL) SALDO_ACTUAL, (0) SALDO_PROMEDIO, NVL(B.EXCHANGE_RATE,1) EXCHANGE_RATE, " & _
                 "         A.ID_TIPO_CUENTA, A.ROL_CUENTA, A.FROM_MOVS " & _
                 " FROM (" & _
                    " select A.ID_MONEDA, NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, NVL(CLAVE_SECTOR,'00') CLAVE_SECTOR, " & _
                    "        NVL(NOMBRE_CUENTA_ESTANDAR, ' ') NOMBRE_CUENTA, " & _
                    "        SALDO_INICIAL, CARGOS, ABONOS, SALDO_ACTUAL , " & _
                    "        NVL(ID_TIPO_CUENTA,0) ID_TIPO_CUENTA, NVL(ROL_CUENTA,' ') ROL_CUENTA, (0) FROM_MOVS " & _
                    " FROM ( " & sSQL2 & " ) CUENTA_STD , ( " & sSQL & " ) A , COF_SECTOR " & _
                    " WHERE (A.NUMERO_CUENTA_ESTANDAR = CUENTA_STD.NUMERO_CUENTA_ESTANDAR) and " & _
                    "       (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) AND " & _
                    "       (ID_CUENTA_AUXILIAR = 0) " & _
                    " ) A , AO_EXCHANGE_RATES B, AO_CURRENCIES C " & _
                  " WHERE (A.ID_MONEDA = C.CURRENCY_ID) and (A.ID_MONEDA <> 1) AND (A.ID_MONEDA <> 28) AND " & _
                  "       (A.ID_MONEDA = B.TO_CURRENCY_ID(+)) and (length(NUMERO_cuenta_ESTANDAR) <= " & Len(sPatron) & ") and (A.clave_sector = '00') and (SALDO_ACTUAL <> 0) AND " & _
                  "       (NVL(B.EXCHANGE_RATE_TYPE_ID," & nExcRateType & ") = " & nExcRateType & ") and " & _
                  "       (NVL(B.FROM_CURRENCY_ID," & cnFromCurrency & ") = " & cnFromCurrency & ") and " & _
                  "       (NVL(FROM_DATE,'" & dExcDate & "') <= '" & dExcDate & "') AND " & _
                  "       (NVL(TO_DATE,'" & dExcDate & "') >= '" & dExcDate & "') " & _
                  "       AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5011%')) AND (NOT(NUMERO_CUENTA_ESTANDAR LIKE '5051%')) " & _
                  " GROUP BY A.ID_MONEDA, A.NATURALEZA, A.NUMERO_CUENTA_ESTANDAR, A.CLAVE_SECTOR, EXCHANGE_RATE, A.NOMBRE_CUENTA, CURRENCY_NAME, A.ID_TIPO_CUENTA, " & _
                  "          A.ROL_CUENTA, A.FROM_MOVS, C.O_ID_MONEDA " & _
                  " ORDER BY NATURALEZA DESC, NUMERO_CUENTA_ESTANDAR, CLAVE_SECTOR, O_ID_MONEDA "
        End If
      End If
    
      With oRs
        .CursorLocation = adUseClient
        .Open sSQL, oConn, adOpenStatic, adLockReadOnly
        Set .ActiveConnection = Nothing
      End With
    Set GeneraRecordset99 = oRs
  Exit Function
ErrHandler:
  RaiseError ClassId, "GeneraRecordset99", Err.Number
End Function


Public Function GeneraRS_5031(sAppServer As String, _
                              dCuteDate As Date, _
                              vGL As Variant) As ADODB.Recordset
  Dim oConn               As New ADODB.Connection
  Dim oCommand            As ADODB.Command
  Dim oRs                 As New ADODB.Recordset

  Dim sSQL                As String
  Dim sSQL1               As String
  Dim sSQL2               As String
'  Dim sSql3               As String
  Dim sSqlAUXData         As String
  
  Dim sGl                 As String
  Dim sAccounts           As String
  '**********************************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    
    If (IsArray(vGL)) Then
      sGl = Join(vGL, ",")
      sGl = "id_mayor in (" & sGl & ") "
    Else
      sGl = vGL
      sGl = "id_mayor = " & sGl
    End If
    
    sAccounts = "'5106-02-01'" & " , " & "'5106-02-02-00-02'" & " , " & _
                "'5108-01-01-00-01'" & " , " & "'5108-01-01-00-02-01'" & " , " & "'5108-01-01-00-02-02'" & " , " & _
                "'5108-05-00-00-01-01'" & " , " & "'5108-05-00-00-02-01'" & " , " & "'5108-05-00-00-03-01'" & " , " & _
                "'5108-90-01-00-01-01'" & " , " & "'5108-90-01-00-01-02'" & " , " & _
                "'5109-01-00-00-01'" & " , " & "'5109-02-00-00-01'" & " , " & "'5109-90-00-00-02-01'" & " , " & _
                "'5109-90-00-00-03-01'" & " , " & _
                "'6390-00-00-00-80-10'"

    sSQL1 = " SELECT COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) as Cargos, (0) as Abonos " & _
            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) and " & _
            "       (COF_SALDO_INICIAL." & sGl & ") " & _
            " GROUP BY COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR "
    
'            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND "
    sSQL2 = " SELECT COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
            "        (0) as SI,  " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'D',MONTO,0))) as Cargos, " & _
            "        (sum(decode(TIPO_MOVIMIENTO,'H',MONTO,0))) AS Abonos " & _
            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) " & _
            "        and (COF_TRANSACCION." & sGl & ") " & _
            " GROUP BY COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "

      sSqlAUXData = SQLCreaCamposAux(sAppServer, "")
      sSQL = " SELECT nvl(B.RFC,' ') RFC, nvl(B.CURP,' ') CURP, NOMBRE_CUENTA_AUXILIAR , " & _
             "        nvl(B.CLAVE_RETENCION,' ') CLAVE_RETENCION, " & _
             "        (sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)) TOT_BASE, " & _
             "        (sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)) TOT_BASE_IVA," & _
             "        ((sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)) * 0.10 ) TOT_RET , " & _
             "        (DECODE(ESTA_GRAVADO,1,sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS),0) * (NVL(IVA,100)/100)) TOT_IVA " & _
             " FROM ( " & _
             sSQL1 & _
             " UNION " & _
             sSQL2 & _
             " ) A  , ( " & sSqlAUXData & " ) B , COF_CUENTA_AUXILIAR " & _
             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) and " & _
             "       (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
             " GROUP BY  RFC, CURP, NOMBRE_CUENTA_AUXILIAR, CLAVE_RETENCION, ESTA_GRAVADO, ISR , IVA " & _
             " HAVING (sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)) <> 0 " & _
             " ORDER BY RFC, NOMBRE_CUENTA_AUXILIAR "

     With oRs
      .CursorLocation = adUseClient
      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
      Set .ActiveConnection = Nothing
    End With
    Set GeneraRS_5031 = oRs
  Exit Function
ErrHandler:
   RaiseError ClassId, "GeneraRS_5031", Err.Number
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

'''''Public Function GetRS_7014(sAppServer As String, _
'''''                                 dCuteDate As Date, _
'''''                                 sAccounts As String, nGL As Long) As ADODB.Recordset
'''''  Dim oConn                As New ADODB.Connection
'''''  Dim oCommand             As ADODB.Command
'''''  Dim oRs                  As New ADODB.Recordset
'''''
'''''  Dim sSQL                 As String
'''''  Dim ssql1                As String
'''''  Dim ssql2                As String
'''''  Dim sSql3                As String
'''''  Dim sSqlAux              As String
'''''  Dim sSqlComp             As String
'''''  '**********************************************************************************
'''''  On Error GoTo ErrHandler
'''''    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
'''''    oConn.Open
'''''
'''''    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
'''''            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
'''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
'''''            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) as Cargos, (0) as Abonos " & _
'''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
'''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'''''            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
'''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
'''''            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGL & ") " & _
'''''            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
'''''                    " COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
'''''            "       COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
'''''
'''''    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
'''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
'''''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  (0) as Cargos, Sum(COF_MOVIMIENTO.MONTO) AS Abonos " & _
'''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
'''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
'''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
'''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
'''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
'''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
'''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
'''''            "        and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
'''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
'''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
'''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
'''''
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
'''''
'''''      sSQL = " SELECT A.Cuenta, NOMBRE_CUENTA_ESTANDAR, " & _
'''''                      " A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOM_SECTOR, nvl(CLAVE_SECTOR, '00') NUM_SECTOR, " & _
'''''                      " A.ID_CUENTA_AUXILIAR, nvl(COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR,' ') NUMERO_CUENTA_AUXILIAR, nvl(COF_CUENTA_AUXILIAR.NOMBRE_CUENTA_AUXILIAR, ' ') NOMBRE_CUENTA_AUXILIAR  , " & _
'''''             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, sum(A.SI) SI, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
'''''             "        decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
'''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) SA " & _
'''''             " FROM ( " & _
'''''             ssql1 & _
'''''             " UNION " & _
'''''             ssql2 & _
'''''             " UNION " & _
'''''             sSql3 & _
'''''             " ) A  , COF_SECTOR , COF_CUENTA_AUXILIAR" & _
'''''             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) " & _
'''''                 "  AND (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
'''''             " GROUP BY  A.Cuenta, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
'''''                       " NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
'''''             " HAVING decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
'''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
'''''             " ORDER BY A.Cuenta, CLAVE_SECTOR , NUMERO_CUENTA_AUXILIAR "
'''''
'''''    ' Para obtener los campos de los auxiliares
'''''
'''''    sSqlAux = SQLCreaCamposAux(sAppServer, "")
'''''
'''''    sSqlComp = " SELECT A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
'''''                        " SUM(A.SA) AS SA, A.FECHA_ULT, " & _
'''''                        " A.NOM_SECTOR, A.NUM_SECTOR , A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
'''''                        "  nvl(B.DESCRIPCION,' ') DESCRIPCION  " & _
'''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
'''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
'''''               " GROUP BY A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.ID_CUENTA_AUXILIAR, " & _
'''''                       " A.FECHA_ULT, A.NOM_SECTOR, A.NUM_SECTOR, A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, " & _
'''''                       " B.DESCRIPCION "
'''''    With oRs
'''''      .CursorLocation = adUseClient
'''''      .Open sSqlComp, oConn, adOpenStatic, adLockReadOnly
'''''      Set .ActiveConnection = Nothing
'''''    End With
'''''    Set GetRS_7014 = oRs
'''''  Exit Function
'''''ErrHandler:
'''''   RaiseError ClassId, "GetRS_7014", Err.Number
'''''End Function

''''''Public Function GeneraRS_7021(sAppServer As String, _
''''''                              dCuteDate As Date, _
''''''                              vGL As Variant) As ADODB.Recordset
''''''  Dim oConn               As New ADODB.Connection
''''''  Dim oRs                 As New ADODB.Recordset
''''''
''''''  Dim sSQL                As String
''''''  Dim ssql1               As String
''''''  Dim ssql2               As String
''''''  Dim sSql3               As String
''''''
''''''  Dim sGL                 As String
''''''  Dim sAccounts           As String
''''''  Dim sAccounts1           As String
''''''  Dim sAccounts2           As String
''''''  '**********************************************************************************
''''''  On Error GoTo ErrHandler
''''''    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
''''''    oConn.Open
''''''
''''''    If (IsArray(vGL)) Then
''''''      sGL = Join(vGL, ",")
''''''      sGL = "id_mayor in (" & sGL & ") "
''''''    Else
''''''      sGL = vGL
''''''      sGL = "id_mayor = " & sGL
''''''    End If
''''''
''''''    sAccounts1 = "'1350-00-00-00-04-01-01','1350-00-00-00-04-01-02','1350-00-00-00-04-01-03','1350-00-00-00-04-01-04'," & _
''''''                 "'1350-00-00-00-04-02-01','1350-00-00-00-04-03-01','1350-00-00-00-04-04-01','1350-00-00-00-06-02-01'," & _
''''''                 "'1350-00-00-00-04-02-04','1350-00-00-00-07-02-01','1350-00-00-00-08-02-01'," & _
''''''                 "'1350-00-00-00-04-02-30','1350-00-00-00-04-03-30','1350-00-00-00-04-04-30','1350-00-00-00-06-02-30'," & _
''''''                 "'1350-00-00-00-07-02-30','1350-00-00-00-08-02-30'," & _
''''''                 "'1319-52-00-00-04-01-02','1319-52-00-00-04-02-00','1319-52-00-00-04-03-00','1319-52-00-00-04-04-00'," & _
''''''                 "'1319-52-00-00-06-02-00','1319-52-00-00-07-02-00','1319-52-00-00-08-02-00'," & _
''''''                 "'1319-30-52-00-04-01-00','1319-30-52-00-04-02-00','1319-30-52-00-04-03-00','1319-30-52-00-06-01-00'," & _
''''''                 "'1319-30-52-00-07-01-00','1319-30-52-00-08-01-00'," & _
''''''                 "'1319-30-52-00-04-04-00','1319-30-52-00-04-05-00','1319-30-52-00-04-06-00','1319-30-52-00-06-02-00'," & _
''''''                 "'1319-30-52-00-07-02-00','1319-30-52-00-08-02-00'," & _
''''''                 "'1314-01-50-00-04-02-00','1314-01-50-00-04-03-00','1314-01-50-00-04-04-00','1314-01-50-00-06-00-00'," & _
''''''                 "'1314-01-50-00-07-00-00','1314-01-50-00-08-00-00'," & _
''''''                 "'1314-02-50-00-04-02-00','1314-02-50-00-04-03-00','1314-02-50-00-04-04-00','1314-02-50-00-06-00-00'," & _
''''''                 "'1314-02-50-00-07-00-00','1314-02-50-00-08-00-00'," & _
''''''                 "'1320-50-00-00-04-02-02','1320-50-00-00-04-03-02','1320-50-00-00-04-04-02','1320-50-00-00-06-02-00'," & _
''''''                 "'1320-50-00-00-07-02-00','1320-50-00-00-08-02-00'," & _
''''''                 "'1320-50-00-00-04-02-03','1320-50-00-00-04-03-03','1320-50-00-00-04-04-03','1320-50-00-00-06-03-00'," & _
''''''                 "'1320-50-00-00-07-03-00','1320-50-00-00-08-03-00'," & _
''''''                 "'6378-02-02-00-04-19-00','6378-02-02-00-04-20-00','6378-02-02-00-04-21-00','6378-02-02-00-04-25-00'," & _
''''''                 "'6378-02-02-00-04-26-00','6378-02-02-00-04-27-00'," & _
''''''                 "'6378-02-02-00-04-13-00','6378-02-02-00-04-14-00','6378-02-02-00-04-15-00','6378-02-02-00-04-31-00'," & _
''''''                 "'6378-02-02-00-04-32-00','6378-02-02-00-04-33-00'," & _
''''''                 "'6378-02-02-00-04-22-00','6378-02-02-00-04-23-00','6378-02-02-00-04-24-00','6378-02-02-00-04-28-00'," & _
''''''                 "'6378-02-02-00-04-29-00','6378-02-02-00-04-30-00',"
''''''
''''''    sAccounts2 = "'6378-02-02-00-04-16-00','6378-02-02-00-04-17-00','6378-02-02-00-04-18-00','6378-02-02-00-04-34-00'," & _
''''''                 "'6378-02-02-00-04-35-00','6378-02-02-00-04-36-00'," & _
''''''                 "'6378-02-02-00-04-01-00','6378-02-02-00-04-11-00','6378-02-02-00-04-12-00','6378-02-02-00-04-06-00'," & _
''''''                 "'6378-02-02-00-04-07-00','6378-02-02-00-04-08-00'"
''''''
''''''    sAccounts = sAccounts1 & sAccounts2
''''''
''''''    ssql1 = " SELECT COF_CUENTA.ID_CUENTA,COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
''''''            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) as Cargos, (0) as Abonos " & _
''''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
''''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''''            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
''''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) and " & _
''''''            "       (COF_SALDO_INICIAL." & sGL & ") " & _
''''''            " GROUP BY COF_CUENTA.ID_CUENTA,COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR "
''''''
''''''    ssql2 = " SELECT COF_CUENTA.ID_CUENTA,COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''''            "        (0) as SI, " & _
''''''            "        (sum(decode(TIPO_MOVIMIENTO,'D',Monto,0))) as Cargos, " & _
''''''            "        (sum(decode(TIPO_MOVIMIENTO,'H',Monto,0))) as Abonos " & _
''''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
''''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
''''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
''''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
''''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
''''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
''''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
''''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
''''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR in (" & sAccounts & ")) " & _
''''''            "        and (COF_TRANSACCION." & sGL & ") " & _
''''''            " GROUP BY COF_CUENTA.ID_CUENTA,COF_MOVIMIENTO.ID_CUENTA_AUXILIAR "
''''''
''''''    sSQL = " SELECT A.ID_CUENTA,ID_CUENTA_AUXILIAR, " & _
''''''           "        (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
''''''           "                              (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) )) SALDO " & _
''''''           " FROM ( " & _
''''''           ssql1 & _
''''''           " UNION " & _
''''''           ssql2 & _
''''''           " ) A  , COF_CUENTA B, COF_CUENTA_ESTANDAR C " & _
''''''           " WHERE (A.ID_CUENTA = B.ID_CUENTA) and " & _
''''''           "       (B.ID_CUENTA_ESTANDAR = C.ID_CUENTA_ESTANDAR) " & _
''''''           " GROUP BY  A.ID_CUENTA, ID_CUENTA_AUXILIAR, NATURALEZA " & _
''''''           " HAVING (decode(NATURALEZA,'D',(sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS)), " & _
''''''           "                              (sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS))) <> 0 ) "
''''''
''''''    ssql1 = "select NUMERO_CUENTA_AUXILIAR AUX, NOMBRE_CUENTA_AUXILIAR ACREDITADO, " & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1350-00-00-00-04-01-01',SALDO,'1350-00-00-00-04-01-02',SALDO,'1350-00-00-00-04-01-03',SALDO," & _
''''''                            "'1350-00-00-00-04-01-04',SALDO,'1350-00-00-00-04-02-01',SALDO,'1350-00-00-00-04-03-01',SALDO," & _
''''''                            "'1350-00-00-00-04-04-01',SALDO,'1350-00-00-00-06-02-01',SALDO,'1350-00-00-00-04-02-04',SALDO," & _
''''''                            "'1350-00-00-00-07-02-01',SALDO,'1350-00-00-00-08-02-01',SALDO)),0) CAPITAL_VIG," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1350-00-00-00-04-02-30',SALDO,'1350-00-00-00-04-03-30',SALDO,'1350-00-00-00-04-04-30',SALDO," & _
''''''                            "'1350-00-00-00-06-02-30',SALDO,'1350-00-00-00-07-02-30',SALDO,'1350-00-00-00-08-02-30',SALDO)),0) AMORT_X_TRA," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1319-52-00-00-04-01-02',SALDO,'1319-52-00-00-04-02-00',SALDO,'1319-52-00-00-04-03-00',SALDO," & _
''''''                            "'1319-52-00-00-04-04-00',SALDO,'1319-52-00-00-06-02-00',SALDO,'1319-52-00-00-07-02-00',SALDO," & _
''''''                            "'1319-52-00-00-08-02-00',SALDO)),0) INT_VIGENTE," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1319-30-52-00-04-01-00',SALDO,'1319-30-52-00-04-02-00',SALDO,'1319-30-52-00-04-03-00',SALDO," & _
''''''                            "'1319-30-52-00-06-01-00',SALDO,'1319-30-52-00-07-01-00',SALDO,'1319-30-52-00-08-01-00',SALDO)),0) INTS_NOR_X_TRASP," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1319-30-52-00-04-04-00',SALDO,'1319-30-52-00-04-05-00',SALDO,'1319-30-52-00-04-06-00',SALDO," & _
''''''                            "'1319-30-52-00-06-02-00',SALDO,'1319-30-52-00-07-02-00',SALDO,'1319-30-52-00-08-02-00',SALDO)),0) INTS_COR_X_TRASP," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1314-01-50-00-04-02-00',SALDO,'1314-01-50-00-04-03-00',SALDO,'1314-01-50-00-04-04-00',SALDO," & _
''''''                            "'1314-01-50-00-06-00-00',SALDO,'1314-01-50-00-07-00-00',SALDO,'1314-01-50-00-08-00-00',SALDO)),0) CAP_VENC_EXIGIBLE," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1314-02-50-00-04-02-00',SALDO,'1314-02-50-00-04-03-00',SALDO,'1314-02-50-00-04-04-00',SALDO," & _
''''''                            "'1314-02-50-00-06-00-00',SALDO,'1314-02-50-00-07-00-00',SALDO,'1314-02-50-00-08-00-00',SALDO)),0) CAP_VENC_NO_EXIGIBLE,"
''''''    ssql2 = "        NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1320-50-00-00-04-02-02',SALDO,'1320-50-00-00-04-03-02',SALDO,'1320-50-00-00-04-04-02',SALDO," & _
''''''                            "'1320-50-00-00-06-02-00',SALDO,'1320-50-00-00-07-02-00',SALDO,'1320-50-00-00-08-02-00',SALDO)),0) INTERESES_VENCIDS," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'1320-50-00-00-04-02-03',SALDO,'1320-50-00-00-04-03-03',SALDO,'1320-50-00-00-04-04-03',SALDO," & _
''''''                            "'1320-50-00-00-06-03-00',SALDO,'1320-50-00-00-07-03-00',SALDO,'1320-50-00-00-08-03-00',SALDO)),0) INTERESES_CORRIDS," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'6378-02-02-00-04-19-00',SALDO,'6378-02-02-00-04-20-00',SALDO,'6378-02-02-00-04-21-00',SALDO," & _
''''''                            "'6378-02-02-00-04-25-00',SALDO,'6378-02-02-00-04-26-00',SALDO,'6378-02-02-00-04-27-00',SALDO)),0) INTER_NORM_EJ_CSO," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'6378-02-02-00-04-13-00',SALDO,'6378-02-02-00-04-14-00',SALDO,'6378-02-02-00-04-15-00',SALDO," & _
''''''                            "'6378-02-02-00-04-31-00',SALDO,'6378-02-02-00-04-32-00',SALDO,'6378-02-02-00-04-33-00',SALDO)),0) INTER_NORM_EJ_ANT," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'6378-02-02-00-04-22-00',SALDO,'6378-02-02-00-04-23-00',SALDO,'6378-02-02-00-04-24-00',SALDO," & _
''''''                            "'6378-02-02-00-04-28-00',SALDO,'6378-02-02-00-04-29-00',SALDO,'6378-02-02-00-04-30-00',SALDO)),0) INTER_CORR_EJ_CSO," & _
''''''                   " NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR, " & _
''''''                            "'6378-02-02-00-04-16-00',SALDO,'6378-02-02-00-04-17-00',SALDO,'6378-02-02-00-04-18-00',SALDO," & _
''''''                            "'6378-02-02-00-04-34-00',SALDO,'6378-02-02-00-04-35-00',SALDO,'6378-02-02-00-04-36-00',SALDO)),0) INTER_CORR_EJ_ANT,"
''''''    sSql3 = "        NVL(SUM(DECODE(NUMERO_CUENTA_ESTANDAR," & _
''''''                            "'6378-02-02-00-04-01-00',SALDO,'6378-02-02-00-04-11-00',SALDO,'6378-02-02-00-04-12-00',SALDO," & _
''''''                            "'6378-02-02-00-04-06-00',SALDO,'6378-02-02-00-04-07-00',SALDO,'6378-02-02-00-04-08-00',SALDO)),0) INTER_MORATOR " & _
''''''            " FROM ( " & sSQL & " ) A, cof_cuenta B,cof_cuenta_estandar C, COF_CUENTA_AUXILIAR D " & _
''''''            " WHERE (A.id_cuenta = B.id_cuenta) and " & _
''''''            "       (B.id_cuenta_estandar = C.id_cuenta_estandar) and " & _
''''''            "       (A.id_cuenta_auxiliar = D.id_cuenta_auxiliar) " & _
''''''            " GROUP BY NUMERO_CUENTA_AUXILIAR , NOMBRE_CUENTA_AUXILIAR " & _
''''''            " ORDER BY NUMERO_CUENTA_AUXILIAR "
''''''
''''''    sSQL = ssql1 & ssql2 & sSql3
''''''
''''''    With oRs
''''''      .CursorLocation = adUseClient
''''''      .Open sSQL, oConn, adOpenStatic, adLockReadOnly
''''''      Set .ActiveConnection = Nothing
''''''    End With
''''''    Set GeneraRS_7021 = oRs
''''''  Exit Function
''''''ErrHandler:
''''''   RaiseError ClassId, "GeneraRS_7021", Err.Number
''''''End Function

''''Public Function GetRS_7024(sAppServer As String, _
''''                                 dCuteDate As Date, _
''''                                 sAccounts As String, nGL As Long) As ADODB.Recordset
''''  Dim oConn                As New ADODB.Connection
''''  Dim oCommand             As ADODB.Command
''''  Dim oRs                  As New ADODB.Recordset
''''
''''  Dim sSQL                 As String
''''  Dim ssql1                As String
''''  Dim ssql2                As String
''''  Dim sSql3                As String
''''  Dim sSqlAux              As String
''''  Dim sSqlComp             As String
''''  '**********************************************************************************
''''  On Error GoTo ErrHandler
''''    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
''''    oConn.Open
''''
''''    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''''            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
''''            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) as Cargos, (0) as Abonos " & _
''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
''''            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGL & ") " & _
''''            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
''''                    " COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
''''            "       COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
''''
''''    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  (0) as Cargos, Sum(COF_MOVIMIENTO.MONTO) AS Abonos " & _
''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
''''            "        and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
''''
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
''''
''''      sSQL = " SELECT A.Cuenta, NOMBRE_CUENTA_ESTANDAR, " & _
''''                      " A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOM_SECTOR, nvl(CLAVE_SECTOR, '00') NUM_SECTOR, " & _
''''                      " A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , " & _
''''             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, sum(A.SI) SI, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
''''             "        decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) SA " & _
''''             " FROM ( " & _
''''             ssql1 & _
''''             " UNION " & _
''''             ssql2 & _
''''             " UNION " & _
''''             sSql3 & _
''''             " ) A  , COF_SECTOR , COF_CUENTA_AUXILIAR" & _
''''             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) " & _
''''                 "  AND (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
''''             " GROUP BY  A.Cuenta, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
''''                       " NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
''''             " HAVING decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
''''             " ORDER BY A.Cuenta, CLAVE_SECTOR , NUMERO_CUENTA_AUXILIAR "
''''
''''    ' Para obtener los campos de los auxiliares
''''
''''    sSqlAux = SQLCreaCamposAux(sAppServer, True)
''''
''''    sSqlComp = " SELECT A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''               " A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, SUM(A.SA) AS SA, " & _
''''               " NVL(B.RFC,' ') RFC, NVL(B.DOMICILIO,' ') DOMICILIO, NVL(B.COLONIA,' ') AS COLONIA, NVL(B.ESTADO,' ') ESTADO " & _
''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) AND (A.ID_CUENTA_AUXILIAR <> 0)" & _
''''               " GROUP BY A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''                        " A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, B.RFC, B.DOMICILIO, B.COLONIA, B.ESTADO " & _
''''               " ORDER BY A.CUENTA, A.NUMERO_CUENTA_AUXILIAR "
''''
''''
''''    With oRs
''''      .CursorLocation = adUseClient
''''      .Open sSqlComp, oConn, adOpenStatic, adLockReadOnly
''''      Set .ActiveConnection = Nothing
''''    End With
''''    Set GetRS_7024 = oRs
''''  Exit Function
''''ErrHandler:
''''   RaiseError ClassId, "GetRS_7024", Err.Number
''''End Function

''''Public Function GetRS_7028(sAppServer As String, _
''''                                 dCuteDate As Date, _
''''                                 sAccounts As String, nGL As Long) As ADODB.Recordset
''''  Dim oConn                As New ADODB.Connection
''''  Dim oCommand             As ADODB.Command
''''  Dim oRs                  As New ADODB.Recordset
''''
''''  Dim sSQL                 As String
''''  Dim ssql1                As String
''''  Dim ssql2                As String
''''  Dim sSql3                As String
''''  Dim sSqlAux              As String
''''  Dim sSqlComp             As String
''''  '**********************************************************************************
''''  On Error GoTo ErrHandler
''''    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
''''    oConn.Open
''''
''''    ssql1 = " SELECT COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''''            "        COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(FECHA_SALDO_INICIAL) FECHA, " & _
''''            "        sum(COF_SALDO_INICIAL.SALDO_INICIAL) SI, (0) as Cargos, (0) as Abonos " & _
''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_SALDO_INICIAL " & _
''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''            "       (COF_CUENTA.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA) and " & _
''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) and " & _
''''            "       (COF_SALDO_INICIAL.ID_MAYOR = " & nGL & ") " & _
''''            " GROUP BY COF_SALDO_INICIAL.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
''''                    " COF_SALDO_INICIAL.ID_SECTOR, COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR, " & _
''''            "       COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
''''
''''    ssql2 = " SELECT COF_MOVIMIENTO.ID_MONEDA, NUMERO_CUENTA_ESTANDAR AS Cuenta, " & _
''''            "        NOMBRE_CUENTA_ESTANDAR, " & _
''''            "        COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''            "        COF_CUENTA_ESTANDAR.NATURALEZA, max(nvl(FECHA_MOVIMIENTO,'01/01/1990')) FECHA, (0) as SI,  (0) as Cargos, Sum(COF_MOVIMIENTO.MONTO) AS Abonos " & _
''''            " FROM COF_CUENTA, COF_CUENTA_ESTANDAR, COF_MOVIMIENTO, COF_TRANSACCION, COF_SALDO_INICIAL  " & _
''''            " WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_CUENTA_ESTANDAR.ID_CUENTA_ESTANDAR) and " & _
''''            "       (COF_CUENTA.ID_CUENTA = COF_MOVIMIENTO.ID_CUENTA) and " & _
''''            "       (COF_MOVIMIENTO.ID_TRANSACCION = COF_TRANSACCION.ID_TRANSACCION) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA = COF_SALDO_INICIAL.ID_CUENTA(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_SECTOR = COF_SALDO_INICIAL.ID_SECTOR(+)) and " & _
''''            "       (COF_MOVIMIENTO.ID_CUENTA_AUXILIAR = COF_SALDO_INICIAL.ID_CUENTA_AUXILIAR(+)) and " & _
''''            "       (COF_MOVIMIENTO.TIPO_MOVIMIENTO='H') AND " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION <= '" & dCuteDate & "')  and " & _
''''            "       (COF_TRANSACCION.FECHA_AFECTACION >= nvl(FECHA_SALDO_INICIAL,'01/01/1990')) and " & _
''''            "       (COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR like ('" & sAccounts & "%')) " & _
''''            "        and (COF_TRANSACCION.ID_MAYOR = " & nGL & ") " & _
''''            " GROUP BY COF_MOVIMIENTO.ID_MONEDA, COF_CUENTA_ESTANDAR.NUMERO_CUENTA_ESTANDAR, " & _
''''                    " COF_MOVIMIENTO.ID_SECTOR, COF_MOVIMIENTO.ID_CUENTA_AUXILIAR, " & _
''''                    " COF_CUENTA_ESTANDAR.NATURALEZA, NOMBRE_CUENTA_ESTANDAR "
''''
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
''''
''''      sSQL = " SELECT A.Cuenta, NOMBRE_CUENTA_ESTANDAR, " & _
''''                      " A.ID_SECTOR, nvl(NOMBRE_SECTOR, 'Sin Sector') NOM_SECTOR, nvl(CLAVE_SECTOR, '00') NUM_SECTOR, " & _
''''                      " A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, NOMBRE_CUENTA_AUXILIAR , " & _
''''             "        A.NATURALEZA, max(A.FECHA) FECHA_ULT, sum(A.SI) SI, sum(A.Cargos) Cargos, sum(A.Abonos) Abonos, " & _
''''             "        decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) SA " & _
''''             " FROM ( " & _
''''             ssql1 & _
''''             " UNION " & _
''''             ssql2 & _
''''             " UNION " & _
''''             sSql3 & _
''''             " ) A  , COF_SECTOR , COF_CUENTA_AUXILIAR" & _
''''             " WHERE (A.ID_CUENTA_AUXILIAR = COF_CUENTA_AUXILIAR.ID_CUENTA_AUXILIAR(+)) " & _
''''                 "  AND (A.ID_SECTOR = COF_SECTOR.ID_SECTOR(+)) " & _
''''             " GROUP BY  A.Cuenta, A.ID_SECTOR, A.ID_CUENTA_AUXILIAR, NUMERO_CUENTA_AUXILIAR, " & _
''''                       " NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, NOMBRE_SECTOR, CLAVE_SECTOR, NOMBRE_CUENTA_ESTANDAR " & _
''''             " HAVING decode(A.NATURALEZA,'D',sum(A.SI)+sum(A.CARGOS)-sum(A.ABONOS), " & _
''''             "                                sum(A.SI)-sum(A.CARGOS)+sum(A.ABONOS)) <> 0 " & _
''''             " ORDER BY A.Cuenta, CLAVE_SECTOR , NUMERO_CUENTA_AUXILIAR "
''''
''''    ' Para obtener los campos de los auxiliares
''''
''''    sSqlAux = SQLCreaCamposAux(sAppServer, "")
''''
''''    sSqlComp = " SELECT A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''               " A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, SUM(A.SA) AS SA, " & _
''''               " NVL(B.FECHA_OFICIO, TO_DATE('')) FECHA_OFICIO, NVL(B.NUMERO_OFICIO,' ') NUMERO_OFICIO  " & _
''''               " FROM  ( " & sSQL & " ) A, ( " & sSqlAux & " ) B " & _
''''               " WHERE (A.ID_CUENTA_AUXILIAR = B.ID_CUENTA_AUXILIAR(+)) " & _
''''               " GROUP BY A.Cuenta, A.NOMBRE_CUENTA_ESTANDAR, A.ID_SECTOR, A.NOM_SECTOR, A.NUM_SECTOR, " & _
''''                        " A.ID_CUENTA_AUXILIAR, A.NUMERO_CUENTA_AUXILIAR, A.NOMBRE_CUENTA_AUXILIAR, A.NATURALEZA, " & _
''''                        " B.FECHA_OFICIO, B.NUMERO_OFICIO " & _
''''               " ORDER BY B.FECHA_OFICIO, A.CUENTA, a.numero_cuenta_auxiliar "
''''
''''
''''    With oRs
''''      .CursorLocation = adUseClient
''''      .Open sSqlComp, oConn, adOpenStatic, adLockReadOnly
''''      Set .ActiveConnection = Nothing
''''    End With
''''    Set GetRS_7028 = oRs
''''  Exit Function
''''ErrHandler:
''''   RaiseError ClassId, "GetRS_7028", Err.Number
''''End Function

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

Public Function TypeStdAccount(sAppServer As String, nGl As Long) As Long
  Dim oConn             As New ADODB.Connection
  Dim oRs               As New ADODB.Recordset
  Dim sSQLTypeStdAcc    As String
  '**********************************************************************
  On Error GoTo ErrHandler
    oConn.ConnectionString = GetConstant(sConstantName:=sAppServer)
    oConn.Open
    sSQLTypeStdAcc = " SELECT ID_TIPO_CUENTAS_STD FROM COF_MAYOR WHERE (ID_MAYOR = " & nGl & ")"
    
    With oRs
      .Open sSQLTypeStdAcc, oConn, adOpenStatic, adLockReadOnly
      TypeStdAccount = !id_tipo_cuentas_std
      .Close
      Set .ActiveConnection = Nothing
    End With
  Exit Function
ErrHandler:
  RaiseError ClassId, "TypeStdAccount", Err.Number
End Function

Private Function LikeSQLString(sFieldName, vLowBound As Variant, vSupBound As Variant, _
                               Optional bFormatAsTextField As Boolean = True) As String
  Dim sTemp As String
  '*************************************************************************************
  On Error GoTo ErrHandler
    If (Len(vLowBound) <> 0) And (Len(vSupBound) <> 0) Then
      sTemp = "(" & sFieldName & " >= '" & vLowBound & "' AND " & sFieldName & " <= '" & vSupBound & "')"
    ElseIf (Len(vLowBound) <> 0) And (Len(vSupBound) = 0) Then
      If HasWildCharacters(vLowBound) Then
        sTemp = "(" & sFieldName & " LIKE '" & vLowBound & "')"
      Else
        sTemp = "(" & sFieldName & " >= '" & vLowBound & "')"
      End If
    ElseIf (Len(vLowBound) = 0) And (Len(vSupBound) <> 0) Then
      If HasWildCharacters(vSupBound) Then
        sTemp = "(" & sFieldName & " LIKE '" & vSupBound & "')"
      Else
        sTemp = "(" & sFieldName & " <= '" & vSupBound & "')"
      End If
    Else
      sTemp = ""
    End If
    sTemp = Replace(sTemp, "*", "%")
    sTemp = Replace(sTemp, "?", "_")
    If (Not bFormatAsTextField) Then
      sTemp = Replace(sTemp, "'", "")
    End If
    LikeSQLString = sTemp
  Exit Function
ErrHandler:
  RaiseError ClassId, "LikeSQLString", Err.Number
End Function

Private Function HasWildCharacters(vExpression As Variant) As Boolean
  On Error GoTo ErrHandler
    HasWildCharacters = True
    If (InStr(1, vExpression, "*") > 0) Then
      Exit Function
    End If
    If (InStr(1, vExpression, "?") > 0) Then
      Exit Function
    End If
    If (InStr(1, vExpression, "%") > 0) Then
      Exit Function
    End If
    If (InStr(1, vExpression, "_") > 0) Then
      Exit Function
    End If
    HasWildCharacters = False
  Exit Function
ErrHandler:
  RaiseError ClassId, "HasWildCharacters", Err.Number
End Function
