Eliminar TMP_ABONOSPERIODTBL
Eliminar TMP_ABONOSTBL
Eliminar TMP_CARGOSPERIODTBL
Eliminar TMP_CARGOSTBL
Eliminar TMP_SALDOSINICIALESTBL
Eliminar TMP_SI_TBL
Eliminar TMP_TOTABONOSSUMPERTBL  
Eliminar TMP_TOTABONOSTBL
Eliminar TMP_TOTCARGOSSUMPERTBL  
Eliminar TMP_TOTCARGOSTBL
Eliminar TMP_TOTSI_TBL

CREATE GLOBAL TEMPORARY TABLE TMP_Balances
(
 ID_MAYOR        		NUMBER(12) ,
 ID_MONEDA        		NUMBER(12) ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256) ,
 ID_SECTOR		  	NUMBER(12) ,
 ID_CUENTA_AUXILIAR		NUMBER(12) ,
 NUMERO_CUENTA_AUXILIAR		VARCHAR2(64) ,
 NOMBRE_CUENTA_AUXILIAR		VARCHAR2(256) ,
 SALDO_INICIAL			NUMBER ,
 CARGOS		  		NUMBER ,
 ABONOS		  		NUMBER ,
 SALDO_ACTUAL	  		NUMBER ,
 SALDO_NV			NUMBER ,
 FECHA_ULTIMO_MOVIMIENTO	DATE
) ON COMMIT PRESERVE ROWS;

CREATE GLOBAL TEMPORARY TABLE TMP_Balances_With_Vouchers
(
 ID_MAYOR        		NUMBER(12) ,
 ID_TRANSACCION			NUMBER(12) ,
 NUMERO_TRANSACCION		VARCHAR2(16) ,
 FECHA_AFECTACION		DATE ,
 ID_MONEDA        		NUMBER(12) ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256) ,
 ID_SECTOR		  	NUMBER(12) ,
 ID_CUENTA_AUXILIAR		NUMBER(12) ,
 NUMERO_CUENTA_AUXILIAR		VARCHAR2(64) ,
 NOMBRE_CUENTA_AUXILIAR		VARCHAR2(256) ,
 SALDO_INICIAL			NUMBER ,
 CARGOS		  		NUMBER ,
 ABONOS		  		NUMBER ,
 SALDO_ACTUAL	  		NUMBER ,
 SALDO_NV			NUMBER ,
 FECHA_ULTIMO_MOVIMIENTO	DATE
) ON COMMIT PRESERVE ROWS;

CREATE GLOBAL TEMPORARY TABLE TMP_Average_Balances
(
 FECHA_SALDO			DATE , 
 ID_MAYOR        		NUMBER(12) ,
 ID_MONEDA        		NUMBER(12) ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256) ,
 ID_SECTOR		  	NUMBER(12) ,
 ID_CUENTA_AUXILIAR		NUMBER(12) ,
 NUMERO_CUENTA_AUXILIAR		VARCHAR2(64) ,
 NOMBRE_CUENTA_AUXILIAR		VARCHAR2(256) ,
 SALDO_INICIAL			NUMBER ,
 CARGOS		  		NUMBER ,
 ABONOS		  		NUMBER ,
 SALDO_ACTUAL	  		NUMBER ,
 SALDO_NV			NUMBER ,
 SALDO_PROMEDIO			NUMBER ,
 SALDO_PROMEDIO_NV		NUMBER 
) ON COMMIT PRESERVE ROWS;

CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_MAYOR
(
 ID_MAYOR        		NUMBER(12) ,
 ID_MONEDA        		NUMBER(12) ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256) ,
 ID_SECTOR		  	NUMBER(12) ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12) ,
 NUMERO_CUENTA_AUXILIAR		VARCHAR2(256) ,
 SALDO_INICIAL			NUMBER ,
 CARGOS		  		NUMBER ,
 ABONOS		  		NUMBER ,
 SALDO_ACTUAL	  		NUMBER
) ON COMMIT PRESERVE ROWS;

CREATE GLOBAL TEMPORARY TABLE TMP_MOVSTBL
(
 ID_CUENTA	          NUMBER(12)  ,	 
 ID_SECTOR	          NUMBER(12)  ,	 
 ID_CUENTA_AUXILIAR       NUMBER(12)  ,	 
 TIPO_MOVIMIENTO          CHAR(1)  ,	 
 ID_MONEDA                NUMBER(12)  ,	 
 MONTO                    NUMBER(25,10)  ,	 
 MONTO_MONEDA_BASE        NUMBER(25,10)  ,	 
 SALDO_INICIAL		  NUMBER,
 FECHA_SALDO_INICIAL	  DATE,
 ID_MAYOR                 NUMBER(12)  ,	 
 FECHA_AFECTACION         DATE  ,
 NUMERO_TRANSACCION       VARCHAR2(16) ,
 FROM_MOVS 	 	  NUMBER(1)  ,
 CONCEPTO_TRANSACCION     VARCHAR2(1024)
) ON COMMIT PRESERVE ROWS;


CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256)  ,
 ID_SECTOR		  	NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12)  ,
 SALDO_INICIAL			NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER
) ON COMMIT PRESERVE ROWS;


CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_B
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256)  ,
 ID_SECTOR		  	NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12)  ,
 SALDO_INICIAL			NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER
) ON COMMIT PRESERVE ROWS;


CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_EXT
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR	  	VARCHAR2(256)  ,
 ID_SECTOR	  		NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR  		NUMBER(12)  ,
 SALDO_INICIAL	  		NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER,
 SALDO_INICIAL_MONEDA_BASE	NUMBER,
 CARGOS_MONEDA_BASE  		NUMBER,
 ABONOS_MONEDA_BASE  		NUMBER,
 SALDO_ACTUAL_MONEDA_BASE	NUMBER
) ON COMMIT PRESERVE ROWS;

CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_EXT_PROM
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR	  	VARCHAR2(256)  ,
 ID_SECTOR	  		NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR  		NUMBER(12)  ,
 FECHA_REV			DATE,
 SALDO_INICIAL	  		NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER,
 SALDO_INICIAL_MONEDA_BASE	NUMBER,
 CARGOS_MONEDA_BASE  		NUMBER,
 ABONOS_MONEDA_BASE  		NUMBER,
 SALDO_ACTUAL_MONEDA_BASE	NUMBER
) ON COMMIT PRESERVE ROWS;


CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_PROM
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256)  ,
 ID_SECTOR		  	NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12)  ,
 FECHA_REV			DATE,
 SALDO_INICIAL			NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER
) ON COMMIT PRESERVE ROWS;


CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_PROM_B
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256)  ,
 ID_SECTOR		  	NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12)  ,
 FECHA_REV			DATE,
 SALDO_INICIAL			NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_INICIAL_ALL		NUMBER,
 CARGOS_ALL	  		NUMBER,
 ABONOS_ALL	  		NUMBER
) ON COMMIT PRESERVE ROWS;


/*****SIGRO*******/

OOJJOO: ¿Esta tabla se utiliza? 	OOJJOO
CREATE TABLE SCON_SALDOS
(
 scon_anio			NUMBER(4) NOT NULL,
 scon_mes			NUMBER(2) NOT NULL,
 scon_area			VARCHAR2(8) NOT NULL,
 scon_moneda                    NUMBER(3) NOT NULL,
 scon_numero_mayor		VARCHAR2(12) NOT NULL,
 scon_cuenta			VARCHAR2(256) NOT NULL,
 scon_sector			VARCHAR2(2) NOT NULL,
 scon_auxiliar			VARCHAR2(20) NOT NULL,
 scon_fecha_ultimo_movimiento   DATE,
 scon_saldo			NUMBER,
 scon_moneda_origen		NUMBER,
 scon_naturaleza_cuenta		NUMBER,
 scon_saldo_promedio            NUMBER,
 scon_monto_debito		NUMBER,
 scon_monto_credito		NUMBER,
 scon_saldo_anterior		NUMBER,
 scon_empresa			NUMBER
) ;

CREATE GLOBAL TEMPORARY TABLE TMP_SALDOS_PROM_SIGRO
(
 ID_MONEDA        		NUMBER(12)  ,
 NUMERO_CUENTA_ESTANDAR		VARCHAR2(256)  ,
 ID_SECTOR		  	NUMBER(12)  ,	  
 ID_CUENTA_AUXILIAR		NUMBER(12)  ,
 FECHA_REV			DATE,
 SALDO_INICIAL			NUMBER,
 CARGOS		  		NUMBER,
 ABONOS		  		NUMBER,
 SALDO_ACTUAL	  		NUMBER
) ON COMMIT PRESERVE ROWS;



SELECT (0) id_mayor, id_transaccion, numero_transaccion, fecha_afectacion, id_moneda, 
       numero_cuenta_estandar, id_sector, A.id_cuenta_auxiliar, 
       MAX(D.numero_cuenta_auxiliar) numero_cuenta_auxiliar, MAX(D.nombre_cuenta_auxiliar) nombre_cuenta_auxiliar, 
       SUM(saldo_inicial) saldo_inicial, SUM(cargos) cargos, SUM(abonos) abonos, 
       (DECODE(naturaleza, 'D', (SUM(saldo_inicial) + SUM(cargos) - SUM(abonos)), 
       (SUM(saldo_inicial) - SUM(cargos) + SUM(abonos)) )) saldo_actual, 
       (DECODE(naturaleza, 'D', (SUM(saldo_inicial) + SUM(cargos) - SUM(abonos)), 
       (SUM(saldo_inicial) - SUM(cargos) + SUM(abonos)) )) saldo_nv, 
       MAX(fecha_ultimo_movimiento) fecha_ultimo_movimiento 
FROM 
    (
       (
          SELECT (0) id_transaccion, ('') numero_transaccion, (to_date('1/1/1990')) fecha_afectacion, 
              A.id_moneda, A.id_cuenta, A.id_sector, A.id_cuenta_auxiliar, 
              (DECODE(naturaleza, 'D', (SUM(saldo_inicial) + SUM(cargos) - SUM(abonos)), 
              (SUM(saldo_inicial) - SUM(cargos) + SUM(abonos)))) saldo_inicial, 
              (0) cargos, (0) abonos, MAX(fecha_ultimo_movimiento) fecha_ultimo_movimiento 
          FROM 
              (
                 (SELECT id_moneda, id_cuenta, id_sector, id_cuenta_auxiliar, 
                     SUM(saldo_inicial) saldo_inicial, (0) cargos, (0) abonos, 
                     MAX(fecha_saldo_inicial) fecha_ultimo_movimiento 
                  FROM COF_Saldo_Inicial 
                  WHERE (id_mayor = 271) AND (fecha_saldo_inicial <= '31/01/2001') 
                  GROUP BY id_moneda, id_cuenta, id_sector, id_cuenta_auxiliar
                 ) 
                 UNION 
                 (SELECT id_moneda, id_cuenta, id_sector, id_cuenta_auxiliar, 
                     (0) saldo_inicial, (SUM(DECODE(tipo_movimiento, 'D', monto , 0))) cargos, 
                     (SUM(DECODE(tipo_movimiento, 'H', monto, 0))) abonos, 
                     MAX(fecha_afectacion) fecha_ultimo_movimiento 
                  FROM COF_Transaccion, COF_Movimiento 
                  WHERE (id_mayor = 271) AND (fecha_afectacion <= '31/12/2000') AND 
                        (COF_Transaccion.id_transaccion = COF_Movimiento.id_transaccion) 
                  GROUP BY id_moneda, id_cuenta, id_sector, id_cuenta_auxiliar
                  )
               ) A, COF_Cuenta B, 
                    (SELECT MAX(id_cuenta_estandar_hist) id_cuenta_estandar_hist, 
                         id_cuenta_estandar, numero_cuenta_estandar, naturaleza, id_tipo_cuenta 
                     FROM COF_Cuenta_Estandar_Hist 
                     WHERE (id_tipo_cuentas_std = 2) AND 
                           (TRUNC(TO_DATE(fecha_inicio)) <= '31/01/2001') AND 
                           (TRUNC(TO_DATE(fecha_fin)) >= '01/01/2001') 
                     GROUP BY id_cuenta_estandar, numero_cuenta_estandar, naturaleza, id_tipo_cuenta
                    ) C 
           WHERE (A.id_cuenta = B.id_cuenta) AND (B.id_cuenta_estandar = C.id_cuenta_estandar) 
           GROUP BY id_transaccion, numero_transaccion, fecha_afectacion, A.id_moneda, naturaleza, 
                    A.id_cuenta, A.id_sector, A.id_cuenta_auxiliar
     ) 
    UNION 
    (
      SELECT COF_Transaccion.id_transaccion, numero_transaccion, fecha_afectacion, id_moneda, 
           id_cuenta, id_sector, id_cuenta_auxiliar, (0) saldo_inicial, 
           (SUM(DECODE(tipo_movimiento, 'D', monto, 0))) cargos, 
           (SUM(DECODE(tipo_movimiento, 'H', monto, 0))) abonos, 
           MAX(fecha_afectacion) fecha_ultimo_movimiento 
      FROM COF_Transaccion, COF_Movimiento 
      WHERE (id_mayor = 271) AND (fecha_afectacion >= '01/01/2001') AND 
            (fecha_afectacion <= '31/01/2001') AND 
            (COF_Movimiento.id_transaccion = COF_Transaccion.id_transaccion) 
      GROUP BY COF_Transaccion.id_transaccion, numero_transaccion, fecha_afectacion, 
               id_moneda, id_cuenta, id_sector, id_cuenta_auxiliar
    )
) A, COF_Cuenta B, 
  (SELECT MAX(id_cuenta_estandar_hist) id_cuenta_estandar_hist, id_cuenta_estandar, 
        numero_cuenta_estandar, naturaleza, id_tipo_cuenta 
   FROM COF_Cuenta_Estandar_Hist 
   WHERE (id_tipo_cuentas_std = 2) AND (TRUNC(TO_DATE(fecha_inicio)) <= '31/01/2001') AND 
        (TRUNC(TO_DATE(fecha_fin)) >= '01/01/2001') 
   GROUP BY id_cuenta_estandar, numero_cuenta_estandar, naturaleza, id_tipo_cuenta
   ) C , COF_Cuenta_Auxiliar D 
WHERE (A.id_cuenta = B.id_cuenta) AND (B.id_cuenta_estandar = C.id_cuenta_estandar) AND 
  (A.id_cuenta_auxiliar = D.id_cuenta_auxiliar(+))
GROUP BY id_transaccion, numero_transaccion, fecha_afectacion, id_moneda, 
         numero_cuenta_estandar, naturaleza, id_sector, A.id_cuenta_auxiliar
/

