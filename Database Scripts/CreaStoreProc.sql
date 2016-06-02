
CREATE PROCEDURE pcodelcategoria (Pid_categoria NUMBER,
                                                                            PResultado OUT NUMBER ) AS
  Subcategorias NUMBER;
BEGIN
  SELECT COUNT(*) INTO Subcategorias FROM CON_categoria 
       					 WHERE ( CON_categoria.id_categoria_padre = Pid_categoria );

  IF ( Subcategorias = 0 ) THEN
      DELETE FROM CON_categoria WHERE ( CON_categoria.id_categoria  = Pid_categoria);
      PResultado := 0;
      Return;
  END IF;
  PResultado := 1 ;
END;

/


CREATE PROCEDURE pcodelcuenta (Pid_cuenta NUMBER) AS
BEGIN
	DELETE FROM CON_cuenta WHERE ( CON_cuenta.id_cuenta = Pid_cuenta );
END;

/

CREATE PROCEDURE pcodelcuentaestandar (Pid_cuenta_estandar NUMBER) AS
BEGIN
	DELETE FROM CON_cuenta_estandar 
               WHERE ( CON_cuenta_estandar.id_cuenta_estandar = Pid_cuenta_estandar);
END;

/

CREATE PROCEDURE pcodelcuentaestandarsector (Pid_cuenta_estandar NUMBER, Pid_sector NUMBER) AS
BEGIN
  DELETE FROM CON_cuenta_est_sector 
         WHERE (CON_cuenta_est_sector.id_cuenta_estandar = Pid_cuenta_estandar) AND
               (CON_cuenta_est_sector.id_sector = Pid_sector);
END;

/

CREATE PROCEDURE pcodelfuente (Pid_fuente NUMBER,
                                                                        PResultado OUT NUMBER ) AS
  Relaciones NUMBER;
BEGIN
  SELECT COUNT(*) INTO Relaciones FROM CON_fuente_mayor
       			  	WHERE ( CON_fuente_mayor.id_fuente = Pid_fuente );

  IF ( Relaciones = 0 ) THEN
      DELETE FROM CON_fuente WHERE (CON_fuente.id_fuente  = Pid_fuente);
      PResultado := 0;
      Return;
  END IF;
  PResultado := 1 ;
END;

/

CREATE PROCEDURE pcodelfuentemayor (Pid_fuente NUMBER, Pid_mayor NUMBER) AS
BEGIN
  DELETE FROM CON_fuente_mayor 
         WHERE (CON_fuente_mayor.id_fuente = Pid_fuente) AND
               (CON_fuente_mayor.id_mayor = Pid_mayor);
END;

/


CREATE PROCEDURE pcodelmayor (Pid_mayor NUMBER) AS
BEGIN
	DELETE FROM CON_mayor WHERE (CON_mayor.id_mayor = Pid_mayor );
END;

/

CREATE PROCEDURE pcodelmovimiento (Pid_movimiento NUMBER) AS
BEGIN
	DELETE FROM CON_movimiento WHERE ( CON_movimiento.id_movimiento = Pid_movimiento);
END;

/

CREATE PROCEDURE pcodelmovimientotmp (Pid_movimiento NUMBER) AS
BEGIN
  DELETE FROM CON_movimiento_tmp WHERE ( CON_movimiento_tmp.id_movimiento = Pid_movimiento);
END;

/

CREATE PROCEDURE pcodelperiodo (Pid_periodo NUMBER) AS
BEGIN
	DELETE FROM CON_periodo	WHERE (CON_periodo.id_periodo = Pid_periodo);
END;

/

CREATE PROCEDURE pcodelsector (Pid_sector NUMBER) AS
BEGIN
  DELETE FROM CON_sector WHERE (CON_sector.id_sector = Pid_sector);
END;

/

CREATE PROCEDURE pcodeltipocambio (Pid_tipo_cambio NUMBER) AS
BEGIN
	DELETE FROM CON_tipo_cambio WHERE (CON_tipo_cambio.id_tipo_cambio = Pid_tipo_cambio );
END;

/

CREATE PROCEDURE pcodeltransaccion (Pid_transaccion NUMBER) AS  
  EstaAbierta NUMBER;
BEGIN
  SELECT esta_abierta INTO EstaAbierta FROM CON_transaccion 
                                       WHERE (CON_transaccion.id_transaccion = Pid_transaccion ); 
  IF ( EstaAbierta = 0) THEN
    DELETE FROM CON_movimiento WHERE (CON_movimiento.id_transaccion = Pid_transaccion);
  ELSE
    DELETE FROM CON_movimiento_tmp WHERE (CON_movimiento_tmp.id_transaccion = Pid_transaccion);
  END IF;

  DELETE FROM CON_transaccion WHERE (CON_transaccion.id_transaccion = Pid_transaccion);
END;

/

CREATE PROCEDURE pcodelmoneda (Pid_moneda NUMBER,
                                                                            PResultado OUT NUMBER ) AS
  CuantosTipoCambio NUMBER;
BEGIN
  SELECT COUNT(*) INTO CuantosTipoCambio FROM CON_tipo_cambio 
       					 WHERE (CON_tipo_cambio.id_moneda_base = Pid_moneda) OR 
             				       (CON_tipo_cambio.id_moneda_derivada = Pid_moneda);
  IF (CuantosTipoCambio = 0) THEN
      DELETE FROM CON_moneda WHERE ( CON_moneda.id_moneda  = Pid_moneda);
      PResultado := 0;
      Return;
  END IF;
  PResultado := 1 ;
END;

/

