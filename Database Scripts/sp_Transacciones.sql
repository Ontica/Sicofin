CREATE PROCEDURE sp_val_posting_values (p_id_cuenta NUMBER, p_sector VARCHAR2, p_cuenta_auxiliar VARCHAR2, 
				 	p_result OUT NUMBER) AS
  nRes NUMBER;
  nStdAccountId NUMBER;
  nSectorId NUMBER;
BEGIN
      
  IF (Length(p_sector) = 0) THEN
    SELECT COUNT(COF_CUENTA.ID_CUENTA_ESTANDAR) INTO nRes FROM COF_MAPEO_SECTOR, COF_CUENTA 
    WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR) AND
    (COF_CUENTA.ID_CUENTA = p_id_cuenta);

      If (nRes != 0) Then
          p_result := -1;
          return;        
      End If;
  End If;

    IF (Length(p_sector) != 0) Then
      SELECT COUNT(*) INTO nRes FROM COF_CUENTA, COF_MAPEO_SECTOR, COF_SECTOR 
       WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR) AND 
              (COF_SECTOR.ID_SECTOR = COF_MAPEO_SECTOR.ID_SECTOR) AND
              (COF_CUENTA.ID_CUENTA =  p_id_cuenta );
                
        If (nRes = 0) Then
          p_result := -2;
          return;
        Else
          SELECT COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR INTO nStdAccountId FROM COF_CUENTA, COF_MAPEO_SECTOR, COF_SECTOR 
          WHERE (COF_CUENTA.ID_CUENTA_ESTANDAR = COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR) AND 
              (COF_SECTOR.ID_SECTOR = COF_MAPEO_SECTOR.ID_SECTOR) AND
              (COF_CUENTA.ID_CUENTA =  p_id_cuenta );          
        End If;
      End If;
            
      If (Length(p_sector) != 0) Then
        SELECT COUNT(COF_MAPEO_SECTOR.ID_SECTOR) INTO nRes FROM COF_MAPEO_SECTOR, COF_SECTOR
               WHERE (COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR = nStdAccountId ) AND
               (COF_MAPEO_SECTOR.ID_SECTOR = COF_SECTOR.ID_SECTOR) AND 
               (COF_SECTOR.CLAVE_SECTOR = p_sector);
        
        If (nRes = 0) Then
          p_result := -3;
          return;
        Else
          SELECT COF_MAPEO_SECTOR.ID_SECTOR INTO nSectorId FROM COF_MAPEO_SECTOR, COF_SECTOR
               WHERE (COF_MAPEO_SECTOR.ID_CUENTA_ESTANDAR = nStdAccountId ) AND
               (COF_MAPEO_SECTOR.ID_SECTOR = COF_SECTOR.ID_SECTOR) AND 
               (COF_SECTOR.CLAVE_SECTOR = p_sector);
        End If;
      End If;
           
      If (Length(p_cuenta_auxiliar) = 0) Then
        SELECT count(*) INTO nRes FROM COF_MAPEO_MAYOR_AUXILIAR
               WHERE (COF_MAPEO_MAYOR_AUXILIAR.ID_CUENTA = p_id_cuenta) AND
               (COF_MAPEO_MAYOR_AUXILIAR.ID_SECTOR = nSectorId);

        If (nRes != 0) Then
          p_result := -4;
          return;
        End If;
      End If;
  
      If (Length(p_cuenta_auxiliar) != 0) Then
        SELECT count(*) into nRes FROM COF_MAPEO_MAYOR_AUXILIAR 
               WHERE (COF_MAPEO_MAYOR_AUXILIAR.ID_CUENTA = p_id_cuenta) AND 
               (COF_MAPEO_MAYOR_AUXILIAR.ID_SECTOR = nSectorId);
        
        If (nRes = 0) Then
          p_result := -5;
          return;
        End If;
      End If;
      
      If (Length(p_cuenta_auxiliar) != 0) Then
        SELECT COUNT(*) INTO nRes FROM COF_MAPEO_MAYOR_AUXILIAR, COF_CUENTA_AUXILIAR 
               WHERE (COF_MAPEO_MAYOR_AUXILIAR.ID_CUENTA = p_id_cuenta) AND
               (COF_MAPEO_MAYOR_AUXILIAR.ID_SECTOR = nSectorId) AND
               (COF_MAPEO_MAYOR_AUXILIAR.ID_MAYOR_AUXILIAR = COF_CUENTA_AUXILIAR.ID_MAYOR_AUXILIAR) AND 
               (COF_CUENTA_AUXILIAR.NUMERO_CUENTA_AUXILIAR = p_cuenta_auxiliar);

        If (nRes = 0) Then
          p_result := -6;
          return;
        End If;
      End If;

p_result := 1;
END;                               	
/

CREATE PROCEDURE sp_post_transaction(p_id_transaccion NUMBER, p_numero_transaccion VARCHAR2, p_id_autorizada_por NUMBER, p_result OUT NUMBER) AS
  nRes  	   NUMBER;
  nGLId            NUMBER;
  nVoucherTypeId   NUMBER;
BEGIN        
        LOCK TABLE COF_TRANSACCION IN EXCLUSIVE MODE;

	SELECT id_mayor INTO nGLId FROM COF_TRANSACCION
        WHERE (id_transaccion = p_id_transaccion);

        SELECT id_tipo_poliza INTO nVoucherTypeId FROM COF_TRANSACCION
        WHERE (id_transaccion = p_id_transaccion);        
	
	SELECT COUNT(*) INTO nRes FROM COF_TRANSACCION
	WHERE (numero_transaccion = p_numero_transaccion) AND (id_mayor = nGLId);

	If (nRes != 0) Then
           ROLLBACK;
	   p_result := -1;
           RETURN;
	End If;

        If (nVoucherTypeId != 28) Then
		UPDATE COF_TRANSACCION
	        SET esta_abierta = 0, numero_transaccion = p_numero_transaccion, 
        	    id_autorizada_por = p_id_autorizada_por, fecha_registro = SYSDATE
	        WHERE (id_transaccion = p_id_transaccion);
	Else
		UPDATE COF_TRANSACCION
	        SET esta_abierta = 0, numero_transaccion = p_numero_transaccion, 
        	    id_autorizada_por = p_id_autorizada_por
	        WHERE (id_transaccion = p_id_transaccion);	
	End If;

	SELECT COUNT(*) INTO nRes FROM COF_TRANSACCION
	WHERE (numero_transaccion = p_numero_transaccion) AND (id_mayor = nGLId);

	If (nRes > 1) Then
           ROLLBACK;
	   p_result := -1;
           RETURN;
	End If;
	
	INSERT INTO COF_MOVIMIENTO
		(SELECT id_movimiento_tmp id_movimiento, id_transaccion, id_cuenta, id_cuenta_auxiliar, 
		 id_sector, id_movimiento_referencia, id_area_responsabilidad, 
		 clave_presupuestal, clave_disponibilidad, numero_verificacion, 
		 tipo_movimiento, fecha_movimiento, concepto_movimiento, 
		 id_moneda, monto, monto_moneda_base
		 FROM COF_MOVIMIENTO_TMP 
		 WHERE (id_transaccion = p_id_transaccion)
                );
           		
	DELETE FROM COF_MOVIMIENTO_TMP WHERE (id_transaccion = p_id_transaccion);
         
	COMMIT;
        p_result := 1;
        RETURN;
EXCEPTION
   WHEN OTHERS THEN
      p_result := 0;
      ROLLBACK;
      RETURN;
END;
/


CREATE PROCEDURE sp_del_transaction(p_id_transaccion NUMBER, p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN

    SELECT COUNT(ID_TRANSACCION) INTO nRes FROM COF_TRANSACCION
    WHERE (id_transaccion = p_id_transaccion) AND (esta_abierta = 1);
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;

    DELETE FROM cof_movimiento_tmp WHERE (id_transaccion = p_id_transaccion);
    DELETE FROM cof_transaccion WHERE (id_transaccion = p_id_transaccion);
    commit;

    p_result := 1;
END;
/

CREATE PROCEDURE sp_del_posting(p_id_movimiento_tmp NUMBER, p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN

    SELECT COUNT(ID_MOVIMIENTO_TMP) INTO nRes FROM COF_MOVIMIENTO_TMP
    WHERE (ID_MOVIMIENTO_TMP = p_id_movimiento_tmp);
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;

    DELETE FROM cof_movimiento_tmp WHERE (id_movimiento_tmp = p_id_movimiento_tmp);
    commit;

    p_result := 1;
END;
/

CREATE PROCEDURE sp_apd_transaction(p_numero_transaccion VARCHAR2, p_id_mayor NUMBER, p_id_fuente NUMBER, 
				    p_id_tipo_transaccion NUMBER, p_id_tipo_poliza NUMBER, p_concepto_transaccion VARCHAR2, 
			       	    p_fecha_afectacion DATE, p_fecha_registro DATE, 
                               	    p_id_elaborada_por NUMBER, p_id_autorizada_por NUMBER, 
				    p_result OUT NUMBER) AS 
  nId NUMBER;
BEGIN

   SELECT SEC_ID_TRANSACCION.NEXTVAL INTO nId FROM DUAL;
   
   INSERT INTO COF_Transaccion(id_transaccion, numero_transaccion, id_mayor, id_fuente, 
			       id_tipo_transaccion, id_tipo_poliza, concepto_transaccion, 
			       fecha_afectacion, fecha_registro, 
                               id_elaborada_por, id_autorizada_por, esta_abierta)
   	               VALUES (nId, p_numero_transaccion, p_id_mayor, p_id_fuente, 
			       p_id_tipo_transaccion, p_id_tipo_poliza, p_concepto_transaccion, 
			       p_fecha_afectacion, p_fecha_registro, 
                               p_id_elaborada_por, p_id_autorizada_por, 1);

   p_result := nId;
   commit;
END;
/

CREATE PROCEDURE sp_upd_transaction(p_id_transaccion NUMBER, p_numero_transaccion VARCHAR2, p_id_mayor NUMBER, p_id_fuente NUMBER, 
				    p_id_tipo_transaccion NUMBER, p_id_tipo_poliza NUMBER, p_concepto_transaccion VARCHAR2, 
			       	    p_fecha_afectacion DATE, p_fecha_registro DATE, 
                               	    p_id_elaborada_por NUMBER, p_id_autorizada_por NUMBER, 
				    p_result OUT NUMBER) AS 
  nRes NUMBER;
BEGIN

   SELECT COUNT(id_transaccion) INTO nRes FROM COF_Transaccion 
          WHERE (id_transaccion = p_id_transaccion) AND (esta_abierta = 1); 
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;
   
   UPDATE COF_Transaccion SET id_mayor = p_id_mayor, id_fuente = p_id_fuente, id_tipo_transaccion = p_id_tipo_transaccion, 
          id_tipo_poliza = p_id_tipo_poliza, concepto_transaccion = p_concepto_transaccion, fecha_afectacion = p_fecha_afectacion, 
	  fecha_registro = p_fecha_registro, id_elaborada_por = p_id_elaborada_por, id_autorizada_por = p_id_autorizada_por, 
          esta_abierta = 1          
   WHERE id_transaccion = p_id_transaccion AND esta_abierta = 1;

   commit;
   p_result := 1;
END;
/

CREATE PROCEDURE sp_apd_posting(p_id_transaccion NUMBER, p_id_cuenta NUMBER, p_id_cuenta_auxiliar NUMBER, 
				p_id_sector NUMBER, p_id_movimiento_referencia NUMBER, 
				p_id_area_responsabilidad NUMBER, p_clave_presupuestal VARCHAR2, 
				p_clave_disponibilidad VARCHAR2, p_numero_verificacion VARCHAR2, 
				p_tipo_movimiento CHAR, p_fecha_movimiento DATE, p_concepto_movimiento VARCHAR2,
				p_id_moneda NUMBER, p_monto NUMBER, p_monto_moneda_base NUMBER, p_protegido NUMBER,
				p_result OUT NUMBER) AS 
  nId NUMBER;
BEGIN

   SELECT SEC_ID_MOVIMIENTO_TMP.NEXTVAL INTO nId FROM DUAL;
   
   INSERT INTO COF_Movimiento_Tmp(id_movimiento_tmp, id_transaccion, id_cuenta, id_cuenta_auxiliar, 
			          id_sector, id_movimiento_referencia, id_area_responsabilidad, 
				  clave_presupuestal, clave_disponibilidad, numero_verificacion, 
			          tipo_movimiento, fecha_movimiento, concepto_movimiento, 
			          id_moneda, monto, monto_moneda_base, protegido) 
   	                  VALUES (nId, p_id_transaccion, p_id_cuenta, p_id_cuenta_auxiliar, 
			          p_id_sector, p_id_movimiento_referencia, p_id_area_responsabilidad, 
				  p_clave_presupuestal, p_clave_disponibilidad, p_numero_verificacion, 
			          p_tipo_movimiento, p_fecha_movimiento, p_concepto_movimiento, 
			          p_id_moneda, p_monto, p_monto_moneda_base, p_protegido); 
	
   p_result := nId;
   commit;
END;
/

CREATE PROCEDURE sp_upd_posting(p_id_transaccion NUMBER, p_id_movimiento_tmp NUMBER, p_id_cuenta NUMBER, p_id_cuenta_auxiliar NUMBER, 
				p_id_sector NUMBER, p_id_movimiento_referencia NUMBER, 
				p_id_area_responsabilidad NUMBER, p_clave_presupuestal VARCHAR2,
				p_clave_disponibilidad VARCHAR2, p_numero_verificacion VARCHAR2, 
				p_tipo_movimiento CHAR, p_fecha_movimiento DATE, p_concepto_movimiento VARCHAR2,
				p_id_moneda NUMBER, p_monto NUMBER, p_monto_moneda_base NUMBER, p_protegido NUMBER,
				p_result OUT NUMBER) AS 
  nRes NUMBER;
BEGIN

   SELECT COUNT(id_transaccion) INTO nRes FROM COF_Transaccion 
          WHERE (id_transaccion = p_id_transaccion) AND (esta_abierta = 1); 
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;
 
   UPDATE COF_Movimiento_Tmp SET id_cuenta = p_id_cuenta, id_cuenta_auxiliar = p_id_cuenta_auxiliar, id_sector = p_id_sector, 
          id_movimiento_referencia = p_id_movimiento_referencia, id_area_responsabilidad = p_id_area_responsabilidad, 
	  clave_presupuestal = p_clave_presupuestal, clave_disponibilidad = p_clave_disponibilidad, numero_verificacion = p_numero_verificacion, 
          tipo_movimiento = p_tipo_movimiento, fecha_movimiento = p_fecha_movimiento, concepto_movimiento = p_concepto_movimiento, 
	  id_moneda = p_id_moneda, monto = p_monto, monto_moneda_base = p_monto_moneda_base, protegido = p_protegido
   WHERE id_transaccion = p_id_transaccion AND id_movimiento_tmp = p_id_movimiento_tmp;

   commit;
   p_result := 1;
END;
/


CREATE PROCEDURE sp_upd_transaction_check_user(p_id_transaccion NUMBER, p_id_autorizada_por NUMBER, p_result OUT NUMBER) AS 
  nRes NUMBER;
BEGIN

   SELECT COUNT(id_transaccion) INTO nRes FROM COF_Transaccion 
          WHERE (id_transaccion = p_id_transaccion) AND (esta_abierta = 1); 
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;
 
   UPDATE COF_Transaccion SET id_autorizada_por = p_id_autorizada_por 
   WHERE id_transaccion = p_id_transaccion;

   commit;
   p_result := 1;
END;
/

CREATE PROCEDURE sp_upd_transaction_user(p_id_transaccion NUMBER, p_id_elaborada_por NUMBER, p_result OUT NUMBER) AS 
  nRes NUMBER;
BEGIN

   SELECT COUNT(id_transaccion) INTO nRes FROM COF_Transaccion 
          WHERE (id_transaccion = p_id_transaccion) AND (esta_abierta = 1); 
      If (nRes = 0) Then
          p_result := -1;
          return;
      End If;
 
   UPDATE COF_Transaccion SET id_elaborada_por = p_id_elaborada_por
   WHERE id_transaccion = p_id_transaccion;

   commit;
   p_result := 1;
END;
/

