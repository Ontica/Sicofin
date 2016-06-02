
CREATE PROCEDURE sp_del_Address (p_address_id NUMBER,
				  p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_mayor WHERE (COF_mayor.id_domicilio = p_address_id AND COF_mayor.eliminado = 0);
 
  IF (Num = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_domicilio WHERE (COF_domicilio.id_domicilio = p_address_id AND COF_domicilio.eliminado = 0);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      ELSE
        UPDATE COF_domicilio SET Eliminado = 1 WHERE (Id_Domicilio = p_address_id);
        p_result :=  Num;
        Return;       
      END IF; 
  END IF;
    p_result :=  -1;
END;

/

CREATE PROCEDURE sp_del_calendar (p_calendar_id NUMBER, 
                                                                              p_result OUT NUMBER) AS
   Num NUMBER;
   Num2 NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM AO_periods WHERE (AO_periods.calendar_id = p_calendar_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mayor WHERE (COF_mayor.id_calendario = p_calendar_id) AND (COF_mayor.eliminado = 0);

  IF (Num = 0 AND Num2 = 0) THEN
    SELECT COUNT(*) INTO Num FROM AO_calendars WHERE (AO_calendars.calendar_id = p_calendar_id);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      ELSE
        DELETE FROM AO_calendars WHERE ( AO_calendars.calendar_id = p_calendar_id );
        p_result :=  Num;
        Return;       
      END IF; 
  END IF;
  p_result := -1;
END;

/

CREATE PROCEDURE sp_del_Currency (p_currency_id NUMBER,
                                                                                  p_result OUT NUMBER ) AS
  nExcRate NUMBER;
  nMapCurrency NUMBER;
  nGenLed NUMBER;
  Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO nExcRate FROM AO_exchange_rates WHERE (AO_exchange_rates.from_currency_id = p_currency_id) 
							OR (AO_exchange_rates.to_currency_id = p_currency_id);

  SELECT COUNT(*) INTO nMapCurrency FROM COF_mapeo_moneda WHERE (COF_mapeo_moneda.id_moneda = p_currency_id);
  SELECT COUNT(*) INTO nGenLed FROM COF_mayor WHERE (COF_mayor.id_moneda_base = p_currency_id);

  IF (nExcRate = 0 AND nMapCurrency = 0 AND nGenLed = 0) THEN
        SELECT COUNT(*) INTO Num FROM AO_currencies WHERE (AO_currencies.currency_id = p_currency_id);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
             DELETE FROM AO_currencies WHERE ( AO_currencies.currency_id  = p_currency_id);
        p_result := Num;
        Return;
      END IF; 
  END IF;  
     p_result  := -1 ;
END;

/

CREATE PROCEDURE sp_del_Exchange_Rate (p_exchange_rate_id NUMBER, 
					p_result OUT NUMBER) AS
Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_exchange_rates WHERE (AO_exchange_rates.exchange_rate_id = p_exchange_rate_id);
    IF (Num = 0) then
      p_result := 0;
      Return;
    END IF;
      DELETE FROM AO_exchange_rates WHERE (AO_exchange_rates .exchange_rate_id = p_exchange_rate_id);
      p_result := Num;      
END;

/


CREATE PROCEDURE sp_del_Organization (p_organization_id NUMBER,
				          p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_mayor WHERE (COF_mayor.id_empresa = p_organization_id AND COF_mayor.eliminado = 0);
 
  IF (Num = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_empresa WHERE (COF_empresa.id_empresa = p_organization_id AND COF_empresa.eliminada = 0);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      ELSE
        UPDATE COF_empresa SET Eliminada = 1 WHERE (Id_Empresa = p_organization_id);
        p_result :=  Num;
        Return;       
      END IF; 
  END IF;
    p_result :=  -1;
END;

/

CREATE PROCEDURE sp_del_Budget_Struct (p_budget_struct_id NUMBER,
				            p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_movimiento WHERE (COF_movimiento.id_clave_presupuestal = p_budget_struct_id);
 
  IF (Num = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_estr_presupuestal WHERE (COF_estr_presupuestal.id_clave_padre = p_budget_struct_id
                                                                                                                                            AND COF_estr_presupuestal.eliminada = 0);
      IF (Num != 0) THEN
          p_result := -1;
          Return;
      END IF;

    SELECT COUNT(*) INTO Num FROM COF_estr_presupuestal WHERE (COF_estr_presupuestal.id_clave_presupuestal = p_budget_struct_id
                                                                                                                                            AND COF_estr_presupuestal.eliminada = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_estr_presupuestal SET Eliminada = 1 WHERE (id_clave_presupuestal = p_budget_struct_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -2;
END;

/

CREATE PROCEDURE sp_del_Responsability_Area (p_responsability_area_id NUMBER,
				                           p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_movimiento WHERE (COF_movimiento.id_area_resp = p_responsability_area_id);
 
  IF (Num = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_area_resp WHERE (COF_area_resp.id_area_resp = p_responsability_area_id
                                                                                                                                            AND COF_area_resp.eliminada = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_area_resp SET Eliminada = 1 WHERE (id_area_resp = p_responsability_area_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -1;
END;
/

CREATE PROCEDURE sp_del_Source (p_source_id NUMBER,
				p_result OUT NUMBER) AS
   Num   NUMBER;
   Num2 NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_transaccion WHERE (COF_transaccion.id_fuente = p_source_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mapeo_fuente WHERE (COF_mapeo_fuente.id_fuente = p_source_id);
 
  IF (Num = 0 AND Num2 = 0) THEN

    SELECT COUNT(*) INTO Num FROM COF_fuente WHERE (COF_fuente.id_fuente = p_source_id
                                                                                                                       AND COF_fuente.eliminada = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_fuente SET Eliminada = 1 WHERE (id_fuente = p_source_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -1;
END;

/

CREATE PROCEDURE sp_del_Standard_Account (p_standard_account_id NUMBER, p_delete_date DATE,
					   p_result OUT NUMBER) AS
BEGIN

        DELETE FROM COF_cuenta_estandar WHERE (id_cuenta_estandar = p_standard_account_id);
        UPDATE COF_cuenta_estandar_hist SET fecha_fin = p_delete_date WHERE (id_cuenta_estandar = p_standard_account_id) AND (fecha_fin = '31/12/2049');
        p_result := 1;
END;

/


CREATE PROCEDURE sp_del_General_Ledger(p_general_ledger_id NUMBER, p_delete_date DATE,
				       p_result OUT NUMBER) AS
  nCounter NUMBER;
BEGIN
  SELECT COUNT(*) INTO nCounter FROM COF_Cuenta WHERE (id_mayor = p_general_ledger_id);
  IF (nCounter = 0) THEN
     DELETE FROM MHParticipantObjects WHERE (entityId = 9) AND (objectId = p_general_ledger_id);
     DELETE FROM COF_Cuenta_Auxiliar 
     WHERE (id_mayor_auxiliar IN
         (SELECT id_mayor_auxiliar FROM COF_Mayor_Auxiliar
          WHERE (id_mayor = p_general_ledger_id)
         )
     );
     DELETE FROM COF_Mayor_Auxiliar WHERE (id_mayor = p_general_ledger_id);
     DELETE FROM COF_Elemento_Grupo_Mayor WHERE (id_mayor = p_general_ledger_id);
     DELETE FROM COF_Mayor WHERE (id_mayor = p_general_ledger_id);
  END IF;
  p_result := -1;
END;
/

