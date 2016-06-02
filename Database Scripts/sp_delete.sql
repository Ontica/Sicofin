CREATE PROCEDURE sp_del_sector_mapping(p_standard_account_id NUMBER, p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM COF_Mapeo_Sector WHERE (id_cuenta_estandar = p_standard_account_id); 
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM COF_Mapeo_Sector WHERE (id_cuenta_estandar = p_standard_account_id);
        p_result :=  Num;
        Return;
END;
/

CREATE PROCEDURE sp_del_General_Ledger (p_general_ledger_id NUMBER, p_delete_date DATE,
					p_result OUT NUMBER) AS
  Num  NUMBER;
  Num2  NUMBER;
  Num3  NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_transaccion WHERE (COF_transaccion.id_mayor = p_general_ledger_id);
  SELECT COUNT(*) INTO Num2 FROM COF_cuenta WHERE (COF_cuenta.id_mayor = p_general_ledger_id);
  SELECT COUNT(*) INTO Num3 FROM COF_mapeo_tipo_transaccion WHERE (COF_mapeo_tipo_transaccion.id_mayor = p_general_ledger_id);
 
  IF (Num = 0 AND Num2 = 0 AND Num3 = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_mayor WHERE (COF_mayor.id_mayor = p_general_ledger_id
           AND COF_mayor.eliminado = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_mayor SET Eliminado = 1, Fecha_Cierre = p_delete_date  WHERE (id_mayor = p_general_ledger_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -1;
END;
/

CREATE PROCEDURE sp_del_Subsidiary_Ledger(p_subsidiary_ledger_id NUMBER, p_result OUT NUMBER) AS
  Num  NUMBER;
  Num2  NUMBER;
BEGIN

  SELECT COUNT(*) INTO Num FROM COF_cuenta_auxiliar WHERE (id_mayor_auxiliar = p_subsidiary_ledger_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mapeo_mayor_auxiliar WHERE (id_mayor_auxiliar = p_subsidiary_ledger_id); 
 
  IF (Num = 0 AND Num2 = 0) THEN
    SELECT COUNT(*) INTO Num FROM COF_mayor_auxiliar WHERE (id_mayor_auxiliar = p_subsidiary_ledger_id) AND (eliminado = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_mayor_auxiliar SET Eliminado = 1  WHERE (id_mayor_auxiliar = p_subsidiary_ledger_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -1;
END;
/


CREATE PROCEDURE sp_del_Subsidiary_Account (p_subsidiary_account_id NUMBER, p_result OUT NUMBER) AS
   Num  NUMBER;
BEGIN
 
    SELECT COUNT(*) INTO Num FROM COF_cuenta_auxiliar WHERE (id_cuenta_auxiliar = p_subsidiary_account_id);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
	UPDATE COF_cuenta_auxiliar SET Eliminada = 1 WHERE (id_cuenta_auxiliar = p_subsidiary_account_id);
        p_result := Num;
        Return;
      END IF;
END;
/

CREATE PROCEDURE sp_del_Sector (p_sector_id NUMBER, p_result OUT NUMBER) AS
   Num  NUMBER;
  Num2 NUMBER;
  Num3 NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_movimiento WHERE (COF_movimiento.id_sector = p_sector_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mapeo_sector WHERE (COF_mapeo_sector.id_sector = p_sector_id);
  SELECT COUNT(*) INTO Num3 FROM COF_mapeo_mayor_auxiliar WHERE (COF_mapeo_mayor_auxiliar.id_sector = p_sector_id);
 
  IF (Num = 0 AND Num2 = 0 AND Num3 = 0) THEN

    SELECT COUNT(*) INTO Num FROM COF_sector WHERE (COF_sector.id_sector = p_sector_id AND COF_sector.eliminado = 0);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        UPDATE COF_sector SET Eliminado = 1 WHERE (id_sector = p_sector_id);
        p_result := Num;
        Return;       
      END IF; 
  END IF;
    p_result := -1;
END;
/

CREATE PROCEDURE sp_del_Standard_Account (p_standard_account_id NUMBER, p_delete_date DATE,
					   p_result OUT NUMBER) AS
   Num  NUMBER;
  Num2 NUMBER;
  Num3 NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM COF_cuenta WHERE (COF_cuenta.id_cuenta_estandar = p_standard_account_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mapeo_sector WHERE (COF_mapeo_sector.id_cuenta_estandar = p_standard_account_id);
  SELECT COUNT(*) INTO Num3 FROM COF_mapeo_moneda WHERE (COF_mapeo_moneda.id_cuenta_estandar = p_standard_account_id);
 
  IF (Num = 0 AND Num2 = 0 AND Num3 = 0) THEN

    SELECT COUNT(*) INTO Num FROM COF_cuenta_estandar WHERE (COF_cuenta_estandar.id_cuenta_estandar = p_standard_account_id);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
        DELETE FROM COF_cuenta_estandar WHERE (id_cuenta_estandar = p_standard_account_id);
        UPDATE COF_cuenta_estandar_hist SET fecha_fin = p_delete_date WHERE (id_cuenta_estandar = p_standard_account_id) AND (fecha_fin = '31/12/2999');
        p_result := Num;
        Return;
      END IF;
  END IF;
    p_result := -1;
END;
/

CREATE PROCEDURE sp_del_Std_Account_Categ(p_category_id NUMBER, p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_Categories WHERE (category_id = p_category_id) AND (parent_id = 4);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM AO_Categories WHERE (category_id = p_category_id) AND (parent_id = 4);
        p_result :=  Num;
        Return;      
END;
/

CREATE PROCEDURE sp_del_Entity(p_entity_id NUMBER, p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_Entities WHERE (entity_id = p_entity_id);
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM AO_Entities WHERE (entity_id = p_entity_id);
        p_result :=  Num;
        Return;
END;
/

CREATE PROCEDURE sp_del_calendar (p_calendar_id NUMBER, 
                                                                              p_result OUT NUMBER) AS
   Num NUMBER;
   Num2 NUMBER;
   Num3 NUMBER;
  Num4 NUMBER;
BEGIN
  SELECT COUNT(*) INTO Num FROM AO_periods WHERE (AO_periods.calendar_id = p_calendar_id);
  SELECT COUNT(*) INTO Num2 FROM COF_mayor WHERE (COF_mayor.id_calendario = p_calendar_id);
  SELECT COUNT(*) INTO Num3 FROM AO_holidays WHERE (AO_holidays.calendar_id = p_calendar_id);
  SELECT COUNT(*) INTO Num4 FROM AO_not_labor_weekdays WHERE (AO_not_labor_weekdays.calendar_id = p_calendar_id);

  IF (Num = 0 AND Num2 = 0 AND Num3 = 0 AND Num4 = 0) THEN
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


CREATE PROCEDURE sp_del_period (p_period_id NUMBER, 
  	                                                       p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_periods WHERE (AO_periods.period_id = p_period_id);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM AO_periods WHERE ( AO_periods.period_id = p_period_id );
        p_result :=  Num;
        Return;      
END;
/

CREATE PROCEDURE sp_del_holiday (p_holiday_id NUMBER, 
                                                                           p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_holidays WHERE (AO_holidays.holiday_id = p_holiday_id);   
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM AO_holidays WHERE ( AO_holidays.holiday_id = p_holiday_id );
        p_result :=  Num;
        Return;      
END;
/

CREATE PROCEDURE sp_del_not_labor_weekdays(p_calendar_id NUMBER, p_result OUT NUMBER) AS
   Num NUMBER;
BEGIN
    SELECT COUNT(*) INTO Num FROM AO_not_labor_weekdays WHERE (calendar_id = p_calendar_id); 
      IF (Num = 0) THEN
        p_result :=  0;
        Return;
      END IF;
        DELETE FROM AO_not_labor_weekdays WHERE (calendar_id = p_calendar_id);
        p_result :=  Num;
        Return;
END;
/

CREATE PROCEDURE sp_del_Currency (p_currency_id NUMBER, p_result OUT NUMBER ) AS
  nExcRate NUMBER;
  nMapCurrency NUMBER;
  nGenLed NUMBER;
  Num NUMBER;
BEGIN
  SELECT COUNT(*) INTO nExcRate FROM AO_exchange_rates WHERE (AO_exchange_rates.from_currency_id = p_currency_id) 
							OR (AO_exchange_rates.to_currency_id = p_currency_id);

  SELECT COUNT(*) INTO nMapCurrency FROM COF_mapeo_moneda WHERE (COF_mapeo_moneda.id_moneda = p_currency_id);
  SELECT COUNT(*) INTO nGenLed FROM COF_mayor WHERE (COF_mayor.id_moneda_base = p_currency_id) AND (COF_mayor.eliminado = 0);

  IF (nExcRate = 0 AND nMapCurrency = 0 AND nGenLed = 0) THEN
        SELECT COUNT(*) INTO Num FROM AO_currencies WHERE (AO_currencies.currency_id = p_currency_id);
      IF (Num = 0) THEN
        p_result := 0;
        Return;
      ELSE
             UPDATE AO_currencies SET Deleted = 1 WHERE ( AO_currencies.currency_id  = p_currency_id);
        p_result := Num;
        Return;
      END IF; 
  END IF;  
     p_result  := -1 ;
END;
/

CREATE PROCEDURE sp_del_Exchange_Rate (p_exchange_rate_id NUMBER, p_result OUT NUMBER) AS
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


