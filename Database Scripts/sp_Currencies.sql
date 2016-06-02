
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


