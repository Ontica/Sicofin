CREATE PROCEDURE sp_val_exchange_rate (p_from_currency_id NUMBER, p_to_currency_id NUMBER,
				              p_from_date DATE, p_to_date DATE,
			                                 p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(*) INTO nRes FROM AO_Exchange_Rates WHERE (from_currency_id = p_from_currency_id) 
							AND (to_currency_id = p_to_currency_id)
							AND (from_date = p_from_date)
							AND (to_date = p_to_date);
  IF (nRes = 0) THEN
    SELECT COUNT(currency_id) INTO nRes FROM AO_Currencies WHERE ((currency_id = p_from_currency_id) 
							  OR (currency_id = p_to_currency_id)) AND (deleted = 0);
    IF ( nRes != 2 ) THEN
      p_result := -1;
      RETURN;
    END IF; 
    p_result := nRes;
    Return;
  END IF;
  p_result := -2;
END;                               	
/



