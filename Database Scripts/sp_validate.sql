
CREATE PROCEDURE sp_val_Sector_For_Upd (p_sector_id NUMBER, p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(*) INTO nRes FROM COF_sector
                                WHERE (id_sector = p_sector_id) AND (eliminado = 0);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END;                              	
/

CREATE PROCEDURE sp_val_Calendar_For_Upd (p_calendar_id NUMBER, p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(*) INTO nRes FROM AO_calendars
                                WHERE (calendar_id = p_calendar_id);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END; 
/

CREATE PROCEDURE sp_val_period (p_period_calendar_id NUMBER,
			                 p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(period_id) INTO nRes FROM AO_Periods
                                WHERE (period_id = p_period_calendar_id);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END;                               	
/

CREATE PROCEDURE sp_val_Period_For_Upd (p_period_id NUMBER,
  					 p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(*) INTO nRes FROM AO_periods
                                WHERE (period_id = p_period_id);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END;                               	
/

CREATE PROCEDURE sp_val_date_in_period (p_period_id NUMBER, p_date DATE,
			                 p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
    
  IF (p_period_id != 0) then
    SELECT COUNT(period_id) INTO nRes FROM AO_Periods
                                WHERE (period_id = p_period_id) AND (is_open<>0) AND (from_date <= p_date) AND (to_date >= p_date);
    IF ( nRes = 0 ) THEN
      p_result := 0;
      return;
    ELSE
      p_result := nRes;
      return;
    END IF; 
  ELSE
    SELECT COUNT(period_id) INTO nRes FROM AO_Periods
                                WHERE (is_open<>0) AND (from_date <= p_date) AND (to_date >= p_date);
      IF ( nRes = 0 ) THEN
        p_result := 0;
        return;
      ELSE
        p_result := nRes;
        return;
      END IF; 
  END IF;
  p_result := -1;
END;                               	
/

CREATE PROCEDURE sp_upd_period_to_close (p_period_id NUMBER,
   				                 p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(period_id) INTO nRes FROM AO_Periods
                                WHERE (period_id = p_period_id);
  IF ( nRes = 0 ) THEN
    p_result := 0;
    RETURN;
  ELSE
    UPDATE AO_Periods SET Is_Open = 0 WHERE (period_id = p_period_id);
    p_result := nRes;
    RETURN;
  END IF; 
  p_result := -1;
END;                               	
/

CREATE PROCEDURE sp_upd_period_to_open (p_period_id NUMBER,
   				                  p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(period_id) INTO nRes FROM AO_Periods
                                WHERE (period_id = p_period_id);
  IF ( nRes = 0 ) THEN
    p_result := 0;
    RETURN;
  ELSE
    UPDATE AO_Periods SET Is_Open = 1 WHERE (period_id = p_period_id);
    p_result := nRes;
    RETURN;
  END IF; 
  p_result := -1;
END;                               	
/

CREATE PROCEDURE sp_val_Holiday (p_holiday_calendar_id NUMBER,
				p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(calendar_id) INTO nRes FROM AO_Calendars
                                WHERE (calendar_id = p_holiday_calendar_id);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END;                               	
/

CREATE PROCEDURE sp_val_Holiday_For_Upd (p_holiday_id NUMBER,
  					    p_result OUT NUMBER) AS
  nRes NUMBER;
BEGIN
  SELECT COUNT(*) INTO nRes FROM AO_holidays
                                WHERE (holiday_id = p_holiday_id);
  IF ( nRes = 0 ) THEN
    p_result := -1;
    RETURN;
  END IF; 
  p_result := nRes;
END;                               	
/



