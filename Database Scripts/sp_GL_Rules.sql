CREATE PROCEDURE spLastChildPosition(pRuleId NUMBER, pResult OUT NUMBER) AS 
  nRuleDefId NUMBER;
  nLevel     NUMBER;
  nPosition  NUMBER;
BEGIN 
   SELECT id_regla_contable INTO nRuleDefId
   FROM COF_Grupo_Cuenta
   WHERE (id_grupo_cuenta = pRuleId);

   SELECT nivel INTO nLevel
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);

   SELECT posicion INTO nPosition
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);
 
   SELECT MIN(posicion) INTO pResult
   FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis
   WHERE (id_regla_contable = nRuleDefId) AND 
         (nivel <= nLevel) AND (posicion > nPosition) AND 
         (COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta);
   IF pResult IS NULL THEN
      SELECT MAX(posicion) INTO pResult
      FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis
      WHERE (id_regla_contable = nRuleDefId) AND          
            (COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta);
   ELSE
      pResult := pResult - 1;
   END IF;

   IF pResult IS NULL THEN
      pResult := 0;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := 0;
END;
/

CREATE PROCEDURE spRuleDefName(pRuleDefId NUMBER, pResult OUT VARCHAR2) AS 
BEGIN   
   SELECT nombre_regla_contable INTO pResult    
   FROM COF_Regla_Contable
   WHERE (id_regla_contable = pRuleDefId);
   IF pResult IS NULL THEN
      pResult := '';
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := '';
END;
/

CREATE PROCEDURE spRuleDefStdAccountTypeId(pRuleDefId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT id_tipo_cuentas_std INTO pResult    
   FROM COF_Regla_Contable
   WHERE (id_regla_contable = pRuleDefId);
   IF pResult IS NULL THEN
      pResult := -1;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := -1;
END;
/

CREATE PROCEDURE spRuleDefType(pRuleDefId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT tipo_regla INTO pResult    
   FROM COF_Regla_Contable
   WHERE (id_regla_contable = pRuleDefId);
   IF pResult IS NULL THEN
      pResult := -1;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := -1;
END;
/

CREATE PROCEDURE spRuleChildsCount (pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN
   SELECT COUNT(*) INTO pResult 
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta_padre = pRuleId); 
END;
/

CREATE PROCEDURE spRuleChildsType (pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
  SELECT DISTINCT tipo_grupo_cuenta INTO pResult 
  FROM COF_Grupo_Cuenta, COF_Grupo_Cuenta_Bis
  WHERE (id_grupo_cuenta_padre = pRuleId) AND
        (COF_Grupo_Cuenta.id_grupo_cuenta = COF_Grupo_Cuenta_Bis.id_grupo_cuenta);
  IF pResult IS NULL THEN
    pResult := -1;
  END IF;
EXCEPTION
  WHEN NO_DATA_FOUND THEN
    pResult := -1;
END;
/

CREATE PROCEDURE spRuleLevel(pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT nivel INTO pResult
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);
   IF pResult IS NULL THEN
      pResult := 0;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := 0;
END;
/


CREATE PROCEDURE spRuleParentId(pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT id_grupo_cuenta_padre INTO pResult
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);
   IF pResult IS NULL THEN
      pResult := 0;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := 0;
END;
/

CREATE PROCEDURE spRulePosition(pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT posicion INTO pResult
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);
   IF pResult IS NULL THEN
      pResult := 0;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := 0;
END;
/

CREATE PROCEDURE spRuleType(pRuleId NUMBER, pResult OUT NUMBER) AS 
BEGIN   
   SELECT tipo_grupo_cuenta INTO pResult
   FROM COF_Grupo_Cuenta
   WHERE (id_grupo_cuenta = pRuleId);
   IF pResult IS NULL THEN
      pResult := -1;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pResult := -1;
END;
/

CREATE PROCEDURE spDelRule(pRuleId NUMBER) AS 
  nRuleDefId NUMBER;
  nPosition  NUMBER;
BEGIN   

   SELECT id_regla_contable INTO nRuleDefId
   FROM COF_Grupo_Cuenta
   WHERE (id_grupo_cuenta = pRuleId);

   SELECT posicion INTO nPosition
   FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);

   DELETE FROM COF_Grupo_Cuenta
   WHERE (id_grupo_cuenta = pRuleId);
   
   DELETE FROM COF_Grupo_Cuenta_Bis
   WHERE (id_grupo_cuenta = pRuleId);

   UPDATE COF_Grupo_Cuenta_Bis
   SET posicion = posicion - 1
   WHERE posicion >= nPosition AND id_grupo_cuenta IN
    (SELECT id_grupo_cuenta
     FROM COF_Grupo_Cuenta
     WHERE (id_regla_contable = nRuleDefId)
    );

   COMMIT;
END;
/

CREATE PROCEDURE spUpdRulePositions(pRuleDefId NUMBER, pStartOrder NUMBER) AS 
BEGIN
  UPDATE COF_Grupo_Cuenta_Bis
  SET posicion = posicion + 1
  WHERE posicion >= pStartOrder AND id_grupo_cuenta IN
   (SELECT id_grupo_cuenta
    FROM COF_Grupo_Cuenta
    WHERE (id_regla_contable = pRuleDefId)
   );

  COMMIT;
END;
/

