CREATE PROCEDURE spDoLogin (pCredentialAttr NUMBER, pUserName VARCHAR2, pPassword VARCHAR2, pUserId OUT NUMBER) AS 
BEGIN   
   SELECT EWFParticipants.participantId INTO pUserId
   FROM EWFParticipants, EWFParticipantAttrs
   WHERE (EWFParticipants.participantId = EWFParticipantAttrs.participantId) AND 
   (EWFParticipants.participantKey = pUserName) AND (EWFParticipants.participantType = 'U') AND 
   (EWFParticipants.status = 'A') AND 
   (EWFParticipants.fromDate <= sysDate) AND (EWFParticipants.toDate = '31/12/2049') AND 
   (EWFParticipantAttrs.entityAttrDefId = pCredentialAttr) AND (EWFParticipantAttrs.ValString = pPassword);
   IF pUserId IS NULL THEN
      pUserId := 0;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
      pUserId := 0;
END;
/

CREATE PROCEDURE spIsTaskAssignedToUser(pTaskId NUMBER, pUserId NUMBER, pResult OUT NUMBER) AS
BEGIN
     SELECT pTaskId INTO pResult
     FROM EWFParticipantTasks
     WHERE (taskId = pTaskId) AND (participantId = pUserId);

   IF pResult IS NULL THEN
      pResult := 0;
   ELSE
      pResult := 1;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
     pResult := 0;
END;
/

CREATE PROCEDURE spCheckUserCredential (pCredentialAttr NUMBER, pUserId NUMBER, pCredential VARCHAR2, pResult OUT NUMBER) AS
BEGIN
   IF pCredentialAttr <> 0 THEN
     SELECT ParticipantAttrId INTO pResult
     FROM EWFParticipantAttrs
     WHERE (participantId = pUserId) AND (entityAttrDefId = pCredentialAttr) AND (ValString = pCredential);
   ELSE
     SELECT participantId INTO pResult
     FROM EWFParticipants
     WHERE (participantId = pUserId) AND (participantKey = pCredential);
   END IF;

   IF pResult IS NULL THEN
      pResult := 0;
   ELSE
      pResult := 1;
   END IF;
EXCEPTION
   WHEN NO_DATA_FOUND THEN
     pResult := 0;
END;
/

CREATE PROCEDURE spUpdUserCredential (pCredentialAttr NUMBER, pUserId NUMBER, pOldCredential VARCHAR2, pNewCredential VARCHAR2) AS
BEGIN
   IF (pCredentialAttr <> 0) THEN
     UPDATE EWFParticipantAttrs
     SET ValString = pNewCredential
     WHERE (participantId = pUserId) AND (entityAttrDefId = pCredentialAttr) AND (ValString = pOldCredential);
   ELSE
     UPDATE EWFParticipants
     SET participantKey = pNewCredential
     WHERE (participantId = pUserId) AND (participantKey = pOldCredential);
   END IF;
   COMMIT;
   RETURN;
EXCEPTION
   WHEN OTHERS THEN
   ROLLBACK;
   RETURN;   
END;
/

