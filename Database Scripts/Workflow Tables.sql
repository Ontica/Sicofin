/***********************  TABLAS  *******************************/

CREATE TABLE EWFApplications
(
 ApplicationId			NUMBER(12) PRIMARY KEY,
 AppName			VARCHAR2(64) NOT NULL ,
 AppDescription			VARCHAR2(1024) ,
 AppIcon			VARCHAR2(128) ,
 AppDocumentation		VARCHAR2(128) ,
 AppPath			VARCHAR2(128) ,
 AppParametersDef		VARCHAR2(255) ,
 IsWebApp			NUMBER(1) NOT NULL ,
 AppStatus			CHAR(1)
);


CREATE TABLE EWFParticipants
(
 ParticipantId			NUMBER(12) PRIMARY KEY , 
 ParticipantName		VARCHAR2(64) NOT NULL , 
 ParticipantKey			VARCHAR2(24) NOT NULL , 
 Description			VARCHAR2(512) , 
 ParticipantType                CHAR(1) NOT NULL , 
 Status                		CHAR(1) NOT NULL , 
 FromDate			DATE NOT NULL , 
 ToDate				DATE NOT NULL
); 


CREATE TABLE EWFParticipantAttrs
(
 ParticipantAttrId              NUMBER(12) PRIMARY KEY , 
 ParticipantId                  NUMBER(12) NOT NULL , 
 EntityAttrDefId                NUMBER(12) NOT NULL , 
 ValString			VARCHAR2(2048) ,
 ValNumeric			NUMBER ,
 ValDate			DATE ,
 FromDate			DATE NOT NULL ,
 ToDate				DATE NOT NULL
); 


CREATE TABLE MHParticipantObjects
(
 ParticipantObjectId		NUMBER(12) PRIMARY KEY , 
 ParticipantId			NUMBER(12) NOT NULL , 
 EntityId			NUMBER(12) NOT NULL , 
 ObjectId			NUMBER(12) NOT NULL ,
 FromDate			DATE NOT NULL ,
 ToDate				DATE NOT NULL
);


CREATE TABLE MHParticipantsRelations
(
 ParticipantsRelationId         NUMBER(12) PRIMARY KEY , 
 RelationTypeId			NUMBER(12) NOT NULL ,
 ParticipantAId			NUMBER(12) NOT NULL , 
 ParticipantBId			NUMBER(12) NOT NULL , 
 FromDate			DATE NOT NULL,
 ToDate				DATE NOT NULL
);

CREATE TABLE EWFParticipantTasks
(
 ParticipantId         		NUMBER(12) NOT NULL , 
 TaskId				NUMBER(12) NOT NULL
);

CREATE TABLE EWFTasks
(
 TaskId				NUMBER(12) PRIMARY KEY ,
 TaskName			VARCHAR2(255) NOT NULL , 
 TaskShortName			VARCHAR2(48) , 
 TaskDescription		VARCHAR2(1024) , 
 TaskIcon			VARCHAR2(128) ,
 TaskDocumentation		VARCHAR2(128) , 
 ApplicationId			NUMBER(12) NOT NULL ,  
 ApplicationInitPars		VARCHAR2(255) , 
 TaskKind			CHAR(1) , 
 Implementation			CHAR(1) , 
 StartMode			CHAR(1) , 
 FinishMode			CHAR(1) , 
 AccessRestrictionListId	NUMBER(12) NOT NULL ,
 DurationUnit			CHAR(1) ,
 Limit				NUMBER(6) ,  
 DefaultPriority		NUMBER(1) NOT NULL ,
 Instantation			NUMBER(6) NOT NULL , 
 AverageWorkingTime		NUMBER(6) ,  	
 AverageWaitingTime		NUMBER(6) ,  
 AverageCost			NUMBER(25,10)
); 

CREATE TABLE EWFTaskBars
(
 ParticipantId			NUMBER(12) NOT NULL ,
 TaskId				NUMBER(12) NOT NULL ,
 TaskBarType			CHAR(1) NOT NULL ,
 ItemLabel			VARCHAR2(64) , 
 PresentationOrder		NUMBER(4) NOT NULL 
);

CREATE TABLE MHInboxes
(
 InboxItemId			NUMBER(12) PRIMARY KEY ,
 ItemType			CHAR(1) NOT NULL ,			
 Subject			VARCHAR2(2048) NOT NULL , 
 Observations			VARCHAR2(2048) , 
 ItemId				NUMBER(12) NOT NULL ,
 FromParticipantId		NUMBER(12) NOT NULL ,
 ToParticipantId		NUMBER(12) NOT NULL ,
 ReceivedDate			DATE NOT NULL ,
 HasAttachments			NUMBER(1) NOT NULL ,
 Priority			NUMBER(1) NOT NULL ,
 Parameters			VARCHAR2(2048) , 
 Status				CHAR(1) NOT NULL
);

/***********************    SECUENCIAS *******************************/

CREATE SEQUENCE seqEWFParticipantId
INCREMENT BY 1
START WITH 696
NOCACHE;

CREATE SEQUENCE seqEWFParticipantObjectId
INCREMENT BY 1
START WITH 9482
NOCACHE;

CREATE SEQUENCE seqEWFParticipantRelationId
INCREMENT BY 1
START WITH 721
NOCACHE;

CREATE SEQUENCE seqEWFParticipantAttrId
INCREMENT BY 1
START WITH 3199
NOCACHE;

CREATE SEQUENCE seqEWFTask
INCREMENT BY 1
START WITH 101
NOCACHE;

CREATE SEQUENCE seqMHInboxItemId
INCREMENT BY 1
START WITH 101
NOCACHE;