/******Patterson Lab*********************************************/

EXECUTE BLOCK
AS
DECLARE ASSAY_ID_VAR SMALLINT;
DECLARE SUBJECT_NUM_VAR VARCHAR(10);
DECLARE ASSAY_DATE_VAR DATE;
DECLARE FLUID_TYPE_ID_VAR SMALLINT;

DECLARE A CURSOR FOR (  /*returns assay records out of imported records */
   SELECT G.SUBJECT_NUM, G.ASSAY_DATE, G.FLUID_TYPE_ID
   FROM   TMP_GC_MS_IMPORT1 G
   GROUP
      BY G.SUBJECT_NUM, G.FLUID_TYPE_ID, G.ASSAY_DATE
);

DECLARE TP_ID_VAR SMALLINT;

DECLARE NUM_200_VAR INTEGER;
DECLARE NUM_205_VAR INTEGER;
DECLARE MP_VAR NUMERIC(5, 4);
DECLARE MPE_VAR NUMERIC(5, 3);
DECLARE NUM205_200 NUMERIC(5, 4);
DECLARE NUM_TTR_BY_CAL_VAR NUMERIC(10, 7);
DECLARE NUM_TTR_BY_CAL_SUBS_R0_VAR NUMERIC(10, 7);

DECLARE B CURSOR FOR ( /*returns assay_sample records out of imported records*/
   SELECT G.SUBJECT_NUM, G.FLUID_TYPE_ID, G.TP_ID,
   G.NUM_200, G.NUM_205, G.MP, G.MPE, G.NUM205_200, G.NUM_TTR_BY_CAL, G.NUM_TTR_BY_CAL_SUBS_R0

   FROM   TMP_GC_MS_IMPORT1 G
   WHERE G.ASSAY_DATE = :ASSAY_DATE_VAR AND G.SUBJECT_NUM = :SUBJECT_NUM_VAR
);

DECLARE SAMPLE_ID_VAR INTEGER;
DECLARE C CURSOR FOR ( /*returns SAMPLE_ID by subject, fluid, timepoint*/
  SELECT smpl.SAMPLE_ID
  FROM SAMPLE smpl
  WHERE smpl.SUBJECT_ID = 
  ( select subj.SUBJECT_ID
    from SUBJECT subj
    where ( subj.SUBJECT_NUM = :SUBJECT_NUM_VAR ) )  AND
  ( smpl.FLUID_TYPE_ID = :FLUID_TYPE_ID_VAR )        AND
  ( smpl.TIME_POINT_ID = :TP_ID_VAR )
);

DECLARE ASSAY_SAMPLE_ID_VAR BIGINT;
DECLARE ASSAY_DATA_ID_VAR BIGINT;
DECLARE FLUID_TYPE_STR VARCHAR(7); 
DECLARE UNDERSCORE VARCHAR(1);

BEGIN

UNDERSCORE = '_';

OPEN A;
   WHILE (1=1) DO
   BEGIN
	 FETCH A INTO :SUBJECT_NUM_VAR, :ASSAY_DATE_VAR, FLUID_TYPE_ID_VAR;
   if (row_count = 0) then leave; 
   if (FLUID_TYPE_ID_VAR = 1) then FLUID_TYPE_STR = 'hCSF';
   if (FLUID_TYPE_ID_VAR = 2) then FLUID_TYPE_STR = 'hPlasma';
   UPDATE OR INSERT INTO ASSAY
   (ASSAY_DATE, ASSAY_TYPE_ID, ASSAY_DONE_BY, SAMPLE_PROCESSED_BY,
   DATA_PROCESS_BY, DESCRIP)
   VALUES (:ASSAY_DATE_VAR,
        8,                   /* tBDMS GC/MS    Patterson Lab */
        0,0,0,               /*Not a Lab Member*/
        :ASSAY_DATE_VAR || :UNDERSCORE || :FLUID_TYPE_STR || '_13C6-Leu-tBDMS_' || :SUBJECT_NUM_VAR 
   )
   matching (ASSAY_DATE, ASSAY_TYPE_ID, DESCRIP)
   returning ASSAY_ID into :ASSAY_ID_VAR;
   
   OPEN B;
      WHILE (1=1) DO
      BEGIN
      FETCH B INTO :SUBJECT_NUM_VAR, :FLUID_TYPE_ID_VAR, :TP_ID_VAR,
      :NUM_200_VAR, :NUM_205_VAR, :MP_VAR, :MPE_VAR, :NUM205_200, :NUM_TTR_BY_CAL_VAR, :NUM_TTR_BY_CAL_SUBS_R0_VAR;
      
      if (row_count = 0) then leave;
      
      SAMPLE_ID_VAR = NULL;
      OPEN C;
        FETCH C INTO :SAMPLE_ID_VAR;
      CLOSE C;
      if (SAMPLE_ID_VAR IS NULL ) then 
            EXCEPTION SAMPLE_NOT_FOUND 'Cannot find sample - ' || :SUBJECT_NUM_VAR ||' - '|| :FLUID_TYPE_ID_VAR ||' - '|| :TP_ID_VAR ;
            
      ASSAY_SAMPLE_ID_VAR = GEN_ID(ASSAY_SAMPLE_ID_GEN, 0);
      ASSAY_DATA_ID_VAR   = GEN_ID(ASSAY_DATA_ASSAY_DATA_ID_GEN, 0);


      INSERT INTO ASSAY_SAMPLE (ASSAY_SAMPLE_TYPE_ID, ASSAY_ID,
      VOLUME, VOLUME_UNIT_ID, DESCRIP, SAMPLE_ID)
      VALUES (        /*ASSAY_SAMPLE_ID from ASSAY_SAMPLE_ID_GEN*/
        0,            /*Unknown*/
        :ASSAY_ID_VAR,
        0.5,          /*Volume*/
        5,            /*Volume unit 5 - mL*/
      :SUBJECT_NUM_VAR || :UNDERSCORE || :FLUID_TYPE_STR || :UNDERSCORE || :TP_ID_VAR || :UNDERSCORE || 'hr', 
      :SAMPLE_ID_VAR
      );
      
      
      ASSAY_SAMPLE_ID_VAR = ASSAY_SAMPLE_ID_VAR + 1;
      INSERT INTO ASSAY_DATA ( ASSAY_SAMPLE_ID, ANALYTE_ID, QUANT_TYPE_ID)
      VALUES (                      /* ASSAY_DATA_ID from ASSAY_DATA_ID_GEN */ 
              :ASSAY_SAMPLE_ID_VAR, /*Syncronized with ASSAY_SAMPLE_ID_GEN*/
              11,                   /*LeuC13*/
              1                     /*Relative*/     );
      
      
      ASSAY_DATA_ID_VAR = ASSAY_DATA_ID_VAR + 1;        
      INSERT INTO DATA_GCMS ( ASSAY_DATA_ID, INT_200, INT_205, NUM_MP, NUM_MPE,
      NUM_205_200, NUM_TTR, NUM_TTR_BY_CAL, NUM_TTR_BY_CAL_SUBS_R0 )
      VALUES ( :ASSAY_DATA_ID_VAR,  /*Syncronized with ASSAY_DATA_ID_GEN*/
               :NUM_200_VAR,
               :NUM_205_VAR,
               :MP_VAR,
               :MPE_VAR,
               :NUM205_200,
			   :NUM_TTR_BY_CAL_SUBS_R0_VAR * 100,
               :NUM_TTR_BY_CAL_VAR,
			   :NUM_TTR_BY_CAL_SUBS_R0_VAR);
        
              
      
      END
   
   CLOSE B;
   
   END

CLOSE A;

END;