
RECREATE TABLE TMP_XEVO_IMPORT (
  EXPER_FILE VARCHAR(200) CHARACTER SET NONE COLLATE NONE,
  ASSAY_DATE DATE,
  ASSAY_TYPE_ID SMALLINT,
  ANTIBODY_ID SMALLINT,
  ENZYME_ID SMALLINT,
  QUAN_TYPE_ID SMALLINT,
  SAMPLE_PROCESS_BY SMALLINT,
  ASSAY_DONE_BY SMALLINT,
  DATA_PROCESS_BY SMALLINT,
  FILENAME_DESC VARCHAR(200) CHARACTER SET NONE COLLATE NONE,
  SAMPLE_TYPE_ID SMALLINT,
  LEVEL_TP_ID SMALLINT,
  NUM_SPEC_CONC NUMERIC(10, 5),
  SUBJECT_NUM VARCHAR(10) CHARACTER SET NONE COLLATE NONE,
  FLUID_TYPE_ID SMALLINT,
  ANYLYTE_NAME VARCHAR(20) CHARACTER SET NONE COLLATE NONE,
  AREA NUMERIC(10, 3),
  HEIGHT INTEGER,
  RT NUMERIC(4, 2),
  ISTD_AREA NUMERIC(10, 3),
  ISTD_HEIGHT INTEGER,
  ISTD_RT NUMERIC(4, 2),
  RESPONSE NUMERIC(10, 5),
  CURVE VARCHAR(15) CHARACTER SET NONE COLLATE NONE,
  WEIGHTING VARCHAR(15) CHARACTER SET NONE COLLATE NONE,
  ORIGIN VARCHAR(15) CHARACTER SET NONE COLLATE NONE,
  EQUATION VARCHAR(30) CHARACTER SET NONE COLLATE NONE,
  R NUMERIC(10, 5),
  R_SQR NUMERIC(10, 5),
  NUM_SLOPE NUMERIC(10, 5),
  NUM_INTERSECT NUMERIC(10, 5),
  NUM_ANAL_CONC NUMERIC(10, 5),
  CONC_DEV_PERC NUMERIC(10, 5),
  SIGNOISE NUMERIC(5, 1),
  CHROMNOISE NUMERIC(7, 2),
  INJECTION_ID SMALLINT    
  );


