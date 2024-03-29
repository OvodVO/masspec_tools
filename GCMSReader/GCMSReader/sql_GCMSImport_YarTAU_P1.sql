/*---GCMC import  YarasheskiTAU Lab   */

/*---Creating temporary tables   */
RECREATE TABLE TMP_GC_MS_IMPORT
( SUBJECT_NUM VARCHAR(10) NOT NULL,
  FLUID_TYPE_ID SMALLINT NOT NULL,
  ASSAY_DATE DATE NOT NULL,
  TP_ID SMALLINT NOT NULL,
  NUM_349 INTEGER NOT NULL,
  NUM_355 INTEGER NOT NULL,
  MP NUMERIC(5, 4) NOT NULL,
  MPE NUMERIC(5, 3) NOT NULL,
  NUM355_349 NUMERIC(5, 4) NOT NULL,
  TTR NUMERIC(5, 2) NOT NULL,
  CALC_TTR NUMERIC(10, 7) NOT NULL );
 