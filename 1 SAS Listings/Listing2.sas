/* If indication is MEdical History but dates are not consistent */

/*DEfining file paths in a macro this time */
%let mh_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_MH.xlsx";
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS02.xlsx";

/* Importing CM and MH data */
proc import datafile=&cm_file out=cm dbms=xlsx replace;
getnames=yes;
run;

proc import datafile=&mh_file out=mh dbms=xlsx replace;
getnames=yes;
run;

data cm;
    set cm;
    format CMSTDAT CMENDAT date9.;
run;

data mh;
    set mh;
    format MHSTDAT MHENDAT date9.;
run;

/* Merging CM and MH on USUBJID */
proc sort data=cm; by USUBJID; run;
proc sort data=mh; by USUBJID; run;

data merged;
    merge cm(in=a) mh(in=b);
    by USUBJID;
run;

/*  Adding missing columns and default values */
data processed;
    set merged;
    length SITEID SITENAME SUBJID VISITID FORMID CHKREF VISITSEQ REVINST MSG $200;
    
    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "GENERAL CONMED";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = CMSEQ;
    if missing(CHKREF) then CHKREF = "CM-MH Recon";
    if missing(VISITSEQ) then VISITSEQ = "NA";
    
    format EXEC_DTTM datetime20.;
    EXEC_DTTM = datetime();
    REVINST = "If Indication for the Medication is reported as 'Medical History', then the Medication Start and End Date/Time should be consistent with the corresponding Medical History Start and End Date.";
    MSG = "Indication in Concomitant Therapy form is reported as 'Medical History'. However, concomitant therapy start date is prior to the Medical history date, or the concomitant therapy end date is after the Medical history end date. Please correct else clarify.";
run;

/* Apply logical filters/ check logic */



data filtered;
    set processed;
         if CMINDC = "MEDICAL HISTORY" and 
       (not missing(MHDSL1) or not missing(MHDSL2) or not missing(MHDSL3) or not missing(MHDSL4) or not missing(MHDSL5)) and
       (CMSTDAT < MHSTDAT or CMSTDAT > MHENDAT or CMENDAT > MHENDAT);
run;

/*Renaming columns as per Spec */
data final;
	retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM CMSEQ CMTRT CMINDC CMSTDAT CMENDAT
           MHDSL1 MHDSL2 MHDSL3 MHDSL4 MHDSL5 MHSEQ MHCAT MHSTDAT MHENDAT REVINST MSG;

    set filtered;
    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM CMSEQ CMTRT CMINDC CMSTDAT CMENDAT
         MHDSL1 MHDSL2 MHDSL3 MHDSL4 MHDSL5 MHSEQ MHCAT MHSTDAT MHENDAT REVINST MSG;

    label
        STUDYID = "STUDYID"
        SITEID = "SITENUMBER"
        SITENAME = "SITE"
        SUBJID = "SUBJECT"
        VISITID = "FOLDERNAME"
        VISITSEQ = "VISITSEQ"
        FORMID = "FORMID"
        FORMSEQ = "FORMSEQ"
        CHKREF = "CHKREF"
        EXEC_DTTM = "Execution Date/Time"
        CMSEQ = "CM Recordposition"
        CMTRT = "Medication or Therapy"
        CMINDC = "Indication"
        CMSTDAT = "CM Start Date"
        CMENDAT = "End Date"
        MHDSL1 = "MH#-Term_1"
        MHDSL2 = "MH#-Term_2"
        MHDSL3 = "MH#-Term_3"
        MHDSL4 = "MH#-Term_4"
        MHDSL5 = "MH#-Term_5"
        MHSEQ = "MH Recordposition"
        MHCAT = "MH Category"
        MHSTDAT = "MH Start Date"
        MHENDAT = "End Date"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
        
run;

proc export data=final
    outfile=&output_file
    dbms=xlsx
    replace;
run;
