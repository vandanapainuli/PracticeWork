/* Medication for AE, dates inconsistent */
%let ae_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_AE.xlsx";
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS06.xlsx";


proc import datafile=&ae_file out=ae dbms=xlsx replace;
    getnames=yes;
run;

proc import datafile=&cm_file out=cm dbms=xlsx replace;
    getnames=yes;
run;

proc sort data=ae; by USUBJID; run;
proc sort data=cm; by USUBJID; run;

data merged;
    merge ae(in=a) cm(in=b);
    by USUBJID;
run;

data processed;
    set merged;
    length SITEID SITENAME SUBJID VISITID FORMID CHKREF VISITSEQ REVINST MSG $200;
    format EXEC_DTTM yymmdd10.;

    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "AESAE";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = AESEQ;
    if missing(CHKREF) then CHKREF = "AE-CM Recon";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();
    REVINST = "If CM Indication is 'Adverse Event', AE start date should not be after CM Start date.";
    MSG = "CM Indication is Adverse Event; however, AE start date is after CM Start date. Please reconcile.";
run;

/* main logic */
data filtered;
    set processed;
    

    format AESTDAT yymmdd10.;
	

    format CMSTDAT yymmdd10.;
	

  	if upcase(CMINDC) = "ADVERSE EVENT" and 
   	not missing(AESTDAT) and not missing(CMSTDAT) and 
   	AESTDAT gt CMSTDAT;

run;

data final;
  retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           CMTRT CMINDC AEDSL1 AEDSL2 AEDSL3 AEDSL4 AEDSL5 CMSTDAT CMENDAT AESEQ AETERM AESTDAT AEENDTC 
           REVINST MSG;
    set filtered;
      keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         CMTRT CMINDC AEDSL1 AEDSL2 AEDSL3 AEDSL4 AEDSL5 CMSTDAT CMENDAT AESEQ AETERM AESTDAT AEENDTC 
         REVINST MSG;

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
        CMTRT = "Medication or Therapy"
        CMINDC = "Indication"
        AEDSL1 = "AE log line start date and term 1"
        AEDSL2 = "AE log line start date and term 2"
        AEDSL3 = "AE log line start date and term 3"
        AEDSL4 = "AE log line start date and term 4"
        AEDSL5 = "AE log line start date and term 5"
        CMSTDAT = "CM Start Date"
        CMENDAT = "CM End Date"
        AESEQ = "AE Record position"
        AETERM = "What is the adverse event term"
        AESTDAT = "AE Start Date"
        AEENDTC = "AE End Date"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=statistical options(sheet_name="Listing06");

proc report data=final ;
run;

ods excel close;

 