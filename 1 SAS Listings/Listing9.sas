/* Indication is AE butdata inconsistent */
%let ae_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_AE.xlsx";
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS09.xlsx";


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

    REVINST = "If Indication for this CM is 'Adverse Event', there must be an Adverse Event recorded where the question 'concomitant treatment?' is answered yes.";
    MSG = "Indication for this CM is 'Adverse Event', but there is no Adverse Event recorded where the question 'concomitant treatment?' is answered yes. Please check and confirm.";
run;

/*   logic  */
data filtered;
    set processed;
    if upcase(CMINDC) = "ADVERSE EVENT" and 
       upcase(AECONTRT) = "Y" and 
       missing(AEDSL1) and missing(AEDSL2) and missing(AEDSL3) and missing(AEDSL4) and missing(AEDSL5);
run;

data final;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           AETERM AESTDAT AEENDTC AECONTRT CMTRT CMINDC CMSTDAT AEDSL1 AEDSL2 AEDSL3 AEDSL4 AEDSL5 
           REVINST MSG;
    set filtered;

    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         AETERM AESTDAT AEENDTC AECONTRT CMTRT CMINDC CMSTDAT AEDSL1 AEDSL2 AEDSL3 AEDSL4 AEDSL5 
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
        AETERM = "Adverse Event"
        AESTDAT = "Start Date"
        AEENDTC = "End Date"
        AECONTRT = "Concomitant treatment given for AE"
        CMTRT = "Medication or Therapy"
        CMINDC = "Indication"
        CMSTDAT = "CM Start Date"
        AEDSL1 = "AE log line start date and term 1"
        AEDSL2 = "AE log line start date and term 2"
        AEDSL3 = "AE log line start date and term 3"
        AEDSL4 = "AE log line start date and term 4"
        AEDSL5 = "AE log line start date and term 5"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=statistical options(sheet_name="Listing09");
proc report data=final ;
run;
ods excel close;
