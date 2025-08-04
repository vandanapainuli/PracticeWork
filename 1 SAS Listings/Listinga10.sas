/* CM-dose mentioned but units missing */
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSASa10.xlsx";

proc import datafile=&cm_file out=cm dbms=xlsx replace;
    getnames=yes;
run;

data processed;
    set cm;
    length SITEID SITENAME SUBJID VISITID FORMID CHKREF VISITSEQ REVINST MSG $200;
    format EXEC_DTTM datetime20.;

    if missing(SITEID) then SITEID = "N001";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "CM";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = CMSEQ;
    if missing(CHKREF) then CHKREF = "CM CLEAN UP";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();

    REVINST = "If Dose is not provided, then Dose Unit is also not expected to reported or vice versa.";
    MSG = "Dose is not recorded; however Dose Unit is provided. Please check OR Dose is recorded; however Dose Unit is not provided. Please resolve.";
run;

data filtered;
    set processed;
    if (missing(CMDSTXT) and not missing(CMDOSU)) or 
       (not missing(CMDSTXT) and missing(CMDOSU));
run;

/* Step 5: Keep and rename columns */
data final;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           CMTRT CMDSTXT CMDOSU CMDOSFRQ CMINDC CMSTDAT CMENDAT CMSEQ REVINST MSG;
    set filtered;

    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         CMTRT CMDSTXT CMDOSU CMDOSFRQ CMINDC CMSTDAT CMENDAT CMSEQ REVINST MSG;

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
        CMDSTXT = "Dose"
        CMDOSU = "Dose Units"
        CMDOSFRQ = "Frequency"
        CMINDC = "Indication"
        CMSTDAT = "CM Start Date"
        CMENDAT = "CM End Date"
        CMSEQ = "Recordposition"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=statistical;

proc report data=final;
run;

ods excel close;