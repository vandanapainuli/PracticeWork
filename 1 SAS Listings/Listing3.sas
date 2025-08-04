/* To identify duplicate CM records */
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS03.xlsx";

proc import datafile=&cm_file out=cm dbms=xlsx replace;
    getnames=yes;
run;

data processed;
    set cm;
    length SITEID SITENAME SUBJID VISITID FORMID FORMSEQ CHKREF VISITSEQ REVINST MSG $200;
    format EXEC_DTTM datetime20.;

    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "CM";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = CMSEQ;
    if missing(CHKREF) then CHKREF = "CM Duplicates";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();
    REVINST = "Check and query for all duplicate Concomitant Medications.";
    MSG = "Conmed records are duplicated. Please correct else clarify.";
run;

/* Filtering duplicates */
proc sql;
    create table filtered as
    select a.*
    from processed as a
    inner join (
        select USUBJID, CMTRT, CMDSTXT, CMDOSU, CMDOSFRQ, CMINDC, CMSTDAT, CMENDAT
        from processed
        group by USUBJID, CMTRT, CMDSTXT, CMDOSU, CMDOSFRQ, CMINDC, CMSTDAT, CMENDAT
        having count(*) > 1
    ) as b
    on a.USUBJID = b.USUBJID and
       a.CMTRT = b.CMTRT and
       a.CMDSTXT = b.CMDSTXT and
       a.CMDOSU = b.CMDOSU and
       a.CMDOSFRQ = b.CMDOSFRQ and
       a.CMINDC = b.CMINDC and
       a.CMSTDAT = b.CMSTDAT and
       a.CMENDAT = b.CMENDAT;
quit;

data final;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           CMTRT CMDSTXT CMDOSU CMDOSFRQ CMINDC CMSTDAT CMENDAT CMSEQ REVINST MSG;
    set filtered;

    label
        STUDYID   = "STUDYID"
        SITEID    = "SITENUMBER"
        SITENAME  = "SITE"
        SUBJID    = "SUBJECT"
        VISITID   = "FOLDERNAME"
        VISITSEQ  = "VISITSEQ"
        FORMID    = "FORMID"
        FORMSEQ   = "FORMSEQ"
        CHKREF    = "CHKREF"
        EXEC_DTTM = "Execution Date/Time"
        CMTRT     = "Medication or Therapy"
        CMDSTXT   = "Dose"
        CMDOSU    = "Dose Units"
        CMDOSFRQ  = "Frequency"
        CMINDC    = "Indication"
        CMSTDAT   = "CM Start Date"
        CMENDAT   = "CM End Date"
        CMSEQ     = "Record position"
        REVINST   = "Reviewer Instructions"
        MSG       = "Message";

    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         CMTRT CMDSTXT CMDOSU CMDOSFRQ CMINDC CMSTDAT CMENDAT CMSEQ REVINST MSG;
run;

proc export data=final
    outfile=&output_file
    dbms=xlsx replace;
    sheet="Listing03";
run;
