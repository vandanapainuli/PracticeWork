/* To identify duplicate MH records */
%let mh_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_MH.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS04.xlsx";

proc import datafile=&mh_file out=mh dbms=xlsx replace;
    getnames=yes;
run;

data processed;
    set mh;
    length SITEID SITENAME SUBJID VISITID FORMID FORMSEQ CHKREF VISITSEQ REVINST MSG $200;
    format EXEC_DTTM yymmdd10.;

    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "SCRN";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = MHSEQ;
    if missing(CHKREF) then CHKREF = "MH Duplicates";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();
    REVINST = "Check for duplicate Medical history for the same subject.";
    MSG = "Medical history records are duplicated. Please check.";
run;



/* Step 1: Identify duplicate key combinations */
proc sql;
    create table duplicate_keys as
    select USUBJID, MHTERM, MHSTDAT, MHONGO, MHENDAT, MHTOXGR
    from processed
    group by USUBJID, MHTERM, MHSTDAT, MHONGO, MHENDAT, MHTOXGR
    having count(*) > 1;
quit;

/* Step 2: Join back to get all matching duplicate records */
proc sql;
    create table duplicates as
    select a.*
    from processed as a
    inner join duplicate_keys as b
    on a.USUBJID = b.USUBJID and
       a.MHTERM = b.MHTERM and
       a.MHSTDAT = b.MHSTDAT and
       a.MHONGO = b.MHONGO and
       a.MHENDAT = b.MHENDAT and
       a.MHTOXGR = b.MHTOXGR;
quit;

data final;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
            MHSEQ MHTERM MHSTDAT MHONGO MHENDAT MHTOXGR REVINST MSG;
    set duplicates;

    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
          MHSEQ MHTERM MHSTDAT MHONGO MHENDAT MHTOXGR REVINST MSG;

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
        MHYN = "Abnormalities"
        MHSEQ = "Record position"
        MHTERM = "Verbatim term for MH condition/event"
        MHSTDAT = "MH Start Date"
        MHONGO = "Ongoing"
        MHENDAT = "End Date"
        MHTOXGR = "Toxicity Grade"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=statistical;

proc report data=final;
run;

ods excel close;

