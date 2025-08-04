/*mentioned PROPHYLAXIS in text */
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS07.xlsx";


proc import datafile=&cm_file out=cm dbms=xlsx replace;
    getnames=yes;
run;

data processed;
    set cm;
    length SITEID SITENAME SUBJID VISITID FORMID CHKREF VISITSEQ REVINST MSG $200;
    format EXEC_DTTM datetime20.;

    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "CM";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = CMSEQ;
    if missing(CHKREF) then CHKREF = "CM Recon";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();

    REVINST = "If indication is reported as Prophylaxis or Other, then the Indication specification should not contain the term Prophylaxis.";
    MSG = "Indication is reported as Prophylaxis or Other; however, Indication specification contains the term Prophylaxis. Please check.";
run;

/* Step 4: Apply logic filter */
data filtered;
    set processed;
    if upcase(CMINDC) in ("PROPHYLAXIS", "OTHER") and index(upcase(CMINDDSC), "PROPHYLAXIS") > 0;
run;

data final;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           CMTRT CMINDC CMINDDSC CMSTDAT CMENDAT CMSEQ REVINST MSG;
    set filtered;


    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         CMTRT CMINDC CMINDDSC CMSTDAT CMENDAT CMSEQ REVINST MSG;

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
        CMINDDSC = "If indication is 'Prophylaxis'"
        CMSTDAT = "CM Start Date"
        CMENDAT = "CM End Date"
        CMSEQ = "Record position"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=statistical options(sheet_name="Listing07");

proc report data=final;
run;

ods excel close;
