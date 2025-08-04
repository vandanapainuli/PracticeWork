/* AGE in AE not matching AGE in demographics */
%let ae_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_AE.xlsx";
%let dm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_DM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS08.xlsx";


proc import datafile=&ae_file out=ae dbms=xlsx replace;
    getnames=yes;
run;

proc import datafile=&dm_file out=dm dbms=xlsx replace;
    getnames=yes;
run;

proc sort data=ae; by USUBJID; run;
proc sort data=dm; by USUBJID; run;

data merged;
    merge ae(in=a) dm(in=b);
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
    if missing(CHKREF) then CHKREF = "AE-DM Recon";
    if missing(VISITSEQ) then VISITSEQ = "NA";

    EXEC_DTTM = today();

    REVINST = "'Age at onset of SAE' on AE form should be equal to Age reported in Demographics form.";
    MSG = cats("Age at onset of SAE on AE form (",AEAGE,") is NOT equal to the Age (",AGE, ") reported in Demographics form. Please reconcile.");
run;

data filtered;
    set processed;
    AEAGE_num = input(AEAGE, best.);
    AGE_num = input(AGE, best.);
    if upcase(AESER) = "Y" and not missing(AEAGE_num) and (missing(AGE_num) or AEAGE_num ne AGE_num);
run;

data final;
retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           AETERM AESTDAT AESER AEAGE AGEU RFICDAT AGE REVINST MSG;
    set filtered;
    
    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         AETERM AESTDAT AESER AEAGE AGEU RFICDAT AGE REVINST MSG;

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
        AETERM = "What is the adverse event term"
        AESTDAT = "Start Date"
        AESER = "Serious AE"
        AEAGE = "Age at onset of SAE"
        AGEU = "Age unit"
        RFICDAT = "Informed consent date"
        AGE = "What is the subject's age"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
run;

ods excel file=&output_file style=journal;

proc report data=final;
run;

ods excel close;

