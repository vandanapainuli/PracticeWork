proc import datafile="/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_AE.xlsx"
out=Demo_AE dbms=xlsx replace;
run;

data listing;
    retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM
           AETERM AESTDAT AEENDTC AESER AEOUT REVINST MSG;
    set Demo_AE;

    /* Assigning default values */
    SITEID     = "A000";
    EXEC_DTTM  = PUT(today(),datetime20.);
    FORMID     = DOMAIN;
    FORMSEQ    = AESEQ;
    CHKREF     = "Data Dump";
    VISITSEQ   = "NA";
    SITENAME   = "Remote";
    SUBJID     = USUBJID;
    VISITID    = "AESAE";
    REVINST    = "Review data for completeness ";
    MSG        = "Some data is missing, please review and update. Else Clarify with the site.";

    /* Applying labels */
    label
        STUDYID   = "STUDYID"
        SITEID    = "SITEID"
        SITENAME  = "Site Name"
        SUBJID    = "SUBJID"
        VISITID   = "VISITID"
        VISITSEQ  = "VISITSEQ"
        FORMID    = "FORMID"
        FORMSEQ   = "FORMSEQ"
        CHKREF    = "CHKREF"
        EXEC_DTTM = "Execution Date/Time"
        AETERM    = "What is the adverse event term"
        AESTDAT   = "Start Date"
        AEENDTC   = "End Date"
        AESER     = "IS AE Serious?"
        AEOUT     = "Outcome"
        REVINST   = "Reviewer Instructions"
        MSG       = "Message";

    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM
         AETERM AESTDAT AEENDTC AESER AEOUT REVINST MSG;
run;

ods excel file="/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS01.xlsx" style=statistical;

proc report data=listing;
run;
ods excel close;