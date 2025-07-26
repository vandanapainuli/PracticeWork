
/* Step 1: Import the Excel file */
proc import datafile="/home/vandanapainuli0/EPG/DEMO_AE.xlsx"
    out=demo_ae dbms=xlsx replace;
    sheet="AE";
    getnames=yes;
run;

/* Step 2: Create derived variables and assign default values */
data listing1;
    set demo_ae;

    /* Create missing variables if they don't exist */
    length SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           REVINST MSG $200;

    /* Assign default or derived values */
    SITEID     = "A000";
    EXEC_DTTM  = put(today(), yymmdd10.);
    FORMID     = DOMAIN;
    FORMSEQ    = AESEQ;
    CHKREF     = "Data Dump";
    VISITSEQ   = "NA";
    SITENAME   = "Remote";
    SUBJID     = USUBJID;
    VISITID    = "AESAE";
    REVINST    = "Review data for completeness ";
    MSG        = "Some data is missing, please review and update. Else Clarify with the site.";

    /* Apply labels to match Python column headers */
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

    /* Keep variables in the exact order */
    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM
         AETERM AESTDAT AEENDTC AESER AEOUT REVINST MSG;
run;

/* Step 3: Export to Excel */
proc export data=listing1
    outfile="/home/vandanapainuli0/EPG/ListingSAS1.xlsx"
    dbms=xlsx replace;
    sheet="Listing1";
run;
