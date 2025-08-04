/* Date inconsisteny for medication taken for medical history */
%let cm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_CM.xlsx";
%let mh_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_MH.xlsx";
%let dm_file = "/home/vandanapainuli0/EPG/ProjectDatasets/DEMO_DM.xlsx";
%let output_file = "/home/vandanapainuli0/EPG/ListingOUTPUTS/ListingSAS05.xlsx";

proc import datafile=&cm_file out=cm dbms=xlsx replace;
    getnames=yes;
run;

proc import datafile=&mh_file out=mh dbms=xlsx replace;
    getnames=yes;
run;

proc import datafile=&dm_file out=dm dbms=xlsx replace;
    getnames=yes;
run;

/* Step 3: Sort and merge datasets */
proc sort data=cm; by USUBJID; run;
proc sort data=mh; by USUBJID; run;
proc sort data=dm; by USUBJID; run;

data merged;
    merge cm(in=a) mh(in=b) dm(in=c);
    by USUBJID;
run;

/* Step 4: Add derived columns and default values */
data processed;
    set merged;
    length SITEID SITENAME SUBJID VISITID FORMID CHKREF VISITSEQ REVINST  $200;
    length MSG $200;
    format EXEC_DTTM date9.;

    if missing(SITEID) then SITEID = "A000";
    if missing(SITENAME) then SITENAME = "Remote";
    if missing(SUBJID) then SUBJID = USUBJID;
    if missing(VISITID) then VISITID = "CM";
    if missing(FORMID) then FORMID = DOMAIN;
    if missing(FORMSEQ) then FORMSEQ = CMSEQ;
    if missing(CHKREF) then CHKREF = "MH-CM-DM Recon";
    if missing(VISITSEQ) then VISITSEQ = "NA";
	
    EXEC_DTTM = today();

    REVINST = "If Indication is reported as 'Medical History' in Concomitant Therapy form, then the Start date of Medical History and Concomitant Medication should be before the date of Informed consent date.";
	MSG = "Indication in Concomitant Therapy form is reported as 'Medical History'. However, the Start date of Medical history or Concomitant Therapy is not before the date of Informed consent. Please check.";

run;

data filtered;
    set processed;
    format CMSTDAT MHSTDAT RFICDAT date9.;
    if upcase(CMINDC) = "MEDICAL HISTORY" and 
       (MHSTDAT >= RFICDAT or CMSTDAT >= RFICDAT);
run;

data final;
        retain STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
           CMTRT CMINDC CMSTDAT MHSEQ MHCAT MHSTDAT MHENDAT MHONGO MHTERM FOLDERNAME RFICDAT 
           REVINST MSG;
    set filtered;
    keep STUDYID SITEID SITENAME SUBJID VISITID VISITSEQ FORMID FORMSEQ CHKREF EXEC_DTTM 
         CMTRT CMINDC CMSTDAT MHSEQ MHCAT MHSTDAT MHENDAT MHONGO MHTERM FOLDERNAME RFICDAT 
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
        CMSTDAT = "CM Start Date"
        MHSEQ = "MH Record position"
        MHCAT = "MH Category"
        MHSTDAT = "MH Start Date"
        MHENDAT = "MH End Date"
        MHONGO = "Ongoing"
        MHTERM = "Verbatim term for MH condition/event"
        FOLDERNAME = "FOLDERNAME_DM"
        RFICDAT = "Informed consent date"
        REVINST = "Reviewer Instructions"
        MSG = "Message";
        
run;


ods excel file=&output_file style=statistical options(sheet_name="Listing04");

proc report data=final nowd;

run;

ods excel close;
