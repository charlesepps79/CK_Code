*** M MORGAN 5/3/18                                                ***;  
*** THIS PROGRAM MATCHES ANN'S CREDIT KARMA LEADS FILE WITH        ***;
*** BRIAN'S APPS FILE (WHICH IS LIMITED TO JUST APPS FOR THE SSN'S ***;
*** ANN PROVIDES HIM), THEN PRODUCES A FILE WITH THE CREDIT KARMA  ***;
*** ID, LEAD INFO AND APP FLAGS. WILL BE RUN WEEKLY ON MONDAYS AND ***; 
*** OUTPUT CSV SENT TO CREDIT KARMA IN EXACT LAYOUT AND FILE       ***;
*** NAMING CONVENTION. EMAIL THE CSV TO                            ***;
*** PLOANREPORTING@CREDITKARMA.COM.  STILL NEED TO...              ***;
*** - CREATE NAMING CONVENTION FOR INCOMING FILES (LEADS FILE AND  ***;
*** APPS FILE).                                                    ***;
*** - CREATE A MACRO FOR NAMING THE OUTPUT CSV FILE.               ***;
*** CREDIT KARMA USES THIS FILE TO FINE TUNE THEIR LEAD TARGETING  ***;
*** SO THEY CAN IDENTIFY THE LEADS WITH THE BEST CHANCE OF         ***;
*** SUCCESS. ANY CHANGES MADE TO THIS CODE SHOULD ALSO BE MADE TO  ***;
*** THE SAS CODE FOR OUR INTERNAL REPORT. ------------------------ ***;

OPTIONS MPRINT MLOGIC SYMBOLGEN; /* SET DEBUGGING OPTIONS */

*** M DRIVE ------------------------------------------------------ ***;
LIBNAME MDRIVE '\\rmc.locAl\dfsroot\Dept\MArketINg\AnAlytics';

*** WEB REPORT MAIN DIRECTORY LOCATION --------------------------- ***;
%LET TCI_LOC = 
"\\rmc.locAl\dfsroot\Dept\MArketINg\AnAlytics\DigITAl\CredIT KArmA\INput LEAD AND App Files\CK LEADs 20180522.xlsx";
%LET APPFILE_LOC = 
"\\rmc.locAl\dfsroot\Dept\MArketINg\AnAlytics\DigITAl\CredIT KArmA\INput LEAD AND App Files\CredITKArmA523.xlsx";
%LET CK_OUTPUT = 
\\rmc.locAl\dfsroot\Dept\MArketINg\AnAlytics\DigITAl\CredIT KArmA\OUTput Result Files;
%LET CK_SUMMARY = 
\\rmc.locAl\dfsroot\Dept\MArketINg\AnAlytics\DigITAl\CredIT KArmA\CredIT KArmA Reports;

*** LEADS                                                          ***;
*** READ IN THE FILE WITH RECENT LEADS THAT ANN PROVIDED FROM TCI  ***;
PROC IMPORT 
	DATAFILE = &TCI_LOC.
	DBMS = EXCEL
	OUT = TCI 
	REPLACE;
RUN;

*** APPS FOR THOSE LEADS                                           ***;
*** READ IN THE FILE THAT BRIAN KILLEN CREATED OF NLS APPS BASED   ***;
*** ON THE SSNS IN ANN'S TCI FILE. SOON HE WILL JUST GIVE US THE   ***;
*** SQL QUERY ---------------------------------------------------- ***;
PROC IMPORT 
	DATAFILE = &APPFILE_LOC.
	DBMS = EXCEL
	OUT = APPFILE 
	REPLACE ;
RUN;

DATA LEADS;   
	SET TCI ;
	*** THE VAR NAME IN THE LEADS FILE HAS SPACES IN IT. ALSO IT'S ***;
	*** CHARACTER. CONVERTING TO A NUMERIC VAR NAMED SSN SO I CAN  ***;
	*** MATCH TO APPS FILE --------------------------------------- ***;
	SSN = 'ApplicAnt ssn'N * 1;  
	LEADDATE = INT('APPLICATION DATE'N);
	LEAD_SCORE = 'APPLICANT CREDIT SCORE'N;
	PQ_DECISION = 'DECISION STATUS'N;
	LEAD_TYPE = 'LOAN TYPE'N;
	LEAD_STATE = 'APPLICANT ADDRESS STATE'N;
	AMT_APPLIED = 'AMT# FIN#'N;
	CK_TRACKING_ID = 'SUB ID'N;
	IS_PREQUAL = 0;
	IF PQ_DECISION = 'Auto Approved' THEN IS_PREQUAL = 1;
	FORMAT LEADDATE yymmddd10.; /* s = slAsh, d = dAsh*/
RUN;

*** NOW THAT I'M GETTING FILES FROM ANN WITH MULTIPLE DAYS, I AM   ***;
*** GETTING MULTIPLE LEADS, WHERE THE FIRST LEAD HAS A CR SCORE    ***;
*** AND SHOWS PQ = 1 WHILE THE REST ARE DECLINED ----------------- ***;

*** OFTEN THE FIRST IS LT, FOLLOWED MINUTES LATER BY A WEB PQ. SO  ***;
*** LET'S DEDUPE AND KEEP THE FIRST LEAD? LEADDATE HAS TIME IN IT. ***;
*** BUT WHEN I SORT BY LEADDATE, IT IGNORES THE TIME VALUE AND     ***;
*** DOES NOT ALWAYS SORT RIGHT. I COULD SORT BY CR SCORE TO GET    ***;
*** THE FIRST LEAD.  OR I COULD JUST REMOVE LEADS WITH A VALUE OF  ***;
*** 'DUPLICATE' FOR DECISION STATUS ------------------------------ ***;
PROC SORT
	DATA = LEADS;
	BY SSN 
	DESCENDING LEAD_SCORE;
RUN;

DATA LEADS2;
	SET LEADS;
	BY SSN;
	*** WILL ONLY KEEP THE FIRST LEAD (SORTED BY CREDIT SCORE)     ***;
	IF FIRST.SSN; 
	*** NOW BRING IN THE APPS                                      ***;
RUN;

DATA APPS;
	SET APPFILE; /* FORMAT FUNDDATE yymmddd10. */
	IF APPLICATIONENTERDATE <  '1JAN2018'D THEN DELETE;
	IF FUNDEDDATE = 'NULL' THEN FUNDEDDATE = "";
	FUNDDATE = INPUT(STRIP(FUNDEDDATE), mmddyy10.);
	FORMAT APPLICATIONENTERDATE yymmddd10.  FUNDDATE mmddyy10.;
RUN;

*** NEED TO DEDUPE APPS. I WANT TO KEEP THE MOST SUCCESSFUL APP,   ***;
*** BUT ALSO WITH AN APP DATE AFTER THE LEAD DATE. SO I WILL       ***;
*** DEDUPE AFTER MERGING WITH LEADS ------------------------------ ***;
PROC SORT 
	DATA = APPS;
	BY SSN;
RUN;

DATA MERGE1;
	MERGE LEADS2(IN = INLEADS) APPS(IN = INAPPS);
	BY SSN;
	IF INLEADS;
	IF INAPPS THEN IS_APPL = 1;

	*** INITIALIZE SOME FUNNEL ATTRIBUTES. LATER CHANGE TO '1'     ***;
	*** BASED ON CONDITIONS -------------------------------------- ***;
	IS_APPL = 0;
	IS_APPROVE = 0;
	IS_FUNDED = 0;
	IF AMOUNTREQUESTED > 0 THEN IS_APPL = 1;

	*** SOME APPS WITH AN APPROVED AMOUNT WERE NOT GETTING FLAGGED ***;
	*** AS APPROVED HERE. TASKSTATUS OF PENDING APPROVAL --------- ***;
	IF TASKSTATUS IN ('FUNDED', 'ELIGIBLE') THEN IS_APPROVE = 1; 
	IF TASKSTATUS IN ('FUNDED') THEN IS_FUNDED = 1;
	AMT_APPROVED = MAX(SMALLAPPROVEDAMOUNT, LARGEAPPROVEDAMOUNT);
	IF LARGETERM > 0 THEN TERM_APPROVED = LARGETERM;
	ELSE TERM_APPROVED = SMALLTERM;
	IF LARGEAPR > 0 THEN RATE_APPROVED = LARGEAPR;
	ELSE RATE_APPROVED = SMALLAPR;
	AMOUNT_FUNDED = FINALLOANAMOUNT;
	TERM_FUNDED = 0;
	IF IS_FUNDED THEN TERM_FUNDED = FINALTERM;
	RATE_FUNDED = 0;
	IF IS_FUNDED = 1 THEN RATE_FUNDED = FINALAPR;
	ISPQAPP = 1;
	IF AMT_APPROVED > 0 AND REASONDECLINED = 'NULL' 
		THEN IS_APPROVE = 1;
	APP_BEFORE_LEAD = 0;

	*** UNDO FLAGS FOR APPS THAT CAME BEFORE THE LEAD DATE ------- ***;
	IF APPLICATIONENTERDATE < LEADDATE THEN DO;
		APP_BEFORE_LEAD = 1;
		APPLICATIONENTERDATE = .;
		FUNDDATE = .;
		IS_APPL = 0;
		IS_APPROVE = 0;
		IS_FUNDED = 0;
		AMT_APPROVED = 0;
		TERM_APPROVED = 0;
		RATE_APPROVED = 0; 
		AMOUNT_FUNDED = 0;
		TERM_FUNDED = 0;
		RATE_FUNDED = 0;
	 /* ISPQAPP=1; */
	END;

*** CALCULATE COST OF LEADS BASED ON AMOUNT FUNDED --------------- ***;
	COST = 0;
	IF 0 < AMOUNT_FUNDED LE 2500 THEN COST = 125;
	ELSE IF AMOUNT_FUNDED > 2500 THEN COST = 200;

	*** I SEE LEADS THAT DID NOT PASS PQ, YET THERE IS AN APPROVAL ***;
	*** ON AN APP FOR SOME --------------------------------------- ***;
 /* IF IS_PREQUAL = 0 AND IS_APPROVE = 1 AND PQ_DECISION = 'DUPLICATE' 
		THEN IS_PREQUAL = 1; */  
	IF IS_PREQUAL = 0 AND IS_APPROVE = 1 THEN IS_PREQUAL = 1;
	IF IS_APPROVE = 1 THEN IS_APPL = 1;
	IF IS_APPL = 1 THEN IS_PREQUAL = 1;
RUN;

*** DEDUPE BUT KEEP THE MOST SUCCESSFUL APP - AS LONG AS THE APP   ***;
*** WAS NOT BEFORE THE LEAD -------------------------------------- ***;
PROC SORT 
	DATA = MERGE1;  
	BY SSN 
	DESCENDING APP_BEFORE_LEAD 
	DESCENDING IS_FUNDED 
	DESCENDING IS_APPROVE;
RUN;

DATA MERGE2;
	SET MERGE1;
	BY SSN;
	IF FIRST.SSN;
RUN;

*** FORMAT STATEMENT BEFORE THE SET STATEMENT WILL CHANGE THE      ***;
*** ORDER OF THE VARIABLES SO THOSE IN THE (OTHERWISE EMPTY)
*** FORMAT STATEMENT APPEAR FIRST. ------------------------------- ***;
DATA LAST(
	KEEP = CK_TRACKING_ID LEADDATE APPLICATIONENTERDATE FUNDDATE 
		   ISPQAPP IS_PREQUAL IS_APPL IS_APPROVE IS_FUNDED AMT_APPLIED
		   AMT_APPROVED TERM_APPROVED RATE_APPROVED AMOUNT_FUNDED 
		   TERM_FUNDED RATE_FUNDED);
	FORMAT CK_TRACKING_ID LEADDATE APPLICATIONENTERDATE FUNDDATE
		   ISPQAPP IS_PREQUAL IS_APPL IS_APPROVE IS_FUNDED AMT_APPLIED 
		   AMT_APPROVED TERM_APPROVED RATE_APPROVED AMOUNT_FUNDED 
		   TERM_FUNDED RATE_FUNDED;
 /* LEAD_SCORE PQ_DECISION  LEAD_TYPE  LEAD_STATE); */
	SET MERGE2;
RUN;

PROC SORT DATA = LAST;
	BY DESCENDING IS_FUNDED;
RUN;

*** LOOKING AT ALL COMBINATIONS OF RESULT FLAGS TO MAKE SURE WE    ***;
*** DON'T HAVE SOME LEADS WITH A ZERO IN THE MIDDLE OF THE FUNNEL  ***; 
*** BUT NOT THE BOTTOM ------------------------------------------- ***;
PROC FREQ;
	TABLES ISPQAPP *
		   IS_PREQUAL *
		   IS_APPL *
		   IS_APPROVE *
		   IS_FUNDED / 
	LIST;
RUN;

*** SEND THIS RESULTS FILE TO CREDIT KARMA EACH FRIDAY ----------- ***;
PROC EXPORT 
	DATA = LAST 
	OUTFILE =  "&CK_OUTPUT\CreditKarma_Regional_05232018.csv"
	DBMS = CSV
	REPLACE;
RUN;

*** NOW CREATE A SUMMARY REPORT BY LEAD DATE FOR ANN & JIM. LATER  ***;
*** CHANGE THIS TO SUMMARIZE BY WEEK INSTEAD OF DATE ------------- ***;
DATA TEMP;
	SET MERGE2;
RUN;

PROC SORT 
	DATA=TEMP;
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR IS_PREQUAL;
	OUTPUT OUT = VAR1OUT N = NUM_LEADS SUM = NUM_PQ;
RUN;

PROC SORT 
	DATA = TEMP;
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR IS_APPL;
	OUTPUT OUT = VAR2OUT SUM = NUM_APPS;
RUN;

PROC SORT 
	DATA = TEMP;
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR IS_APPROVE;                 
	OUTPUT OUT = VAR3OUT SUM = NUM_APPR;
RUN;

PROC SORT 
	DATA = TEMP;       
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR IS_FUNDED;                
	OUTPUT OUT = VAR4OUT SUM = NUM_BOOKED;
RUN;

PROC SORT 
	DATA = TEMP;       
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR AMOUNT_FUNDED;                 
	OUTPUT OUT = VAR5OUT SUM = DOLLARS_BOOKED;
RUN;

PROC SORT 
	DATA = TEMP;       
	BY LEADDATE;
RUN;

PROC SUMMARY 
	DATA = TEMP;  
	CLASS LEADDATE;
	VAR COST;
	OUTPUT OUT = VAR6OUT SUM=LEAD_COST;
RUN;

DATA SUMMARY(DROP = _FREQ_ );
	MERGE VAR1OUT VAR2OUT VAR3OUT VAR4OUT VAR5OUT VAR6OUT;
	BY LEADDATE;
	IF LEADDATE = . THEN DELETE;
	DROP _TYPE_;
RUN;

*** NOW OPEN UP THE ABOVE EXCEL SPREADSHEET AND COPY THE VALUES    ***;
*** INTO THE TEMPLATE REPORT THAT HAS MORE COLUMNS WITH RATIOS     ***;
*** CALCULATED AND HAS A TOTAL ROW AND FORMATTING.                 ***;  
*** SEND THE REPORT TO ANN --------------------------------------- ***;

PROC EXPORT 
	DATA = SUMMARY 
	OUTFILE = "&CK_SUMMARY\CKRPT_05232018.xlsx" 
	REPLACE ;
RUN;
