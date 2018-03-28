
/*Importing Excel Files TO SAS*/

%let location= /folders/myfolders/sasuser.v94/data/CHANNEL WISE INWARDS & NCA/;

%macro Import(file_name,ds_name,db_name);
   proc import
   datafile="&location&file_name"
   out=&ds_name
   dbms=&db_name
   replace;
   run;
%mend;

%import(INW.xls,INWARDS,xls);
%import(NCA.xls,NCA,xls);
%import(ACC.xls,ACC,xls);


/*Creating Macros For dates*/
 
%global day yday month pre_month year pre_year;

%macro dates;
%let day=%sysfunc(today(),date9.);
%let yday=%sysfunc(intnx(day,"&day"d,-1),date9.);
%let month=%sysfunc(today(),monname3.);
%let pre_month=%sysfunc(intnx(month,"&day"d,-1),monname3.);
%let year=%sysfunc(today(),year4.);
%let pre_year=%sysfunc(intnx(year,"&day"d,-1),year4.);
/* %put &day &yday &month &pre_month &year &pre_year; */
%mend;
%dates;


/*Title and Footnote*/

title height=12pt "Channel Wise Inwards and NCA Reports as on &day";
title2" ";
title3 height=12pt "Inwards";

PROC REPORT DATA=WORK.INWARDS LS=132 PS=60  SPLIT="/" CENTER  ;
 COLUMN FUN_TEAM
 		("TOTAL" TOTAL_FTD TOTAL_FTD_LAST TOTAL_MTD TOTAL_MTD_LAST )
 		("RETAIL" RETAIL_FTD RETAIL_FTD_LAST RETAIL_MTD RETAIL_MTD_LAST)
 		("CORPORATE" CORPORATE_FTD CORPORATE_FTD_LAST CORPORATE_MTD CORPORATE_MTD_LAST)
 		("NRI" NRI_FTD NRI_FTD_LAST NRI_MTD NRI_MTD_LAST )
 		("ATS" ATS_ONLINE_FTD  ATS_ONLINE_LAST ATS_ONLINE_MTD ATS_ONLINE_LAST_MTD)
 		("INVESTMENT ACCOUNT" INV_ONLINE_FTD INV_ONLINE_LAST INV_ONLINE_MTD INV_ONLINE_LAST_MTD) 
 		("TRADE RACER" TRADE_RACER_ONLINE_FTD TRADE_RACER_ONLINE_LAST TRADE_RACER_ONLINE_MTD
 		TRADE_RACER_ONLINE_LAST_MTD);
  
 DEFINE  FUN_TEAM / DISPLAY FORMAT= $20. WIDTH=20    SPACING=2   LEFT 
 style(column header)=[foreground=white background=grey] "TEAM" ;
 
 DEFINE  TOTAL_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT  
 style(column header)=[foreground=white background=black] "As on/&day";
 DEFINE  TOTAL_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT 
 style(column header)=[foreground=white background=black] "As on/&yday" ;
 DEFINE  TOTAL_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT 
 style(column header)=[foreground=white background=black] "MTD-/&month" ;
 DEFINE  TOTAL_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT
 style(column header)=[foreground=white background=black]  "MTD-/&pre_month";
 
 DEFINE  RETAIL_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  RETAIL_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  RETAIL_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  RETAIL_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  CORPORATE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  CORPORATE_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  CORPORATE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  CORPORATE_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;

 DEFINE  NRI_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  NRI_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  NRI_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  NRI_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;

 DEFINE  ATS_ONLINE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  ATS_ONLINE_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  ATS_ONLINE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  ATS_ONLINE_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;

 DEFINE  INV_ONLINE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  INV_ONLINE_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  INV_ONLINE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  INV_ONLINE_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;

 DEFINE  TRADE_RACER_ONLINE_FTD / SUM FORMAT= BEST12. WIDTH=12  SPACING=2   RIGHT "As on/&day" ;
 DEFINE  TRADE_RACER_ONLINE_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  TRADE_RACER_ONLINE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  TRADE_RACER_ONLINE_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;

 rbreak after/summarize style={background=grey};
 RUN;

 
/*Report 2*/
title height=12pt "NCA";


 PROC REPORT DATA=WORK.NCA LS=132 PS=60  SPLIT="/" CENTER ;
 COLUMN  TEAM  
 		("TOTAL" TOTAL_FTD  TOTAL_FTD_LAST  TOTAL_MTD  TOTAL_MTD_LAST)
 		("RETAIL" RETAIL_FTD  RETAIL_FTD_LAST  RETAIL_MTD RETAIL_MTD_LAST)
 		("CORPORATE"  CORPORATE_FTD  CORPORATE_FTD_LAST CORPORATE_MTD CORPORATE_MTD_LAST)
 		("NRI" NRI_FTD NRI_FTD_LAST NRI_MTD NRI_MTD_LAST)
 		("ATS" ATS_ONLINE_FTD ATS_ONLINE_LAST  ATS_ONLINE_MTD ATS_ONLINE_LAST_MTD)
 		("INVESTMENT ACCOUNT" INV_ONLINE_FTD INV_ONLINE_LAST INV_ONLINE_MTD INV_ONLINE_LAST_MTD)
 		("TRADER RACER" TRADE_RACER_FTD TRADE_RACER_LAST TRADE_RACER_MTD TRADE_RACER_LAST_MTD);
  
  
 DEFINE  TEAM / DISPLAY FORMAT= $19. WIDTH=19    SPACING=2   LEFT 
  style(column header)=[foreground=white background=grey] "TEAM" ;
 
 DEFINE  TOTAL_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT
  style(column header)=[foreground=white background=black] "As on/&day" ;
 DEFINE  TOTAL_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT
  style(column header)=[foreground=white background=black] "As on/&yday" ;
 DEFINE  TOTAL_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT
  style(column header)=[foreground=white background=black] "MTD-/&month" ;
 DEFINE  TOTAL_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT 
  style(column header)=[foreground=white background=black] "MTD-/&pre_month" ;
 
 DEFINE  RETAIL_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  RETAIL_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  RETAIL_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  RETAIL_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  CORPORATE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  CORPORATE_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  CORPORATE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  CORPORATE_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  NRI_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  NRI_FTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  NRI_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  NRI_MTD_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  ATS_ONLINE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  ATS_ONLINE_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  ATS_ONLINE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  ATS_ONLINE_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  INV_ONLINE_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  INV_ONLINE_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  INV_ONLINE_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  INV_ONLINE_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 DEFINE  TRADE_RACER_FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&day" ;
 DEFINE  TRADE_RACER_LAST / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on/&yday" ;
 DEFINE  TRADE_RACER_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&month" ;
 DEFINE  TRADE_RACER_LAST_MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-/&pre_month" ;
 
 rbreak after/summarize style={background=grey};
 
 compute after;
 if team=" " then team="Total";
 endcomp;
 
 RUN;
 
 /* Report 3*/

title height=12pt "Accounts Opened as on &day";
footnote height=13pt "Note: This is a system generated mail. Please do not respond to this mail.";

 PROC REPORT DATA=WORK.ACC LS=132 PS=60  SPLIT="/" CENTER ;

 COLUMN  TEAM FTD MTD THIS_FY;
  
 DEFINE  TEAM / DISPLAY FORMAT= $19. WIDTH=19    SPACING=2   LEFT 
  style(column header)=[foreground=white background=black] "TEAM" ;
 DEFINE  FTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "As on &day" ;
 DEFINE  MTD / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "MTD-&month" ;
 DEFINE  THIS_FY / SUM FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "YTD &year-&pre_year" ;


 rbreak after/summarize style={background=grey};
 
 compute after;
 if team=" " then team="Total";
 endcomp;
 
 RUN;
