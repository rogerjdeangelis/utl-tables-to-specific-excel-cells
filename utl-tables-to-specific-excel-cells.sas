Tables to specific excel cells
 
  Two Solutions
 
      a. SAS Stack the output (other examples in links)
      b. R XLConnect package (python examples in links)
 
github
https://github.com/rogerjdeangelis/utl-tables-to-specific-excel-cells
 
SAS forum
https://tinyurl.com/6ut6prcu
https://communities.sas.com/t5/SAS-Enterprise-Guide/Exporting-multiple-SAS-tables-to-Specific-Excel-cells/m-p/723423
 
see
https://tinyurl.com/yasxfxwm
https://tinyurl.com/ybfam2ql
https://tinyurl.com/y8sm6xsv
https://tinyurl.com/sg5ohbp
https://tinyurl.com/y7g552rt
*
  __ _     ___  __ _ ___
 / _` |   / __|/ _` / __|
| (_| |_  \__ \ (_| \__ \
 \__,_(_) |___/\__,_|___/
 
;
* this will stack the reports on a sheet called SEX;
 
%utlfkil(d:/xls/sex.xlsx);
 
title;footnote;
ods excel file="d:/xls/sex.xlsx" style=minimal
  Options( sheet_name="sex" sheet_interval="none");
 
proc report data=sashelp.class(
     keep= name sex age
     where=(sex='M')) missing nowd;
run;quit;
 
ods excel Options( sheet_interval="none");
proc report data=sashelp.class(
     keep= name sex age
     where=(sex='F')) missing nowd;
run;quit;
 
ods excel close;
*_        ____
| |__    |  _ \
| '_ \   | |_) |
| |_) |  |  _ <
|_.__(_) |_| \_\
 
;
* This gives you more control
 
options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
  length species $5;
  set sashelp.fish (keep=species weight height width
      where=(species in ('Parkki' ,'Pike' ,'Smelt' ,'Whitefish')));
run;quit;
 
 
%utl_submit_r64('
source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
library(haven);
library(XLConnect);
have<-read_sas("d:/sd1/have.sas7bdat");
have;
White<-have[have$SPECIES=="White",];
parkk <-have[have$SPECIES=="Parkk",];
pike<-have[have$SPECIES=="Pike",];
smelt<-have[have$SPECIES=="Smelt",];
wb <- loadWorkbook("d:/xls/utl-excel-grid-of-four-reports-in-one-sheet.xlsx", create = TRUE);
createSheet(wb, name = "species");
writeWorksheet(wb, parkk , sheet = "species", startRow = 11,startCol = 10, header = TRUE);
writeWorksheet(wb, White, sheet = "species", startRow = 11,startCol = 4, header = TRUE);
writeWorksheet(wb, smelt, sheet = "species", startRow = 25,startCol = 10, header = FALSE);
writeWorksheet(wb, pike, sheet = "species", startRow = 25,startCol = 4, header = FALSE);
saveWorkbook(wb);
 
 
 
