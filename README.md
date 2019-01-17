# utl-excel-import-individual-cells
Read individual cells A20 amd C3 in excel sheet 'class' or named range 'class'
    Excel import individual cells

      Read individual cells A20 amd C3 in excel sheet 'class' or named range 'class'

       1. Datastep Libname engine;
       2. Proc sql
       3. R xlconnect (uses excel cell references)

    github
    https://github.com/rogerjdeangelis/utl-excel-import-individual-cells

    SAS Forum
    https://tinyurl.com/ybnwfhfg
    https://communities.sas.com/t5/SAS-Programming/Pulling-in-Values-from-Specific-Cells-of-Excel-SpreadSheet-and/m-p/527863


    INPUT
    =====

    %utlfkil(d:/xls/class.xlsx);
    libname xel "d:/xls/class.xlsx";
    data xel.class;
      set sashelp.class;
    run;quit;
    libname xel clear;


     d:/xls/class.xlsx   ( read A20 and C3)

         +----------------------------------------------------------------+
         |     A      |    B       |     C      |    D       |    E       |
         +----------------------------------------------------------------+
       1 |  NAME      |   SEX      |    AGE     |  HEHT    |  WEHT    |
         +------------+------------+------------+------------+------------+
       2 | Alfred     |    M       |    14      |    69      |  112.5     |
         +------------+------------+------------+------------+------------+
       3 | Alice      |    F       | ---13------|    68      |  102.5     |
         +------------+------------+------------+------------+------------+
          ...
         +------------+------------+------------+------------+------------+
      20 |---WILLIAM--|    F       |    15      |   66.5     |  112       |
         +------------+------------+------------+------------+------------+




    EXAMPLE OUTPUT (XLconnect output)
    ----------------------------------

      WORK.WANT_R total obs=1

          NAME      AGE

         William     13


    SOLUTIONS
    =========

       1. Datastep Libname engine;
       ---------------------------

          libname xel "d:/xls/class.xlsx";
          data want;
            set xel.'class$A19:A20'n;
            set xel.'class$C2:C3'n;
          run;quit;

          /*
          WORK.WANT total obs=1  (uses the revious row value for name)

            THOMAS     _4

             William    13
          */

       2. Proc sql
       ------------

          libname xel "d:/xls/class.xlsx";
          proc sql;
             select
                 l.*
                ,r.*
             from
                xel.'class$A19:A20'n as l, xel.'class$C2:C3'n as r
          ;quit;
          libname xel clear;

          /*
          WORK.WANT total obs=1  (uses the revious row value for name)

            THOMAS     _4

             William    13
          */

      3. R xlconnect (uses excel cell references)
      --------------------------------------------

       %utl_submit_r64('
       library(XLConnect);
       library(SASxport);
       wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);
       A20<-readWorksheet(wb,sheet="class",region="A20:A20",header=FALSE);
       C3<-readWorksheet(wb,sheet="class",region="C3:C3",header=FALSE);
       want<-as.data.frame(cbind(A20,C3));
       colnames(want)<-c("name","age");
       str(want);
       write.xport(want,file="d:/xpt/want.xpt");
       want;
       ');

       libname xpt xport "d:/xpt/want.xpt";
       data want_r;
         set xpt.want;
       run;quit;
       libname xpt clear;

        WORK.WANT_R total obs=1

            NAME      AGE

           William     13



