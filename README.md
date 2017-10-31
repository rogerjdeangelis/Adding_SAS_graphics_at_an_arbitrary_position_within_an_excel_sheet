# Adding_SAS_graphics_at_an_arbitrary_position_within_an_excel_sheet
Adding SAS graphics at an arbitrary position within an excel sheet

    ```  ODS EXCEL to have multiple tabs and embed image(logo) in title                                                                                               ```
    ```                                                                                                                                                               ```
    ```  Adding SAS graphics at an arbitrary position into existing excel sheets (9.4M2)                                                                              ```
    ```                                                                                                                                                               ```
    ```  related post                                                                                                                                                 ```
    ```                                                                                                                                                               ```
    ```  "According to Chevel Parker(SAS), you cannot insert an image directly with ODS EXCEL. You have to post process."                                             ```
    ```                                                                                                                                                               ```
    ```  /* T008420 Adding SAS graphics at an arbitrary position within an excel sheet;                                                                               ```
    ```                                                                                                                                                               ```
    ```  HAVE                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```   1.  Workbook created using SAS, d:/xls/xlconnect_class.xlsx, with a SAS produced report                                                                     ```
    ```   2,  Have a four panel graph, d:/png/xlconnect_class.png, produced by SAS                                                                                    ```
    ```                                                                                                                                                               ```
    ```  WANT                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```   To place the graph at an arbitrary position. above, below or beside the SAS report.                                                                         ```
    ```                                                                                                                                                               ```
    ```  SOLUTION                                                                                                                                                     ```
    ```                                                                                                                                                               ```
    ```  * 1. SAS: create a excel workbook and a report;                                                                                                              ```
    ```  * 2. SAS: create a SAS png graph to be inserted into excel sheet                                                                                             ```
    ```  * 3. R: Insert graph anywhere on the sheet ie  at col G row 2                                                                                                ```
    ```                                                                                                                                                               ```
    ```  github                                                                                                                                                       ```
    ```  https://goo.gl/YrbkMe                                                                                                                                        ```
    ```  https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets                                               ```
    ```                                                                                                                                                               ```
    ```  see                                                                                                                                                          ```
    ```  https://goo.gl/s4KqY3                                                                                                                                        ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/ODS-EXCEL-to-have-multiple-tabs-and-embed-image-logo-in-title/m-p/409058                               ```
    ```                                                                                                                                                               ```
    ```  https://goo.gl/JBQFnA                                                                                                                                        ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/ODS-EXCEL-AND-HOW-TO-INSERT-AN-IMAGE/td-p/289493                                                       ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * create a workbook and ad report in excel  d:/xls/xlconnect_class.xlsx;                                                                                     ```
    ```  %utlfkil(d:/xls/xlconnect_class.xlsx);                                                                                                                       ```
    ```  libname xls "d:/xls/xlconnect_class.xlsx";                                                                                                                   ```
    ```  data xls.class;                                                                                                                                              ```
    ```   set sashelp.class;                                                                                                                                          ```
    ```  run;quit;                                                                                                                                                    ```
    ```  libname xls clear;                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  * create a four panel graph d:/png/xlconnect_class.png;                                                                                                      ```
    ```  ods listing  gpath="d:/png";                                                                                                                                 ```
    ```  ods graphics on  /                                                                                                                                           ```
    ```             reset=all                                                                                                                                         ```
    ```             reset=index                                                                                                                                       ```
    ```             imagefmt=png                                                                                                                                      ```
    ```             imagename="xlconnect_class"                                                                                                                       ```
    ```             height=400px                                                                                                                                      ```
    ```             width=400px                                                                                                                                       ```
    ```  ;                                                                                                                                                            ```
    ```  PROC TEMPLATE; DEFINE STATGRAPH Panel;                                                                                                                       ```
    ```  BEGINGRAPH;                                                                                                                                                  ```
    ```  ENTRYTITLE "Paneled Display ";                                                                                                                               ```
    ```     LAYOUT LATTICE / ROWS = 2 COLUMNS = 2 ROWGUTTER = 10 COLUMNGUTTER = 10;                                                                                   ```
    ```       LAYOUT OVERLAY; SCATTERPLOT Y = Weight X = Height;                                                                                                      ```
    ```         REGRESSIONPLOT Y = Weight X = Height;                                                                                                                 ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY / XAXISOPTS = (LABEL = "Weight");                                                                                                        ```
    ```         HISTOGRAM Weight;                                                                                                                                     ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY / YAXISOPTS = (LABEL = "Height");                                                                                                        ```
    ```         BOXPLOT Y = Height;                                                                                                                                   ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY; SCATTERPLOT Y = weight X = height /                                                                                                     ```
    ```         GROUP = sex NAME = "Scat";                                                                                                                            ```
    ```         DISCRETELEGEND "Scat"                                                                                                                                 ```
    ```         / TITLE = "Sex";                                                                                                                                      ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```     ENDLAYOUT;                                                                                                                                                ```
    ```   ENDGRAPH;                                                                                                                                                   ```
    ```  END;                                                                                                                                                         ```
    ```  RUN;                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  PROC SGRENDER DATA = Sashelp.Class TEMPLATE = Panel;                                                                                                         ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  ods graphics off;                                                                                                                                            ```
    ```  ods html close;                                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * create a png graphic to add to sheet mtcars below;                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  %utl_submit_r(                                                                                                                                               ```
    ```     library('XLConnect');                                                                                                                                     ```
    ```     wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');                                                                                                        ```
    ```     createName(wb, name = 'class_png', formula = 'class!$G$2');                                                                                               ```
    ```     addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',originalSize = TRUE);                                                            ```
    ```     saveWorkbook(wb);                                                                                                                                         ```
    ```  );                                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  * if you want to know the active sheet and last row and last column;                                                                                         ```
    ```  %utl_submit_r(                                                                                                                                               ```
    ```     library('XLConnect');                                                                                                                                     ```
    ```     wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');                                                                                                        ```
    ```     activeSheetIndex <- getActiveSheetIndex(wb);                                                                                                              ```
    ```     activeSheetName <- getActiveSheetName(wb);                                                                                                                ```
    ```     activeSheetName;                                                                                                                                          ```
    ```     activeSheetIndex;                                                                                                                                         ```
    ```     LastColumn<-getLastColumn(wb, 'class');                                                                                                                   ```
    ```     LastRow   <-getLastRow(wb, 'class');                                                                                                                      ```
    ```     LastColumn;                                                                                                                                               ```
    ```     LastRow;                                                                                                                                                  ```
    ```     createName(wb, name = 'class_png', formula = 'class!$G$2');                                                                                               ```
    ```     addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',originalSize = TRUE);                                                            ```
    ```     saveWorkbook(wb);                                                                                                                                         ```
    ```  );                                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  LOG                                                                                                                                                          ```
    ```                                                                                                                                                               ```
    ```  NOTE: 10 lines were written to file PRINT.                                                                                                                   ```
    ```  Stderr output:                                                                                                                                               ```
    ```  Loading required package: XLConnectJars                                                                                                                      ```
    ```  XLConnect 0.2-12 by Mirai Solutions GmbH [aut],                                                                                                              ```
    ```    Martin Studer [cre],                                                                                                                                       ```
    ```    The Apache Software Foundation [ctb, cph] (Apache POI, Apache Commons                                                                                      ```
    ```      Codec),                                                                                                                                                  ```
    ```    Stephen Colebourne [ctb, cph] (Joda-Time Java library),                                                                                                    ```
    ```    Graph Builder [ctb, cph] (Curvesapi Java library)                                                                                                          ```
    ```  http://www.mirai-solutions.com ,                                                                                                                             ```
    ```  http://miraisolutions.wordpress.com                                                                                                                          ```
    ```  NOTE: 8 records were read from the infile RUT.                                                                                                               ```
    ```        The minimum record length was 2.                                                                                                                       ```
    ```        The maximum record length was 504.                                                                                                                     ```
    ```  NOTE: DATA statement used (Total process time):                                                                                                              ```
    ```        real time           2.88 seconds                                                                                                                       ```
    ```        cpu time            0.04 seconds                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  NOTE: Fileref RUT has been deassigned.                                                                                                                       ```
    ```  NOTE: Fileref R_PGM has been deassigned.                                                                                                                     ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  OUTPUT                                                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```  > library('XLConnect');    wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');    activeSheetIndex <- getActiveSheetIndex(wb);                                ```
    ```      activeSheetName <- getActiveSheetName(wb);                                                                                                               ```
    ```      activeSheetName;    activeSheetIndex;    LastColumn<-getLastColumn(wb, 'class');    LastRow   <-getLastRow(wb, 'class');                                 ```
    ```  LastColumn;    LastRow;    createName(wb, n                                                                                                                  ```
    ```  ame = 'class_png', formula = 'class!$G$2');    addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',                                     ```
    ```  originalSize = TRUE);    saveWorkbook(wb);                                                                                                                   ```
    ```  [1] "class"                                                                                                                                                  ```
    ```  [1] 1                                                                                                                                                        ```
    ```  class columns                                                                                                                                                ```
    ```      5                                                                                                                                                        ```
    ```  class rows                                                                                                                                                   ```
    ```     20                                                                                                                                                        ```
    ```                                                                                                                                                               ```

    ```  ODS EXCEL to have multiple tabs and embed image(logo) in title                                                                                               ```
    ```                                                                                                                                                               ```
    ```  Adding SAS graphics at an arbitrary position into existing excel sheets (9.4M2)                                                                              ```
    ```                                                                                                                                                               ```
    ```  related post                                                                                                                                                 ```
    ```                                                                                                                                                               ```
    ```  "According to Chevel Parker(SAS), you cannot insert an image directly with ODS EXCEL. You have to post process."                                             ```
    ```                                                                                                                                                               ```
    ```  /* T008420 Adding SAS graphics at an arbitrary position within an excel sheet;                                                                               ```
    ```                                                                                                                                                               ```
    ```  HAVE                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```   1.  Workbook created using SAS, d:/xls/xlconnect_class.xlsx, with a SAS produced report                                                                     ```
    ```   2,  Have a four panel graph, d:/png/xlconnect_class.png, produced by SAS                                                                                    ```
    ```                                                                                                                                                               ```
    ```  WANT                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```   To place the graph at an arbitrary position. above, below or beside the SAS report.                                                                         ```
    ```                                                                                                                                                               ```
    ```  SOLUTION                                                                                                                                                     ```
    ```                                                                                                                                                               ```
    ```  * 1. SAS: create a excel workbook and a report;                                                                                                              ```
    ```  * 2. SAS: create a SAS png graph to be inserted into excel sheet                                                                                             ```
    ```  * 3. R: Insert graph anywhere on the sheet ie  at col G row 2                                                                                                ```
    ```                                                                                                                                                               ```
    ```  github                                                                                                                                                       ```
    ```  https://goo.gl/YrbkMe                                                                                                                                        ```
    ```  https://github.com/rogerjdeangelis/utl_adding_SAS_graphics_at_an_arbitrary_position_into_existing_excel_sheets                                               ```
    ```                                                                                                                                                               ```
    ```  see                                                                                                                                                          ```
    ```  https://goo.gl/s4KqY3                                                                                                                                        ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/ODS-EXCEL-to-have-multiple-tabs-and-embed-image-logo-in-title/m-p/409058                               ```
    ```                                                                                                                                                               ```
    ```  https://goo.gl/JBQFnA                                                                                                                                        ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/ODS-EXCEL-AND-HOW-TO-INSERT-AN-IMAGE/td-p/289493                                                       ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * create a workbook and ad report in excel  d:/xls/xlconnect_class.xlsx;                                                                                     ```
    ```  %utlfkil(d:/xls/xlconnect_class.xlsx);                                                                                                                       ```
    ```  libname xls "d:/xls/xlconnect_class.xlsx";                                                                                                                   ```
    ```  data xls.class;                                                                                                                                              ```
    ```   set sashelp.class;                                                                                                                                          ```
    ```  run;quit;                                                                                                                                                    ```
    ```  libname xls clear;                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  * create a four panel graph d:/png/xlconnect_class.png;                                                                                                      ```
    ```  ods listing  gpath="d:/png";                                                                                                                                 ```
    ```  ods graphics on  /                                                                                                                                           ```
    ```             reset=all                                                                                                                                         ```
    ```             reset=index                                                                                                                                       ```
    ```             imagefmt=png                                                                                                                                      ```
    ```             imagename="xlconnect_class"                                                                                                                       ```
    ```             height=400px                                                                                                                                      ```
    ```             width=400px                                                                                                                                       ```
    ```  ;                                                                                                                                                            ```
    ```  PROC TEMPLATE; DEFINE STATGRAPH Panel;                                                                                                                       ```
    ```  BEGINGRAPH;                                                                                                                                                  ```
    ```  ENTRYTITLE "Paneled Display ";                                                                                                                               ```
    ```     LAYOUT LATTICE / ROWS = 2 COLUMNS = 2 ROWGUTTER = 10 COLUMNGUTTER = 10;                                                                                   ```
    ```       LAYOUT OVERLAY; SCATTERPLOT Y = Weight X = Height;                                                                                                      ```
    ```         REGRESSIONPLOT Y = Weight X = Height;                                                                                                                 ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY / XAXISOPTS = (LABEL = "Weight");                                                                                                        ```
    ```         HISTOGRAM Weight;                                                                                                                                     ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY / YAXISOPTS = (LABEL = "Height");                                                                                                        ```
    ```         BOXPLOT Y = Height;                                                                                                                                   ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```       LAYOUT OVERLAY; SCATTERPLOT Y = weight X = height /                                                                                                     ```
    ```         GROUP = sex NAME = "Scat";                                                                                                                            ```
    ```         DISCRETELEGEND "Scat"                                                                                                                                 ```
    ```         / TITLE = "Sex";                                                                                                                                      ```
    ```       ENDLAYOUT;                                                                                                                                              ```
    ```     ENDLAYOUT;                                                                                                                                                ```
    ```   ENDGRAPH;                                                                                                                                                   ```
    ```  END;                                                                                                                                                         ```
    ```  RUN;                                                                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  PROC SGRENDER DATA = Sashelp.Class TEMPLATE = Panel;                                                                                                         ```
    ```  run;quit;                                                                                                                                                    ```
    ```                                                                                                                                                               ```
    ```  ods graphics off;                                                                                                                                            ```
    ```  ods html close;                                                                                                                                              ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  * create a png graphic to add to sheet mtcars below;                                                                                                         ```
    ```                                                                                                                                                               ```
    ```  %utl_submit_r(                                                                                                                                               ```
    ```     library('XLConnect');                                                                                                                                     ```
    ```     wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');                                                                                                        ```
    ```     createName(wb, name = 'class_png', formula = 'class!$G$2');                                                                                               ```
    ```     addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',originalSize = TRUE);                                                            ```
    ```     saveWorkbook(wb);                                                                                                                                         ```
    ```  );                                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```  * if you want to know the active sheet and last row and last column;                                                                                         ```
    ```  %utl_submit_r(                                                                                                                                               ```
    ```     library('XLConnect');                                                                                                                                     ```
    ```     wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');                                                                                                        ```
    ```     activeSheetIndex <- getActiveSheetIndex(wb);                                                                                                              ```
    ```     activeSheetName <- getActiveSheetName(wb);                                                                                                                ```
    ```     activeSheetName;                                                                                                                                          ```
    ```     activeSheetIndex;                                                                                                                                         ```
    ```     LastColumn<-getLastColumn(wb, 'class');                                                                                                                   ```
    ```     LastRow   <-getLastRow(wb, 'class');                                                                                                                      ```
    ```     LastColumn;                                                                                                                                               ```
    ```     LastRow;                                                                                                                                                  ```
    ```     createName(wb, name = 'class_png', formula = 'class!$G$2');                                                                                               ```
    ```     addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',originalSize = TRUE);                                                            ```
    ```     saveWorkbook(wb);                                                                                                                                         ```
    ```  );                                                                                                                                                           ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  LOG                                                                                                                                                          ```
    ```                                                                                                                                                               ```
    ```  NOTE: 10 lines were written to file PRINT.                                                                                                                   ```
    ```  Stderr output:                                                                                                                                               ```
    ```  Loading required package: XLConnectJars                                                                                                                      ```
    ```  XLConnect 0.2-12 by Mirai Solutions GmbH [aut],                                                                                                              ```
    ```    Martin Studer [cre],                                                                                                                                       ```
    ```    The Apache Software Foundation [ctb, cph] (Apache POI, Apache Commons                                                                                      ```
    ```      Codec),                                                                                                                                                  ```
    ```    Stephen Colebourne [ctb, cph] (Joda-Time Java library),                                                                                                    ```
    ```    Graph Builder [ctb, cph] (Curvesapi Java library)                                                                                                          ```
    ```  http://www.mirai-solutions.com ,                                                                                                                             ```
    ```  http://miraisolutions.wordpress.com                                                                                                                          ```
    ```  NOTE: 8 records were read from the infile RUT.                                                                                                               ```
    ```        The minimum record length was 2.                                                                                                                       ```
    ```        The maximum record length was 504.                                                                                                                     ```
    ```  NOTE: DATA statement used (Total process time):                                                                                                              ```
    ```        real time           2.88 seconds                                                                                                                       ```
    ```        cpu time            0.04 seconds                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  NOTE: Fileref RUT has been deassigned.                                                                                                                       ```
    ```  NOTE: Fileref R_PGM has been deassigned.                                                                                                                     ```
    ```                                                                                                                                                               ```
    ```                                                                                                                                                               ```
    ```  OUTPUT                                                                                                                                                       ```
    ```                                                                                                                                                               ```
    ```  > library('XLConnect');    wb <- loadWorkbook('d:/xls/xlconnect_class.xlsx');    activeSheetIndex <- getActiveSheetIndex(wb);                                ```
    ```      activeSheetName <- getActiveSheetName(wb);                                                                                                               ```
    ```      activeSheetName;    activeSheetIndex;    LastColumn<-getLastColumn(wb, 'class');    LastRow   <-getLastRow(wb, 'class');                                 ```
    ```  LastColumn;    LastRow;    createName(wb, n                                                                                                                  ```
    ```  ame = 'class_png', formula = 'class!$G$2');    addImage(wb, filename = 'd:/png/xlconnect_class.png', name = 'class_png',                                     ```
    ```  originalSize = TRUE);    saveWorkbook(wb);                                                                                                                   ```
    ```  [1] "class"                                                                                                                                                  ```
    ```  [1] 1                                                                                                                                                        ```
    ```  class columns                                                                                                                                                ```
    ```      5                                                                                                                                                        ```
    ```  class rows                                                                                                                                                   ```
    ```     20                                                                                                                                                        ```
    ```                                                                                                                                                               ```

