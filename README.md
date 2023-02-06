# Excel auto processing for qRT-PCR
## Table of contents
 * [Module]<br>
    * Loading workbook, sheetname and load data<br>
       * openxyl, os
    * Calculation<br>
    * Output<br><br>
    
 * [Usage]
    1. put excel and script under same folder, and don't name the processing excel as "output.xlsx", otherwise will be skipped.
    2. please name the header as "Target" "Sample" "Cq" in sheet，letter case sensitive. 
    3. in the case of using mac terminal:
    ```python
    cd path
    python3 ExcelAutoProcessing.py
    ```
    4. This script can help to deal with one replication data or three replication data of qRT-PCR raw data based on your choice. 
   
