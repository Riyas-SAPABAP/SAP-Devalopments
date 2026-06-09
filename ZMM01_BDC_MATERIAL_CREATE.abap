*&---------------------------------------------------------------------*
*& Report  ZMM01_BDC_MATERIAL_CREATE
*&---------------------------------------------------------------------*
*& Purpose    : BDC program to create Material Master records (MM01)
*&              from a TAB-delimited input file (input.txt), based on
*&              the SHDB recording of transaction MM01.
*&
*& Input file : TAB-delimited text file with 3 header lines
*&                Line 1 : Field keys
*&                Line 2 : BDC screen field names
*&                Line 3 : Field descriptions
*&              Actual data starts from line 4.
*&
*& Views      : Basic Data 1, Basic Data 2,
*&              Sales: Sales Org. Data 1, Sales: Sales Org. Data 2,
*&              Sales: General/Plant Data, Foreign Trade: Export Data,
*&              Purchasing, MRP 1, MRP 2, MRP 3,
*&              Plant Data / Storage 1, Quality Management,
*&              Accounting 1, Costing 1
*&
*& Notes      : - Material number is left blank -> internal number
*&                assignment as per the recording.
*&              - Amount fields (e.g. 250.00) must match the decimal
*&                notation of the batch-input user (SU3 defaults).
*&              - The view selection indices MSICHTAUSW-KZSEL(nn) and
*&                screen/OK-code sequence correspond to the BDC
*&                recording; verify them against your system if the
*&                view list for material type ROH differs.
*&---------------------------------------------------------------------*
REPORT zmm01_bdc_material_create NO STANDARD PAGE HEADING
                                 LINE-SIZE 255.

*----------------------------------------------------------------------*
* Types - one field per column of input.txt (in file column order)
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_input,
         mbrsh    TYPE c LENGTH 1,   "RMMG1-MBRSH  Industry Sector
         mtart    TYPE c LENGTH 4,   "RMMG1-MTART  Material Type
         werks    TYPE c LENGTH 4,   "RMMG1-WERKS  Plant
         lgort    TYPE c LENGTH 4,   "RMMG1-LGORT  Storage Location
         vkorg    TYPE c LENGTH 4,   "RMMG1-VKORG  Sales Organization
         vtweg    TYPE c LENGTH 2,   "RMMG1-VTWEG  Distribution Channel
         maktx    TYPE c LENGTH 40,  "MAKT-MAKTX   Material Description
         meins    TYPE c LENGTH 3,   "MARA-MEINS   Base Unit of Measure
         matkl    TYPE c LENGTH 9,   "MARA-MATKL   Material Group
         spart    TYPE c LENGTH 2,   "MARA-SPART   Division
         brgew    TYPE c LENGTH 17,  "MARA-BRGEW   Gross Weight
         gewei    TYPE c LENGTH 3,   "MARA-GEWEI   Weight Unit
         ntgew    TYPE c LENGTH 17,  "MARA-NTGEW   Net Weight
         wrkst    TYPE c LENGTH 48,  "MARA-WRKST   Basic Material
         sktof    TYPE c LENGTH 1,   "MVKE-SKTOF   Cash Discount Ind.
         taxkm1   TYPE c LENGTH 1,   "MG03STEUER-TAXKM(01) Tax Class. 1
         taxkm2   TYPE c LENGTH 1,   "MG03STEUER-TAXKM(02) Tax Class. 2
         versg    TYPE c LENGTH 1,   "MVKE-VERSG   Matl Statistics Grp
         kondm    TYPE c LENGTH 2,   "MVKE-KONDM   Matl Pricing Group
         tragr    TYPE c LENGTH 4,   "MARA-TRAGR   Transportation Group
         ladgr    TYPE c LENGTH 4,   "MARC-LADGR   Loading Group
         prctr    TYPE c LENGTH 10,  "MARC-PRCTR   Profit Center
         mtvfp    TYPE c LENGTH 2,   "MARC-MTVFP   Availability Check
         xchpf    TYPE c LENGTH 1,   "MARC-XCHPF   Batch Management Ind.
         steuc    TYPE c LENGTH 16,  "MARC-STEUC   Control Code / HSN
         taxim    TYPE c LENGTH 1,   "MG03STEUMM-TAXIM Tax Ind. Material
         dismm    TYPE c LENGTH 2,   "MARC-DISMM   MRP Type
         beskz    TYPE c LENGTH 1,   "MARC-BESKZ   Procurement Type
         perkz    TYPE c LENGTH 1,   "MARC-PERKZ   Period Indicator
         iprkz    TYPE c LENGTH 1,   "MARA-IPRKZ   SLED Period Indicator
         sled_bbd TYPE c LENGTH 1,   "MARA-SLED_BBD Shelf Life / BBD
         qmpur    TYPE c LENGTH 1,   "MARA-QMPUR   QM Procurement Active
         ssqss    TYPE c LENGTH 8,   "MARC-SSQSS   QM Control Key
         bklas    TYPE c LENGTH 4,   "MBEW-BKLAS   Valuation Class
         vprsv    TYPE c LENGTH 1,   "MBEW-VPRSV   Price Control Ind.
         stprs    TYPE c LENGTH 15,  "MBEW-STPRS   Standard Price
         peinh    TYPE c LENGTH 6,   "MBEW-PEINH   Price Unit
         ekalr    TYPE c LENGTH 1,   "MBEW-EKALR   With Qty Structure
         hkmat    TYPE c LENGTH 1,   "MBEW-HKMAT   Material Origin
         sobsk    TYPE c LENGTH 2,   "MARC-SOBSK   Spec. Proc. Costing
         losgr    TYPE c LENGTH 17,  "MARC-LOSGR   Costing Lot Size
         mlast    TYPE c LENGTH 1,   "CKMLHD-MLAST Price Determination
         stprs1   TYPE c LENGTH 15,  "CKMAT_DISPLAY-STPRS_1 Std Price 1
         stprs2   TYPE c LENGTH 15,  "CKMAT_DISPLAY-STPRS_2 Std Price 2
         stprs3   TYPE c LENGTH 15,  "CKMAT_DISPLAY-STPRS_3 Std Price 3
         peinh1   TYPE c LENGTH 6,   "CKMAT_DISPLAY-PEINH_1 Price Unit 1
         peinh2   TYPE c LENGTH 6,   "CKMAT_DISPLAY-PEINH_2 Price Unit 2
         peinh3   TYPE c LENGTH 6,   "CKMAT_DISPLAY-PEINH_3 Price Unit 3
         verpr    TYPE c LENGTH 15,  "MBEW-VERPR   Moving Average Price
         insmk    TYPE c LENGTH 1,   "MARC-INSMK   Inspection Stock Ind.
         kzdkz    TYPE c LENGTH 1,   "MARC-KZDKZ   Documentation/Insp.
         ncost    TYPE c LENGTH 1,   "MARC-NCOST   Do Not Cost Material
         qpls_arg TYPE c LENGTH 10,  "RMQAM-ARGUMENT Insp. Plan Usage
         qpls_art TYPE c LENGTH 8,   "RMQAM-ART(01)  Inspection Type
       END OF ty_input.

*----------------------------------------------------------------------*
* Global data
*----------------------------------------------------------------------*
DATA: gt_input   TYPE STANDARD TABLE OF ty_input,
      gs_input   TYPE ty_input,
      gt_bdcdata TYPE STANDARD TABLE OF bdcdata,
      gs_bdcdata TYPE bdcdata,
      gt_msg     TYPE STANDARD TABLE OF bdcmsgcoll,
      gs_msg     TYPE bdcmsgcoll,
      gv_msgtext TYPE string,
      gv_error   TYPE abap_bool,
      gv_success TYPE i,
      gv_failed  TYPE i,
      gv_recno   TYPE i.

*----------------------------------------------------------------------*
* Selection screen
*----------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
PARAMETERS: p_file  TYPE rlgrap-filename
                    DEFAULT 'C:\input.txt' OBLIGATORY,
            p_mode  TYPE ctu_mode   DEFAULT 'N',  "A/E/N
            p_updat TYPE ctu_update DEFAULT 'S'.  "S/A/L
SELECTION-SCREEN END OF BLOCK b1.

*----------------------------------------------------------------------*
* F4 help for file name
*----------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  CALL FUNCTION 'F4_FILENAME'
    EXPORTING
      program_name  = sy-repid
      dynpro_number = sy-dynnr
      field_name    = 'P_FILE'
    IMPORTING
      file_name     = p_file.

*----------------------------------------------------------------------*
* Main processing
*----------------------------------------------------------------------*
START-OF-SELECTION.
  PERFORM f_upload_file.
  PERFORM f_process_data.

END-OF-SELECTION.
  PERFORM f_display_summary.

*&---------------------------------------------------------------------*
*& Form F_UPLOAD_FILE
*&---------------------------------------------------------------------*
*& Upload TAB-delimited file and remove the 3 header lines
*&---------------------------------------------------------------------*
FORM f_upload_file.

  DATA lv_file TYPE string.

  lv_file = p_file.

  CALL FUNCTION 'GUI_UPLOAD'
    EXPORTING
      filename            = lv_file
      filetype            = 'ASC'
      has_field_separator = 'X'
    TABLES
      data_tab            = gt_input
    EXCEPTIONS
      file_open_error     = 1
      file_read_error     = 2
      no_batch            = 3
      invalid_type        = 4
      unknown_error       = 5
      OTHERS              = 6.

  IF sy-subrc <> 0.
    MESSAGE e000(8i) WITH 'Error uploading file' p_file.
  ENDIF.

* Lines 1-3 are header lines (keys, BDC field names, descriptions)
  DELETE gt_input FROM 1 TO 3.

* Remove completely empty lines (e.g. trailing line in the file)
  DELETE gt_input WHERE table_line IS INITIAL.

  IF gt_input[] IS INITIAL.
    MESSAGE e000(8i) WITH 'No data records found in file' p_file.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form F_PROCESS_DATA
*&---------------------------------------------------------------------*
*& Build the BDC data per record and call transaction MM01
*&---------------------------------------------------------------------*
FORM f_process_data.

  LOOP AT gt_input INTO gs_input.
    gv_recno = sy-tabix.

    REFRESH: gt_bdcdata, gt_msg.
    PERFORM f_build_bdc.
    PERFORM f_call_transaction.

  ENDLOOP.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form F_BUILD_BDC
*&---------------------------------------------------------------------*
*& BDC flow as per the MM01 recording
*&---------------------------------------------------------------------*
FORM f_build_bdc.

*----------------------------------------------------------------------*
* Initial screen - industry sector / material type
* (Material number blank -> internal number assignment)
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '0060'.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'RMMG1-MTART'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '=AUSW'.
  PERFORM bdc_field  USING 'RMMG1-MBRSH' gs_input-mbrsh.
  PERFORM bdc_field  USING 'RMMG1-MTART' gs_input-mtart.

*----------------------------------------------------------------------*
* View selection
* NOTE: KZSEL indices must match the position of the views in the
*       selection list of your system (as in the BDC recording).
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '0070'.
  PERFORM bdc_field  USING 'BDC_OKCODE'           '=ENTR'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' 'X'. "Basic Data 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' 'X'. "Basic Data 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(03)' 'X'. "Sales Org. 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' 'X'. "Sales Org. 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' 'X'. "Sales Gen./Plant
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' 'X'. "For.Trade Export
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(07)' 'X'. "Purchasing
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' 'X'. "MRP 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(09)' 'X'. "MRP 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(10)' 'X'. "MRP 3
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(11)' 'X'. "Plant/Storage 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(12)' 'X'. "Quality Mgmt
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(13)' 'X'. "Accounting 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(14)' 'X'. "Costing 1

*----------------------------------------------------------------------*
* Organizational levels
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '0080'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '=ENTR'.
  PERFORM bdc_field  USING 'RMMG1-WERKS' gs_input-werks.
  PERFORM bdc_field  USING 'RMMG1-LGORT' gs_input-lgort.
  PERFORM bdc_field  USING 'RMMG1-VKORG' gs_input-vkorg.
  PERFORM bdc_field  USING 'RMMG1-VTWEG' gs_input-vtweg.

*----------------------------------------------------------------------*
* Basic Data 1
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MAKT-MAKTX'  gs_input-maktx.
  PERFORM bdc_field  USING 'MARA-MEINS'  gs_input-meins.
  PERFORM bdc_field  USING 'MARA-MATKL'  gs_input-matkl.
  PERFORM bdc_field  USING 'MARA-SPART'  gs_input-spart.
  PERFORM bdc_field  USING 'MARA-BRGEW'  gs_input-brgew.
  PERFORM bdc_field  USING 'MARA-GEWEI'  gs_input-gewei.
  PERFORM bdc_field  USING 'MARA-NTGEW'  gs_input-ntgew.

*----------------------------------------------------------------------*
* Basic Data 2
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARA-WRKST'  gs_input-wrkst.

*----------------------------------------------------------------------*
* Sales: Sales Org. Data 1
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'            '/00'.
  PERFORM bdc_field  USING 'MVKE-SKTOF'            gs_input-sktof.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(01)'  gs_input-taxkm1.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(02)'  gs_input-taxkm2.

*----------------------------------------------------------------------*
* Sales: Sales Org. Data 2
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MVKE-VERSG'  gs_input-versg.
  PERFORM bdc_field  USING 'MVKE-KONDM'  gs_input-kondm.

*----------------------------------------------------------------------*
* Sales: General/Plant Data
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARA-TRAGR'  gs_input-tragr.
  PERFORM bdc_field  USING 'MARC-LADGR'  gs_input-ladgr.
  PERFORM bdc_field  USING 'MARC-PRCTR'  gs_input-prctr.
  PERFORM bdc_field  USING 'MARC-MTVFP'  gs_input-mtvfp.
  PERFORM bdc_field  USING 'MARC-XCHPF'  gs_input-xchpf.

*----------------------------------------------------------------------*
* Foreign Trade: Export Data (Control code / HSN)
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARC-STEUC'  gs_input-steuc.

*----------------------------------------------------------------------*
* Purchasing (Tax indicator for material)
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'        '/00'.
  PERFORM bdc_field  USING 'MG03STEUMM-TAXIM'  gs_input-taxim.

*----------------------------------------------------------------------*
* MRP 1
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARC-DISMM'  gs_input-dismm.

*----------------------------------------------------------------------*
* MRP 2
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARC-BESKZ'  gs_input-beskz.

*----------------------------------------------------------------------*
* MRP 3
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.
  PERFORM bdc_field  USING 'MARC-PERKZ'  gs_input-perkz.

*----------------------------------------------------------------------*
* Plant Data / Storage 1 (Shelf life data)
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'     '/00'.
  PERFORM bdc_field  USING 'MARA-IPRKZ'     gs_input-iprkz.
  PERFORM bdc_field  USING 'MARA-SLED_BBD'  gs_input-sled_bbd.

*----------------------------------------------------------------------*
* Quality Management
* '=PB29' = "Insp. setup" pushbutton -> opens inspection setup popup
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '=PB29'.
  PERFORM bdc_field  USING 'MARA-QMPUR'  gs_input-qmpur.
  PERFORM bdc_field  USING 'MARC-SSQSS'  gs_input-ssqss.
  PERFORM bdc_field  USING 'MARC-INSMK'  gs_input-insmk.
  PERFORM bdc_field  USING 'MARC-KZDKZ'  gs_input-kzdkz.

* Inspection setup popup (program SAPLQPLS)
  PERFORM bdc_dynpro USING 'SAPLQPLS' '0100'.
  PERFORM bdc_field  USING 'BDC_CURSOR'      'RMQAM-ART(01)'.
  PERFORM bdc_field  USING 'BDC_OKCODE'      '=ENTR'.
  PERFORM bdc_field  USING 'RMQAM-ARGUMENT'  gs_input-qpls_arg.
  PERFORM bdc_field  USING 'RMQAM-ART(01)'   gs_input-qpls_art.

* Back on QM view -> continue to next view
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '/00'.

*----------------------------------------------------------------------*
* Accounting 1 (incl. material ledger / prices in currencies)
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'             '/00'.
  PERFORM bdc_field  USING 'MBEW-BKLAS'             gs_input-bklas.
  PERFORM bdc_field  USING 'MBEW-VPRSV'             gs_input-vprsv.
  PERFORM bdc_field  USING 'MBEW-STPRS'             gs_input-stprs.
  PERFORM bdc_field  USING 'MBEW-PEINH'             gs_input-peinh.
  PERFORM bdc_field  USING 'MBEW-VERPR'             gs_input-verpr.
  PERFORM bdc_field  USING 'CKMLHD-MLAST'           gs_input-mlast.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-STPRS_1'  gs_input-stprs1.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-STPRS_2'  gs_input-stprs2.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-STPRS_3'  gs_input-stprs3.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-PEINH_1'  gs_input-peinh1.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-PEINH_2'  gs_input-peinh2.
  PERFORM bdc_field  USING 'CKMAT_DISPLAY-PEINH_3'  gs_input-peinh3.

*----------------------------------------------------------------------*
* Costing 1 - last view -> save with '=BU'
*----------------------------------------------------------------------*
  PERFORM bdc_dynpro USING 'SAPLMGMM' '4004'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  '=BU'.
  PERFORM bdc_field  USING 'MBEW-EKALR'  gs_input-ekalr.
  PERFORM bdc_field  USING 'MBEW-HKMAT'  gs_input-hkmat.
  PERFORM bdc_field  USING 'MARC-SOBSK'  gs_input-sobsk.
  PERFORM bdc_field  USING 'MARC-LOSGR'  gs_input-losgr.
  PERFORM bdc_field  USING 'MARC-NCOST'  gs_input-ncost.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form F_CALL_TRANSACTION
*&---------------------------------------------------------------------*
FORM f_call_transaction.

  DATA lv_matnr TYPE matnr.

  CALL TRANSACTION 'MM01' USING  gt_bdcdata
                          MODE   p_mode
                          UPDATE p_updat
                          MESSAGES INTO gt_msg.

  gv_error = abap_false.
  CLEAR lv_matnr.

  LOOP AT gt_msg INTO gs_msg.
    IF gs_msg-msgtyp CA 'EA'.
      gv_error = abap_true.
    ENDIF.
*   M3 800: "Material & created"
    IF gs_msg-msgid = 'M3' AND gs_msg-msgnr = '800'.
      lv_matnr = gs_msg-msgv1.
    ENDIF.
  ENDLOOP.

  IF gv_error = abap_false.
    gv_success = gv_success + 1.
    WRITE: / 'Record', gv_recno, ':',
             'Material created successfully -', lv_matnr,
             '(', gs_input-maktx, ')'.
  ELSE.
    gv_failed = gv_failed + 1.
    WRITE: / 'Record', gv_recno, ':',
             'Error while creating material (', gs_input-maktx, ')'.
    LOOP AT gt_msg INTO gs_msg WHERE msgtyp CA 'EA'.
      CLEAR gv_msgtext.
      CALL FUNCTION 'FORMAT_MESSAGE'
        EXPORTING
          id        = gs_msg-msgid
          lang      = sy-langu
          no        = gs_msg-msgnr
          v1        = gs_msg-msgv1
          v2        = gs_msg-msgv2
          v3        = gs_msg-msgv3
          v4        = gs_msg-msgv4
        IMPORTING
          msg       = gv_msgtext
        EXCEPTIONS
          not_found = 1
          OTHERS    = 2.
      IF sy-subrc = 0.
        WRITE: /5 gv_msgtext.
      ENDIF.
    ENDLOOP.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form F_DISPLAY_SUMMARY
*&---------------------------------------------------------------------*
FORM f_display_summary.

  ULINE.
  WRITE: / 'Total records processed :', gv_recno.
  WRITE: / 'Successfully created    :', gv_success.
  WRITE: / 'Failed                  :', gv_failed.
  ULINE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BDC_DYNPRO
*&---------------------------------------------------------------------*
*& Start new screen in BDC table
*&---------------------------------------------------------------------*
FORM bdc_dynpro USING p_program TYPE clike
                      p_dynpro  TYPE clike.

  CLEAR gs_bdcdata.
  gs_bdcdata-program  = p_program.
  gs_bdcdata-dynpro   = p_dynpro.
  gs_bdcdata-dynbegin = 'X'.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BDC_FIELD
*&---------------------------------------------------------------------*
*& Insert field into BDC table.
*& Empty data fields are skipped (NODATA handling) so that screen
*& defaults are not overwritten; BDC_* control fields are always sent.
*&---------------------------------------------------------------------*
FORM bdc_field USING p_fnam TYPE clike
                     p_fval TYPE clike.

  DATA lv_fnam TYPE bdcdata-fnam.

  lv_fnam = p_fnam.

  IF p_fval IS INITIAL AND lv_fnam(4) <> 'BDC_'.
    RETURN.
  ENDIF.

  CLEAR gs_bdcdata.
  gs_bdcdata-fnam = lv_fnam.
  gs_bdcdata-fval = p_fval.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.
