REPORT zmm_mm01_create_bdc.

TYPE-POOLS: slis.

*&---------------------------------------------------------------------*
*& Type Declarations
*&---------------------------------------------------------------------*
TYPES: ty_msgtext TYPE c LENGTH 255,
       ty_char40  TYPE c LENGTH 40.

*& Excel column mapping (54 columns, matching MM01_ROH_BDC_Input_Template):
*& Col  1 MBRSH    Col 14 WRKST    Col 27 DISMM    Col 40 SOBSK
*& Col  2 MTART    Col 15 SKTOF    Col 28 BESKZ    Col 41 LOSGR
*& Col  3 WERKS    Col 16 TAXKM1   Col 29 PERKZ    Col 42 MLAST
*& Col  4 LGORT    Col 17 TAXKM2   Col 30 IPRKZ    Col 43 STPRS1
*& Col  5 VKORG    Col 18 VERSG    Col 31 SLED_BBD Col 44 STPRS2
*& Col  6 VTWEG    Col 19 KONDM    Col 32 QMPUR    Col 45 STPRS3
*& Col  7 MAKTX    Col 20 TRAGR    Col 33 SSQSS    Col 46 PEINH1
*& Col  8 MEINS    Col 21 LADGR    Col 34 BKLAS    Col 47 PEINH2
*& Col  9 MATKL    Col 22 PRCTR    Col 35 VPRSV    Col 48 PEINH3
*& Col 10 SPART    Col 23 MTVFP    Col 36 STPRS    Col 49 VERPR
*& Col 11 BRGEW    Col 24 XCHPF    Col 37 PEINH    Col 50 INSMK
*& Col 12 GEWEI    Col 25 STEUC    Col 38 EKALR    Col 51 KZDKZ
*& Col 13 NTGEW    Col 26 TAXIM    Col 39 HKMAT    Col 52 NCOST
*&                                                  Col 53 QPLS_ARG
*&                                                  Col 54 QPLS_ART

TYPES: BEGIN OF ty_input,
         mbrsh    TYPE c LENGTH 1,   "Col 01 - Industry Sector
         mtart    TYPE c LENGTH 4,   "Col 02 - Material Type
         werks    TYPE c LENGTH 4,   "Col 03 - Plant
         lgort    TYPE c LENGTH 4,   "Col 04 - Storage Location
         vkorg    TYPE c LENGTH 4,   "Col 05 - Sales Organisation
         vtweg    TYPE c LENGTH 2,   "Col 06 - Distribution Channel
         maktx    TYPE c LENGTH 40,  "Col 07 - Material Description
         meins    TYPE c LENGTH 3,   "Col 08 - Base Unit of Measure
         matkl    TYPE c LENGTH 9,   "Col 09 - Material Group
         spart    TYPE c LENGTH 2,   "Col 10 - Division
         brgew    TYPE p DECIMALS 3, "Col 11 - Gross Weight
         gewei    TYPE c LENGTH 3,   "Col 12 - Weight Unit
         ntgew    TYPE p DECIMALS 3, "Col 13 - Net Weight
         wrkst    TYPE c LENGTH 48,  "Col 14 - Basic Material / Raw Material
         sktof    TYPE c LENGTH 1,   "Col 15 - Cash Discount Indicator
         taxkm1   TYPE c LENGTH 1,   "Col 16 - Tax Classification 1
         taxkm2   TYPE c LENGTH 1,   "Col 17 - Tax Classification 2
         versg    TYPE c LENGTH 1,   "Col 18 - Material Statistics Group
         kondm    TYPE c LENGTH 2,   "Col 19 - Material Pricing Group
         tragr    TYPE c LENGTH 4,   "Col 20 - Transportation Group
         ladgr    TYPE c LENGTH 4,   "Col 21 - Loading Group
         prctr    TYPE c LENGTH 10,  "Col 22 - Profit Centre
         mtvfp    TYPE c LENGTH 2,   "Col 23 - Availability Check
         xchpf    TYPE c LENGTH 1,   "Col 24 - Batch Management Indicator
         steuc    TYPE c LENGTH 16,  "Col 25 - Control Code / HSN
         taxim    TYPE c LENGTH 1,   "Col 26 - Tax Indicator Material
         dismm    TYPE c LENGTH 2,   "Col 27 - MRP Type
         beskz    TYPE c LENGTH 1,   "Col 28 - Procurement Type
         perkz    TYPE c LENGTH 1,   "Col 29 - Period Indicator
         iprkz    TYPE c LENGTH 1,   "Col 30 - SLED Period Indicator
         sled_bbd TYPE c LENGTH 1,   "Col 31 - Shelf Life / BBD
         qmpur    TYPE c LENGTH 1,   "Col 32 - QM Procurement Active
         ssqss    TYPE c LENGTH 8,   "Col 33 - QM Control Key
         bklas    TYPE c LENGTH 4,   "Col 34 - Valuation Class
         vprsv    TYPE c LENGTH 1,   "Col 35 - Price Control
         stprs    TYPE p DECIMALS 2, "Col 36 - Standard Price
         peinh    TYPE p DECIMALS 0, "Col 37 - Price Unit
         ekalr    TYPE c LENGTH 1,   "Col 38 - With Quantity Structure
         hkmat    TYPE c LENGTH 1,   "Col 39 - Material Origin
         sobsk    TYPE c LENGTH 2,   "Col 40 - Special Procurement Type
         losgr    TYPE p DECIMALS 0, "Col 41 - Costing Lot Size
         mlast    TYPE c LENGTH 1,   "Col 42 - Price Determination
         stprs1   TYPE p DECIMALS 2, "Col 43 - Std Price Period 1
         stprs2   TYPE p DECIMALS 2, "Col 44 - Std Price Period 2
         stprs3   TYPE p DECIMALS 2, "Col 45 - Std Price Period 3
         peinh1   TYPE p DECIMALS 0, "Col 46 - Price Unit Period 1
         peinh2   TYPE p DECIMALS 0, "Col 47 - Price Unit Period 2
         peinh3   TYPE p DECIMALS 0, "Col 48 - Price Unit Period 3
         verpr    TYPE p DECIMALS 2, "Col 49 - Moving Average Price
         insmk    TYPE c LENGTH 1,   "Col 50 - Inspection Stock Indicator (MARC-INSMK)
         kzdkz    TYPE c LENGTH 1,   "Col 51 - Post to Inspection Stock  (MARC-KZDKZ)
         ncost    TYPE c LENGTH 1,   "Col 52 - Do Not Cost Material      (MARC-NCOST)
         qpls_arg TYPE c LENGTH 10,  "Col 53 - Insp Plan Usage (QPLS popup)
         qpls_art TYPE c LENGTH 4,   "Col 54 - Insp Plan Type  (QPLS popup)
       END OF ty_input.

TYPES: BEGIN OF ty_input_raw,
         mbrsh    TYPE string,
         mtart    TYPE string,
         werks    TYPE string,
         lgort    TYPE string,
         vkorg    TYPE string,
         vtweg    TYPE string,
         maktx    TYPE string,
         meins    TYPE string,
         matkl    TYPE string,
         spart    TYPE string,
         brgew    TYPE string,
         gewei    TYPE string,
         ntgew    TYPE string,
         wrkst    TYPE string,
         sktof    TYPE string,
         taxkm1   TYPE string,
         taxkm2   TYPE string,
         versg    TYPE string,
         kondm    TYPE string,
         tragr    TYPE string,
         ladgr    TYPE string,
         prctr    TYPE string,
         mtvfp    TYPE string,
         xchpf    TYPE string,
         steuc    TYPE string,
         taxim    TYPE string,
         dismm    TYPE string,
         beskz    TYPE string,
         perkz    TYPE string,
         iprkz    TYPE string,
         sled_bbd TYPE string,
         qmpur    TYPE string,
         ssqss    TYPE string,
         bklas    TYPE string,
         vprsv    TYPE string,
         stprs    TYPE string,
         peinh    TYPE string,
         ekalr    TYPE string,
         hkmat    TYPE string,
         sobsk    TYPE string,
         losgr    TYPE string,
         mlast    TYPE string,
         stprs1   TYPE string,
         stprs2   TYPE string,
         stprs3   TYPE string,
         peinh1   TYPE string,
         peinh2   TYPE string,
         peinh3   TYPE string,
         verpr    TYPE string,
         insmk    TYPE string,
         kzdkz    TYPE string,
         ncost    TYPE string,
         qpls_arg TYPE string,
         qpls_art TYPE string,
       END OF ty_input_raw.

TYPES: BEGIN OF ty_record,
         rowno TYPE i,
         input TYPE ty_input,
       END OF ty_record.

TYPES: BEGIN OF ty_error,
         rowno    TYPE i,
         material TYPE matnr,
         msgtyp   TYPE c LENGTH 1,
         msgid    LIKE bdcmsgcoll-msgid,
         msgno    LIKE bdcmsgcoll-msgnr,
         msgtext  TYPE ty_msgtext,
       END OF ty_error.

*&---------------------------------------------------------------------*
*& Constants
*&---------------------------------------------------------------------*
CONSTANTS:
  c_tcode_mm01    TYPE tcode VALUE 'MM01',
  c_expected_cols TYPE i     VALUE 54,     "54 columns in template
  c_update_sync   TYPE c     VALUE 'S',
  c_x             TYPE c     VALUE 'X',

  c_ok_enter      TYPE bdcdata-fval VALUE '/00',    "Data screen Enter
  c_ok_entr       TYPE bdcdata-fval VALUE '=ENTR',  "Initial/Org screen Enter
  c_ok_p_plus     TYPE bdcdata-fval VALUE '=P+',    "Page Down in view list
  c_ok_schl       TYPE bdcdata-fval VALUE '=SCHL',  "Close view selection
  c_ok_yes        TYPE bdcdata-fval VALUE '=YES',   "Confirm save popup

  c_prog_mm       LIKE bdcdata-program VALUE 'SAPLMGMM',
  c_scr_0060      LIKE bdcdata-dynpro  VALUE '0060',
  c_scr_0070      LIKE bdcdata-dynpro  VALUE '0070',
  c_scr_0080      LIKE bdcdata-dynpro  VALUE '0080',
  c_scr_4000      LIKE bdcdata-dynpro  VALUE '4000',
  c_scr_4004      LIKE bdcdata-dynpro  VALUE '4004',

  c_prog_spo1     LIKE bdcdata-program VALUE 'SAPLSPO1',
  c_scr_spo1      LIKE bdcdata-dynpro  VALUE '0300'.

*&---------------------------------------------------------------------*
*& Selection Screen
*&---------------------------------------------------------------------*
PARAMETERS:
  p_file  TYPE rlgrap-filename OBLIGATORY,
  p_mode  TYPE c DEFAULT 'A',
  p_matnr TYPE matnr.         "Leave blank for internal number range

*&---------------------------------------------------------------------*
*& Global Data
*&---------------------------------------------------------------------*
DATA:
  wa_input        TYPE ty_input,
  gs_raw          TYPE ty_input_raw,
  gs_record       TYPE ty_record,
  gt_records      TYPE STANDARD TABLE OF ty_record,

  gt_binary       TYPE solix_tab,
  gv_filelength   TYPE i,
  gv_xstring      TYPE xstring,
  gv_filename     TYPE string,

  go_excel        TYPE REF TO cl_fdt_xl_spreadsheet,
  go_excel_error  TYPE REF TO cx_fdt_excel_core,
  gt_worksheets   TYPE STANDARD TABLE OF string,
  gv_worksheet    TYPE string,
  gr_excel_data   TYPE REF TO data,

  gv_rowno        TYPE i,
  gv_excel_offset TYPE i,
  gv_index        TYPE i,
  gv_len          TYPE i,
  gv_last_pos     TYPE i,
  gv_total        TYPE i,
  gv_success      TYPE i,
  gv_failed       TYPE i,
  gv_call_subrc   TYPE sy-subrc,

  gv_ok           TYPE abap_bool,
  gv_skip         TYPE abap_bool,
  gv_row_failed   TYPE abap_bool,

  gv_token        TYPE string,
  gv_first_cell   TYPE string,
  gv_second_cell  TYPE string,

  gv_text         TYPE ty_msgtext,
  gv_msgtext      TYPE ty_msgtext,
  gv_numtext      TYPE c LENGTH 50,
  gv_subrc_c      TYPE c LENGTH 10,
  gv_material     TYPE matnr,

  gt_bdcdata      TYPE STANDARD TABLE OF bdcdata,
  gs_bdcdata      TYPE bdcdata,
  gt_msgcoll      TYPE STANDARD TABLE OF bdcmsgcoll,
  gs_msgcoll      TYPE bdcmsgcoll,
  gs_ctu_params   TYPE ctu_params,

  gt_errors       TYPE STANDARD TABLE OF ty_error,
  gs_error        TYPE ty_error,

  gt_fieldcat     TYPE slis_t_fieldcat_alv,
  gs_fieldcat     TYPE slis_fieldcat_alv,

  gt_filetab      TYPE filetable,
  gs_filetab      TYPE file_table,
  gv_rc           TYPE i.

FIELD-SYMBOLS:
  <gt_excel_dyn> TYPE STANDARD TABLE,
  <gs_excel_dyn> TYPE any,
  <fs_cell>      TYPE any,
  <fs_raw>       TYPE any.

*&---------------------------------------------------------------------*
*& F4 Help
*&---------------------------------------------------------------------*
AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_file.

*&---------------------------------------------------------------------*
*& Selection Validation
*&---------------------------------------------------------------------*
AT SELECTION-SCREEN.
  PERFORM validate_selection.

*&---------------------------------------------------------------------*
*& Main
*&---------------------------------------------------------------------*
START-OF-SELECTION.

  WRITE: / 'MM01 Material Master Creation BDC'.
  SKIP.

  PERFORM upload_excel_xlsx.

  DESCRIBE TABLE gt_records LINES gv_total.

  IF gt_records IS INITIAL.
    WRITE: / 'No valid input records found.'.
    PERFORM display_errors.
    RETURN.
  ENDIF.

  LOOP AT gt_records INTO gs_record.

    REFRESH: gt_bdcdata, gt_msgcoll.
    CLEAR:   gv_material, gv_call_subrc, gv_row_failed, wa_input.

    wa_input = gs_record-input.

    PERFORM build_bdc USING wa_input.

    CLEAR gs_ctu_params.
    gs_ctu_params-dismode  = p_mode.
    gs_ctu_params-updmode  = c_update_sync.
    gs_ctu_params-defsize  = c_x.
    gs_ctu_params-racommit = c_x.

    CALL TRANSACTION c_tcode_mm01
      USING gt_bdcdata
      OPTIONS FROM gs_ctu_params
      MESSAGES INTO gt_msgcoll.

    gv_call_subrc = sy-subrc.

    PERFORM analyze_bdc_messages USING gs_record-rowno.

  ENDLOOP.

  PERFORM display_summary.
  PERFORM display_errors.

*&---------------------------------------------------------------------*
*& Form F4_FILE
*&---------------------------------------------------------------------*
FORM f4_file.

  REFRESH gt_filetab.
  CLEAR: gs_filetab, gv_rc.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      window_title = 'Select MM01 ROH Excel File'
      file_filter  = 'Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
    CHANGING
      file_table   = gt_filetab
      rc           = gv_rc
    EXCEPTIONS
      OTHERS       = 1.

  IF sy-subrc = 0 AND gv_rc > 0.
    READ TABLE gt_filetab INTO gs_filetab INDEX 1.
    IF sy-subrc = 0.
      p_file = gs_filetab-filename.
    ENDIF.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form VALIDATE_SELECTION
*&---------------------------------------------------------------------*
FORM validate_selection.

  TRANSLATE p_mode TO UPPER CASE.

  IF p_mode <> 'A' AND p_mode <> 'E' AND p_mode <> 'N'.
    MESSAGE 'P_MODE must be A, E, or N.' TYPE 'E'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form UPLOAD_EXCEL_XLSX
*&---------------------------------------------------------------------*
FORM upload_excel_xlsx.

  REFRESH: gt_records, gt_binary, gt_worksheets.
  CLEAR:   gv_filelength, gv_xstring, gv_filename,
           gv_worksheet,  gr_excel_data, gv_excel_offset.

  gv_filename = p_file.

  CALL METHOD cl_gui_frontend_services=>gui_upload
    EXPORTING
      filename                = gv_filename
      filetype                = 'BIN'
    IMPORTING
      filelength              = gv_filelength
    CHANGING
      data_tab                = gt_binary
    EXCEPTIONS
      file_open_error         = 1
      file_read_error         = 2
      no_batch                = 3
      gui_refuse_filetransfer = 4
      invalid_type            = 5
      no_authority            = 6
      unknown_error           = 7
      bad_data_format         = 8
      header_not_allowed      = 9
      separator_not_allowed   = 10
      OTHERS                  = 11.

  IF sy-subrc <> 0.
    gv_text = 'Excel upload failed. Check path and SAP GUI authorisation.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
    RETURN.
  ENDIF.

  IF gt_binary IS INITIAL OR gv_filelength IS INITIAL.
    gv_text = 'Excel file is empty or could not be read.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
    RETURN.
  ENDIF.

  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING
      input_length = gv_filelength
    IMPORTING
      buffer       = gv_xstring
    TABLES
      binary_tab   = gt_binary
    EXCEPTIONS
      failed       = 1
      OTHERS       = 2.

  IF sy-subrc <> 0 OR gv_xstring IS INITIAL.
    gv_text = 'Excel binary-to-XSTRING conversion failed.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
    RETURN.
  ENDIF.

  TRY.
      CREATE OBJECT go_excel
        EXPORTING
          document_name = gv_filename
          xdocument     = gv_xstring.

      CALL METHOD go_excel->if_fdt_doc_spreadsheet~get_worksheet_names
        IMPORTING
          worksheet_names = gt_worksheets.

      IF gt_worksheets IS INITIAL.
        gv_text = 'No worksheet found in Excel file.'.
        PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
        RETURN.
      ENDIF.

      READ TABLE gt_worksheets INTO gv_worksheet INDEX 1.
      IF sy-subrc <> 0 OR gv_worksheet IS INITIAL.
        gv_text = 'Unable to read first worksheet.'.
        PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
        RETURN.
      ENDIF.

      gr_excel_data =
        go_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( gv_worksheet ).

    CATCH cx_fdt_excel_core INTO go_excel_error.
      CLEAR gv_text.
      gv_text = go_excel_error->get_text( ).
      IF gv_text IS INITIAL.
        gv_text = 'Excel parsing failed (CL_FDT_XL_SPREADSHEET).'.
      ENDIF.
      PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
      RETURN.
  ENDTRY.

  IF gr_excel_data IS INITIAL.
    gv_text = 'Excel worksheet data reference is initial.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
    RETURN.
  ENDIF.

  ASSIGN gr_excel_data->* TO <gt_excel_dyn>.
  IF sy-subrc <> 0.
    gv_text = 'Unable to assign Excel worksheet data.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
    RETURN.
  ENDIF.

  PERFORM detect_excel_offset.

  LOOP AT <gt_excel_dyn> ASSIGNING <gs_excel_dyn>.

    gv_rowno = sy-tabix + gv_excel_offset.

    CLEAR: gs_raw, wa_input, gv_ok, gv_skip.
    gv_ok = abap_true.

    DO c_expected_cols TIMES.
      gv_index = sy-index.
      CLEAR gv_token.

      ASSIGN COMPONENT gv_index OF STRUCTURE <gs_excel_dyn> TO <fs_cell>.
      IF sy-subrc = 0.
        gv_token = <fs_cell>.
        PERFORM clean_token CHANGING gv_token.
      ENDIF.

      ASSIGN COMPONENT gv_index OF STRUCTURE gs_raw TO <fs_raw>.
      IF sy-subrc = 0.
        <fs_raw> = gv_token.
      ENDIF.
    ENDDO.

    PERFORM check_skip_row USING gs_raw CHANGING gv_skip.
    IF gv_skip = abap_true. CONTINUE. ENDIF.

    PERFORM validate_required_fields USING gs_raw gv_rowno CHANGING gv_ok.
    IF gv_ok = abap_false.
      gv_failed = gv_failed + 1.
      CONTINUE.
    ENDIF.

    PERFORM convert_raw_to_input
      USING    gs_raw gv_rowno
      CHANGING wa_input gv_ok.

    IF gv_ok = abap_true.
      CLEAR gs_record.
      gs_record-rowno = gv_rowno.
      gs_record-input = wa_input.
      APPEND gs_record TO gt_records.
    ELSE.
      gv_failed = gv_failed + 1.
    ENDIF.

  ENDLOOP.

  IF gt_records IS INITIAL AND gt_errors IS INITIAL.
    gv_text = 'No valid data rows found. Template data should start at row 4.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form DETECT_EXCEL_OFFSET
*& Reads the first cell to determine how many header rows the template
*& has so that row numbers in error messages match Excel row numbers.
*& Template: row1=field names, row2=SAP names, row3=descriptions, row4=data
*&---------------------------------------------------------------------*
FORM detect_excel_offset.

  CLEAR: gv_excel_offset, gv_first_cell, gv_second_cell.

  READ TABLE <gt_excel_dyn> ASSIGNING <gs_excel_dyn> INDEX 1.
  IF sy-subrc <> 0. gv_excel_offset = 0. RETURN. ENDIF.

  ASSIGN COMPONENT 1 OF STRUCTURE <gs_excel_dyn> TO <fs_cell>.
  IF sy-subrc = 0.
    gv_first_cell = <fs_cell>.
    PERFORM clean_token CHANGING gv_first_cell.
  ENDIF.

  TRANSLATE gv_first_cell TO UPPER CASE.

  IF      gv_first_cell = 'MBRSH'.           gv_excel_offset = 0.
  ELSEIF  gv_first_cell = 'RMMG1-MBRSH'.    gv_excel_offset = 1.
  ELSEIF  gv_first_cell = 'INDUSTRY SECTOR'. gv_excel_offset = 2.
  ELSE.                                       gv_excel_offset = 3.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CHECK_SKIP_ROW
*&---------------------------------------------------------------------*
FORM check_skip_row USING    ps_raw  TYPE ty_input_raw
                    CHANGING pv_skip TYPE abap_bool.

  DATA: lv_mbrsh TYPE string,
        lv_mtart TYPE string,
        lv_maktx TYPE string.

  pv_skip = abap_false.

  lv_mbrsh = ps_raw-mbrsh.
  lv_mtart = ps_raw-mtart.
  lv_maktx = ps_raw-maktx.

  TRANSLATE lv_mbrsh TO UPPER CASE.
  TRANSLATE lv_mtart TO UPPER CASE.
  TRANSLATE lv_maktx TO UPPER CASE.

  "Fully blank row
  IF ps_raw-mbrsh IS INITIAL AND ps_raw-mtart IS INITIAL
     AND ps_raw-werks IS INITIAL AND ps_raw-lgort IS INITIAL
     AND ps_raw-maktx IS INITIAL.
    pv_skip = abap_true. RETURN.
  ENDIF.

  "Header rows: field-name row, SAP-name row, description row
  IF lv_mbrsh = 'MBRSH'           OR lv_mtart = 'MTART'
     OR lv_maktx = 'MAKTX'
     OR lv_mbrsh = 'RMMG1-MBRSH'  OR lv_mtart = 'RMMG1-MTART'
     OR lv_mbrsh = 'INDUSTRY SECTOR'
     OR lv_mtart = 'MATERIAL TYPE'
     OR lv_maktx = 'MATERIAL DESCRIPTION'.
    pv_skip = abap_true. RETURN.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form VALIDATE_REQUIRED_FIELDS
*&---------------------------------------------------------------------*
FORM validate_required_fields USING    ps_raw   TYPE ty_input_raw
                                       pv_rowno TYPE i
                              CHANGING pv_ok    TYPE abap_bool.

  pv_ok = abap_true.

  IF ps_raw-mbrsh IS INITIAL.
    gv_text = 'Mandatory: MBRSH (Industry Sector) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  IF ps_raw-mtart IS INITIAL.
    gv_text = 'Mandatory: MTART (Material Type) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  IF ps_raw-werks IS INITIAL.
    gv_text = 'Mandatory: WERKS (Plant) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  IF ps_raw-maktx IS INITIAL.
    gv_text = 'Mandatory: MAKTX (Description) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  IF ps_raw-meins IS INITIAL.
    gv_text = 'Mandatory: MEINS (Base Unit of Measure) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  "*-- DISMM (MRP Type) is mandatory on MRP 1 - SAP msg 072 if blank
  IF ps_raw-dismm IS INITIAL.
    gv_text = 'Mandatory: DISMM (MRP Type, col 27) is blank - SAP msg 072.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  "*-- MTVFP (Availability Check) is mandatory on Sales:General/Plant
  "*-- and MRP 3 screens - SAP msg 278 if blank or msg 298 if wrong format.
  IF ps_raw-mtvfp IS INITIAL.
    gv_text = 'Mandatory: MTVFP (Availability Check, col 23) is blank - SAP msg 278.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.
  "*-- PRCTR (Profit Centre) is mandatory on Sales:General/Plant screen
  "*-- - SAP msg 278 if blank, msg 298 if wrong format.
  IF ps_raw-prctr IS INITIAL.
    gv_text = 'Mandatory: PRCTR (Profit Centre, col 22) is blank - SAP msg 278.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CLEAN_TOKEN
*&---------------------------------------------------------------------*
*& Strips CR/LF, leading/trailing whitespace, and surrounding quotes.
*& Also converts Excel integer-as-float strings (e.g. "2.0", "38089299.0")
*& to plain integer strings ("2", "38089299") so that downstream SAP
*& CHAR fields receive correctly formatted values and avoid msg 298.
*&---------------------------------------------------------------------*
FORM clean_token CHANGING cv_value TYPE string.

  DATA: lv_dot_pos  TYPE i,
        lv_frac_off TYPE i,
        lv_frac     TYPE string.

  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf   IN cv_value WITH space.
  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN cv_value WITH space.

  SHIFT cv_value LEFT  DELETING LEADING  space.
  SHIFT cv_value RIGHT DELETING TRAILING space.

  "*-- Excel stores integers in numeric cells as floating-point.
  "*-- When read via CL_FDT_XL_SPREADSHEET they may arrive as e.g. "2.0"
  "*-- or "38089299.0". Strip the decimal part when the fraction is all zeros
  "*-- so that SAP CHAR fields (MTVFP, STEUC, PRCTR, etc.) receive clean values
  "*-- and do not trigger message 298 (Formatting error).
  FIND FIRST OCCURRENCE OF '.' IN cv_value MATCH OFFSET lv_dot_pos.
  IF sy-subrc = 0.
    lv_frac_off = lv_dot_pos + 1.
    lv_frac     = cv_value+lv_frac_off.
    TRANSLATE lv_frac USING '0 '.   "replace every '0' with space
    CONDENSE lv_frac NO-GAPS.
    IF lv_frac IS INITIAL.          "fraction was all zeros -> integer value
      cv_value = cv_value(lv_dot_pos).
    ENDIF.
  ENDIF.

  gv_len = strlen( cv_value ).

  IF gv_len >= 2 AND cv_value+0(1) = '"'.
    gv_last_pos = gv_len - 1.
    IF cv_value+gv_last_pos(1) = '"'.
      gv_len = gv_len - 2.
      cv_value = cv_value+1(gv_len).
      REPLACE ALL OCCURRENCES OF '""' IN cv_value WITH '"'.
    ENDIF.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CONVERT_RAW_TO_INPUT
*& Maps all 54 Excel columns to typed input fields.
*&---------------------------------------------------------------------*
FORM convert_raw_to_input USING    ps_raw   TYPE ty_input_raw
                                   pv_rowno TYPE i
                          CHANGING ps_input TYPE ty_input
                                   pv_ok    TYPE abap_bool.

  CLEAR ps_input.
  pv_ok = abap_true.

  "--- Character fields (direct move) ---
  ps_input-mbrsh    = ps_raw-mbrsh.    "Col 01
  ps_input-mtart    = ps_raw-mtart.    "Col 02
  ps_input-werks    = ps_raw-werks.    "Col 03
  ps_input-lgort    = ps_raw-lgort.    "Col 04
  ps_input-vkorg    = ps_raw-vkorg.    "Col 05
  ps_input-vtweg    = ps_raw-vtweg.    "Col 06
  ps_input-maktx    = ps_raw-maktx.    "Col 07
  ps_input-meins    = ps_raw-meins.    "Col 08
  ps_input-matkl    = ps_raw-matkl.    "Col 09
  ps_input-spart    = ps_raw-spart.    "Col 10
  ps_input-gewei    = ps_raw-gewei.    "Col 12
  ps_input-wrkst    = ps_raw-wrkst.    "Col 14
  ps_input-sktof    = ps_raw-sktof.    "Col 15
  ps_input-taxkm1   = ps_raw-taxkm1.  "Col 16
  ps_input-taxkm2   = ps_raw-taxkm2.  "Col 17
  ps_input-versg    = ps_raw-versg.    "Col 18
  ps_input-kondm    = ps_raw-kondm.    "Col 19
  ps_input-tragr    = ps_raw-tragr.    "Col 20
  ps_input-ladgr    = ps_raw-ladgr.    "Col 21
  ps_input-prctr    = ps_raw-prctr.    "Col 22
  "*-- MTVFP (Availability Check, col 23): SAP expects a 2-char code such as
  "*-- '02', 'KP', etc.  Excel stores numeric codes (e.g. 2) as a 1-char
  "*-- string.  Left-pad a single digit with '0' to satisfy the field format
  "*-- and avoid SAP message 298 (Formatting error in MARC-MTVFP).
  IF strlen( ps_raw-mtvfp ) = 1 AND ps_raw-mtvfp CO '0123456789'.
    CONCATENATE '0' ps_raw-mtvfp INTO ps_input-mtvfp.
  ELSE.
    ps_input-mtvfp = ps_raw-mtvfp.    "Col 23
  ENDIF.
  "*-- Pre-validate MTVFP against T441 so the BDC is never called with a
  "*-- code that doesn't exist - gives a clear error instead of msg 298/278.
  IF ps_input-mtvfp IS NOT INITIAL.
    PERFORM check_mtvfp_t441 USING ps_input-mtvfp pv_rowno CHANGING pv_ok.
  ENDIF.
  ps_input-xchpf    = ps_raw-xchpf.    "Col 24
  "*-- STEUC (Control Code / HSN, col 25): validated against T604F at the
  "*-- SAP screen level (msg 058 if entry missing).  Pre-validate here so
  "*-- the BDC is never called with an invalid HSN and the user receives a
  "*-- clear message pointing to the correct maintenance transaction.
  "*-- clean_token has already stripped any Excel float '.0' suffix (msg 298).
  ps_input-steuc    = ps_raw-steuc.    "Col 25
  "*-- Pre-validate STEUC against T604F for the plant's country.
  IF ps_input-steuc IS NOT INITIAL.
    PERFORM check_steuc_t604f USING ps_input-steuc ps_input-werks pv_rowno CHANGING pv_ok.
  ENDIF.
  ps_input-taxim    = ps_raw-taxim.    "Col 26
  ps_input-dismm    = ps_raw-dismm.    "Col 27
  ps_input-beskz    = ps_raw-beskz.    "Col 28
  ps_input-perkz    = ps_raw-perkz.    "Col 29
  ps_input-iprkz    = ps_raw-iprkz.    "Col 30
  ps_input-sled_bbd = ps_raw-sled_bbd."Col 31
  ps_input-qmpur    = ps_raw-qmpur.    "Col 32
  ps_input-ssqss    = ps_raw-ssqss.    "Col 33
  ps_input-bklas    = ps_raw-bklas.    "Col 34
  ps_input-vprsv    = ps_raw-vprsv.    "Col 35
  ps_input-ekalr    = ps_raw-ekalr.    "Col 38
  ps_input-hkmat    = ps_raw-hkmat.    "Col 39
  ps_input-sobsk    = ps_raw-sobsk.    "Col 40
  ps_input-mlast    = ps_raw-mlast.    "Col 42
  ps_input-insmk    = ps_raw-insmk.    "Col 50 - MARC-INSMK
  ps_input-kzdkz    = ps_raw-kzdkz.    "Col 51 - MARC-KZDKZ
  ps_input-ncost    = ps_raw-ncost.    "Col 52 - MARC-NCOST
  ps_input-qpls_arg = ps_raw-qpls_arg. "Col 53 - RMQAM-ARGUMENT (QPLS popup)
  ps_input-qpls_art = ps_raw-qpls_art. "Col 54 - RMQAM-ART(01) (QPLS popup)

  "--- Packed/numeric fields (via move_numeric) ---
  PERFORM move_numeric USING ps_raw-brgew  'BRGEW'  pv_rowno CHANGING ps_input-brgew  pv_ok. "Col 11
  PERFORM move_numeric USING ps_raw-ntgew  'NTGEW'  pv_rowno CHANGING ps_input-ntgew  pv_ok. "Col 13
  PERFORM move_numeric USING ps_raw-stprs  'STPRS'  pv_rowno CHANGING ps_input-stprs  pv_ok. "Col 36
  PERFORM move_numeric USING ps_raw-peinh  'PEINH'  pv_rowno CHANGING ps_input-peinh  pv_ok. "Col 37
  PERFORM move_numeric USING ps_raw-losgr  'LOSGR'  pv_rowno CHANGING ps_input-losgr  pv_ok. "Col 41
  PERFORM move_numeric USING ps_raw-stprs1 'STPRS1' pv_rowno CHANGING ps_input-stprs1 pv_ok. "Col 43
  PERFORM move_numeric USING ps_raw-stprs2 'STPRS2' pv_rowno CHANGING ps_input-stprs2 pv_ok. "Col 44
  PERFORM move_numeric USING ps_raw-stprs3 'STPRS3' pv_rowno CHANGING ps_input-stprs3 pv_ok. "Col 45
  PERFORM move_numeric USING ps_raw-peinh1 'PEINH1' pv_rowno CHANGING ps_input-peinh1 pv_ok. "Col 46
  PERFORM move_numeric USING ps_raw-peinh2 'PEINH2' pv_rowno CHANGING ps_input-peinh2 pv_ok. "Col 47
  PERFORM move_numeric USING ps_raw-peinh3 'PEINH3' pv_rowno CHANGING ps_input-peinh3 pv_ok. "Col 48
  PERFORM move_numeric USING ps_raw-verpr  'VERPR'  pv_rowno CHANGING ps_input-verpr  pv_ok. "Col 49

ENDFORM.

*&---------------------------------------------------------------------*
*& Form MOVE_NUMERIC
*&---------------------------------------------------------------------*
FORM move_numeric USING    pv_value TYPE string
                            pv_field TYPE string
                            pv_rowno TYPE i
                   CHANGING cv_num   TYPE any
                            cv_ok    TYPE abap_bool.

  CLEAR gv_numtext.
  gv_numtext = pv_value.

  SHIFT gv_numtext LEFT  DELETING LEADING  space.
  SHIFT gv_numtext RIGHT DELETING TRAILING space.
  CONDENSE gv_numtext NO-GAPS.

  IF gv_numtext IS INITIAL.
    CLEAR cv_num.
    RETURN.
  ENDIF.

  REPLACE ALL OCCURRENCES OF ',' IN gv_numtext WITH '.'.

  TRY.
      cv_num = gv_numtext.
    CATCH cx_sy_conversion_no_number cx_sy_conversion_overflow.
      CLEAR gv_text.
      CONCATENATE 'Invalid numeric value in' pv_field ':' pv_value
             INTO gv_text SEPARATED BY space.
      PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
      cv_ok = abap_false.
  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CHECK_MTVFP_T441
*& Pre-validates the Availability Check Group code against table T441.
*& Avoids SAP messages 298 (Formatting error) + 278 (mandatory field)
*& from the BDC and replaces them with a clear, actionable pre-run error.
*&
*& Root cause of 298/278 for MARC-MTVFP:
*&  - Excel numeric cell (e.g. 2) arrives as "2.0" -> clean_token strips
*&    to "2" -> padded to "02" in convert_raw_to_input.
*&  - If the padded value ("02", "01", etc.) is still not in T441 the
*&    BDC would fail with 298+278.  This form surfaces that early.
*&---------------------------------------------------------------------*
FORM check_mtvfp_t441 USING    pv_mtvfp TYPE marc-mtvfp
                               pv_rowno TYPE i
                      CHANGING pv_ok    TYPE abap_bool.

  DATA lv_mtvfp TYPE t441-mtvfp.

  TRY.
      SELECT SINGLE mtvfp FROM t441 INTO lv_mtvfp WHERE mtvfp = pv_mtvfp.
    CATCH cx_sy_open_sql_error.
      RETURN.  " T441 inaccessible - skip check, let BDC surface the error
  ENDTRY.

  IF sy-subrc <> 0.
    CONCATENATE 'MARC-MTVFP "' pv_mtvfp
                '" does not exist in T441 (Availability Check Groups).'
                ' Correct col 23 in the Excel template or add the code'
                ' in SPRO > MM > Plant Params > Avail. Check.'
      INTO gv_text SEPARATED BY space.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CHECK_STEUC_T604F
*& Pre-validates the HSN / commodity code against table T604F for the
*& plant's country.  Avoids SAP messages 298 + 058 from the BDC and
*& replaces them with a clear, actionable pre-run error that tells the
*& user exactly which entry to create and which transaction to use.
*&
*& Root cause of 298/058 for MARC-STEUC:
*&  - "298 Formatting error ... see next message" is a wrapper message;
*&    the real failure is "058 Entry IN <code> does not exist in T604F".
*&  - The float notation fix in clean_token removes "38089299.0" -> "38089299"
*&    (eliminates 298 from bad format).
*&  - The 058 error itself means the HSN entry is missing in T604F.
*&    Create it via SM30 > V_T604F or transaction VEN3 for country = LAND1
*&    of the plant, then re-run the BDC.
*&---------------------------------------------------------------------*
FORM check_steuc_t604f USING    pv_steuc TYPE marc-steuc
                                pv_werks TYPE marc-werks
                                pv_rowno TYPE i
                       CHANGING pv_ok    TYPE abap_bool.

  DATA: lv_land1  TYPE t001w-land1,
        lv_zollnr TYPE t604f-zollnr,
        lv_dummy  TYPE t604f-zollnr.

  " Determine plant country from T001W
  SELECT SINGLE land1 FROM t001w INTO lv_land1 WHERE werks = pv_werks.
  IF sy-subrc <> 0 OR lv_land1 IS INITIAL.
    RETURN.  " Cannot determine country - let BDC surface any error
  ENDIF.

  lv_zollnr = pv_steuc.  " widen 16-char STEUC to T604F-ZOLLNR field length

  TRY.
      SELECT SINGLE zollnr FROM t604f INTO lv_dummy
        WHERE land1 = lv_land1
          AND zollnr = lv_zollnr.
    CATCH cx_sy_open_sql_error.
      RETURN.  " T604F inaccessible - skip check, let BDC surface the error
  ENDTRY.

  IF sy-subrc <> 0.
    CONCATENATE 'MARC-STEUC (HSN) "' pv_steuc
                '" does not exist in T604F for country' lv_land1
                '(plant' pv_werks ').'
                ' Create the commodity-code entry via SM30 > V_T604F'
                ' or transaction VEN3, then re-run the BDC.'
      INTO gv_text SEPARATED BY space.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BUILD_BDC
*&---------------------------------------------------------------------*
*& Complete screen sequence confirmed from SHDB recording + template:
*&
*&  0060 Initial screen
*&  0070 View page 1  (=P+ scroll)
*&  0070 View page 2  (=SCHL close)
*&  0080 Org levels
*&  4004 Basic Data 1
*&  4004 Basic Data 2             WRKST
*&  4000 Sales Org 1  pass 1/4    SKTOF, TAXKM1, TAXKM2
*&  4000 Sales Org 1  pass 2/4    re-display: SKTOF, TAXKM2
*&  4000 Sales Org 1  pass 3/4    re-display: SKTOF, TAXKM2
*&  4000 Sales Org 1  pass 4/4    navigate to Org 2
*&  4000 Sales Org 2              VERSG, KONDM
*&  4004 Sales Gen/Plant          MTVFP, XCHPF, TRAGR, LADGR, PRCTR
*&  4004 Int'l Trade Export       STEUC
*&  4000 Purchasing               MG03STEUMM-TAXIM (no index suffix)
*&  4000 MRP 1                    DISMM
*&  4000 MRP 2                    BESKZ
*&  4000 MRP 3                    PERKZ, MTVFP
*&  4000 MRP 4                    navigation only
*&  4000 Gen Plant/Storage 1      IPRKZ, SLED_BBD
*&  4000 Gen Plant/Storage 2      navigation only
*&  4000 Quality Management       QMPUR, INSMK, KZDKZ, SSQSS
*&  [QPLS popup - conditional]    QPLS_ARG, QPLS_ART (if filled)
*&  4000 Costing 1  pass 1/3      MLAST, STPRS1, PEINH1
*&  4000 Costing 1  pass 2/3      MLAST, all STPRS/PEINH
*&  4000 Costing 1  pass 3/3      BKLAS + all STPRS/PEINH
*&  4000 Costing navigation       navigation only
*&  4000 Costing 2                EKALR, HKMAT, SOBSK, LOSGR, NCOST
*&  4000 Accounting 1             VPRSV, BKLAS, STPRS, PEINH (VERPR if V)
*&  4000 Accounting 2             navigation only (view is selected)
*&  0300 Save popup               =YES
*&
*& COSTING FIELD NOTE: Screen field is CKMMAT_DISPLAY (double-M) from
*&  SHDB recording. The Excel template header shows CKMAT_DISPLAY
*&  (single-M) as a documentation label only - it does not affect the
*&  column-to-variable mapping which is done by position.
*&
*& QPLS POPUP NOTE: If QPLS_ARG / QPLS_ART are filled, the QM screen
*&  triggers an inspection-plan popup. Add the exact program/dynpro
*&  from your SHDB re-recording if you need this popup handled.
*&---------------------------------------------------------------------*
FORM build_bdc USING ps_input TYPE ty_input.

  REFRESH gt_bdcdata.

  "===============================================================
  " 0060 - Initial Screen
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0060.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'RMMG1-MTART'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_entr.
  PERFORM bdc_field  USING 'RMMG1-MBRSH' ps_input-mbrsh.
  PERFORM bdc_field  USING 'RMMG1-MTART' ps_input-mtart.
  IF p_matnr IS NOT INITIAL.
    PERFORM bdc_field USING 'RMMG1-MATNR' p_matnr.
  ENDIF.

  "===============================================================
  " 0070 - View Selection  PAGE 1  (absolute rows 1-10 visible)
  "
  " Row  View                       Selected
  "  01  Basic Data 1               YES
  "  02  Basic Data 2               YES
  "  03  Classification             skip
  "  04  Sales: Org Data 1          YES
  "  05  Sales: Org Data 2          YES
  "  06  Sales: General/Plant       YES
  "  07  Extended SPP Basic Data    skip
  "  08  Int'l Trade: Export        YES
  "  09  Sales Text                 skip  (no text screen needed)
  "  10  Purchasing                 YES
  "  11  Int'l Trade: Import        skip  (visible but not selected)
  "  12  Purchase Order Text        skip  (visible but not selected)
  "
  " Cursor at last selected visible row (10). Scroll with =P+.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR'           'MSICHTAUSW-DYTXT(10)'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.  "Basic Data 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' c_x.  "Basic Data 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.  "Sales: Org Data 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' c_x.  "Sales: Org Data 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' c_x.  "Sales: General/Plant
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.  "Int'l Trade: Export
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(10)' c_x.  "Purchasing
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_p_plus. "scroll

  "===============================================================
  " 0070 - View Selection  PAGE 2  (relative rows restart at MRP 1)
  "
  " Rel  Absolute view               Selected
  "  01  MRP 1                       YES
  "  02  MRP 2                       YES
  "  03  MRP 3                       YES
  "  04  MRP 4                       YES
  "  05  Advanced Planning           skip
  "  06  Extended SPP                skip
  "  07  Forecasting                 skip
  "  08  Gen Plant Data/Storage 1    YES
  "  09  Gen Plant Data/Storage 2    YES
  "  10  Warehouse Mgmt 1            skip
  "  11  Warehouse Mgmt 2            skip
  "  12  Quality Management          YES
  "  13  Accounting 1                YES
  "  14  Accounting 2                YES  (confirmed from image)
  "  15  Costing 1                   YES
  "  16  Costing 2                   YES
  "  17  WM Execution                skip
  "
  " Cursor at last selected row (16). =SCHL closes the dialog.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR'           'MSICHTAUSW-DYTXT(16)'.
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.  "MRP 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' c_x.  "MRP 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(03)' c_x.  "MRP 3
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.  "MRP 4
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.  "Gen Plant/Storage 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(09)' c_x.  "Gen Plant/Storage 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(12)' c_x.  "Quality Management
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(13)' c_x.  "Accounting 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(14)' c_x.  "Accounting 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(15)' c_x.  "Costing 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(16)' c_x.  "Costing 2
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_schl. "close

  "===============================================================
  " 0080 - Organisational Levels
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0080.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'RMMG1-VTWEG'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_entr.
  PERFORM bdc_field  USING 'RMMG1-WERKS' ps_input-werks.
  PERFORM bdc_field  USING 'RMMG1-LGORT' ps_input-lgort.
  PERFORM bdc_field  USING 'RMMG1-VKORG' ps_input-vkorg.
  PERFORM bdc_field  USING 'RMMG1-VTWEG' ps_input-vtweg.

  "===============================================================
  " 4004 - Basic Data 1
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARA-NTGEW'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MAKT-MAKTX'  ps_input-maktx.
  PERFORM bdc_field  USING 'MARA-MEINS'  ps_input-meins.
  PERFORM bdc_field  USING 'MARA-MATKL'  ps_input-matkl.
  PERFORM bdc_field  USING 'MARA-SPART'  ps_input-spart.
  PERFORM bdc_num    USING 'MARA-BRGEW'  ps_input-brgew.
  PERFORM bdc_field  USING 'MARA-GEWEI'  ps_input-gewei.
  PERFORM bdc_num    USING 'MARA-NTGEW'  ps_input-ntgew.

  "===============================================================
  " 4004 - Basic Data 2  (WRKST - Basic/Raw Material)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARA-WRKST'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARA-WRKST'  ps_input-wrkst.

  "===============================================================
  " 4000 - Sales Org 1  pass 1/4  (initial data entry)
  " SAP re-displays 3 more times to confirm tax category entries.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'            'MG03STEUER-TAXKM(02)'.
  PERFORM bdc_field  USING 'BDC_OKCODE'            c_ok_enter.
  PERFORM bdc_field  USING 'MVKE-SKTOF'            ps_input-sktof.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(01)'  ps_input-taxkm1.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(02)'  ps_input-taxkm2.

  "===============================================================
  " 4000 - Sales Org 1  pass 2/4  (re-display - confirm TAXKM2)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'            'MG03STEUER-TAXKM(02)'.
  PERFORM bdc_field  USING 'BDC_OKCODE'            c_ok_enter.
  PERFORM bdc_field  USING 'MVKE-SKTOF'            ps_input-sktof.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(02)'  ps_input-taxkm2.

  "===============================================================
  " 4000 - Sales Org 1  pass 3/4  (re-display - confirm again)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'            'MG03STEUER-TAXKM(02)'.
  PERFORM bdc_field  USING 'BDC_OKCODE'            c_ok_enter.
  PERFORM bdc_field  USING 'MVKE-SKTOF'            ps_input-sktof.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(02)'  ps_input-taxkm2.

  "===============================================================
  " 4000 - Sales Org 1  pass 4/4  (navigate to Sales Org 2)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MVKE-SKTOF'  ps_input-sktof.

  "===============================================================
  " 4000 - Sales Org 2  (VERSG, KONDM)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MVKE-KONDM'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MVKE-VERSG'  ps_input-versg.
  PERFORM bdc_field  USING 'MVKE-KONDM'  ps_input-kondm.

  "===============================================================
  " 4004 - Sales: General / Plant  (MTVFP, XCHPF, TRAGR, LADGR, PRCTR)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARC-XCHPF'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARC-MTVFP'  ps_input-mtvfp.
  PERFORM bdc_field  USING 'MARC-XCHPF'  ps_input-xchpf.
  PERFORM bdc_field  USING 'MARA-TRAGR'  ps_input-tragr.
  PERFORM bdc_field  USING 'MARC-LADGR'  ps_input-ladgr.
  PERFORM bdc_field  USING 'MARC-PRCTR'  ps_input-prctr.

  "===============================================================
  " 4004 - International Trade: Export  (STEUC - Control Code/HSN)
  " Note: msg 058 'Entry XX nnnn does not exist in T604F' is a master-data
  " error, not a code error. The HSN code must exist in T604F for country XX
  " (e.g. via transaction VEN3 or SM30 on view V_T604F) before this BDC runs.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARC-STEUC'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARC-STEUC'  ps_input-steuc.

  "===============================================================
  " 4000 - Purchasing  (MG03STEUMM-TAXIM - no (01) suffix)
  " Sales Text and PO Text are NOT selected so no text screens appear.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'       'MG03STEUMM-TAXIM'.
  PERFORM bdc_field  USING 'BDC_OKCODE'       c_ok_enter.
  PERFORM bdc_field  USING 'MG03STEUMM-TAXIM' ps_input-taxim.

  "===============================================================
  " 4000 - MRP 1  (DISMM - MRP Type)
  " Cursor at T438T-DIBEZ (MRP type text field, from recording).
  " SAP fires msg 072 'Enter the MRP type' when DISMM is blank.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'T438T-DIBEZ'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARC-DISMM'  ps_input-dismm.

  "===============================================================
  " 4000 - MRP 2  (BESKZ - Procurement Type)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARC-BESKZ'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARC-BESKZ'  ps_input-beskz.

  "===============================================================
  " 4000 - MRP 3  (PERKZ, MTVFP - MTVFP also appears here per recording)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARC-PERKZ'  ps_input-perkz.
  PERFORM bdc_field  USING 'MARC-MTVFP'  ps_input-mtvfp.

  "===============================================================
  " 4000 - MRP 4  (navigation only)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.

  "===============================================================
  " 4000 - General Plant Data / Storage 1  (IPRKZ, SLED_BBD)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'     'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'     c_ok_enter.
  PERFORM bdc_field  USING 'MARA-IPRKZ'     ps_input-iprkz.
  PERFORM bdc_field  USING 'MARA-SLED_BBD'  ps_input-sled_bbd.

  "===============================================================
  " 4000 - General Plant Data / Storage 2  (navigation only)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.

  "===============================================================
  " 4000 - Quality Management
  "   MARA-QMPUR  QM Procurement Active       (Col 32)
  "   MARC-INSMK  Inspection Stock Indicator  (Col 50)  <-- NEW
  "   MARC-KZDKZ  Post to Inspection Stock    (Col 51)  <-- NEW
  "   MARC-SSQSS  QM Control Key              (Col 33)
  " NOTE: If QPLS_ARG/QPLS_ART are filled an inspection-plan popup
  " may appear after this screen. See the conditional block below.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARC-SSQSS'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MARA-QMPUR'  ps_input-qmpur.
  PERFORM bdc_field  USING 'MARC-INSMK'  ps_input-insmk.
  PERFORM bdc_field  USING 'MARC-KZDKZ'  ps_input-kzdkz.
  PERFORM bdc_field  USING 'MARC-SSQSS'  ps_input-ssqss.

  "===============================================================
  " QPLS Inspection Plan Popup  (conditional)
  " Appears when SSQSS requires an inspection plan selection.
  " Exact program/screen must be confirmed with an SHDB re-recording
  " that includes the QPLS popup interaction. Uncomment and fill in
  " the correct c_prog_xxx/c_scr_xxx values after re-recording.
  "===============================================================
  IF ps_input-qpls_arg IS NOT INITIAL OR ps_input-qpls_art IS NOT INITIAL.
*   PERFORM bdc_dynpro USING '<PROG>' '<SCR>'.     "<-- get from SHDB
*   PERFORM bdc_field  USING 'RMQAM-ARGUMENT' ps_input-qpls_arg.
*   PERFORM bdc_field  USING 'RMQAM-ART(01)'  ps_input-qpls_art.
*   PERFORM bdc_field  USING 'BDC_OKCODE'     c_ok_enter.
  ENDIF.

  "===============================================================
  " 4000 - Costing 1  pass 1/3  (initial: MLAST + STPRS1/PEINH1)
  " Screen field: CKMMAT_DISPLAY (double-M, from SHDB recording).
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'                'CKMMAT_DISPLAY-STPRS_1'.
  PERFORM bdc_field  USING 'BDC_OKCODE'                c_ok_enter.
  PERFORM bdc_field  USING 'CKMLHD-MLAST'              ps_input-mlast.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_1'    ps_input-stprs1.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_1'    ps_input-peinh1.

  "===============================================================
  " 4000 - Costing 1  pass 2/3  (confirm all period prices)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'                'MBEW-BKLAS'.
  PERFORM bdc_field  USING 'BDC_OKCODE'                c_ok_enter.
  PERFORM bdc_field  USING 'CKMLHD-MLAST'              ps_input-mlast.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_1'    ps_input-stprs1.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_2'    ps_input-stprs2.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_3'    ps_input-stprs3.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_1'    ps_input-peinh1.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_2'    ps_input-peinh2.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_3'    ps_input-peinh3.

  "===============================================================
  " 4000 - Costing 1  pass 3/3  (set BKLAS + confirm all prices)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'                'MBEW-BKLAS'.
  PERFORM bdc_field  USING 'BDC_OKCODE'                c_ok_enter.
  PERFORM bdc_field  USING 'MBEW-BKLAS'                ps_input-bklas.
  PERFORM bdc_field  USING 'CKMLHD-MLAST'              ps_input-mlast.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_1'    ps_input-stprs1.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_2'    ps_input-stprs2.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-STPRS_3'    ps_input-stprs3.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_1'    ps_input-peinh1.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_2'    ps_input-peinh2.
  PERFORM bdc_num    USING 'CKMMAT_DISPLAY-PEINH_3'    ps_input-peinh3.

  "===============================================================
  " 4000 - Costing navigation  (navigation only)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.

  "===============================================================
  " 4000 - Costing 2  (EKALR, HKMAT, SOBSK, LOSGR, NCOST)
  "   MARC-NCOST  Do Not Cost Material  (Col 52)  <-- NEW
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MARC-SOBSK'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MBEW-EKALR'  ps_input-ekalr.
  PERFORM bdc_field  USING 'MBEW-HKMAT'  ps_input-hkmat.
  IF ps_input-sobsk IS NOT INITIAL.
    PERFORM bdc_field USING 'MARC-SOBSK' ps_input-sobsk.
  ENDIF.
  PERFORM bdc_num    USING 'MARC-LOSGR'  ps_input-losgr.
  PERFORM bdc_field  USING 'MARC-NCOST'  ps_input-ncost.

  "===============================================================
  " 4000 - Accounting 1  (VPRSV, BKLAS, STPRS, PEINH)
  " VERPR posted only when VPRSV = 'V' (moving average price).
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MBEW-BKLAS'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.
  PERFORM bdc_field  USING 'MBEW-VPRSV'  ps_input-vprsv.
  PERFORM bdc_field  USING 'MBEW-BKLAS'  ps_input-bklas.
  PERFORM bdc_num    USING 'MBEW-STPRS'  ps_input-stprs.
  PERFORM bdc_num    USING 'MBEW-PEINH'  ps_input-peinh.
  IF ps_input-vprsv = 'V'.
    PERFORM bdc_num USING 'MBEW-VERPR' ps_input-verpr.
  ENDIF.

  "===============================================================
  " 4000 - Accounting 2  (navigation only)
  " Accounting 2 is selected in the view dialog so SAP displays it.
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'MAKT-MAKTX'.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_enter.

  "===============================================================
  " SAPLSPO1 0300 - Save confirmation popup  -> YES
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_spo1 c_scr_spo1.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_yes.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BDC_DYNPRO
*&---------------------------------------------------------------------*
FORM bdc_dynpro USING pv_program LIKE bdcdata-program
                      pv_dynpro  LIKE bdcdata-dynpro.

  CLEAR gs_bdcdata.
  gs_bdcdata-program  = pv_program.
  gs_bdcdata-dynpro   = pv_dynpro.
  gs_bdcdata-dynbegin = c_x.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BDC_FIELD
*& Appends only when fval is non-initial, or fnam = BDC_OKCODE/CURSOR.
*&---------------------------------------------------------------------*
FORM bdc_field USING pv_fnam LIKE bdcdata-fnam
                     pv_fval TYPE any.

  CLEAR gs_bdcdata.
  gs_bdcdata-fnam = pv_fnam.
  gs_bdcdata-fval = pv_fval.

  IF gs_bdcdata-fval IS NOT INITIAL
     OR pv_fnam = 'BDC_OKCODE'
     OR pv_fnam = 'BDC_CURSOR'.
    APPEND gs_bdcdata TO gt_bdcdata.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BDC_NUM
*& Formats a packed/numeric value as plain text and appends it.
*& Zero or blank values are skipped to avoid overwriting SAP defaults.
*&---------------------------------------------------------------------*
FORM bdc_num USING pv_fnam LIKE bdcdata-fnam
                   pv_num  TYPE any.

  CLEAR gv_numtext.

  WRITE pv_num TO gv_numtext NO-GROUPING LEFT-JUSTIFIED.
  CONDENSE gv_numtext NO-GAPS.

  IF gv_numtext IS INITIAL OR gv_numtext = '0'.
    RETURN.
  ENDIF.

  CLEAR gs_bdcdata.
  gs_bdcdata-fnam = pv_fnam.
  gs_bdcdata-fval = gv_numtext.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form ANALYZE_BDC_MESSAGES
*&---------------------------------------------------------------------*
FORM analyze_bdc_messages USING pv_rowno TYPE i.

  CLEAR: gv_row_failed, gv_material.
  gv_row_failed = abap_false.
  gv_material   = p_matnr.

  LOOP AT gt_msgcoll INTO gs_msgcoll.

    IF gs_msgcoll-msgtyp = 'S' AND gv_material IS INITIAL.
      IF gs_msgcoll-msgv1 IS NOT INITIAL.
        gv_material = gs_msgcoll-msgv1.
      ENDIF.
    ENDIF.

    IF gs_msgcoll-msgtyp = 'E' OR gs_msgcoll-msgtyp = 'A'.
      gv_row_failed = abap_true.
      PERFORM get_message_text USING gs_msgcoll CHANGING gv_msgtext.
      PERFORM add_error USING pv_rowno gv_material
                              gs_msgcoll-msgtyp
                              gs_msgcoll-msgid
                              gs_msgcoll-msgnr
                              gv_msgtext.
    ENDIF.

  ENDLOOP.

  IF gv_call_subrc <> 0 AND gv_row_failed = abap_false.
    gv_row_failed = abap_true.
    CLEAR gv_subrc_c.
    WRITE gv_call_subrc TO gv_subrc_c LEFT-JUSTIFIED.
    CONDENSE gv_subrc_c.
    CLEAR gv_msgtext.
    CONCATENATE 'CALL TRANSACTION SY-SUBRC =' gv_subrc_c
           INTO gv_msgtext SEPARATED BY space.
    PERFORM add_error USING pv_rowno gv_material 'E' 'LOCAL' '000' gv_msgtext.
  ENDIF.

  IF gv_row_failed = abap_true.
    gv_failed  = gv_failed  + 1.
  ELSE.
    gv_success = gv_success + 1.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form GET_MESSAGE_TEXT
*&---------------------------------------------------------------------*
FORM get_message_text USING    ps_msg  TYPE bdcmsgcoll
                      CHANGING pv_text TYPE ty_msgtext.

  CLEAR pv_text.
  MESSAGE ID ps_msg-msgid TYPE ps_msg-msgtyp NUMBER ps_msg-msgnr
          WITH ps_msg-msgv1 ps_msg-msgv2 ps_msg-msgv3 ps_msg-msgv4
          INTO pv_text.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form ADD_ERROR
*&---------------------------------------------------------------------*
FORM add_error USING pv_rowno    TYPE i
                     pv_material TYPE matnr
                     pv_msgtyp   TYPE c
                     pv_msgid    LIKE bdcmsgcoll-msgid
                     pv_msgno    LIKE bdcmsgcoll-msgnr
                     pv_msgtext  TYPE ty_msgtext.

  CLEAR gs_error.
  gs_error-rowno    = pv_rowno.
  gs_error-material = pv_material.
  gs_error-msgtyp   = pv_msgtyp.
  gs_error-msgid    = pv_msgid.
  gs_error-msgno    = pv_msgno.
  gs_error-msgtext  = pv_msgtext.
  APPEND gs_error TO gt_errors.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form DISPLAY_SUMMARY
*&---------------------------------------------------------------------*
FORM display_summary.

  ULINE.
  WRITE: / 'MM01 ROH BDC Processing Summary'.
  ULINE.
  WRITE: / 'Valid records processed :', gv_total.
  WRITE: / 'Success count           :', gv_success.
  WRITE: / 'Failure count           :', gv_failed.
  ULINE.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form DISPLAY_ERRORS
*&---------------------------------------------------------------------*
FORM display_errors.

  IF gt_errors IS INITIAL.
    WRITE: / 'No errors found.'.
    RETURN.
  ENDIF.

  PERFORM build_fieldcat.

  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      i_callback_program = sy-repid
      it_fieldcat        = gt_fieldcat
    TABLES
      t_outtab           = gt_errors
    EXCEPTIONS
      program_error      = 1
      OTHERS             = 2.

  IF sy-subrc <> 0.
    WRITE: / 'ALV display failed. Debug GT_ERRORS table directly.'.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BUILD_FIELDCAT
*&---------------------------------------------------------------------*
FORM build_fieldcat.

  REFRESH gt_fieldcat.
  PERFORM add_fieldcat USING 'ROWNO'    'Row#'     8.
  PERFORM add_fieldcat USING 'MATERIAL' 'Material' 18.
  PERFORM add_fieldcat USING 'MSGTYP'   'Type'     5.
  PERFORM add_fieldcat USING 'MSGID'    'Msg ID'   20.
  PERFORM add_fieldcat USING 'MSGNO'    'Msg No'   8.
  PERFORM add_fieldcat USING 'MSGTEXT'  'Message'  120.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form ADD_FIELDCAT
*&---------------------------------------------------------------------*
FORM add_fieldcat USING pv_field TYPE slis_fieldname
                        pv_text  TYPE ty_char40
                        pv_len   TYPE i.

  CLEAR gs_fieldcat.
  gs_fieldcat-fieldname = pv_field.
  gs_fieldcat-seltext_l = pv_text.
  gs_fieldcat-seltext_m = pv_text.
  gs_fieldcat-seltext_s = pv_text.
  gs_fieldcat-outputlen = pv_len.
  APPEND gs_fieldcat TO gt_fieldcat.

ENDFORM.
