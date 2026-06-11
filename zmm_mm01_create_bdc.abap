
REPORT zmm_mm01_create_bdc.

TYPE-POOLS: slis.

*&---------------------------------------------------------------------*
*& Type Declarations
*&---------------------------------------------------------------------*
TYPES: ty_msgtext TYPE c LENGTH 255,
       ty_char40  TYPE c LENGTH 40.

TYPES: BEGIN OF ty_input,
         mbrsh    TYPE c LENGTH 1,
         mtart    TYPE c LENGTH 4,
         werks    TYPE c LENGTH 4,
         lgort    TYPE c LENGTH 4,
         vkorg    TYPE c LENGTH 4,
         vtweg    TYPE c LENGTH 2,
         maktx    TYPE c LENGTH 40,
         meins    TYPE c LENGTH 3,
         matkl    TYPE c LENGTH 9,
         spart    TYPE c LENGTH 2,
         brgew    TYPE p DECIMALS 3,
         gewei    TYPE c LENGTH 3,
         ntgew    TYPE p DECIMALS 3,
         wrkst    TYPE c LENGTH 48,
         sktof    TYPE c LENGTH 1,
         taxkm1   TYPE c LENGTH 1,
         taxkm2   TYPE c LENGTH 1,
         versg    TYPE c LENGTH 1,
         kondm    TYPE c LENGTH 2,
         tragr    TYPE c LENGTH 4,
         ladgr    TYPE c LENGTH 4,
         prctr    TYPE c LENGTH 10,
         mtvfp    TYPE c LENGTH 2,
         xchpf    TYPE c LENGTH 1,
         steuc    TYPE c LENGTH 16,
         taxim    TYPE c LENGTH 1,
         dismm    TYPE c LENGTH 2,
         beskz    TYPE c LENGTH 1,
         perkz    TYPE c LENGTH 1,
         iprkz    TYPE c LENGTH 1,
         sled_bbd TYPE c LENGTH 1,
         qmpur    TYPE c LENGTH 1,
         ssqss    TYPE c LENGTH 8,
         bklas    TYPE c LENGTH 4,
         vprsv    TYPE c LENGTH 1,
         stprs    TYPE p DECIMALS 2,
         peinh    TYPE p DECIMALS 0,
         ekalr    TYPE c LENGTH 1,
         hkmat    TYPE c LENGTH 1,
         sobsk    TYPE c LENGTH 2,
         losgr    TYPE p DECIMALS 0,
         mlast    TYPE c LENGTH 1,
         stprs1   TYPE p DECIMALS 2,
         stprs2   TYPE p DECIMALS 2,
         stprs3   TYPE p DECIMALS 2,
         peinh1   TYPE p DECIMALS 0,
         peinh2   TYPE p DECIMALS 0,
         peinh3   TYPE p DECIMALS 0,
         verpr    TYPE p DECIMALS 2,
         insmk    TYPE c LENGTH 1,
         kzdkz    TYPE c LENGTH 1,
         ncost    TYPE c LENGTH 1,
         qpls_arg TYPE c LENGTH 10,
         qpls_art TYPE c LENGTH 4,
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
  c_expected_cols TYPE i     VALUE 54,
  c_update_sync   TYPE c     VALUE 'S',
  c_x             TYPE c     VALUE 'X',

  c_ok_enter      TYPE bdcdata-fval VALUE '/00',
  c_ok_entr       TYPE bdcdata-fval VALUE '=ENTR',
  c_ok_schl       TYPE bdcdata-fval VALUE '=SCHL',
  c_ok_yes        TYPE bdcdata-fval VALUE '=YES',

  c_prog_mm       LIKE bdcdata-program VALUE 'SAPLMGMM',
  c_scr_0060      LIKE bdcdata-dynpro  VALUE '0060',
  c_scr_0070      LIKE bdcdata-dynpro  VALUE '0070',
  c_scr_0080      LIKE bdcdata-dynpro  VALUE '0080',
  c_scr_4000      LIKE bdcdata-dynpro  VALUE '4000',
  c_scr_4004      LIKE bdcdata-dynpro  VALUE '4004',

  "Final save confirmation popup (from recording).
  c_prog_spo1     LIKE bdcdata-program VALUE 'SAPLSPO1',
  c_scr_spo1      LIKE bdcdata-dynpro  VALUE '0300'.

*&---------------------------------------------------------------------*
*& Selection Screen
*&---------------------------------------------------------------------*
PARAMETERS:
  p_file  TYPE rlgrap-filename OBLIGATORY,
  p_mode  TYPE c DEFAULT 'A',
  p_matnr TYPE matnr.              "Blank = internal number range

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

  WRITE: / 'Material Master Creation BDC'.
  SKIP.

  PERFORM upload_excel_xlsx.

  DESCRIBE TABLE gt_records LINES gv_total.

  IF gt_records IS INITIAL.
    WRITE: / 'No valid input records found.'.
    PERFORM display_errors.
    RETURN.
  ENDIF.

  LOOP AT gt_records INTO gs_record.

    REFRESH:
      gt_bdcdata,
      gt_msgcoll.

    CLEAR:
      gv_material,
      gv_call_subrc,
      gv_row_failed,
      wa_input.

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
  CLEAR:
    gs_filetab,
    gv_rc.

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

  REFRESH:
    gt_records,
    gt_binary,
    gt_worksheets.

  CLEAR:
    gv_filelength,
    gv_xstring,
    gv_filename,
    gv_worksheet,
    gr_excel_data,
    gv_excel_offset.

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
    gv_text = 'Excel file upload failed. Check file path and SAP GUI authorization.'.
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
    gv_text = 'Excel binary conversion to XSTRING failed.'.
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
        gv_text = 'Unable to read first worksheet from Excel file.'.
        PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
        RETURN.
      ENDIF.

      gr_excel_data = go_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( gv_worksheet ).

    CATCH cx_fdt_excel_core INTO go_excel_error.

      CLEAR gv_text.
      gv_text = go_excel_error->get_text( ).

      IF gv_text IS INITIAL.
        gv_text = 'Excel parsing failed using CL_FDT_XL_SPREADSHEET.'.
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

    CLEAR:
      gs_raw,
      wa_input,
      gv_ok,
      gv_skip.

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

    IF gv_skip = abap_true.
      CONTINUE.
    ENDIF.

    PERFORM validate_required_fields USING gs_raw gv_rowno CHANGING gv_ok.

    IF gv_ok = abap_false.
      gv_failed = gv_failed + 1.
      CONTINUE.
    ENDIF.

    PERFORM convert_raw_to_input USING gs_raw
                                       gv_rowno
                              CHANGING wa_input
                                       gv_ok.

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
    gv_text = 'No valid Excel data rows found. Please check that actual data starts from row 4.'.
    PERFORM add_error USING 0 p_matnr 'E' 'LOCAL' '000' gv_text.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form DETECT_EXCEL_OFFSET
*&---------------------------------------------------------------------*
FORM detect_excel_offset.

  CLEAR:
    gv_excel_offset,
    gv_first_cell,
    gv_second_cell.

  READ TABLE <gt_excel_dyn> ASSIGNING <gs_excel_dyn> INDEX 1.
  IF sy-subrc <> 0.
    gv_excel_offset = 0.
    RETURN.
  ENDIF.

  ASSIGN COMPONENT 1 OF STRUCTURE <gs_excel_dyn> TO <fs_cell>.
  IF sy-subrc = 0.
    gv_first_cell = <fs_cell>.
    PERFORM clean_token CHANGING gv_first_cell.
  ENDIF.

  ASSIGN COMPONENT 2 OF STRUCTURE <gs_excel_dyn> TO <fs_cell>.
  IF sy-subrc = 0.
    gv_second_cell = <fs_cell>.
    PERFORM clean_token CHANGING gv_second_cell.
  ENDIF.

  TRANSLATE gv_first_cell TO UPPER CASE.
  TRANSLATE gv_second_cell TO UPPER CASE.

  IF gv_first_cell = 'MBRSH'.
    gv_excel_offset = 0.
  ELSEIF gv_first_cell = 'RMMG1-MBRSH'.
    gv_excel_offset = 1.
  ELSEIF gv_first_cell = 'INDUSTRY SECTOR'.
    gv_excel_offset = 2.
  ELSE.
    gv_excel_offset = 3.
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

  IF ps_raw-mbrsh IS INITIAL
     AND ps_raw-mtart IS INITIAL
     AND ps_raw-werks IS INITIAL
     AND ps_raw-lgort IS INITIAL
     AND ps_raw-maktx IS INITIAL.
    pv_skip = abap_true.
    RETURN.
  ENDIF.

  IF lv_mbrsh = 'MBRSH'
     OR lv_mtart = 'MTART'
     OR lv_maktx = 'MAKTX'.
    pv_skip = abap_true.
    RETURN.
  ENDIF.

  IF lv_mbrsh = 'RMMG1-MBRSH'
     OR lv_mtart = 'RMMG1-MTART'.
    pv_skip = abap_true.
    RETURN.
  ENDIF.

  IF lv_mbrsh = 'INDUSTRY SECTOR'
     OR lv_mtart = 'MATERIAL TYPE'
     OR lv_maktx = 'MATERIAL DESCRIPTION'.
    pv_skip = abap_true.
    RETURN.
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
    gv_text = 'Mandatory field MBRSH is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  IF ps_raw-mtart IS INITIAL.
    gv_text = 'Mandatory field MTART is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  IF ps_raw-werks IS INITIAL.
    gv_text = 'Mandatory field WERKS is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  IF ps_raw-maktx IS INITIAL.
    gv_text = 'Mandatory field MAKTX is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  IF ps_raw-meins IS INITIAL.
    gv_text = 'Mandatory field MEINS is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  "FIX: MTVFP (Availability Check) and PRCTR (Profit Center) are
  "     mandatory on the Sales: General/Plant screen.  Catching them
  "     here produces a clear error row instead of a BDC formatting
  "     error that is harder to trace.
  IF ps_raw-mtvfp IS INITIAL.
    gv_text = 'Mandatory field MTVFP (Availability Check) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

  IF ps_raw-prctr IS INITIAL.
    gv_text = 'Mandatory field PRCTR (Profit Center) is blank.'.
    PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.
    pv_ok = abap_false.
  ENDIF.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form CLEAN_TOKEN
*&---------------------------------------------------------------------*
FORM clean_token CHANGING cv_value TYPE string.

  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>cr_lf   IN cv_value WITH space.
  REPLACE ALL OCCURRENCES OF cl_abap_char_utilities=>newline IN cv_value WITH space.

  SHIFT cv_value LEFT  DELETING LEADING space.
  SHIFT cv_value RIGHT DELETING TRAILING space.

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
*&---------------------------------------------------------------------*
FORM convert_raw_to_input USING    ps_raw   TYPE ty_input_raw
                                   pv_rowno TYPE i
                          CHANGING ps_input TYPE ty_input
                                   pv_ok    TYPE abap_bool.

  "Local variables for field-specific formatting
  DATA: lv_mtvfp_n TYPE n LENGTH 2,
        lv_prctr   TYPE string,
        lv_steuc   TYPE string.

  CLEAR ps_input.
  pv_ok = abap_true.

  ps_input-mbrsh    = ps_raw-mbrsh.
  ps_input-mtart    = ps_raw-mtart.
  ps_input-werks    = ps_raw-werks.
  ps_input-lgort    = ps_raw-lgort.
  ps_input-vkorg    = ps_raw-vkorg.
  ps_input-vtweg    = ps_raw-vtweg.
  ps_input-maktx    = ps_raw-maktx.
  ps_input-meins    = ps_raw-meins.
  ps_input-matkl    = ps_raw-matkl.
  ps_input-spart    = ps_raw-spart.
  ps_input-gewei    = ps_raw-gewei.
  ps_input-wrkst    = ps_raw-wrkst.
  ps_input-sktof    = ps_raw-sktof.
  ps_input-taxkm1   = ps_raw-taxkm1.
  ps_input-taxkm2   = ps_raw-taxkm2.
  ps_input-versg    = ps_raw-versg.
  ps_input-kondm    = ps_raw-kondm.
  ps_input-tragr    = ps_raw-tragr.
  ps_input-ladgr    = ps_raw-ladgr.
  ps_input-xchpf    = ps_raw-xchpf.
  ps_input-taxim    = ps_raw-taxim.
  ps_input-dismm    = ps_raw-dismm.
  ps_input-beskz    = ps_raw-beskz.
  ps_input-perkz    = ps_raw-perkz.
  ps_input-iprkz    = ps_raw-iprkz.
  ps_input-sled_bbd = ps_raw-sled_bbd.
  ps_input-qmpur    = ps_raw-qmpur.
  ps_input-ssqss    = ps_raw-ssqss.
  ps_input-bklas    = ps_raw-bklas.
  ps_input-vprsv    = ps_raw-vprsv.
  ps_input-ekalr    = ps_raw-ekalr.
  ps_input-hkmat    = ps_raw-hkmat.
  ps_input-sobsk    = ps_raw-sobsk.
  ps_input-mlast    = ps_raw-mlast.
  ps_input-insmk    = ps_raw-insmk.
  ps_input-kzdkz    = ps_raw-kzdkz.
  ps_input-ncost    = ps_raw-ncost.
  ps_input-qpls_arg = ps_raw-qpls_arg.
  ps_input-qpls_art = ps_raw-qpls_art.

  "FIX 1: MARC-MTVFP (Availability Check)
  "  SAP requires a 2-digit zero-padded key, e.g. '02' not '2'.
  "  Assigning via TYPE n LENGTH 2 auto-pads on the left with zeros.
  IF ps_raw-mtvfp IS NOT INITIAL.
    TRY.
        lv_mtvfp_n     = ps_raw-mtvfp.
        ps_input-mtvfp = lv_mtvfp_n.
      CATCH cx_sy_conversion_no_number cx_sy_conversion_overflow.
        "Non-numeric value: pass as-is and let SAP validate on screen
        ps_input-mtvfp = ps_raw-mtvfp.
    ENDTRY.
  ENDIF.

  "FIX 2: MARC-PRCTR (Profit Center)
  "  Condense internal whitespace and convert to uppercase so the BDC
  "  value matches the profit-centre master exactly as stored in CEPC.
  lv_prctr = ps_raw-prctr.
  CONDENSE lv_prctr NO-GAPS.
  TRANSLATE lv_prctr TO UPPER CASE.
  ps_input-prctr = lv_prctr.

  "FIX 3: MARC-STEUC (HSN / Control Code)
  "  Excel often stores the value with a country prefix, e.g. 'IN 38089299'.
  "  T604F is keyed on LAND1 + STEUC; the BDC field expects only the
  "  commodity code part.  Strip the leading 'CC ' prefix when present
  "  (2-char country code followed by a single space).
  lv_steuc = ps_raw-steuc.
  SHIFT lv_steuc LEFT  DELETING LEADING  space.
  SHIFT lv_steuc RIGHT DELETING TRAILING space.
  IF strlen( lv_steuc ) > 3 AND lv_steuc+2(1) = ' '.
    lv_steuc = lv_steuc+3.
    SHIFT lv_steuc LEFT  DELETING LEADING  space.
    SHIFT lv_steuc RIGHT DELETING TRAILING space.
  ENDIF.
  ps_input-steuc = lv_steuc.

  PERFORM move_numeric USING ps_raw-brgew  'BRGEW'  pv_rowno CHANGING ps_input-brgew  pv_ok.
  PERFORM move_numeric USING ps_raw-ntgew  'NTGEW'  pv_rowno CHANGING ps_input-ntgew  pv_ok.
  PERFORM move_numeric USING ps_raw-stprs  'STPRS'  pv_rowno CHANGING ps_input-stprs  pv_ok.
  PERFORM move_numeric USING ps_raw-peinh  'PEINH'  pv_rowno CHANGING ps_input-peinh  pv_ok.
  PERFORM move_numeric USING ps_raw-losgr  'LOSGR'  pv_rowno CHANGING ps_input-losgr  pv_ok.
  PERFORM move_numeric USING ps_raw-stprs1 'STPRS1' pv_rowno CHANGING ps_input-stprs1 pv_ok.
  PERFORM move_numeric USING ps_raw-stprs2 'STPRS2' pv_rowno CHANGING ps_input-stprs2 pv_ok.
  PERFORM move_numeric USING ps_raw-stprs3 'STPRS3' pv_rowno CHANGING ps_input-stprs3 pv_ok.
  PERFORM move_numeric USING ps_raw-peinh1 'PEINH1' pv_rowno CHANGING ps_input-peinh1 pv_ok.
  PERFORM move_numeric USING ps_raw-peinh2 'PEINH2' pv_rowno CHANGING ps_input-peinh2 pv_ok.
  PERFORM move_numeric USING ps_raw-peinh3 'PEINH3' pv_rowno CHANGING ps_input-peinh3 pv_ok.
  PERFORM move_numeric USING ps_raw-verpr  'VERPR'  pv_rowno CHANGING ps_input-verpr  pv_ok.

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

  SHIFT gv_numtext LEFT  DELETING LEADING space.
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
      CONCATENATE 'Invalid numeric value in field'
                  pv_field
                  ':'
                  pv_value
             INTO gv_text SEPARATED BY space.

      PERFORM add_error USING pv_rowno p_matnr 'E' 'LOCAL' '000' gv_text.

      cv_ok = abap_false.

  ENDTRY.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form BUILD_BDC
*&---------------------------------------------------------------------*
*& Screen sequence derived from SHDB recording.
*&
*& Fixes applied in this version:
*&  1. Removed the duplicate hard-coded screen 0060 block that appeared
*&     after the data-driven one.  Having two dynpro entries for screen
*&     0060 caused the BDC engine to present 0060 data when SAP had
*&     already advanced to screen 0070 (view selection), leaving all
*&     checkboxes unset and triggering M3/196 "Select at least one view".
*&  2. Costing 1 price fields corrected to CKMAT_DISPLAY (single M) to
*&     match the actual screen field names from the SHDB recording.
*&  3. Purchasing tax indicator corrected to MG03STEUMM-TAXIM (no row
*&     index) to match the field as it appears on the Purchasing screen.
*&  4. Quality Management screen extended with MARC-INSMK, MARC-KZDKZ,
*&     MARC-NCOST, RMQAM-ARGUMENT and RMQAM-ART(01) so that all fields
*&     listed in the upload template are posted to the transaction.
*&---------------------------------------------------------------------*
FORM build_bdc USING ps_input TYPE ty_input.

  REFRESH gt_bdcdata.

  "===============================================================
  " 0060 - Initial Screen  (single block – duplicate removed)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0060.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'RMMG1-MTART'.
  PERFORM bdc_field  USING 'RMMG1-MBRSH' ps_input-mbrsh.
  PERFORM bdc_field  USING 'RMMG1-MTART' ps_input-mtart.

  IF p_matnr IS NOT INITIAL.
    PERFORM bdc_field USING 'RMMG1-MATNR' p_matnr.
  ENDIF.

  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_entr.

  "===============================================================
  " 0070 - Select View(s) - Page 1
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR' 'MSICHTAUSW-DYTXT(10)'.

  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.  "Basic Data 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' c_x.  "Basic Data 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.  "Sales: Sales Org. Data 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(05)' c_x.  "Sales: Sales Org. Data 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(06)' c_x.  "Sales: General/Plant Data
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.  "International Trade: Export
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(10)' c_x.  "Purchasing

  "Scroll down to next list page
  PERFORM bdc_field  USING 'BDC_OKCODE' '=P+'.

  "===============================================================
  " 0070 - Select View(s) - Page 2 (after scroll)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0070.
  PERFORM bdc_field  USING 'BDC_CURSOR' 'MSICHTAUSW-DYTXT(16)'.

  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(01)' c_x.  "MRP 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(02)' c_x.  "MRP 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(03)' c_x.  "MRP 3
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(04)' c_x.  "MRP 4

  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(08)' c_x.  "General Plant Data / Storage 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(09)' c_x.  "General Plant Data / Storage 2

  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(12)' c_x.  "Quality Management
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(13)' c_x.  "Accounting 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(14)' c_x.  "Accounting 2
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(15)' c_x.  "Costing 1
  PERFORM bdc_field  USING 'MSICHTAUSW-KZSEL(16)' c_x.  "Costing 2

  "Confirm selected views and proceed to organisational levels
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_schl.

  "===============================================================
  " 0080 - Organizational Levels
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_0080.
  PERFORM bdc_field  USING 'BDC_CURSOR'  'RMMG1-VTWEG'.
  PERFORM bdc_field  USING 'RMMG1-WERKS' ps_input-werks.
  PERFORM bdc_field  USING 'RMMG1-LGORT' ps_input-lgort.
  PERFORM bdc_field  USING 'RMMG1-VKORG' ps_input-vkorg.
  PERFORM bdc_field  USING 'RMMG1-VTWEG' ps_input-vtweg.
  PERFORM bdc_field  USING 'BDC_OKCODE'  c_ok_entr.

  "===============================================================
  " Basic Data 1  (4004)
  " Fields: MAKT-MAKTX, MARA-MEINS, MARA-MATKL, MARA-SPART,
  "         MARA-BRGEW, MARA-GEWEI, MARA-NTGEW
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'MAKT-MAKTX' ps_input-maktx.
  PERFORM bdc_field  USING 'MARA-MEINS' ps_input-meins.
  PERFORM bdc_field  USING 'MARA-MATKL' ps_input-matkl.
  PERFORM bdc_field  USING 'MARA-SPART' ps_input-spart.
  PERFORM bdc_num    USING 'MARA-BRGEW' ps_input-brgew.
  PERFORM bdc_field  USING 'MARA-GEWEI' ps_input-gewei.
  PERFORM bdc_num    USING 'MARA-NTGEW' ps_input-ntgew.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Basic Data 2  (4004)
  " Fields: MARA-WRKST
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'MARA-WRKST' ps_input-wrkst.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Sales: Sales Org 1  (4000)
  " Fields: MVKE-SKTOF, MG03STEUER-TAXKM(01), MG03STEUER-TAXKM(02)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MVKE-SKTOF'           ps_input-sktof.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(01)' ps_input-taxkm1.
  PERFORM bdc_field  USING 'MG03STEUER-TAXKM(02)' ps_input-taxkm2.
  PERFORM bdc_field  USING 'BDC_OKCODE'           c_ok_enter.

  "===============================================================
  " Sales: Sales Org 2  (4000)
  " Fields: MVKE-VERSG, MVKE-KONDM
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MVKE-VERSG' ps_input-versg.
  PERFORM bdc_field  USING 'MVKE-KONDM' ps_input-kondm.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Sales: General / Plant Data  (4004)
  " Fields: MARC-MTVFP, MARC-XCHPF, MARA-TRAGR, MARC-LADGR,
  "         MARC-PRCTR
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'MARC-MTVFP' ps_input-mtvfp.
  PERFORM bdc_field  USING 'MARC-XCHPF' ps_input-xchpf.
  PERFORM bdc_field  USING 'MARA-TRAGR' ps_input-tragr.
  PERFORM bdc_field  USING 'MARC-LADGR' ps_input-ladgr.
  PERFORM bdc_field  USING 'MARC-PRCTR' ps_input-prctr.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " International Trade: Export  (4004)
  " Fields: MARC-STEUC
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4004.
  PERFORM bdc_field  USING 'MARC-STEUC' ps_input-steuc.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Purchasing  (4000)
  " Fields: MG03STEUMM-TAXIM  (no row index – single field on screen)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MG03STEUMM-TAXIM' ps_input-taxim.
  PERFORM bdc_field  USING 'BDC_OKCODE'        c_ok_enter.

  "===============================================================
  " Costing 1  (4000)
  " Fields: MBEW-BKLAS, CKMLHD-MLAST,
  "         CKMAT_DISPLAY-STPRS_1/2/3, CKMAT_DISPLAY-PEINH_1/2/3
  " NOTE: field structure is CKMAT_DISPLAY (single M) not CKMMAT_DISPLAY
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MBEW-BKLAS'            ps_input-bklas.
  PERFORM bdc_field  USING 'CKMLHD-MLAST'          ps_input-mlast.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-STPRS_1'  ps_input-stprs1.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-STPRS_2'  ps_input-stprs2.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-STPRS_3'  ps_input-stprs3.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-PEINH_1'  ps_input-peinh1.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-PEINH_2'  ps_input-peinh2.
  PERFORM bdc_num    USING 'CKMAT_DISPLAY-PEINH_3'  ps_input-peinh3.
  PERFORM bdc_field  USING 'BDC_OKCODE'             c_ok_enter.

  "===============================================================
  " Costing navigation  (4000) – no user-data fields
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Costing 2  (4000)
  " Fields: MBEW-EKALR, MBEW-HKMAT, MARC-SOBSK, MARC-LOSGR
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MBEW-EKALR' ps_input-ekalr.
  PERFORM bdc_field  USING 'MBEW-HKMAT' ps_input-hkmat.
  IF ps_input-sobsk IS NOT INITIAL.
    PERFORM bdc_field USING 'MARC-SOBSK' ps_input-sobsk.
  ENDIF.
  PERFORM bdc_num    USING 'MARC-LOSGR' ps_input-losgr.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Accounting 1 / Valuation  (4000)
  " Fields: MBEW-VPRSV, MBEW-STPRS, MBEW-PEINH
  "         MBEW-VERPR only when price control = 'V' (moving average)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MBEW-VPRSV' ps_input-vprsv.
  PERFORM bdc_num    USING 'MBEW-STPRS' ps_input-stprs.
  PERFORM bdc_num    USING 'MBEW-PEINH' ps_input-peinh.
  IF ps_input-vprsv = 'V'.
    PERFORM bdc_num USING 'MBEW-VERPR' ps_input-verpr.
  ENDIF.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " MRP 1  (4000)
  " Fields: MARC-DISMM
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MARC-DISMM' ps_input-dismm.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " MRP 2  (4000)
  " Fields: MARC-BESKZ
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MARC-BESKZ' ps_input-beskz.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " MRP 3  (4000)
  " Fields: MARC-PERKZ
  " Note: MARC-MTVFP is set on the Sales: General/Plant Data screen
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MARC-PERKZ' ps_input-perkz.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " MRP 4  (4000) – no user-data fields
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " General Plant Data / Storage 1  (4000)
  " Fields: MARA-IPRKZ, MARA-SLED_BBD
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MARA-IPRKZ'    ps_input-iprkz.
  PERFORM bdc_field  USING 'MARA-SLED_BBD' ps_input-sled_bbd.
  PERFORM bdc_field  USING 'BDC_OKCODE'    c_ok_enter.

  "===============================================================
  " General Plant Data / Storage 2  (4000) – no user-data fields
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'BDC_OKCODE' c_ok_enter.

  "===============================================================
  " Quality Management  (4000)
  " Fields: MARA-QMPUR, MARC-INSMK, MARC-SSQSS, MARC-KZDKZ,
  "         MARC-NCOST, RMQAM-ARGUMENT, RMQAM-ART(01)
  "===============================================================
  PERFORM bdc_dynpro USING c_prog_mm c_scr_4000.
  PERFORM bdc_field  USING 'MARA-QMPUR'     ps_input-qmpur.
  PERFORM bdc_field  USING 'MARC-INSMK'     ps_input-insmk.
  PERFORM bdc_field  USING 'MARC-SSQSS'     ps_input-ssqss.
  PERFORM bdc_field  USING 'MARC-KZDKZ'     ps_input-kzdkz.
  PERFORM bdc_field  USING 'MARC-NCOST'     ps_input-ncost.
  PERFORM bdc_field  USING 'RMQAM-ARGUMENT' ps_input-qpls_arg.
  PERFORM bdc_field  USING 'RMQAM-ART(01)'  ps_input-qpls_art.
  PERFORM bdc_field  USING 'BDC_OKCODE'     c_ok_enter.

  "===============================================================
  " SAPLSPO1 0300 – Save confirmation popup
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
*&---------------------------------------------------------------------*
FORM bdc_num USING pv_fnam LIKE bdcdata-fnam
                   pv_num  TYPE any.

  CLEAR gv_numtext.

  WRITE pv_num TO gv_numtext NO-GROUPING LEFT-JUSTIFIED.
  CONDENSE gv_numtext NO-GAPS.

  CLEAR gs_bdcdata.
  gs_bdcdata-fnam = pv_fnam.
  gs_bdcdata-fval = gv_numtext.
  APPEND gs_bdcdata TO gt_bdcdata.

ENDFORM.

*&---------------------------------------------------------------------*
*& Form ANALYZE_BDC_MESSAGES
*&---------------------------------------------------------------------*
FORM analyze_bdc_messages USING pv_rowno TYPE i.

  CLEAR:
    gv_row_failed,
    gv_material.

  gv_row_failed = abap_false.
  gv_material   = p_matnr.

  LOOP AT gt_msgcoll INTO gs_msgcoll.

    IF gs_msgcoll-msgtyp = 'S' AND gv_material IS INITIAL.
      IF gs_msgcoll-msgv1 IS NOT INITIAL.
        gv_material = gs_msgcoll-msgv1.
      ENDIF.
    ENDIF.

    IF gs_msgcoll-msgtyp = 'E'
       OR gs_msgcoll-msgtyp = 'A'.

      gv_row_failed = abap_true.

      PERFORM get_message_text USING gs_msgcoll CHANGING gv_msgtext.

      PERFORM add_error USING pv_rowno
                              gv_material
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
    CONCATENATE 'CALL TRANSACTION returned SY-SUBRC ='
                gv_subrc_c
           INTO gv_msgtext SEPARATED BY space.

    PERFORM add_error USING pv_rowno
                            gv_material
                            'E'
                            'LOCAL'
                            '000'
                            gv_msgtext.

  ENDIF.

  IF gv_row_failed = abap_true.
    gv_failed = gv_failed + 1.
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

  MESSAGE ID ps_msg-msgid
          TYPE ps_msg-msgtyp
          NUMBER ps_msg-msgnr
          WITH ps_msg-msgv1
               ps_msg-msgv2
               ps_msg-msgv3
               ps_msg-msgv4
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
    WRITE: / 'ALV display failed. Please debug GT_ERRORS internal table.'.
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
